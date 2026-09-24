"""Editable DS-021-0046 Excel layer.

Catalogue selection is unchanged. This module copies the already-fetched
DS-021-0046 report into an editable workbook, re-imports confirmed values,
and never writes catalogue data back over user edits.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from .catalogue import MISSING, clone_report, display_value

SHEET_DS = "DS-021-0046"
SHEET_EDITABLE = "Editable_Report"
SHEET_ORIGINAL = "Original_Data"
SHEET_CHANGES = "Change_Log"
SHEET_SOURCE = "Source_Info"

HEADER_ROW = 4
DATA_START_ROW = 5

# DS-021-0046 page layout (matches the PDF, not a data table).
PDF_FONT = "Arial"
DS_TITLE_ROW = 1
DS_BAR_ROW = 2
DS_CATALOG_LABEL_ROW = 4
DS_CATALOG_VALUE_ROW = 5
DS_FAMILY_ROW = 7
DS_PRODUCT_DATA_ROW = 9
DS_FIELD_START_ROW = 11

# Identity rows required by DS-021-0046 (above the numbered Product Data list).
IDENTITY_SPECS = [
    {"key": "catalog_designation", "label": "Catalog Designation", "unit": ""},
    {"key": "product_family", "label": "Product family line", "unit": ""},
    {"key": "document_id", "label": "Document number", "unit": ""},
    {"key": "disclaimer", "label": "Disclaimer", "unit": ""},
]

# Numeric Product Data fields (unit shown on DS-021-0046). Extra wording is
# allowed for fields that the sample sheet writes as mixed text.
STRICT_NUMERIC_KEYS = {
    "rated_motor_power",
    "rated_motor_speed",
    "overall_gear_ratio",
    "output_speed",
    "output_torque",
    "service_factor",
    "permitted_output_overhung_load",
    "number_of_poles",
    "frequency",
}

OPTIONAL_BLANK_KEYS = {"additional_features"}

DROPDOWNS = {
    "mounting_position": ["M1", "M2", "M3", "M4", "M5", "M6"],
    "position_of_terminal_box": ["KP1", "KP2", "KP3", "KP4"],
    "cable_entry_position": ["A", "B", "C", "D"],
    "output_shaft_type": [
        "Hollow with keyway",
        "Solid with key",
        "Solid with key (double extended)",
        "Hollow with shrink disc",
    ],
    "motor_type": ["Standard", "Brake"],
    "ce_marking": ["Yes"],
    "efficiency_class": ["IE2", "IE3"],
    "insulation_class": ["F"],
    "protection_class": ["IP 55"],
    "design_requirement": ["Europe (CE)"],
    "direction_of_rotation": ["Bi-directional"],
    "lubricant_type": ["ISO VG 320"],
}

NAVY = "003366"
BLUE = "0055A5"
LIGHT = "E8EEF5"
AMBER = "FFF3CD"
GREEN = "E8F5E9"
THIN = Border(
    left=Side(style="thin", color="8AA0B8"),
    right=Side(style="thin", color="8AA0B8"),
    top=Side(style="thin", color="8AA0B8"),
    bottom=Side(style="thin", color="8AA0B8"),
)


def identity_rows(report: dict[str, Any]) -> list[dict[str, Any]]:
    values = {
        "catalog_designation": report.get("catalog_designation", ""),
        "product_family": report.get("product_family", ""),
        "document_id": report.get("document_id", ""),
        "disclaimer": report.get("disclaimer", ""),
    }
    rows = []
    for spec in IDENTITY_SPECS:
        rows.append(
            {
                "key": spec["key"],
                "label": spec["label"],
                "unit": spec["unit"],
                "value": values[spec["key"]],
                "source": "DS-021-0046",
            }
        )
    return rows


def all_report_rows(report: dict[str, Any]) -> list[dict[str, Any]]:
    return identity_rows(report) + list(report["fields"])


def safe_filename_part(text: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", str(text))
    cleaned = re.sub(r"\s+", "_", cleaned).strip("._")
    return cleaned or "Product"


def default_excel_name(report: dict[str, Any]) -> str:
    internal = report.get("internal") or {}
    model = safe_filename_part(internal.get("model") or report.get("catalog_designation") or "Product")
    power = ""
    ratio = ""
    for field in report.get("fields", []):
        if field["key"] == "rated_motor_power":
            power = safe_filename_part(str(field["value"]))
        if field["key"] == "overall_gear_ratio":
            ratio = safe_filename_part(str(field["value"]))
    return f"MGM_Varvel_{model}_{power}kW_i{ratio}_DS-021-0046.xlsx"


def default_pdf_name(report: dict[str, Any]) -> str:
    return default_excel_name(report).replace("_DS-021-0046.xlsx", "_Final.pdf")


def selection_fingerprint(report: dict[str, Any]) -> str:
    internal = report.get("internal") or {}
    parts = [
        str(internal.get("series", "")),
        str(internal.get("model", "")),
        str(internal.get("product_id", "")),
    ]
    for key in (
        "rated_motor_power",
        "overall_gear_ratio",
        "mounting_position",
        "output_shaft_type",
        "position_of_terminal_box",
        "cable_entry_position",
        "motor_type",
    ):
        parts.append(_field_value(report, key))
    return "|".join(parts)


def _field_value(report: dict[str, Any], key: str) -> str:
    for field in report.get("fields", []):
        if field["key"] == key:
            return str(field.get("value", ""))
    return ""


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-12:
            return str(int(round(value)))
        text = f"{value:.12g}"
        return text
    return str(value).strip()


def _style_header_cell(cell, fill_hex: str = NAVY) -> None:
    cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=fill_hex)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = THIN


def _style_body(cell, *, wrap: bool = True) -> None:
    cell.font = Font(name="Calibri", size=11)
    cell.alignment = Alignment(vertical="center", wrap_text=wrap)
    cell.border = THIN


def _status_formula(row: int) -> str:
    return (
        f'=IF(AND(C{row}&""={SHEET_ORIGINAL}!C{row}&"",'
        f'D{row}&""={SHEET_ORIGINAL}!D{row}&""),"Auto","Edited")'
    )


def _pdf_font(size: float, *, bold: bool = False, italic: bool = False, color: str = "333333") -> Font:
    return Font(name=PDF_FONT, size=size, bold=bold, italic=italic, color=color)


def _write_ds_layout(ws: Worksheet, report: dict[str, Any]) -> None:
    """Render the DS-021-0046 page: same type, colours, and stacking as the PDF."""
    hairline = Border(bottom=Side(style="thin", color="8AA0B8"))
    none_border = Border()

    ws.sheet_view.showGridLines = False
    ws.sheet_view.view = "pageLayout"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.55
    ws.page_margins.right = 0.55
    ws.page_margins.top = 0.47
    ws.page_margins.bottom = 0.55
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    # Match the source page: labels and values occupy the left portion, while
    # the clear area on the right is reserved for the source document's diagrams.
    ws.column_dimensions["A"].width = 31
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 31
    ws.print_options.horizontalCentered = True

    ws.merge_cells("A1:C1")
    ws["A1"] = "Product Information"
    ws["A1"].font = _pdf_font(16, bold=True, color=NAVY)
    ws["A1"].alignment = Alignment(vertical="bottom")

    logo_path = Path(__file__).resolve().parent.parent / "static" / "MGM Varvel Logo.jpg"
    if logo_path.is_file():
        logo = ExcelImage(logo_path)
        scale = 130 / logo.width
        logo.width = 130
        logo.height *= scale
        # Anchor within the merged title cell, aligned to the page's right edge.
        logo.anchor = OneCellAnchor(
            _from=AnchorMarker(col=1, colOff=210 * 9525, row=0, rowOff=0),
            ext=XDRPositiveSize2D(cx=logo.width * 9525, cy=logo.height * 9525),
        )
        ws.add_image(logo)
    ws.row_dimensions[DS_TITLE_ROW].height = 22

    ws.merge_cells("A2:C2")
    ws["A2"].fill = PatternFill("solid", fgColor=BLUE)
    ws.row_dimensions[DS_BAR_ROW].height = 4

    ws.row_dimensions[3].height = 10

    ws.merge_cells("A4:C4")
    ws["A4"] = "Catalog Designation"
    ws["A4"].font = _pdf_font(8.5, bold=True, color="333333")
    ws.row_dimensions[DS_CATALOG_LABEL_ROW].height = 14

    ws.merge_cells("A5:C5")
    ws["A5"] = _as_text(report.get("catalog_designation", ""))
    ws["A5"].font = _pdf_font(8, color="222222")
    ws["A5"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[DS_CATALOG_VALUE_ROW].height = 32

    ws.row_dimensions[6].height = 6

    ws.merge_cells("A7:C7")
    ws["A7"] = _as_text(report.get("product_family", ""))
    ws["A7"].font = _pdf_font(8, color="222222")
    ws["A7"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[DS_FAMILY_ROW].height = 16

    ws.row_dimensions[8].height = 10

    ws.merge_cells("A9:B9")
    ws["A9"] = "Product Data"
    ws["A9"].font = _pdf_font(8.5, bold=True, color="333333")
    ws["A9"].fill = PatternFill("solid", fgColor=LIGHT)
    ws["A9"].alignment = Alignment(vertical="center", indent=1)
    ws["C9"] = _as_text(report.get("document_id", "DS-021-0046"))
    ws["C9"].font = _pdf_font(8.5, bold=True, color=NAVY)
    ws["C9"].fill = PatternFill("solid", fgColor=LIGHT)
    ws["C9"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[DS_PRODUCT_DATA_ROW].height = 18

    ws.row_dimensions[10].height = 6

    fields = list(report.get("fields") or [])
    for offset, field in enumerate(fields):
        row = DS_FIELD_START_ROW + offset
        shown = display_value(field)
        missing = field.get("value") == MISSING or str(field.get("value", "")).startswith("Field required")
        a = ws.cell(row, 1, field.get("label", ""))
        b = ws.cell(row, 2, shown)
        a.font = _pdf_font(8.5, color="333333")
        b.font = _pdf_font(8.5, bold=True, color="AA3333" if missing else NAVY)
        a.alignment = Alignment(vertical="center")
        b.alignment = Alignment(vertical="center", wrap_text=True)
        a.border = hairline
        b.border = hairline
        # Column C stays open beside the fields, preserving the diagram area
        # visible on DS-021-0046 and keeping text from running into it.
        ws.row_dimensions[row].height = 16.5

    last = DS_FIELD_START_ROW + len(fields) - 1
    disc_row = last + 2
    ws.merge_cells(start_row=disc_row, start_column=1, end_row=disc_row, end_column=3)
    disc = ws.cell(disc_row, 1, _as_text(report.get("disclaimer", "")))
    disc.font = _pdf_font(7, italic=True, color="444444")
    disc.alignment = Alignment(wrap_text=True, vertical="top")
    disc.border = none_border
    ws.row_dimensions[disc_row].height = 42
    ws.print_area = f"A1:C{disc_row}"


def _write_report_sheet(ws: Worksheet, report: dict[str, Any], *, formulas: bool) -> None:
    ws["A1"] = "MGM-VARVEL  |  DS-021-0046 editable product data"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=NAVY)
    ws.merge_cells("A1:F1")
    ws["A2"] = (
        "Edit Value and Unit only. Status becomes Edited when a cell differs from Original_Data. "
        "Do not rename Field labels or delete rows. Original catalogue values are kept on Original_Data."
    )
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="444444")
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:F2")
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 32

    headers = ["Key", "Field", "Value", "Unit", "Source", "Status"]
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(HEADER_ROW, col, title)
        _style_header_cell(cell)

    rows = all_report_rows(report)
    for offset, item in enumerate(rows):
        row = DATA_START_ROW + offset
        values = [
            item["key"],
            item["label"],
            _as_text(item.get("value", "")),
            _as_text(item.get("unit", "")),
            _as_text(item.get("source", "")),
            "Auto" if not formulas else _status_formula(row),
        ]
        for col, val in enumerate(values, start=1):
            cell = ws.cell(row, col, val)
            _style_body(cell, wrap=col in (2, 3, 5))
            if col in (1, 3, 4):
                cell.number_format = "@"
            if col == 2:
                cell.font = Font(name="Calibri", size=11, bold=True)
                cell.fill = PatternFill("solid", fgColor=LIGHT)
            if col == 3:
                cell.fill = PatternFill("solid", fgColor="FFFFFF")
            if col == 6 and not formulas:
                cell.fill = PatternFill("solid", fgColor=GREEN)
        if formulas:
            ws.cell(row, 6).fill = PatternFill("solid", fgColor=GREEN)

        key = item["key"]
        if key in DROPDOWNS:
            options = ",".join(DROPDOWNS[key])
            dv = DataValidation(
                type="list",
                formula1=f'"{options}"',
                allow_blank=True,
                showDropDown=False,
                showErrorMessage=False,
                showInputMessage=True,
                promptTitle="Catalogue / DS values",
                prompt="Select a known value, or type a correction.",
            )
            dv.add(f"C{row}")
            ws.add_data_validation(dv)

    last = DATA_START_ROW + len(rows) - 1
    ws.auto_filter.ref = f"A{HEADER_ROW}:F{last}"
    ws.freeze_panes = "A5"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 62
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["A"].hidden = True
    ws.row_dimensions[HEADER_ROW].height = 20
    ws.sheet_view.showGridLines = False

    if formulas:
        for row in range(DATA_START_ROW, last + 1):
            for col in (1, 2, 5, 6):
                ws.cell(row, col).protection = Protection(locked=True)
            for col in (3, 4):
                ws.cell(row, col).protection = Protection(locked=False)
        ws.protection.sheet = True
        ws.protection.enable()
        ws.protection.autoFilter = True
        ws.protection.sort = True
        ws.protection.selectUnlockedCells = True


def _write_change_log(ws: Worksheet, n_rows: int) -> None:
    ws["A1"] = "Change log (Auto vs Edited)"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=NAVY)
    ws.merge_cells("A1:F1")
    ws["A2"] = "This sheet does not appear on the final DS-021-0046 PDF. It is for review only."
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="444444")
    headers = ["Field", "Original Value", "Original Unit", "Edited Value", "Edited Unit", "Status"]
    for col, title in enumerate(headers, start=1):
        _style_header_cell(ws.cell(HEADER_ROW, col, title), BLUE)
    last = DATA_START_ROW + n_rows - 1
    for i in range(n_rows):
        row = DATA_START_ROW + i
        mapping = [
            f"={SHEET_EDITABLE}!B{row}",
            f"={SHEET_ORIGINAL}!C{row}",
            f"={SHEET_ORIGINAL}!D{row}",
            f"={SHEET_EDITABLE}!C{row}",
            f"={SHEET_EDITABLE}!D{row}",
            f"={SHEET_EDITABLE}!F{row}",
        ]
        for col, formula in enumerate(mapping, start=1):
            cell = ws.cell(row, col, formula)
            _style_body(cell)
            if col == 6:
                cell.fill = PatternFill("solid", fgColor=AMBER)
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A{HEADER_ROW}:F{last}"
    for letter, width in zip("ABCDEF", (38, 40, 14, 40, 14, 12)):
        ws.column_dimensions[letter].width = width


def _write_source_info(ws: Worksheet, report: dict[str, Any]) -> None:
    internal = report.get("internal") or {}
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [
        ("Product series", internal.get("series", "")),
        ("Catalogue model", internal.get("model", "")),
        ("Catalogue record id", internal.get("product_id", "")),
        ("Selected power", _field_value(report, "rated_motor_power")),
        ("Selected ratio", _field_value(report, "overall_gear_ratio")),
        ("Catalogue page", internal.get("catalogue_page", "")),
        ("Catalogue table", internal.get("catalogue_table", "")),
        ("Catalogue reference", "RF-RK Catalogue"),
        ("Output document", "DS-021-0046"),
        ("Generation date/time", generated),
        ("Weight note", internal.get("weight_note", "")),
    ]
    ws["A1"] = "Source / selection traceability"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=NAVY)
    ws.merge_cells("A1:B1")
    ws["A2"] = "Not included in the final PDF unless DS-021-0046 requires it."
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="444444")
    for col, title in enumerate(("Item", "Value"), start=1):
        _style_header_cell(ws.cell(HEADER_ROW, col, title))
    for i, (label, value) in enumerate(rows):
        row = DATA_START_ROW + i
        a = ws.cell(row, 1, label)
        b = ws.cell(row, 2, _as_text(value))
        _style_body(a)
        _style_body(b)
        a.font = Font(name="Calibri", size=11, bold=True)
        a.fill = PatternFill("solid", fgColor=LIGHT)
        b.number_format = "@"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 80
    ws.freeze_panes = "A5"
    ws.protection.sheet = True
    ws.protection.enable()


def generate_excel(
    report: dict[str, Any],
    output_path: str | Path,
    *,
    include_edit_sheets: bool = False,
) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws_ds = wb.active
    ws_ds.title = SHEET_DS
    _write_ds_layout(ws_ds, report)

    if include_edit_sheets:
        ws_edit = wb.create_sheet(SHEET_EDITABLE)
        ws_orig = wb.create_sheet(SHEET_ORIGINAL)
        ws_log = wb.create_sheet(SHEET_CHANGES)
        ws_src = wb.create_sheet(SHEET_SOURCE)
        _write_report_sheet(ws_edit, report, formulas=True)
        _write_report_sheet(ws_orig, report, formulas=False)
        n_rows = len(all_report_rows(report))
        _write_change_log(ws_log, n_rows)
        _write_source_info(ws_src, report)
        for col in range(1, 7):
            for row in range(HEADER_ROW, HEADER_ROW + n_rows + 1):
                ws_orig.cell(row, col).protection = Protection(locked=True)
        ws_orig.protection.sheet = True
        ws_orig.protection.enable()

    wb.properties.title = f"Product Data {report.get('document_id', 'DS-021-0046')}"
    wb.properties.creator = "MGM-Varvel Product Selector"
    wb.save(output_path)
    return output_path


def _read_sheet_rows(ws: Worksheet) -> list[dict[str, str]]:
    headers = [ _as_text(ws.cell(HEADER_ROW, c).value).lower() for c in range(1, 7) ]
    expected = ["key", "field", "value", "unit", "source", "status"]
    if headers != expected:
        raise ValueError(
            f"Sheet '{ws.title}' is missing the required header row "
            "(Key, Field, Value, Unit, Source, Status)."
        )
    rows = []
    row = DATA_START_ROW
    while True:
        key = _as_text(ws.cell(row, 1).value)
        field = _as_text(ws.cell(row, 2).value)
        if not key and not field:
            break
        rows.append(
            {
                "key": key,
                "label": field,
                "value": _as_text(ws.cell(row, 3).value),
                "unit": _as_text(ws.cell(row, 4).value),
                "source": _as_text(ws.cell(row, 5).value),
                "status": _as_text(ws.cell(row, 6).value) or "Auto",
            }
        )
        row += 1
        if row > DATA_START_ROW + 80:
            break
    return rows


def _compute_status(current: str, original: str) -> str:
    return "Auto" if _as_text(current) == _as_text(original) else "Edited"


def _contains_number(text: str) -> bool:
    return re.search(r"[-+]?\d+(?:[.,]\d+)?", text.replace(" ", "")) is not None


def validate_rows(
    editable_rows: list[dict[str, str]],
    original_rows: list[dict[str, str]],
    required_keys: list[str],
) -> dict[str, Any]:
    errors: list[str] = []
    warnings: list[str] = []
    by_key = {r["key"]: r for r in editable_rows if r["key"]}
    orig_by_key = {r["key"]: r for r in original_rows if r["key"]}

    missing_keys = [k for k in required_keys if k not in by_key]
    if missing_keys:
        labels = ", ".join(missing_keys)
        errors.append(f"Required fields are missing from the Excel: {labels}.")

    identity_keys = [s["key"] for s in IDENTITY_SPECS]
    for spec_key in identity_keys + required_keys:
        row = by_key.get(spec_key)
        if not row:
            continue
        label = row["label"] or spec_key
        if spec_key not in OPTIONAL_BLANK_KEYS and not row["value"]:
            errors.append(f"{label} is missing. Please review the Excel before generating the PDF.")
        orig = orig_by_key.get(spec_key, {})
        expected_unit = orig.get("unit", "")
        if expected_unit and not row["unit"]:
            errors.append(f"{label} has no unit. Please restore the unit before generating the PDF.")
        if spec_key in STRICT_NUMERIC_KEYS and row["value"] and not _contains_number(row["value"]):
            errors.append(
                f"{label} must contain a numeric value (found “{row['value']}”)."
            )
        if row["value"] == MISSING:
            warnings.append(f"{label} still has no catalogue value.")

    edited = []
    for key, row in by_key.items():
        orig = orig_by_key.get(key, {})
        orig_val = orig.get("value", "")
        orig_unit = orig.get("unit", "")
        status_val = _compute_status(row["value"], orig_val)
        status_unit = _compute_status(row["unit"], orig_unit)
        status = "Edited" if (status_val == "Edited" or status_unit == "Edited") else "Auto"
        row["status"] = status
        row["original_value"] = orig_val
        row["original_unit"] = orig_unit
        if status == "Edited":
            edited.append(
                {
                    "key": key,
                    "label": row["label"],
                    "original": f"{orig_val} {orig_unit}".strip(),
                    "edited": f"{row['value']} {row['unit']}".strip(),
                }
            )

    product_ok = bool(by_key.get("catalog_designation", {}).get("value"))
    if not product_ok:
        errors.append("Product identity (Catalog Designation) is missing.")

    return {
        "ok": not errors,
        "errors": errors,
        "warnings": warnings,
        "edited": edited,
        "rows": list(by_key.values()),
    }


def load_excel(path: str | Path) -> dict[str, Any]:
    path = Path(path)
    wb = load_workbook(path, data_only=False)
    if SHEET_EDITABLE not in wb.sheetnames or SHEET_ORIGINAL not in wb.sheetnames:
        raise ValueError(
            f"Workbook must contain sheets {SHEET_EDITABLE} and {SHEET_ORIGINAL}."
        )
    editable = _read_sheet_rows(wb[SHEET_EDITABLE])
    original = _read_sheet_rows(wb[SHEET_ORIGINAL])
    source_info: dict[str, str] = {}
    if SHEET_SOURCE in wb.sheetnames:
        ws = wb[SHEET_SOURCE]
        row = DATA_START_ROW
        while True:
            label = _as_text(ws.cell(row, 1).value)
            if not label:
                break
            source_info[label] = _as_text(ws.cell(row, 2).value)
            row += 1
            if row > 40:
                break
    return {
        "path": str(path),
        "editable": editable,
        "original": original,
        "source_info": source_info,
    }


def apply_excel_to_report(base_report: dict[str, Any], loaded: dict[str, Any]) -> dict[str, Any]:
    """Overlay Editable_Report values onto a clone of the catalogue report.

    Catalogue values are kept as original_value. Current value/unit come only
    from the workbook. Nothing is written back from the catalogue.
    """
    report = clone_report(base_report)
    required = [f["key"] for f in report["fields"]]
    identity_keys = [s["key"] for s in IDENTITY_SPECS]
    check = validate_rows(loaded["editable"], loaded["original"], required)
    if not check["ok"]:
        raise ValueError("\n".join(check["errors"]))

    by_key = {r["key"]: r for r in check["rows"]}

    def overlay_identity(attr: str, key: str) -> None:
        row = by_key.get(key)
        if row:
            report[attr] = row["value"]

    overlay_identity("catalog_designation", "catalog_designation")
    overlay_identity("product_family", "product_family")
    overlay_identity("document_id", "document_id")
    overlay_identity("disclaimer", "disclaimer")

    for field in report["fields"]:
        row = by_key.get(field["key"])
        if not row:
            continue
        field["original_value"] = row.get("original_value", field.get("value"))
        field["original_unit"] = row.get("original_unit", field.get("unit"))
        field["value"] = row["value"]
        field["unit"] = row["unit"]
        field["status"] = row["status"]
        if row.get("source"):
            field["source"] = row["source"]

    report["excel_meta"] = {
        "path": loaded.get("path"),
        "edited_count": len(check["edited"]),
        "edits": check["edited"],
        "warnings": check["warnings"],
        "source_info": loaded.get("source_info") or {},
        "identity_keys": identity_keys,
    }
    return report


def workbook_has_edits(path: str | Path) -> bool:
    loaded = load_excel(path)
    required = [r["key"] for r in loaded["editable"] if r["key"] not in {s["key"] for s in IDENTITY_SPECS}]
    check = validate_rows(loaded["editable"], loaded["original"], required)
    return bool(check["edited"])


def report_from_excel_only(loaded: dict[str, Any], fallback_report: dict[str, Any] | None = None) -> dict[str, Any]:
    """Build the PDF report dict from confirmed Excel rows.

    Uses the Excel as the data source. Structure (labels/order) follows DS fields
    from the fallback catalogue report when provided, otherwise from the sheet.
    """
    if fallback_report is not None:
        return apply_excel_to_report(fallback_report, loaded)
    raise ValueError("A catalogue report is required to keep DS-021-0046 field order.")


def stamp_originals(report: dict[str, Any]) -> dict[str, Any]:
    """Mark every field as Auto and store original_* from current catalogue values."""
    out = clone_report(report)
    for field in out["fields"]:
        field["original_value"] = field.get("value")
        field["original_unit"] = field.get("unit")
        field["status"] = "Auto"
    out["excel_meta"] = {
        "path": None,
        "edited_count": 0,
        "edits": [],
        "warnings": [],
        "source_info": {},
    }
    return out
