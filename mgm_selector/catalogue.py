"""RF-RK catalogue product database and DS-021-0046 field mapping."""

from __future__ import annotations

import json
import re
from copy import deepcopy
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
SCRATCH = ROOT / "scratch"

MISSING = "Required value not found in catalogue."
GEAR_MASS_NOTE = (
    "Catalogue mass is gear unit only (without motor). "
    "Complete unit weight is not tabulated."
)

NOT_FOUND_MSG = "No exact catalogue match found for the selected combination."


def _load_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as f:
        return json.load(f)


class Catalogue:
    def __init__(self) -> None:
        self.lookups = _load_json(DATA / "lookups.json")
        self.ds_fields = _load_json(DATA / "ds_fields.json")
        self.field_map = _load_json(DATA / "field_map.json")
        self.products = self._load_products()

    def _load_products(self) -> list[dict[str, Any]]:
        path = DATA / "products.json"
        if not path.exists():
            path = SCRATCH / "products_flat.json"
        rows = _load_json(path)
        out = []
        for idx, row in enumerate(rows):
            rec = dict(row)
            rec["id"] = idx
            rec["size_code"] = self._size_code(rec.get("model", ""))
            rec["iec_frame"] = self._iec_frame(rec.get("model", ""))
            rec["stages"] = self._stages(rec.get("model", ""))
            out.append(rec)
        return out

    @staticmethod
    def _size_code(model: str) -> str:
        m = re.search(r"\bR[FK]\s*(\d{2,3})", model.replace("-", " "))
        if not m:
            return ""
        num = m.group(1)
        if len(num) == 3:
            return num[:2]
        return num

    @staticmethod
    def _stages(model: str) -> str:
        m = re.search(r"\bR[FK]\s*\d{2}(\d)", model.replace("-", " "))
        return m.group(1) if m else ""

    @staticmethod
    def _iec_frame(model: str) -> str:
        m = re.search(r"IEC\s*(\d+)", model, re.I)
        return m.group(1) if m else ""

    def series_list(self) -> list[str]:
        return sorted({p["series"] for p in self.products})

    def powers_for_series(self, series: str) -> list[float]:
        vals = sorted({p["P1_kW"] for p in self.products if p["series"] == series})
        return vals

    def ratios_for(self, series: str, power: float) -> list[float]:
        vals = sorted(
            {
                p["i"]
                for p in self.products
                if p["series"] == series and abs(p["P1_kW"] - power) < 1e-9
            }
        )
        return vals

    def matches(self, series: str, power: float, ratio: float) -> list[dict[str, Any]]:
        return [
            p
            for p in self.products
            if p["series"] == series
            and abs(p["P1_kW"] - power) < 1e-9
            and abs(p["i"] - ratio) < 1e-9
        ]

    def get_by_id(self, rec_id: int) -> dict[str, Any] | None:
        for p in self.products:
            if p["id"] == rec_id:
                return p
        return None


def _fmt_num(value: float, nd: int = 2) -> str:
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    text = f"{value:.{nd}f}".rstrip("0").rstrip(".")
    return text


def _n1_rpm(i: float, n2: float) -> float:
    return round(i * n2, 1)


def _shaft_text(lookups: dict, series: str, size: str, shaft_code: str) -> tuple[str, str]:
    types = {
        "AC": "Hollow with keyway",
        "AS": "Solid with key",
        "AD": "Solid with key (double extended)",
        "ACC": "Hollow with shrink disc",
    }
    stype = types.get(shaft_code, types["AC"])
    if shaft_code in ("AC",):
        dmap = lookups["hollow_keyway_D_mm"].get(series, {})
        d = dmap.get(size)
        if d is None:
            return stype, MISSING
        return stype, f"ø{d}H7"
    if shaft_code == "ACC":
        dmap = lookups["shrink_disc_D"].get(series, {})
        d = dmap.get(size)
        if not d:
            return stype, MISSING
        return stype, d
    if shaft_code in ("AS", "AD"):
        if series != "RK":
            return stype, (
                "Field required by DS-021-0046 but not available/identified "
                "in RF-RK catalogue (RF solid-shaft table not found)."
            )
        d = lookups["rk_solid_AS"].get(size)
        if not d:
            return stype, MISSING
        return stype, d
    return stype, MISSING


def build_report(
    catalogue: Catalogue,
    product: dict[str, Any],
    *,
    mounting: str = "M1",
    shaft_code: str = "AC",
    motor_kind: str = "SMX",
    terminal_box: str = "KP1",
    cable_entry: str = "A",
) -> dict[str, Any]:
    """Filter a catalogue record to DS-021-0046 fields only."""
    L = catalogue.lookups
    size = product["size_code"]
    series = product["series"]
    motor = L["motor_common"]

    n1 = _n1_rpm(product["i"], product["n2_rpm"])
    poles = "4" if 1300 <= n1 <= 1500 else MISSING

    stype, sdim = _shaft_text(L, series, size, shaft_code)
    flange = L["output_flange_P_mm"].get(series, {}).get(size)
    oil_q = (
        L["oil_quantity_litres"]
        .get(series, {})
        .get(size, {})
        .get(mounting)
    )

    motor_type = motor["motor_type_smx"] if motor_kind == "SMX" else motor["motor_type_bax"]
    extra = ""
    if motor_kind == "BAX":
        extra = "BAX brake motor"

    family = (
        "Parallel Helical Gear Motors (RF Series) + AC motors SMX.. (IE3)"
        if series == "RF"
        else "Helical-Bevel Gear Motors (RK Series) + AC motors SMX.. (IE3)"
    )
    if motor_kind == "BAX":
        family = family.replace("SMX", "BAX")

    designation = (
        f"GEARBOX {product['model']} 1: { _fmt_num(product['i'], 2) } "
        f"- MOUNTING POSITION {mounting} - OUTPUT FLANGE D{flange if flange else '?'} "
        f"+ {motor_kind} IEC {product['iec_frame']} B5"
    )

    weight_value = (
        f"{_fmt_num(product['mass_gear_unit_kg'], 1)} (gear unit only; motor mass not tabulated)"
    )

    paint = MISSING
    efficiency = MISSING
    current = MISSING
    pf = MISSING

    values = {
        "rated_motor_power": _fmt_num(product["P1_kW"], 2),
        "rated_motor_speed": _fmt_num(n1, 1),
        "overall_gear_ratio": _fmt_num(product["i"], 2),
        "output_speed": _fmt_num(product["n2_rpm"], 1),
        "output_torque": _fmt_num(product["M2_Nm"], 1),
        "service_factor": _fmt_num(product["FS"], 2),
        "output_shaft_type": stype,
        "output_shaft_dimensions": sdim,
        "permitted_output_overhung_load": _fmt_num(product["Fr2_N"], 0),
        "direction_of_rotation": L["direction_of_rotation"]["value"],
        "mounting_position": mounting,
        "flange_diameter": _fmt_num(flange, 0) if flange else MISSING,
        "lubricant_type": L["lubricant_type"]["value"],
        "lubricant_quantity": _fmt_num(oil_q, 2) if oil_q is not None else MISSING,
        "position_of_terminal_box": terminal_box,
        "cable_entry_position": cable_entry,
        "finished_paint_colour": paint,
        "motor_type": motor_type,
        "number_of_poles": poles,
        "duration_factor": motor["duration_factor"],
        "efficiency_class": motor["efficiency_class"],
        "efficiency": efficiency,
        "ce_marking": motor["ce_marking"],
        "motor_voltage": motor["voltage_V"],
        "frequency": motor["frequency_Hz"],
        "rated_current": current,
        "power_factor": pf,
        "insulation_class": motor["insulation_class"],
        "protection_class": motor["protection_class"],
        "design_requirement": motor["design_requirement"],
        "weight": weight_value,
        "additional_features": extra,
    }

    fields_out = []
    for spec in catalogue.ds_fields["fields"]:
        fields_out.append(
            {
                "key": spec["key"],
                "label": spec["label"],
                "unit": spec["unit"],
                "value": values[spec["key"]],
                "source": catalogue.field_map[spec["key"]]["catalogue_source"],
            }
        )

    return {
        "document_id": catalogue.ds_fields["document_id"],
        "product_information": "Product Information",
        "catalog_designation": designation,
        "product_family": family,
        "disclaimer": catalogue.ds_fields["disclaimer"],
        "fields": fields_out,
        "internal": {
            "product_id": product["id"],
            "series": series,
            "model": product["model"],
            "catalogue_page": product.get("catalogue_page"),
            "catalogue_table": product.get("catalogue_table"),
            "weight_note": GEAR_MASS_NOTE,
        },
    }


def display_value(field: dict[str, Any]) -> str:
    unit = field.get("unit") or ""
    value = field["value"]
    if value == MISSING or str(value).startswith("Field required"):
        return str(value)
    if unit:
        return f"({unit}) : {value}"
    return f": {value}"


def clone_report(report: dict[str, Any]) -> dict[str, Any]:
    return deepcopy(report)
