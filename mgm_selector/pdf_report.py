"""DS-021-0046 PDF generator. Emits only mapped DS fields."""

from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from .catalogue import MISSING, display_value

NAVY = colors.HexColor("#003366")
BLUE = colors.HexColor("#0055a5")
LIGHT = colors.HexColor("#e8eef5")
RULE = colors.HexColor("#8aa0b8")


def generate_pdf(report: dict, output_path: str | Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=14 * mm,
        title=f"Product Data {report['document_id']}",
        author="MGM-Varvel Product Selector",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "TitleBar",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=16,
        textColor=NAVY,
        spaceAfter=2,
    )
    label = ParagraphStyle(
        "Lab",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        textColor=colors.HexColor("#333333"),
        leading=11,
    )
    value_style = ParagraphStyle(
        "Val",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        textColor=NAVY,
        leading=11,
    )
    small = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#222222"),
    )
    disc = ParagraphStyle(
        "Disc",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=7,
        leading=9,
        textColor=colors.HexColor("#444444"),
    )

    story = []
    story.append(Paragraph("Product Information", title))
    story.append(
        Table(
            [[""]],
            colWidths=[182 * mm],
            rowHeights=[2],
            style=TableStyle([("BACKGROUND", (0, 0), (-1, -1), BLUE)]),
        )
    )
    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>Catalog Designation</b>", label))
    story.append(Paragraph(report["catalog_designation"], small))
    story.append(Spacer(1, 3))
    story.append(Paragraph(report["product_family"], small))
    story.append(Spacer(1, 8))

    head = Table(
        [[
            Paragraph("<b>Product Data</b>", label),
            Paragraph(f"<b>{report['document_id']}</b>", value_style),
        ]],
        colWidths=[140 * mm, 42 * mm],
    )
    head.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
            ]
        )
    )
    story.append(head)
    story.append(Spacer(1, 4))

    rows = []
    for field in report["fields"]:
        shown = display_value(field)
        if field["value"] == MISSING or str(field["value"]).startswith("Field required"):
            shown_html = f'<font color="#a33">{shown}</font>'
        else:
            shown_html = shown
        rows.append(
            [
                Paragraph(field["label"], label),
                Paragraph(shown_html, value_style),
            ]
        )

    table = Table(rows, colWidths=[78 * mm, 104 * mm])
    table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2.2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2),
                ("LINEBELOW", (0, 0), (-1, -2), 0.25, RULE),
                ("BACKGROUND", (0, 0), (-1, 0), colors.white),
            ]
        )
    )
    story.append(table)
    story.append(Spacer(1, 10))
    story.append(Paragraph(report["disclaimer"], disc))
    doc.build(story)
    return output_path
