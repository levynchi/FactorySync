"""Fabric production order PDF for Bilal: white interlock + white rib."""
from __future__ import annotations

import os
from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[1]
OUT_PDF = ROOT / "exports" / "Bilal_Akdem_white_interlock_rib_order.pdf"

PAGE_W, PAGE_H = A4
NAVY = HexColor("#1B2A4A")
INK = HexColor("#1C1917")
MUTED = HexColor("#57534E")
RULE = HexColor("#D6D3D1")
PAPER = HexColor("#FAF7F2")
ROW_ALT = HexColor("#F3EFE8")
GOLD = HexColor("#B0894A")
QTY_BG = HexColor("#1B2A4A")
SWATCH = HexColor("#F7F4EE")

DATE = "23 August 2026"
TOTAL_PAGES = 3

LINES = [
    {
        "key": "interlock",
        "name": "WHITE INTERLOCK",
        "short": "Interlock",
        "code": "110257",
        "qty": 800,
        "yarn": "30/1",
        "form": "Tubular",
        "brush": "One side, twice (as usual)",
        "gsm": "220",
        "width": "85 cm",
        "design": "Solid color only",
        "note": "Same white as previous production. Color code 110257. Do not substitute off-white.",
    },
    {
        "key": "rib",
        "name": "WHITE RIB 1/1",
        "short": "Rib 1/1",
        "code": "110257",
        "qty": 500,
        "yarn": "1/1",
        "form": "Tubular",
        "brush": "Without brush",
        "gsm": "180",
        "width": "85 cm",
        "design": "Solid color only",
        "note": "Same white as the interlock (110257). Match previous white rib production.",
    },
]

FONT = "BilalSans"
FONT_B = "BilalSansBold"


def _register_fonts() -> None:
    regular = next(
        (
            p
            for p in (
                r"C:\Windows\Fonts\calibri.ttf",
                r"C:\Windows\Fonts\arial.ttf",
                r"C:\Windows\Fonts\segoeui.ttf",
            )
            if os.path.exists(p)
        ),
        None,
    )
    bold = next(
        (
            p
            for p in (
                r"C:\Windows\Fonts\calibrib.ttf",
                r"C:\Windows\Fonts\arialbd.ttf",
                r"C:\Windows\Fonts\segoeuib.ttf",
            )
            if os.path.exists(p)
        ),
        regular,
    )
    if not regular:
        return
    pdfmetrics.registerFont(TTFont(FONT, regular))
    pdfmetrics.registerFont(TTFont(FONT_B, bold or regular))


def _header_bar(c: canvas.Canvas, subtitle: str) -> None:
    c.setFillColor(NAVY)
    c.rect(0, PAGE_H - 28 * mm, PAGE_W, 28 * mm, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.rect(0, PAGE_H - 29.2 * mm, PAGE_W, 1.2 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT, 9)
    c.drawString(18 * mm, PAGE_H - 11 * mm, "ARYEH BABY CLOTHES  ·  LEVI MOSHE")
    c.setFont(FONT_B, 13)
    c.drawString(18 * mm, PAGE_H - 19 * mm, "FABRIC PRODUCTION ORDER")
    c.setFont(FONT, 9)
    c.drawRightString(PAGE_W - 18 * mm, PAGE_H - 15 * mm, subtitle)


def _footer(c: canvas.Canvas, page_no: int) -> None:
    c.setFillColor(NAVY)
    c.rect(0, 0, PAGE_W, 12 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT, 8)
    c.drawString(18 * mm, 4.5 * mm, "Confidential  ·  For Akdem production  ·  Bilal")
    c.drawRightString(PAGE_W - 18 * mm, 4.5 * mm, f"{page_no} / {TOTAL_PAGES}")


def _draw_cover(c: canvas.Canvas) -> None:
    _header_bar(c, DATE)
    y = PAGE_H - 42 * mm

    c.setFillColor(PAPER)
    c.roundRect(18 * mm, y - 38 * mm, PAGE_W - 36 * mm, 38 * mm, 3, fill=1, stroke=0)

    meta = [
        ("From", "Lior  ·  Aryeh Baby Clothes / Levi Moshe"),
        ("To", "Bilal  ·  Akdem  ·  Bursa"),
        ("Date", DATE),
        ("Subject", "White Interlock 800 kg  +  White Rib 1/1 500 kg"),
    ]
    row_y = y - 8 * mm
    for label, value in meta:
        c.setFillColor(GOLD)
        c.setFont(FONT_B, 8)
        c.drawString(24 * mm, row_y, label.upper())
        c.setFillColor(INK)
        c.setFont(FONT, 11)
        c.drawString(52 * mm, row_y, value)
        row_y -= 8 * mm

    y = y - 54 * mm
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 12)
    c.drawString(18 * mm, y, "ORDER SUMMARY")
    c.setStrokeColor(GOLD)
    c.setLineWidth(1)
    c.line(18 * mm, y - 2 * mm, 58 * mm, y - 2 * mm)

    headers = ["#", "FABRIC", "COLOR", "AKDEM CODE", "GSM", "WIDTH", "QTY"]
    widths = [10, 38, 28, 32, 18, 22, 22]
    table_x = 18 * mm
    table_w = sum(w * mm for w in widths)
    y -= 10 * mm
    row_h = 11 * mm

    c.setFillColor(NAVY)
    c.rect(table_x, y - row_h, table_w, row_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT_B, 7.5)
    x = table_x
    for h, w in zip(headers, widths):
        c.drawString(x + 2 * mm, y - 7 * mm, h)
        x += w * mm

    for i, item in enumerate(LINES):
        y -= row_h
        c.setFillColor(ROW_ALT if i % 2 else white)
        c.rect(table_x, y - row_h, table_w, row_h, fill=1, stroke=0)
        vals = [
            str(i + 1),
            item["short"],
            "WHITE",
            item["code"],
            item["gsm"],
            item["width"],
            f"{item['qty']} kg",
        ]
        c.setFillColor(INK)
        c.setFont(FONT, 9)
        x = table_x
        for j, (val, w) in enumerate(zip(vals, widths)):
            if j == 2:
                c.setFillColor(SWATCH)
                c.circle(x + 3.6 * mm, y - 5.5 * mm, 2.0 * mm, fill=1, stroke=0)
                c.setStrokeColor(RULE)
                c.setLineWidth(0.5)
                c.circle(x + 3.6 * mm, y - 5.5 * mm, 2.0 * mm, fill=0, stroke=1)
                c.setFillColor(INK)
                c.drawString(x + 7.5 * mm, y - 7 * mm, val)
            else:
                c.drawString(x + 2 * mm, y - 7 * mm, val)
            x += w * mm

    y -= row_h + 3 * mm
    c.setFillColor(QTY_BG)
    c.roundRect(table_x, y - 16 * mm, table_w, 16 * mm, 2, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT, 10)
    c.drawString(table_x + 4 * mm, y - 10 * mm, "TOTAL QUANTITY")
    c.setFont(FONT_B, 16)
    c.drawRightString(table_x + table_w - 4 * mm, y - 10.5 * mm, "1,300 kg")

    y -= 32 * mm
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 12)
    c.drawString(18 * mm, y, "PLEASE PRODUCE AS USUAL")
    c.setStrokeColor(GOLD)
    c.line(18 * mm, y - 2 * mm, 78 * mm, y - 2 * mm)

    notes = [
        "1.  Color: our regular WHITE — Akdem color code 110257 for both fabrics.",
        "2.  Interlock: 30/1, tubular, 220 GSM, width 85 cm, brushed twice on one side.",
        "3.  Rib 1/1: tubular, 180 GSM, width 85 cm, without brush.",
        "4.  Solid color only — no print, no design.",
        "5.  Please match previous white production. Do not use off-white / cream.",
    ]
    y -= 12 * mm
    c.setFillColor(PAPER)
    c.roundRect(18 * mm, y - 48 * mm, PAGE_W - 36 * mm, 50 * mm, 2, fill=1, stroke=0)
    c.setFillColor(INK)
    c.setFont(FONT, 10)
    ty = y - 8 * mm
    for line in notes:
        c.drawString(22 * mm, ty, line)
        ty -= 8 * mm

    c.setFillColor(MUTED)
    c.setFont(FONT, 9)
    c.drawString(
        18 * mm,
        22 * mm,
        "Please confirm receipt, production start date, and expected ready date. Thank you, Lior.",
    )
    _footer(c, 1)


def _draw_line_page(c: canvas.Canvas, item: dict, index: int) -> None:
    _header_bar(c, f"Line {index} of 2")

    left = 18 * mm
    top = PAGE_H - 42 * mm

    c.setFillColor(SWATCH)
    c.roundRect(left, top - 8 * mm, 12 * mm, 12 * mm, 1.5, fill=1, stroke=0)
    c.setStrokeColor(RULE)
    c.setLineWidth(0.6)
    c.roundRect(left, top - 8 * mm, 12 * mm, 12 * mm, 1.5, fill=0, stroke=1)

    c.setFillColor(NAVY)
    c.setFont(FONT_B, 20)
    c.drawString(left + 16 * mm, top - 1 * mm, item["name"])
    c.setFillColor(MUTED)
    c.setFont(FONT, 11)
    c.drawString(left + 16 * mm, top - 9 * mm, f"Akdem color code  {item['code']}")

    swatch_y = top - 28 * mm
    swatch_h = 62 * mm
    c.setFillColor(PAPER)
    c.roundRect(left, swatch_y - swatch_h, PAGE_W - 36 * mm, swatch_h, 3, fill=1, stroke=0)
    c.setFillColor(SWATCH)
    c.roundRect(left + 8 * mm, swatch_y - swatch_h + 8 * mm, PAGE_W - 52 * mm, swatch_h - 16 * mm, 2, fill=1, stroke=0)
    c.setStrokeColor(RULE)
    c.setLineWidth(0.5)
    c.roundRect(left + 8 * mm, swatch_y - swatch_h + 8 * mm, PAGE_W - 52 * mm, swatch_h - 16 * mm, 2, fill=0, stroke=1)
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 16)
    c.drawCentredString(PAGE_W / 2, swatch_y - swatch_h / 2 + 6 * mm, "WHITE")
    c.setFont(FONT, 11)
    c.setFillColor(MUTED)
    c.drawCentredString(PAGE_W / 2, swatch_y - swatch_h / 2 - 4 * mm, f"Akdem  {item['code']}")

    spec_y = swatch_y - swatch_h - 14 * mm
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 11)
    c.drawString(left, spec_y, "PRODUCTION SPEC")
    c.setStrokeColor(GOLD)
    c.setLineWidth(1)
    c.line(left, spec_y - 2 * mm, left + 42 * mm, spec_y - 2 * mm)

    rows = [
        ("Type", item["short"]),
        ("Yarn", item["yarn"]),
        ("Form", item["form"]),
        ("Brush", item["brush"]),
        ("GSM", item["gsm"]),
        ("Width", item["width"]),
        ("Design", item["design"]),
        ("Quantity", f"{item['qty']} kg"),
    ]
    box_y = spec_y - 8 * mm
    box_h = 42 * mm
    c.setFillColor(PAPER)
    c.roundRect(left, box_y - box_h, PAGE_W - 36 * mm, box_h, 2, fill=1, stroke=0)

    col_w = (PAGE_W - 44 * mm) / 4
    for i, (k, v) in enumerate(rows):
        col = i % 4
        row = i // 4
        x = left + 4 * mm + col * col_w
        yy = box_y - 12 * mm - row * 18 * mm
        c.setFillColor(MUTED)
        c.setFont(FONT, 7.5)
        c.drawString(x, yy + 6 * mm, k.upper())
        c.setFillColor(INK if k != "Quantity" else NAVY)
        c.setFont(FONT_B, 11)
        c.drawString(x, yy - 1 * mm, v)

    note_y = box_y - box_h - 10 * mm
    c.setFillColor(HexColor("#FEF3C7"))
    c.roundRect(left, note_y - 18 * mm, PAGE_W - 36 * mm, 18 * mm, 2, fill=1, stroke=0)
    c.setFillColor(HexColor("#92400E"))
    c.setFont(FONT_B, 8.5)
    c.drawString(left + 4 * mm, note_y - 7 * mm, "IMPORTANT")
    c.setFont(FONT, 9)
    c.drawString(left + 4 * mm, note_y - 13.5 * mm, item["note"])

    _footer(c, index + 1)


def build_pdf() -> Path:
    _register_fonts()
    OUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT_PDF), pagesize=A4)
    c.setTitle("Fabric Production Order — White Interlock + White Rib — Bilal / Akdem")
    c.setAuthor("Lior / Aryeh Baby Clothes")

    _draw_cover(c)
    c.showPage()
    for i, item in enumerate(LINES, start=1):
        _draw_line_page(c, item, i)
        c.showPage()
    c.save()
    return OUT_PDF


def main() -> None:
    path = build_pdf()
    print(f"Wrote {path}")


if __name__ == "__main__":
    main()
