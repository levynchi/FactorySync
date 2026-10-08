# -*- coding: utf-8 -*-
"""Garment production order PDF for Omer — solid interlock, 4 products."""
from __future__ import annotations

import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from omer_solids_order_data import (
    BABY_QTY,
    BABY_SIZES,
    COLORS as COLOR_ROWS,
    DATE,
    FABRIC,
    NEWBORN_QTY,
    OUT_PDF_NAME,
    OVERALL_QTY,
    PAJAMA_QTY,
    PAJAMA_SIZES,
    PRODUCTS,
    totals,
)

ROOT = Path(__file__).resolve().parents[1]
OUT_PDF = ROOT / "exports" / OUT_PDF_NAME
SUPPLIER_DIR = ROOT / "supplier_documents" / "4"
SUPPLIERS_JSON = ROOT / "suppliers.json"

PAGE_W, PAGE_H = A4
NAVY = HexColor("#1B2A4A")
INK = HexColor("#1C1917")
MUTED = HexColor("#57534E")
RULE = HexColor("#D6D3D1")
PAPER = HexColor("#FAF7F2")
ROW_ALT = HexColor("#F3EFE8")
GOLD = HexColor("#B0894A")
QTY_BG = HexColor("#1B2A4A")
AMBER_BG = HexColor("#FEF3C7")
AMBER_INK = HexColor("#92400E")
NOTE_BG = HexColor("#E8EEF6")

TOTAL_PAGES = 5
FONT = "OmerSans"
FONT_B = "OmerSansBold"

COLORS = [
    {**row, "swatch": HexColor(hex_code)}
    for row, hex_code in zip(
        COLOR_ROWS,
        ("#1E3A5F", "#D48AA3", "#7A7E84", "#E8E0D4"),
    )
]


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
    c.drawString(18 * mm, PAGE_H - 19 * mm, "GARMENT PRODUCTION ORDER")
    c.setFont(FONT, 9)
    c.drawRightString(PAGE_W - 18 * mm, PAGE_H - 15 * mm, subtitle)


def _footer(c: canvas.Canvas, page_no: int) -> None:
    c.setFillColor(NAVY)
    c.rect(0, 0, PAGE_W, 12 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT, 8)
    c.drawString(18 * mm, 4.5 * mm, "Confidential  ·  For Omer / Nibby Baby  ·  Turkey")
    c.drawRightString(PAGE_W - 18 * mm, 4.5 * mm, f"{page_no} / {TOTAL_PAGES}")


def _section_title(c: canvas.Canvas, x: float, y: float, text: str, rule_w: float = 44) -> None:
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 12)
    c.drawString(x, y, text)
    c.setStrokeColor(GOLD)
    c.setLineWidth(1)
    c.line(x, y - 2 * mm, x + rule_w * mm, y - 2 * mm)


def _swatch_dot(c: canvas.Canvas, x: float, y: float, color, r: float = 2.0 * mm) -> None:
    c.setFillColor(color)
    c.circle(x, y, r, fill=1, stroke=0)
    c.setStrokeColor(RULE)
    c.setLineWidth(0.4)
    c.circle(x, y, r, fill=0, stroke=1)


def _table(
    c: canvas.Canvas,
    x: float,
    y: float,
    headers: list[str],
    widths_mm: list[float],
    rows: list[list],
    row_h: float = 8.2 * mm,
    header_font: float = 7.5,
    body_font: float = 8.5,
    color_col: int | None = None,
) -> float:
    table_w = sum(w * mm for w in widths_mm)
    c.setFillColor(NAVY)
    c.rect(x, y - row_h, table_w, row_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT_B, header_font)
    cx = x
    for h, w in zip(headers, widths_mm):
        c.drawString(cx + 1.8 * mm, y - 5.6 * mm, h)
        cx += w * mm

    for i, row in enumerate(rows):
        y -= row_h
        last = i == len(rows) - 1 and row and str(row[0]).upper() in ("TOTAL", "GRAND TOTAL")
        if last:
            c.setFillColor(QTY_BG)
        else:
            c.setFillColor(ROW_ALT if i % 2 else white)
        c.rect(x, y - row_h, table_w, row_h, fill=1, stroke=0)
        c.setFillColor(white if last else INK)
        c.setFont(FONT_B if last else FONT, body_font)
        cx = x
        for j, (val, w) in enumerate(zip(row, widths_mm)):
            text = "" if val is None else str(val)
            if color_col is not None and j == color_col and not last:
                color = next((col["swatch"] for col in COLORS if col["name"] == text), None)
                if color is not None:
                    _swatch_dot(c, cx + 3.4 * mm, y - 4.1 * mm, color, 1.8 * mm)
                    c.setFillColor(white if last else INK)
                    c.drawString(cx + 6.6 * mm, y - 5.6 * mm, text)
                else:
                    c.drawString(cx + 1.8 * mm, y - 5.6 * mm, text)
            else:
                c.drawString(cx + 1.8 * mm, y - 5.6 * mm, text)
            cx += w * mm
    return y - row_h


def _note_box(c: canvas.Canvas, y: float, title: str, lines: list[str], height_mm: float = 16) -> float:
    c.setFillColor(PAPER)
    c.roundRect(18 * mm, y - height_mm * mm, PAGE_W - 36 * mm, height_mm * mm, 2, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 8.5)
    c.drawString(22 * mm, y - 6 * mm, title)
    c.setFillColor(MUTED)
    c.setFont(FONT, 8.5)
    ny = y - 11.5 * mm
    for line in lines:
        c.drawString(22 * mm, ny, line)
        ny -= 4.4 * mm
    return y - height_mm * mm


def _draw_cover(c: canvas.Canvas) -> None:
    t = totals()
    _header_bar(c, DATE)
    y = PAGE_H - 40 * mm

    c.setFillColor(PAPER)
    c.roundRect(18 * mm, y - 38 * mm, PAGE_W - 36 * mm, 38 * mm, 3, fill=1, stroke=0)
    meta = [
        ("From", "Lior  ·  Aryeh Baby Clothes / Levi Moshe"),
        ("To", "Omer  ·  Nibby Baby  ·  Turkey"),
        ("Date", DATE),
        ("Subject", "Solid interlock garments — 4 products, 4 colors  ·  price request"),
    ]
    row_y = y - 8 * mm
    for label, value in meta:
        c.setFillColor(GOLD)
        c.setFont(FONT_B, 8)
        c.drawString(24 * mm, row_y, label.upper())
        c.setFillColor(INK)
        c.setFont(FONT, 11)
        c.drawString(48 * mm, row_y, value)
        row_y -= 8 * mm

    y = y - 50 * mm
    _section_title(c, 18 * mm, y, "FABRIC SPECIFICATION", 52)
    y -= 12 * mm
    col_w = (PAGE_W - 36 * mm) / 3
    for i, (k, v) in enumerate(FABRIC):
        col = i % 3
        row = i // 3
        x = 18 * mm + col * col_w
        yy = y - row * 12 * mm
        c.setFillColor(MUTED)
        c.setFont(FONT, 8)
        c.drawString(x, yy + 4.5 * mm, k.upper())
        c.setFillColor(INK)
        c.setFont(FONT_B, 11)
        c.drawString(x, yy - 1 * mm, v)

    y = y - 28 * mm
    _section_title(c, 18 * mm, y, "COLORS  ·  MATCH AKDEM CODES", 68)
    y -= 10 * mm
    y = _table(
        c,
        18 * mm,
        y,
        ["#", "COLOR", "AKDEM CODE", "NOTE"],
        [10, 38, 32, 94],
        [
            [str(i), col["name"], col["code"], col["note"]]
            for i, col in enumerate(COLORS, 1)
        ],
        color_col=1,
    )

    y -= 8 * mm
    c.setFillColor(NOTE_BG)
    c.roundRect(18 * mm, y - 32 * mm, PAGE_W - 36 * mm, 32 * mm, 2, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 9)
    c.drawString(22 * mm, y - 6 * mm, "FABRIC SOURCE")
    c.setFillColor(INK)
    c.setFont(FONT, 8.5)
    notes = [
        "Bilal / Akdem has not been available lately for these solid colors.",
        "You may source the same fabric from another mill: 100% cotton interlock, 30/1, 220 GSM, brushed one side.",
        "Colors must match the Akdem codes / attached simulations. Keep leftover fabric at your place.",
        "I will place further orders from that stock. After some time, leftover remaining I will buy from you.",
    ]
    ny = y - 12 * mm
    for line in notes:
        c.drawString(22 * mm, ny, line)
        ny -= 4.4 * mm

    y = y - 42 * mm
    _section_title(c, 18 * mm, y, "PRODUCTS IN THIS ORDER", 56)
    y -= 9 * mm
    qty_map = {
        "A": f"{t['overall']:,} pcs",
        "B": f"{t['pajama']:,} sets",
        "C": f"{t['baby']:,} sets",
        "D": f"{t['newborn']:,} sets",
    }
    model_rows = [
        [p["code"], p["short"], p["construction"], p["sizes"], qty_map[p["code"]]]
        for p in PRODUCTS
    ]
    y = _table(
        c,
        18 * mm,
        y,
        ["#", "PRODUCT", "CONSTRUCTION", "SIZES", "QTY"],
        [10, 40, 62, 38, 24],
        model_rows,
        row_h=8.4 * mm,
    )

    y -= 8 * mm
    c.setFillColor(MUTED)
    c.setFont(FONT, 8.5)
    c.drawString(18 * mm, y, "Please quote FOB USD per piece (A) / per set (B, C, D), excluding Turkish VAT, using the markers.")
    c.drawString(18 * mm, y - 4.5 * mm, "Please also confirm brand / labels (Arye or Baby Basic) before production. Thank you, Lior.")
    _footer(c, 1)


def _model_block(
    c: canvas.Canvas,
    y: float,
    code: str,
    title: str,
    construction: str,
    headers: list[str],
    widths: list[float],
    rows: list[list],
    color_col: int = 1,
    row_h: float = 8.2 * mm,
) -> float:
    c.setFillColor(NAVY)
    c.roundRect(18 * mm, y - 11 * mm, PAGE_W - 36 * mm, 11 * mm, 2, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.setFont(FONT_B, 11)
    c.drawString(22 * mm, y - 7.4 * mm, code)
    c.setFillColor(white)
    c.setFont(FONT_B, 11)
    c.drawString(30 * mm, y - 7.4 * mm, title)
    c.setFont(FONT, 8)
    c.drawRightString(PAGE_W - 22 * mm, y - 7.2 * mm, construction)
    y -= 16 * mm
    return _table(c, 18 * mm, y, headers, widths, rows, row_h=row_h, color_col=color_col)


def _draw_overall(c: canvas.Canvas) -> None:
    t = totals()
    _header_bar(c, "A  ·  Overall")
    y = PAGE_H - 40 * mm
    rows = []
    for size, qty in OVERALL_QTY.items():
        for col in COLORS:
            rows.append([size, col["name"], col["code"], f"{qty}", "pcs"])
    rows.append(["TOTAL", "", "", f"{t['overall']:,}", "pcs"])
    y = _model_block(
        c,
        y,
        "A",
        "OVERALL  /  FOOTED ROMPER",
        "Closed feet  ·  double zipper  ·  0-3, 3-6, 6-12",
        ["SIZE", "COLOR", "AKDEM CODE", "QTY", "UNIT"],
        [28, 42, 42, 32, 30],
        rows,
    )
    y -= 10 * mm
    _note_box(
        c,
        y,
        "PATTERN STATUS  ·  NEW CUT",
        [
            "This is a new overall / sleepsuit pattern. You have not sewn this cut yet.",
            "One garment per unit. Closed feet and double zipper on every size.",
            "Please calculate fabric from the markers. Quote FOB USD per piece, excluding Turkish VAT.",
        ],
        height_mm=22,
    )
    _footer(c, 2)


def _draw_pajama(c: canvas.Canvas) -> None:
    t = totals()
    _header_bar(c, "B  ·  Pajama set")
    y = PAGE_H - 40 * mm
    rows = []
    for size in PAJAMA_SIZES:
        for col in COLORS:
            rows.append([size, col["name"], col["code"], str(PAJAMA_QTY), "sets"])
    rows.append(["TOTAL", "", "", f"{t['pajama']:,}", "sets"])
    y = _model_block(
        c,
        y,
        "B",
        "PAJAMA SET",
        "New long-sleeve shirt + 202 open gatkes  ·  12-18 / 18-24 / 24-30",
        ["SIZE", "COLOR", "AKDEM CODE", "QTY", "UNIT"],
        [28, 42, 42, 32, 30],
        rows,
    )
    y -= 10 * mm
    _note_box(
        c,
        y,
        "PATTERN STATUS  ·  MIXED",
        [
            "One set = 1 long-sleeve shirt + 1 open gatkes (model 202, open-foot pants).",
            "Model 202 you already sewed. The shirt is a new pattern — you have not sewn this cut yet.",
            "Please calculate fabric from the markers. Quote FOB USD per set, excluding Turkish VAT.",
        ],
        height_mm=22,
    )
    _footer(c, 3)


def _draw_baby(c: canvas.Canvas) -> None:
    t = totals()
    _header_bar(c, "C  ·  Baby set")
    y = PAGE_H - 40 * mm
    rows = []
    for size in BABY_SIZES:
        for col in COLORS:
            rows.append([size, col["name"], col["code"], str(BABY_QTY), "sets"])
    rows.append(["TOTAL", "", "", f"{t['baby']:,}", "sets"])
    y = _model_block(
        c,
        y,
        "C",
        "BABY SET",
        "201 wrap bodysuit + footed pants  ·  3-6 / 6-12",
        ["SIZE", "COLOR", "AKDEM CODE", "QTY", "UNIT"],
        [28, 42, 42, 32, 30],
        rows,
    )
    y -= 10 * mm
    _note_box(
        c,
        y,
        "PATTERN STATUS  ·  ALREADY SEWN",
        [
            "One set = 1 envelope wrap bodysuit (201) + 1 footed pants.",
            "You already sewed both patterns (201 wrap and the footed pants).",
            "Please calculate fabric from the markers. Quote FOB USD per set, excluding Turkish VAT.",
        ],
        height_mm=22,
    )
    _footer(c, 4)


def _draw_newborn_and_totals(c: canvas.Canvas) -> None:
    t = totals()
    _header_bar(c, "D  ·  Newborn set  ·  totals")
    y = PAGE_H - 40 * mm
    rows = []
    for col in COLORS:
        rows.append(["0-3", col["name"], col["code"], str(NEWBORN_QTY[col["key"]]), "sets"])
    rows.append(["TOTAL", "", "", f"{t['newborn']:,}", "sets"])
    y = _model_block(
        c,
        y,
        "D",
        "NEWBORN SET",
        "201 wrap bodysuit + footed pants  ·  0-3 only",
        ["SIZE", "COLOR", "AKDEM CODE", "QTY", "UNIT"],
        [28, 42, 42, 32, 30],
        rows,
        row_h=7.6 * mm,
    )
    y -= 8 * mm
    _note_box(
        c,
        y,
        "PATTERN STATUS  ·  ALREADY SEWN",
        [
            "Same set as Baby set: 201 envelope wrap + footed pants. You already sewed both.",
            "Gray 0-3 is 160 sets; Blue Navy / Pink / Off White are 120 sets each.",
        ],
        height_mm=16,
    )

    y -= 24 * mm
    _section_title(c, 18 * mm, y, "ORDER TOTALS", 38)
    y -= 8 * mm
    y = _table(
        c,
        18 * mm,
        y,
        ["#", "PRODUCT", "QTY", "UNIT"],
        [12, 110, 28, 24],
        [
            ["A", "Overall / footed romper", f"{t['overall']:,}", "pcs"],
            ["B", "Pajama set (new shirt + 202)", f"{t['pajama']:,}", "sets"],
            ["C", "Baby set (201 + footed pants)", f"{t['baby']:,}", "sets"],
            ["D", "Newborn set 0-3 (201 + footed pants)", f"{t['newborn']:,}", "sets"],
            ["TOTAL", "A pieces + B/C/D sets", f"{t['units']:,}", ""],
        ],
        row_h=7.4 * mm,
    )

    y -= 8 * mm
    c.setFillColor(AMBER_BG)
    c.roundRect(18 * mm, y - 26 * mm, PAGE_W - 36 * mm, 26 * mm, 2, fill=1, stroke=0)
    c.setFillColor(AMBER_INK)
    c.setFont(FONT_B, 9)
    c.drawString(22 * mm, y - 7 * mm, "PLEASE QUOTE")
    c.setFont(FONT, 8.5)
    c.drawString(22 * mm, y - 13 * mm, "FOB USD per piece (A) and per set (B, C, D), excluding Turkish VAT. Use the markers for yield.")
    c.drawString(22 * mm, y - 18.2 * mm, "Include cutting, sewing, zipper / snaps, labels, packing.")
    c.drawString(22 * mm, y - 23.2 * mm, "Please confirm brand / labels (Arye or Baby Basic) before you start. Thank you, Lior.")
    _footer(c, 5)


def build_pdf() -> Path:
    _register_fonts()
    OUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT_PDF), pagesize=A4)
    c.setTitle("Garment Production Order — Solid Interlock — Omer / Nibby Baby")
    c.setAuthor("Lior / Aryeh Baby Clothes")
    _draw_cover(c)
    c.showPage()
    _draw_overall(c)
    c.showPage()
    _draw_pajama(c)
    c.showPage()
    _draw_baby(c)
    c.showPage()
    _draw_newborn_and_totals(c)
    c.showPage()
    c.save()
    return OUT_PDF


def register_document(src: Path, filename: str) -> Path:
    SUPPLIER_DIR.mkdir(parents=True, exist_ok=True)
    dest = SUPPLIER_DIR / filename
    shutil.copy2(src, dest)
    with open(SUPPLIERS_JSON, "r", encoding="utf-8") as f:
        suppliers = json.load(f)
    omer = next(s for s in suppliers if s.get("id") == 4)
    docs = omer.setdefault("documents", [])
    if not any(d.get("filename") == filename for d in docs):
        next_id = max((d.get("id", 0) for d in docs), default=0) + 1
        docs.append({
            "id": next_id,
            "filename": filename,
            "original_name": filename,
            "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        with open(SUPPLIERS_JSON, "w", encoding="utf-8") as f:
            json.dump(suppliers, f, ensure_ascii=False, indent=2)
            f.write("\n")
    return dest


def main() -> None:
    path = build_pdf()
    dest = register_document(path, OUT_PDF_NAME)
    print(f"Wrote {path}")
    print(f"Copied {dest}")


if __name__ == "__main__":
    main()
