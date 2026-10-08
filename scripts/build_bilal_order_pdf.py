"""One-off: fabric production order PDF for Bilal (Akdem), 4 solid interlock colors."""
from __future__ import annotations

import os
import urllib.request
from pathlib import Path

import pypdfium2 as pdfium
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "exports" / "_bilal_order_assets"
OUT_PDF = ROOT / "exports" / "Bilal_Akdem_interlock_solids_order.pdf"

PAGE_W, PAGE_H = A4
NAVY = HexColor("#1B2A4A")
NAVY_MID = HexColor("#2C4068")
INK = HexColor("#1C1917")
MUTED = HexColor("#57534E")
RULE = HexColor("#D6D3D1")
PAPER = HexColor("#FAF7F2")
ROW_ALT = HexColor("#F3EFE8")
GOLD = HexColor("#B0894A")
QTY_BG = HexColor("#1B2A4A")

COLORS = [
    {
        "key": "gray",
        "name": "GRAY",
        "code": "131994",
        "code_note": None,
        "pdf_name": "GRAY.pdf",
        "url": (
            "https://v5.airtableusercontent.com/v3/u/56/56/1787234400000/"
            "sUBu_fdlArWTzgjgNiFjEA/epk5NMsWoMXQCLx4wnCK2nUYy3gEkqZ149FKUZpw2BNFa8EPPMz5z0Qqba4rLIHS"
            "1NLAv9fjmPeL45halvVhnQOH-eo7Nc66Ne8yFA285MxX5yiNNC5RwquJe1yt65DWKlWH4FZ5Zy0BLqA_azBwKw/"
            "AmN7Sl8Plyzqpj5C91w1HIDgK7XINlUFa7SqYkzzUd0"
        ),
        "swatch": HexColor("#7A7E84"),
    },
    {
        "key": "pink",
        "name": "PINK",
        "code": "216902",
        "code_note": None,
        "pdf_name": "PINK.pdf",
        "url": (
            "https://v5.airtableusercontent.com/v3/u/56/56/1787234400000/"
            "BbCFER0ulWU2xqWEHMkmPw/hFQd_GetjsUzxnwigu0ly1pbSYUziYvYAWFt02b59MU5f8saursboBR3hciDOWbL"
            "yQmqj7d2OWxFqpzkqVtQQOXLyJZMDjyLA1C7Du1WZ8HKW6lbwnm8zmmR8iCzTC78g4qXkpcVME7MGRfnY9sOXA/"
            "0uj6Ud0NEvaTUuEA_0Nua6WLQCv7wCIbLhpapauDOoU"
        ),
        "swatch": HexColor("#D48AA3"),
    },
    {
        "key": "navy",
        "name": "BLUE NAVY",
        "code": "230416",
        "code_note": None,
        "pdf_name": "BLUE NAVY.pdf",
        "url": (
            "https://v5.airtableusercontent.com/v3/u/56/56/1787234400000/"
            "SWgp-dlagOGWadRTbFHplw/POOMAuIi797Qcw1fATrakpuhstUpp2uLLxGhMpAHIbLFsgrWQ_wvbZioUU4dSjMA"
            "MthIUWzbzH89oSf7NY4tpijLOUJcoMPOXU5bs5nevKShDtizEvIxnWirIOVHAaqq15aHoL1ZIqlF3JX9XnoBDcW"
            "MkcOfNXPUSCN-m21neQs/1TZlLCRsbbB30Yjk-_6zwUj-YugEikYOPEaGr7_ZCwg"
        ),
        "swatch": HexColor("#1E3A5F"),
    },
    {
        "key": "offwhite",
        "name": "OFF WHITE",
        "code": "—",
        "code_note": "No Akdem color code on file. Match this simulation / previous production.",
        "pdf_name": "OFF WHITE.pdf",
        "url": (
            "https://v5.airtableusercontent.com/v3/u/56/56/1787234400000/"
            "JXzsQvn6HzmR76ujbpEVtQ/hMJAAn_CWlxr4CyJXTl3jprl9S9meDCh-8DfyIOaC8yTSlQ3Fe6rbgASUO3czeNT"
            "C_5kLH_PcC85tAmh9BSKgYytsnla4Dz-Z1nX8y5NhrOlxB6UdfLUQTg8vPGUzj2pLGA3fgxc4hw3WFEqOVBhbyc"
            "8zj3cfkQbunTHVf5Z1G0/5mKP_dW5089MsnYyDDaq-VlCNty2gNh2myHboxU1ZrU"
        ),
        "swatch": HexColor("#E8E0D4"),
    },
]

KG_EACH = 150
FABRIC = {
    "Type of fabric": "Interlock",
    "Yarn number": "30/1",
    "Form": "Tubular",
    "Brush": "One side",
    "GSM": "220",
    "Width": "85 cm",
    "Finish": "Solid color only",
}

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


def _download(url: str, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as resp, open(dest, "wb") as f:
        f.write(resp.read())


def _render_pdf_preview(pdf_path: Path, png_path: Path, scale: float = 3.0) -> None:
    doc = pdfium.PdfDocument(str(pdf_path))
    page = doc[0]
    bitmap = page.render(scale=scale)
    pil = bitmap.to_pil()
    pil.save(png_path, "PNG")
    doc.close()


def prepare_assets() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    for item in COLORS:
        pdf_path = ASSETS / item["pdf_name"]
        png_path = ASSETS / f"{item['key']}.png"
        if not pdf_path.exists() or pdf_path.stat().st_size < 1000:
            print(f"Downloading {item['pdf_name']}...")
            _download(item["url"], pdf_path)
        print(f"Rendering {item['pdf_name']} -> {png_path.name}")
        _render_pdf_preview(pdf_path, png_path)


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


def _footer(c: canvas.Canvas, page_no: int, total: int) -> None:
    c.setFillColor(NAVY)
    c.rect(0, 0, PAGE_W, 12 * mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT, 8)
    c.drawString(18 * mm, 4.5 * mm, "Confidential  ·  For Akdem production  ·  Bilal")
    c.drawRightString(PAGE_W - 18 * mm, 4.5 * mm, f"{page_no} / {total}")


def _draw_cover(c: canvas.Canvas) -> None:
    _header_bar(c, "20 August 2026")
    y = PAGE_H - 42 * mm

    c.setFillColor(PAPER)
    c.roundRect(18 * mm, y - 38 * mm, PAGE_W - 36 * mm, 38 * mm, 3, fill=1, stroke=0)

    meta = [
        ("From", "Lior  ·  Aryeh Baby Clothes / Levi Moshe"),
        ("To", "Bilal  ·  Akdem  ·  Bursa"),
        ("Date", "20 August 2026"),
        ("Subject", "Solid color Interlock — 4 colors, 150 kg each"),
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

    y = y - 52 * mm
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 12)
    c.drawString(18 * mm, y, "COMMON FABRIC SPECIFICATION")
    c.setStrokeColor(GOLD)
    c.setLineWidth(1)
    c.line(18 * mm, y - 2 * mm, 92 * mm, y - 2 * mm)

    y -= 12 * mm
    specs = list(FABRIC.items())
    col_w = (PAGE_W - 36 * mm) / 2
    for i, (k, v) in enumerate(specs):
        col = i % 2
        row = i // 2
        x = 18 * mm + col * col_w
        yy = y - row * 11 * mm
        c.setFillColor(MUTED)
        c.setFont(FONT, 8)
        c.drawString(x, yy + 4.5 * mm, k.upper())
        c.setFillColor(INK)
        c.setFont(FONT_B, 11)
        c.drawString(x, yy - 1 * mm, v)

    y = y - 42 * mm
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 12)
    c.drawString(18 * mm, y, "ORDER SUMMARY")
    c.setStrokeColor(GOLD)
    c.line(18 * mm, y - 2 * mm, 58 * mm, y - 2 * mm)

    headers = ["#", "COLOR", "AKDEM CODE", "FABRIC", "GSM", "WIDTH", "QTY"]
    widths = [10, 32, 32, 38, 18, 22, 22]
    table_x = 18 * mm
    table_w = sum(w * mm for w in widths)
    y -= 10 * mm
    row_h = 9 * mm

    c.setFillColor(NAVY)
    c.rect(table_x, y - row_h, table_w, row_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT_B, 7.5)
    x = table_x
    for h, w in zip(headers, widths):
        c.drawString(x + 2 * mm, y - 6 * mm, h)
        x += w * mm

    for i, item in enumerate(COLORS):
        y -= row_h
        c.setFillColor(ROW_ALT if i % 2 else white)
        c.rect(table_x, y - row_h, table_w, row_h, fill=1, stroke=0)
        vals = [
            str(i + 1),
            item["name"],
            item["code"],
            "Interlock 30/1",
            "220",
            "85 cm",
            f"{KG_EACH} kg",
        ]
        c.setFillColor(INK)
        c.setFont(FONT, 8.5)
        x = table_x
        for j, (val, w) in enumerate(zip(vals, widths)):
            if j == 1:
                c.setFillColor(item["swatch"])
                c.circle(x + 3.6 * mm, y - 4.5 * mm, 2.0 * mm, fill=1, stroke=0)
                c.setStrokeColor(RULE)
                c.setLineWidth(0.4)
                c.circle(x + 3.6 * mm, y - 4.5 * mm, 2.0 * mm, fill=0, stroke=1)
                c.setFillColor(INK)
                c.drawString(x + 7.5 * mm, y - 6 * mm, val)
            else:
                c.drawString(x + 2 * mm, y - 6 * mm, val)
            x += w * mm

    y -= row_h + 2 * mm
    c.setFillColor(QTY_BG)
    c.roundRect(table_x, y - 14 * mm, table_w, 14 * mm, 2, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont(FONT, 10)
    c.drawString(table_x + 4 * mm, y - 9 * mm, "TOTAL QUANTITY")
    c.setFont(FONT_B, 16)
    c.drawRightString(table_x + table_w - 4 * mm, y - 9.5 * mm, "600 kg")

    y -= 28 * mm
    c.setFillColor(HexColor("#FEF3C7"))
    c.roundRect(18 * mm, y - 22 * mm, PAGE_W - 36 * mm, 22 * mm, 2, fill=1, stroke=0)
    c.setFillColor(HexColor("#92400E"))
    c.setFont(FONT_B, 9)
    c.drawString(22 * mm, y - 8 * mm, "OFF WHITE — NO AKDEM COLOR CODE")
    c.setFont(FONT, 9)
    c.drawString(
        22 * mm,
        y - 15 * mm,
        "Please match the attached simulation and previous production. Do not substitute another cream/white.",
    )

    c.setFillColor(MUTED)
    c.setFont(FONT, 9)
    c.drawString(
        18 * mm,
        22 * mm,
        "Please confirm receipt, production start date, and expected ready date. Thank you, Lior.",
    )
    _footer(c, 1, 5)


def _draw_color_page(c: canvas.Canvas, item: dict, index: int) -> None:
    _header_bar(c, f"Color {index} of 4")

    left = 18 * mm
    top = PAGE_H - 40 * mm

    c.setFillColor(item["swatch"])
    c.roundRect(left, top - 8 * mm, 10 * mm, 10 * mm, 1.5, fill=1, stroke=0)
    c.setStrokeColor(RULE)
    c.setLineWidth(0.4)
    c.roundRect(left, top - 8 * mm, 10 * mm, 10 * mm, 1.5, fill=0, stroke=1)

    c.setFillColor(NAVY)
    c.setFont(FONT_B, 22)
    c.drawString(left + 14 * mm, top - 3 * mm, item["name"])
    c.setFillColor(MUTED)
    c.setFont(FONT, 10)
    code_line = f"Akdem color code  {item['code']}"
    c.drawString(left + 14 * mm, top - 10 * mm, code_line)

    img_path = ASSETS / f"{item['key']}.png"
    img = ImageReader(str(img_path))
    iw, ih = img.getSize()
    max_w = PAGE_W - 36 * mm
    max_h = 118 * mm
    scale = min(max_w / iw, max_h / ih)
    dw, dh = iw * scale, ih * scale
    img_x = (PAGE_W - dw) / 2
    img_y = top - 18 * mm - dh

    c.setFillColor(PAPER)
    c.roundRect(left, img_y - 4 * mm, PAGE_W - 36 * mm, dh + 8 * mm, 3, fill=1, stroke=0)
    c.drawImage(img, img_x, img_y, width=dw, height=dh, mask="auto")

    spec_y = img_y - 16 * mm
    c.setFillColor(NAVY)
    c.setFont(FONT_B, 11)
    c.drawString(left, spec_y, "PRODUCTION SPEC")
    c.setStrokeColor(GOLD)
    c.setLineWidth(1)
    c.line(left, spec_y - 2 * mm, left + 42 * mm, spec_y - 2 * mm)

    rows = [
        ("Type", "Interlock"),
        ("Yarn", "30/1"),
        ("Form", "Tubular"),
        ("Brush", "One side"),
        ("GSM", "220"),
        ("Width", "85 cm"),
        ("Design", "Solid color only"),
        ("Quantity", f"{KG_EACH} kg"),
    ]
    box_y = spec_y - 8 * mm
    box_h = 36 * mm
    c.setFillColor(PAPER)
    c.roundRect(left, box_y - box_h, PAGE_W - 36 * mm, box_h, 2, fill=1, stroke=0)

    col_w = (PAGE_W - 44 * mm) / 4
    for i, (k, v) in enumerate(rows):
        col = i % 4
        row = i // 4
        x = left + 4 * mm + col * col_w
        yy = box_y - 10 * mm - row * 16 * mm
        c.setFillColor(MUTED)
        c.setFont(FONT, 7.5)
        c.drawString(x, yy + 5 * mm, k.upper())
        c.setFillColor(INK if k != "Quantity" else NAVY)
        c.setFont(FONT_B, 12)
        c.drawString(x, yy - 1 * mm, v)

    if item["code_note"]:
        note_y = box_y - box_h - 8 * mm
        c.setFillColor(HexColor("#FEF3C7"))
        c.roundRect(left, note_y - 16 * mm, PAGE_W - 36 * mm, 16 * mm, 2, fill=1, stroke=0)
        c.setFillColor(HexColor("#92400E"))
        c.setFont(FONT_B, 8.5)
        c.drawString(left + 4 * mm, note_y - 6 * mm, "IMPORTANT")
        c.setFont(FONT, 8.5)
        c.drawString(left + 4 * mm, note_y - 12 * mm, item["code_note"])

    _footer(c, index + 1, 5)


def build_pdf() -> Path:
    _register_fonts()
    OUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT_PDF), pagesize=A4)
    c.setTitle("Fabric Production Order — Solid Interlock — Bilal / Akdem")
    c.setAuthor("Lior / Aryeh Baby Clothes")

    _draw_cover(c)
    c.showPage()
    for i, item in enumerate(COLORS, start=1):
        _draw_color_page(c, item, i)
        c.showPage()
    c.save()
    return OUT_PDF


def main() -> None:
    prepare_assets()
    path = build_pdf()
    print(f"Wrote {path}")


if __name__ == "__main__":
    main()
