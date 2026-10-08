# -*- coding: utf-8 -*-
"""דף ייצור להדפסה — כחול ופרפרים (הזמנת עומר 04.10.2026)."""
from __future__ import annotations

import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas

from optitex_analyzer.core.baby_basic_note_pdf import (  # noqa: E402
    _FONT,
    _FONT_BOLD,
    _register_fonts,
    _rtl,
)

PAGE_W, PAGE_H = A4

NAVY = (0.12, 0.18, 0.32)
INK = (0.10, 0.13, 0.18)
MUTED = (0.38, 0.42, 0.48)
LINE = (0.78, 0.80, 0.84)
BLUE = (0.11, 0.31, 0.62)
BLUE_BG = (0.90, 0.94, 0.99)
PINK = (0.55, 0.22, 0.42)
PINK_BG = (0.99, 0.93, 0.96)
HEAD_BG = (0.93, 0.94, 0.96)
TOTAL_BG = (0.12, 0.18, 0.32)
WHITE = (1, 1, 1)
ZEBRA = (0.97, 0.97, 0.98)

# מוצר, מידה, פרפרים, כחול
ROWS = [
    ("אוברול", "0–3", 100, 100),
    ("חליפת שינה", "12–18", 25, 50),
    ("חליפת שינה", "18–24", 25, 50),
    ("חליפת שינה", "24–30", 25, 50),
    ("בגד גוף + רגלית", "3–6", 30, 30),
    ("בגד גוף + רגלית", "6–12", 30, 30),
]


def _shown(text: str) -> str:
    text = str(text)
    if not any('\u0590' <= ch <= '\u05FF' for ch in text):
        return text
    protected = re.sub(r"[0-9][0-9./\-–—]*", lambda m: "\u202a" + m.group(0) + "\u202c", text)
    return _rtl(protected)


def _text(c, text, x, y, font, size, color=INK, align="right"):
    c.setFillColorRGB(*color)
    c.setFont(font, size)
    shown = _shown(text)
    if align == "right":
        c.drawRightString(x, y, shown)
    elif align == "center":
        c.drawCentredString(x, y, shown)
    else:
        c.drawString(x, y, shown)


def build(path: str) -> str:
    _register_fonts()
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle("לייצור — כחול ופרפרים")

    margin = 1.15 * cm
    left, right = margin, PAGE_W - margin

    # כותרת
    c.setFillColorRGB(*NAVY)
    c.rect(0, PAGE_H - 2.55 * cm, PAGE_W, 2.55 * cm, fill=1, stroke=0)
    _text(c, "לייצור — כחול ופרפרים", right - 2 * mm, PAGE_H - 1.25 * cm,
          _FONT_BOLD, 22, WHITE, "right")
    _text(c, "בייבי בייסיק  ·  הזמנת עומר  04.10.2026", right - 2 * mm, PAGE_H - 2.05 * cm,
          _FONT, 11, (0.78, 0.84, 0.92), "right")

    # עמודות מימין לשמאל: מוצר, מידה, פרפרים, כחול, סה״כ, בוצע
    col_w = {
        "product": 5.15 * cm,
        "size": 2.55 * cm,
        "fly": 3.15 * cm,
        "blue": 3.15 * cm,
        "total": 2.55 * cm,
        "done": 2.15 * cm,
    }
    x_right = right
    xs = {}
    for key in ("product", "size", "fly", "blue", "total", "done"):
        xs[key] = (x_right - col_w[key], x_right)
        x_right -= col_w[key]

    table_top = PAGE_H - 3.15 * cm
    head_h = 1.05 * cm
    row_h = 1.55 * cm

    headers = [
        ("product", "מוצר", HEAD_BG, INK),
        ("size", "מידה", HEAD_BG, INK),
        ("fly", "פרפרים", PINK, WHITE),
        ("blue", "כחול", BLUE, WHITE),
        ("total", "סה״כ", HEAD_BG, INK),
        ("done", "בוצע", HEAD_BG, INK),
    ]
    y = table_top
    for key, label, bg, fg in headers:
        x0, x1 = xs[key]
        c.setFillColorRGB(*bg)
        c.rect(x0, y - head_h, x1 - x0, head_h, fill=1, stroke=0)
        _text(c, label, (x0 + x1) / 2, y - head_h + 0.34 * cm, _FONT_BOLD, 12, fg, "center")

    y -= head_h
    for i, (product, size, fly, blue) in enumerate(ROWS):
        y0 = y - row_h
        bg = WHITE if i % 2 == 0 else ZEBRA
        for key in ("product", "size", "total", "done"):
            x0, x1 = xs[key]
            c.setFillColorRGB(*bg)
            c.rect(x0, y0, x1 - x0, row_h, fill=1, stroke=0)
        x0, x1 = xs["fly"]
        c.setFillColorRGB(*PINK_BG)
        c.rect(x0, y0, x1 - x0, row_h, fill=1, stroke=0)
        x0, x1 = xs["blue"]
        c.setFillColorRGB(*BLUE_BG)
        c.rect(x0, y0, x1 - x0, row_h, fill=1, stroke=0)

        mid = y0 + 0.52 * cm
        _text(c, product, xs["product"][1] - 4 * mm, mid, _FONT_BOLD, 13, INK, "right")
        _text(c, size, (xs["size"][0] + xs["size"][1]) / 2, mid, _FONT_BOLD, 14, INK, "center")
        _text(c, str(fly), (xs["fly"][0] + xs["fly"][1]) / 2, mid, _FONT_BOLD, 20, PINK, "center")
        _text(c, str(blue), (xs["blue"][0] + xs["blue"][1]) / 2, mid, _FONT_BOLD, 20, BLUE, "center")
        _text(c, str(fly + blue), (xs["total"][0] + xs["total"][1]) / 2, mid, _FONT_BOLD, 16, INK, "center")

        # ריבוע סימון
        box = 7 * mm
        bx = (xs["done"][0] + xs["done"][1] - box) / 2
        by = y0 + (row_h - box) / 2
        c.setStrokeColorRGB(0.45, 0.48, 0.54)
        c.setLineWidth(1.1)
        c.setFillColorRGB(1, 1, 1)
        c.rect(bx, by, box, box, fill=1, stroke=1)

        c.setStrokeColorRGB(*LINE)
        c.setLineWidth(0.4)
        c.line(left, y0, right, y0)
        y = y0

    fly_sum = sum(r[2] for r in ROWS)
    blue_sum = sum(r[3] for r in ROWS)
    grand = fly_sum + blue_sum

    tot_h = 1.35 * cm
    y0 = y - tot_h
    c.setFillColorRGB(*TOTAL_BG)
    c.rect(left, y0, right - left, tot_h, fill=1, stroke=0)
    mid = y0 + 0.42 * cm
    _text(c, "סה״כ", xs["product"][1] - 4 * mm, mid, _FONT_BOLD, 14, WHITE, "right")
    _text(c, str(fly_sum), (xs["fly"][0] + xs["fly"][1]) / 2, mid, _FONT_BOLD, 18, WHITE, "center")
    _text(c, str(blue_sum), (xs["blue"][0] + xs["blue"][1]) / 2, mid, _FONT_BOLD, 18, WHITE, "center")
    _text(c, str(grand), (xs["total"][0] + xs["total"][1]) / 2, mid, _FONT_BOLD, 18, WHITE, "center")

    # מסגרת
    c.setStrokeColorRGB(*NAVY)
    c.setLineWidth(1.2)
    c.rect(left, y0, right - left, table_top - y0, fill=0, stroke=1)
    for key in xs:
        c.setStrokeColorRGB(0.82, 0.84, 0.88)
        c.setLineWidth(0.4)
        c.line(xs[key][0], y0, xs[key][0], table_top)

    # סיכום לפי מוצר
    groups = [
        ("אוברולים 0–3", 200),
        ("חליפות שינה", 225),
        ("בגד גוף + רגלית", 120),
    ]
    box_top = y0 - 0.85 * cm
    box_h = 2.15 * cm
    gap = 0.35 * cm
    box_w = (right - left - 2 * gap) / 3
    for i, (label, qty) in enumerate(groups):
        x1 = right - i * (box_w + gap)
        x0 = x1 - box_w
        c.setFillColorRGB(0.96, 0.97, 0.98)
        c.setStrokeColorRGB(*LINE)
        c.setLineWidth(0.8)
        c.roundRect(x0, box_top - box_h, box_w, box_h, 4, fill=1, stroke=1)
        _text(c, label, (x0 + x1) / 2, box_top - 0.75 * cm, _FONT, 11, MUTED, "center")
        _text(c, str(qty), (x0 + x1) / 2, box_top - 1.65 * cm, _FONT_BOLD, 22, NAVY, "center")

    note_y = box_top - box_h - 0.9 * cm
    _text(c, "אוברול פרפרים 0–3: יש כבר 40 יח׳ במלאי (נכון ל־04.10).",
          right, note_y, _FONT, 10, MUTED, "right")

    c.showPage()
    c.save()
    return path


if __name__ == "__main__":
    out = os.path.join(ROOT, "exports", "baby_basic_notes", "לייצור_כחול_ופרפרים_20261005.pdf")
    print(build(out))
