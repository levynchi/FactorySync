"""מחולל PDF A4 לתעודת סחורה בייבי בייסיק (עברית RTL)."""
from __future__ import annotations

import os
from typing import Dict, Optional

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

try:
    from bidi.algorithm import get_display
except Exception:
    def get_display(s):
        return s

PAGE_W, PAGE_H = A4
MARGIN = 1.4 * cm

_FONT = 'BbNoteHe'
_FONT_BOLD = 'BbNoteHeBold'
_FONTS_READY = False

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
HEEBO_MEDIUM = os.path.join('assets', 'fonts', 'Heebo-Medium.ttf')
HEEBO_REGULAR = os.path.join('assets', 'fonts', 'Heebo-Regular.ttf')
LOGO_PATH = os.path.join('assets', 'labels', 'logo_baby_basic.png')


def _abs_path(rel: str) -> str:
    rel = str(rel or '').strip()
    if not rel:
        return ''
    if os.path.isabs(rel):
        return rel
    p_cwd = os.path.join(os.getcwd(), rel)
    if os.path.exists(p_cwd):
        return p_cwd
    p_root = os.path.join(_PROJECT_ROOT, rel)
    return p_root if os.path.exists(p_root) else p_cwd


def _register_fonts():
    global _FONTS_READY
    if _FONTS_READY:
        return
    heebo_m = _abs_path(HEEBO_MEDIUM)
    heebo_r = _abs_path(HEEBO_REGULAR)
    if heebo_m and os.path.exists(heebo_m):
        try:
            pdfmetrics.registerFont(TTFont(_FONT, heebo_r if (heebo_r and os.path.exists(heebo_r)) else heebo_m))
            pdfmetrics.registerFont(TTFont(_FONT_BOLD, heebo_m))
            _FONTS_READY = True
            return
        except Exception:
            pass
    candidates_reg = [r'C:\Windows\Fonts\arial.ttf', r'C:\Windows\Fonts\tahoma.ttf']
    candidates_bold = [r'C:\Windows\Fonts\arialbd.ttf', r'C:\Windows\Fonts\tahomabd.ttf']
    reg = next((p for p in candidates_reg if os.path.exists(p)), None)
    bold = next((p for p in candidates_bold if os.path.exists(p)), reg)
    if reg:
        pdfmetrics.registerFont(TTFont(_FONT, reg))
        pdfmetrics.registerFont(TTFont(_FONT_BOLD, bold or reg))
    _FONTS_READY = True


def _has_hebrew(text: str) -> bool:
    return any('\u0590' <= ch <= '\u05FF' for ch in (text or ''))


def _rtl(text) -> str:
    text = str(text if text is not None else '')
    if _has_hebrew(text):
        try:
            return get_display(text)
        except Exception:
            return text
    return text


def _draw_right(c: canvas.Canvas, text: str, x_right: float, y: float, font: str, size: float, color=(0.12, 0.16, 0.23)):
    c.setFillColorRGB(*color)
    c.setFont(font, size)
    c.drawRightString(x_right, y, _rtl(text))


def _draw_center(c: canvas.Canvas, text: str, x: float, y: float, font: str, size: float, color=(0.12, 0.16, 0.23)):
    c.setFillColorRGB(*color)
    c.setFont(font, size)
    c.drawCentredString(x, y, _rtl(text))


def _money(value) -> str:
    try:
        return f"{float(value):,.2f}"
    except Exception:
        return "0.00"


def _qty(value) -> str:
    try:
        n = float(value)
        return str(int(n)) if n == int(n) else f"{n:g}"
    except Exception:
        return "0"


def generate_baby_basic_note_pdf(rec: Dict, file_path: str, biz_name: str = '') -> str:
    """יוצר PDF A4 של תעודת סחורה. מחזיר את הנתיב."""
    _register_fonts()
    os.makedirs(os.path.dirname(file_path) or '.', exist_ok=True)

    left, right = MARGIN, PAGE_W - MARGIN
    usable_w = right - left
    lines = list(rec.get('lines') or [])

    # visual RTL: מוצר | מידה | בד | צבע | ברקוד | יחידות | מחיר | סה״כ
    col_defs = [
        ('name', 4.0 * cm),
        ('size', 1.7 * cm),
        ('fabric', 1.6 * cm),
        ('color', 1.5 * cm),
        ('barcode', 3.1 * cm),
        ('qty', 1.7 * cm),
        ('price', 2.1 * cm),
        ('total', 2.2 * cm),
    ]
    widths = [w for _, w in col_defs]
    leftover = usable_w - sum(widths)
    if leftover != 0:
        widths[0] = max(2.4 * cm, widths[0] + leftover)
    col_rights = []
    x = right
    for w in widths:
        col_rights.append(x)
        x -= w

    headers_he = {
        'name': 'מוצר',
        'size': 'מידה',
        'fabric': 'בד',
        'color': 'צבע',
        'barcode': 'ברקוד',
        'qty': 'יחידות',
        'price': 'מחיר ליחידה',
        'total': 'סה״כ',
    }
    row_h = 0.62 * cm
    header_h = 0.7 * cm
    bottom_limit = MARGIN + 1.4 * cm

    c = canvas.Canvas(file_path, pagesize=A4)

    def draw_header(y: float, first_page: bool) -> float:
        if first_page:
            logo = _abs_path(LOGO_PATH)
            if logo and os.path.exists(logo):
                try:
                    img = ImageReader(logo)
                    iw, ih = img.getSize()
                    h = 1.6 * cm
                    w = h * (iw / float(ih or 1))
                    c.drawImage(img, PAGE_W / 2 - w / 2, y - h + 0.15 * cm, width=w, height=h, mask='auto')
                    y -= h + 0.25 * cm
                except Exception:
                    pass
            title = biz_name.strip() or 'בייבי בייסיק'
            _draw_center(c, title, PAGE_W / 2, y, _FONT_BOLD, 16)
            y -= 0.65 * cm
            _draw_center(c, f"תעודת סחורה #{rec.get('id')}", PAGE_W / 2, y, _FONT_BOLD, 13)
            y -= 0.7 * cm
            _draw_right(c, f"לקוח: {rec.get('customer') or ''}", right, y, _FONT, 11)
            _draw_right(c, f"תאריך: {rec.get('date') or ''}", left + 5.5 * cm, y, _FONT, 11)
            y -= 0.5 * cm
            if rec.get('note'):
                _draw_right(c, f"הערה: {rec.get('note')}", right, y, _FONT, 10, color=(0.4, 0.45, 0.5))
                y -= 0.45 * cm
            c.setStrokeColorRGB(0.23, 0.51, 0.96)
            c.setLineWidth(1.6)
            c.line(left, y, right, y)
            y -= 0.45 * cm
        else:
            _draw_right(c, f"תעודת סחורה #{rec.get('id')} (המשך)", right, y, _FONT_BOLD, 11)
            y -= 0.55 * cm
        return y

    def draw_table_header(y: float) -> float:
        c.setFillColorRGB(0.23, 0.51, 0.96)
        c.rect(left, y - header_h + 0.18 * cm, usable_w, header_h, fill=1, stroke=0)
        for i, (key, _) in enumerate(col_defs):
            _draw_right(c, headers_he[key], col_rights[i] - 0.12 * cm, y - 0.28 * cm, _FONT_BOLD, 8, color=(1, 1, 1))
        return y - header_h - 0.08 * cm

    def draw_footer():
        c.setFillColorRGB(0.4, 0.45, 0.5)
        c.setFont(_FONT, 8)
        c.drawCentredString(PAGE_W / 2, MARGIN * 0.45, _rtl(f"עמוד {c.getPageNumber()}"))

    y = PAGE_H - MARGIN
    y = draw_header(y, True)
    y = draw_table_header(y)

    for idx, line in enumerate(lines):
        if y < bottom_limit + row_h:
            draw_footer()
            c.showPage()
            y = PAGE_H - MARGIN
            y = draw_header(y, False)
            y = draw_table_header(y)
        if idx % 2 == 0:
            c.setFillColorRGB(0.96, 0.97, 0.99)
            c.rect(left, y - row_h + 0.18 * cm, usable_w, row_h, fill=1, stroke=0)
        values = {
            'name': line.get('print_name') or line.get('item_name') or '',
            'size': line.get('size', ''),
            'fabric': line.get('fabric', ''),
            'color': line.get('color', '') or 'לבן',
            'barcode': str(line.get('barcode', '')),
            'qty': _qty(line.get('quantity')),
            'price': _money(line.get('unit_price')),
            'total': _money(line.get('line_total')),
        }
        for i, (key, _) in enumerate(col_defs):
            _draw_right(c, values[key], col_rights[i] - 0.12 * cm, y - 0.22 * cm, _FONT, 8)
        y -= row_h

    y -= 0.35 * cm
    box_h = 1.35 * cm
    if y - box_h < MARGIN:
        draw_footer()
        c.showPage()
        y = PAGE_H - MARGIN - 0.3 * cm
    c.setFillColorRGB(0.93, 0.96, 1.0)
    c.roundRect(left, y - box_h + 0.35 * cm, usable_w, box_h, 5, fill=1, stroke=0)
    _draw_right(
        c,
        f"סה״כ יחידות: {_qty(rec.get('total_quantity'))}",
        right - 0.35 * cm,
        y - 0.05 * cm,
        _FONT_BOLD,
        11,
        color=(0.12, 0.16, 0.23),
    )
    _draw_right(
        c,
        f"סה״כ לתשלום: {_money(rec.get('total_amount'))} ₪",
        right - 0.35 * cm,
        y - 0.6 * cm,
        _FONT_BOLD,
        13,
        color=(0.15, 0.35, 0.75),
    )
    draw_footer()
    c.save()
    return file_path
