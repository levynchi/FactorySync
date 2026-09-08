"""מחולל PDF A4 לתעודת סחורה בייבי בייסיק (עברית RTL)."""
from __future__ import annotations

import os
from typing import Dict, List, Optional

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


def _draw_brand_header(c: canvas.Canvas, y: float, title: str, biz_name: str) -> float:
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
    heading = (biz_name or '').strip() or 'בייבי בייסיק'
    _draw_center(c, heading, PAGE_W / 2, y, _FONT_BOLD, 16)
    y -= 0.65 * cm
    _draw_center(c, title, PAGE_W / 2, y, _FONT_BOLD, 13)
    y -= 0.55 * cm
    c.setStrokeColorRGB(0.23, 0.51, 0.96)
    c.setLineWidth(1.6)
    c.line(MARGIN, y, PAGE_W - MARGIN, y)
    return y - 0.5 * cm


def generate_baby_basic_payment_pdf(rec: Dict, file_path: str, biz_name: str = '',
                                   partner: str = '', supplied=None, paid=None, balance=None) -> str:
    """יוצר PDF A4 של אישור תשלום בודד."""
    _register_fonts()
    os.makedirs(os.path.dirname(file_path) or '.', exist_ok=True)
    left, right = MARGIN, PAGE_W - MARGIN
    usable_w = right - left
    partner = partner or rec.get('customer') or 'בייבי בייסיק'

    c = canvas.Canvas(file_path, pagesize=A4)
    y = PAGE_H - MARGIN
    y = _draw_brand_header(c, y, f"אישור תשלום #{rec.get('id')}", biz_name)

    _draw_right(c, f"שותף: {partner}", right, y, _FONT, 11)
    _draw_right(c, f"תאריך: {rec.get('date') or ''}", left + 5.5 * cm, y, _FONT, 11)
    y -= 0.55 * cm
    if rec.get('created_at'):
        _draw_right(c, f"נרשם ב: {rec.get('created_at')}", right, y, _FONT, 9, color=(0.4, 0.45, 0.5))
        y -= 0.45 * cm

    box_h = 2.4 * cm
    c.setFillColorRGB(0.93, 0.96, 1.0)
    c.roundRect(left, y - box_h, usable_w, box_h, 6, fill=1, stroke=0)
    _draw_center(c, 'סכום ששולם', PAGE_W / 2, y - 0.75 * cm, _FONT, 11, color=(0.4, 0.45, 0.5))
    _draw_center(c, f"{_money(rec.get('amount'))} ₪", PAGE_W / 2, y - 1.55 * cm, _FONT_BOLD, 22, color=(0.15, 0.35, 0.75))
    y -= box_h + 0.55 * cm

    if rec.get('note'):
        _draw_right(c, f"הערה: {rec.get('note')}", right, y, _FONT, 11)
        y -= 0.7 * cm

    if supplied is not None or paid is not None or balance is not None:
        y -= 0.15 * cm
        _draw_right(c, 'מצב חשבון שותף', right, y, _FONT_BOLD, 11)
        y -= 0.5 * cm
        if supplied is not None:
            _draw_right(c, f"סופק: {_money(supplied)} ₪", right, y, _FONT, 10)
            y -= 0.42 * cm
        if paid is not None:
            _draw_right(c, f"שולם כולל תשלום זה: {_money(paid)} ₪", right, y, _FONT, 10)
            y -= 0.42 * cm
        if balance is not None:
            _draw_right(c, f"יתרה: {_money(balance)} ₪", right, y, _FONT_BOLD, 11, color=(0.15, 0.35, 0.75))

    c.setFillColorRGB(0.4, 0.45, 0.5)
    c.setFont(_FONT, 8)
    c.drawCentredString(PAGE_W / 2, MARGIN * 0.45, _rtl(f"עמוד {c.getPageNumber()}"))
    c.save()
    return file_path


def generate_baby_basic_payments_report_pdf(payments: list, file_path: str, biz_name: str = '',
                                           partner: str = '', supplied=None, paid=None, balance=None) -> str:
    """יוצר PDF A4 של כל תשלומי בייבי בייסיק."""
    _register_fonts()
    os.makedirs(os.path.dirname(file_path) or '.', exist_ok=True)
    left, right = MARGIN, PAGE_W - MARGIN
    usable_w = right - left
    partner = partner or 'בייבי בייסיק'
    rows = sorted(list(payments or []), key=lambda p: (str(p.get('date') or ''), int(p.get('id') or 0)))

    col_defs = [
        ('date', 3.4 * cm),
        ('id', 2.2 * cm),
        ('amount', 3.4 * cm),
        ('note', 8.8 * cm),
    ]
    widths = [w for _, w in col_defs]
    leftover = usable_w - sum(widths)
    if leftover != 0:
        widths[-1] = max(3.0 * cm, widths[-1] + leftover)
    col_rights = []
    x = right
    for w in widths:
        col_rights.append(x)
        x -= w
    headers_he = {'date': 'תאריך', 'id': 'מס׳', 'amount': 'סכום ₪', 'note': 'הערה'}
    row_h = 0.62 * cm
    header_h = 0.7 * cm
    bottom_limit = MARGIN + 1.4 * cm

    c = canvas.Canvas(file_path, pagesize=A4)

    def draw_header(y: float, first_page: bool) -> float:
        if first_page:
            y = _draw_brand_header(c, y, 'דוח תשלומים', biz_name)
            _draw_right(c, f"שותף: {partner}", right, y, _FONT, 11)
            y -= 0.55 * cm
        else:
            _draw_right(c, 'דוח תשלומים (המשך)', right, y, _FONT_BOLD, 11)
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

    if not rows:
        _draw_right(c, 'אין תשלומים רשומים.', right, y, _FONT, 11, color=(0.4, 0.45, 0.5))
        y -= 0.8 * cm

    for idx, rec in enumerate(rows):
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
            'date': rec.get('date') or '',
            'id': str(rec.get('id') or ''),
            'amount': _money(rec.get('amount')),
            'note': rec.get('note') or '',
        }
        for i, (key, _) in enumerate(col_defs):
            _draw_right(c, values[key], col_rights[i] - 0.12 * cm, y - 0.22 * cm, _FONT, 8)
        y -= row_h

    y -= 0.4 * cm
    box_h = 1.7 * cm
    if y - box_h < MARGIN:
        draw_footer()
        c.showPage()
        y = PAGE_H - MARGIN - 0.3 * cm
    total_paid = paid if paid is not None else sum(float(p.get('amount') or 0) for p in rows)
    c.setFillColorRGB(0.93, 0.96, 1.0)
    c.roundRect(left, y - box_h + 0.35 * cm, usable_w, box_h, 5, fill=1, stroke=0)
    _draw_right(c, f"מספר תשלומים: {len(rows)}", right - 0.35 * cm, y - 0.05 * cm, _FONT_BOLD, 11)
    _draw_right(c, f"סה״כ שולם: {_money(total_paid)} ₪", right - 0.35 * cm, y - 0.55 * cm, _FONT_BOLD, 13, color=(0.15, 0.35, 0.75))
    if supplied is not None and balance is not None:
        _draw_right(
            c,
            f"סופק: {_money(supplied)} ₪    יתרה: {_money(balance)} ₪",
            right - 0.35 * cm,
            y - 1.05 * cm,
            _FONT,
            10,
        )
    draw_footer()
    c.save()
    return file_path


def generate_baby_basic_room_count_pdf(rec: Dict, file_path: str, biz_name: str = '') -> str:
    """יוצר PDF A4 של ספירת מלאי חדר בייבי בייסיק (דגם / סוג בד / צבע רקע / הדפס / מידה, בלי ברקוד)."""
    _register_fonts()
    os.makedirs(os.path.dirname(file_path) or '.', exist_ok=True)
    left, right = MARGIN, PAGE_W - MARGIN
    usable_w = right - left
    lines = list(rec.get('lines') or [])

    col_defs = [
        ('print_name', 3.2 * cm),
        ('fabric', 2.0 * cm),
        ('color', 2.4 * cm),
        ('print', 2.6 * cm),
        ('size', 1.8 * cm),
        ('area', 1.4 * cm),
        ('box', 1.6 * cm),
        ('qty', 1.8 * cm),
    ]
    widths = [w for _, w in col_defs]
    leftover = usable_w - sum(widths)
    if leftover != 0:
        widths[0] = max(2.2 * cm, widths[0] + leftover)
    col_rights = []
    x = right
    for w in widths:
        col_rights.append(x)
        x -= w
    headers_he = {
        'print_name': 'דגם',
        'fabric': 'סוג בד',
        'color': 'צבע רקע',
        'print': 'הדפס',
        'size': 'מידה',
        'area': 'אזור',
        'box': 'קופסא',
        'qty': 'כמות',
    }
    row_h = 0.62 * cm
    header_h = 0.7 * cm
    bottom_limit = MARGIN + 1.4 * cm

    c = canvas.Canvas(file_path, pagesize=A4)

    def draw_header(y: float, first_page: bool) -> float:
        if first_page:
            y = _draw_brand_header(c, y, f"ספירת מלאי חדר #{rec.get('id')}", biz_name)
            _draw_right(c, f"תאריך: {rec.get('date') or ''}", right, y, _FONT, 11)
            y -= 0.5 * cm
            if rec.get('note'):
                _draw_right(c, f"הערה: {rec.get('note')}", right, y, _FONT, 10, color=(0.4, 0.45, 0.5))
                y -= 0.45 * cm
        else:
            _draw_right(c, f"ספירת מלאי חדר #{rec.get('id')} (המשך)", right, y, _FONT_BOLD, 11)
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

    if not lines:
        _draw_right(c, 'אין שורות בספירה.', right, y, _FONT, 11, color=(0.4, 0.45, 0.5))
        y -= 0.8 * cm

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
            'print_name': line.get('print_name') or '',
            'fabric': line.get('fabric') or '',
            'color': line.get('color') or '',
            'print': line.get('print') or '',
            'size': line.get('size') or '',
            'area': line.get('area') or '',
            'box': line.get('box') or '',
            'qty': _qty(line.get('quantity')),
        }
        for i, (key, _) in enumerate(col_defs):
            _draw_right(c, values[key], col_rights[i] - 0.12 * cm, y - 0.22 * cm, _FONT, 8)
        y -= row_h

    y -= 0.35 * cm
    box_h = 1.15 * cm
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
        y - 0.15 * cm,
        _FONT_BOLD,
        13,
        color=(0.15, 0.35, 0.75),
    )
    draw_footer()
    c.save()
    return file_path


def generate_baby_basic_count_sheet_pdf(
    rows: List[Dict] | None = None,
    file_path: str = '',
    biz_name: str = '',
    location: str = 'חדר בייבי בייסיק',
    area: str = '',
    box_number: str = '',
    extra_blank_rows: int = 0,
    pages: int = 2,
) -> str:
    """גיליון ספירה שחור-לבן, שורות ריקות למילוי ידני. אזור למעלה, מידה ומספר קופסא בעמודות."""
    _register_fonts()
    os.makedirs(os.path.dirname(file_path) or '.', exist_ok=True)
    left, right = MARGIN, PAGE_W - MARGIN
    usable_w = right - left
    location = str(location or '').strip() or 'חדר בייבי בייסיק'
    _ = rows, area, box_number, extra_blank_rows

    col_defs = [
        ('model', 2.6 * cm),
        ('size', 2.0 * cm),
        ('fabric', 2.4 * cm),
        ('color', 2.4 * cm),
        ('print_name', 3.4 * cm),
        ('qty', 1.8 * cm),
        ('box', 2.2 * cm),
    ]
    widths = [w for _, w in col_defs]
    leftover = usable_w - sum(widths)
    if leftover != 0:
        widths[4] = max(2.4 * cm, widths[4] + leftover)
    col_rights = []
    x = right
    for w in widths:
        col_rights.append(x)
        x -= w
    col_lefts = [col_rights[i] - widths[i] for i in range(len(widths))]
    headers_he = {
        'model': 'דגם',
        'size': 'מידה',
        'fabric': 'סוג בד',
        'color': 'צבע רקע',
        'print_name': 'הדפס',
        'qty': 'כמות',
        'box': 'מספר קופסא',
    }
    row_h = 0.78 * cm
    header_h = 0.78 * cm
    bottom_limit = MARGIN + 1.2 * cm
    black = (0, 0, 0)

    c = canvas.Canvas(file_path, pagesize=A4)

    def stroke_black():
        c.setStrokeColorRGB(*black)
        c.setFillColorRGB(*black)
        c.setLineWidth(0.9)

    def draw_header(y: float, first_page: bool) -> float:
        if first_page:
            title = (biz_name or '').strip() or 'בייבי בייסיק'
            _draw_center(c, title, PAGE_W / 2, y, _FONT_BOLD, 14, color=black)
            y -= 0.55 * cm
            _draw_center(c, 'גיליון ספירת מלאי', PAGE_W / 2, y, _FONT_BOLD, 13, color=black)
            y -= 0.55 * cm
            stroke_black()
            c.line(left, y, right, y)
            y -= 0.5 * cm
            _draw_right(c, f"מיקום: {location}", right, y, _FONT_BOLD, 11, color=black)
            y -= 0.48 * cm
            _draw_right(c, "אזור: ________", right, y, _FONT, 11, color=black)
            y -= 0.48 * cm
            _draw_right(c, "תאריך: ________", right, y, _FONT, 10, color=black)
            _draw_right(c, "שם הסופר: ________", left + 7.4 * cm, y, _FONT, 10, color=black)
            y -= 0.5 * cm
        else:
            _draw_right(c, f"גיליון ספירת מלאי — {location} (המשך)", right, y, _FONT_BOLD, 10, color=black)
            y -= 0.42 * cm
            _draw_right(c, "אזור: ________", right, y, _FONT, 10, color=black)
            y -= 0.45 * cm
        return y

    def draw_table_header(y: float) -> float:
        top = y
        bottom = y - header_h
        c.setFillColorRGB(*black)
        c.setStrokeColorRGB(*black)
        c.setLineWidth(0.9)
        c.rect(left, bottom, usable_w, header_h, fill=1, stroke=1)
        text_y = bottom + 0.28 * cm
        for i, (key, _) in enumerate(col_defs):
            _draw_right(c, headers_he[key], col_rights[i] - 0.1 * cm, text_y, _FONT_BOLD, 8, color=(1, 1, 1))
            stroke_black()
            c.line(col_lefts[i], bottom, col_lefts[i], top)
        stroke_black()
        c.line(right, bottom, right, top)
        return bottom

    def draw_empty_row(y: float) -> float:
        top = y
        bottom = y - row_h
        stroke_black()
        c.setFillColorRGB(1, 1, 1)
        c.rect(left, bottom, usable_w, row_h, fill=1, stroke=1)
        for i in range(len(col_defs)):
            c.line(col_lefts[i], bottom, col_lefts[i], top)
        c.line(right, bottom, right, top)
        return bottom

    def draw_footer():
        c.setFillColorRGB(*black)
        c.setFont(_FONT, 8)
        c.drawCentredString(PAGE_W / 2, MARGIN * 0.45, _rtl(f"עמוד {c.getPageNumber()}"))

    def fill_page(y: float):
        while y >= bottom_limit + row_h:
            y = draw_empty_row(y)
        return y

    page_count = max(1, int(pages or 1))
    for page_i in range(page_count):
        if page_i:
            draw_footer()
            c.showPage()
        y = PAGE_H - MARGIN
        y = draw_header(y, page_i == 0)
        y = draw_table_header(y)
        fill_page(y)
    draw_footer()
    c.save()
    return file_path
