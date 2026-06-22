"""מחולל דף מדבקות A4 (3x5) בעיצוב מותג.

פריסת כל מדבקה (בתוך מסגרת מקווקוות):
  - כותרת: לוגו החנות בפינה הימנית-עליונה.
  - גוף: תמונת מוצר משמאל, ומימין שם המוצר, סוג בד, קו מפריד, ומידה.
  - תחתית: תיבה שחורה "N יחידות" משמאל, וברקוד EAN-13 סרוק + מספר מימין.

הקבועים הגאומטריים מרוכזים למעלה לכיול קל אחרי הדפסת ניסיון.
"""
import os

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.barcode import createBarcodeDrawing

try:
    from bidi.algorithm import get_display
except Exception:  # pragma: no cover
    def get_display(s):
        return s

# ===== גאומטריה (תואמת לגיליון המדבקות המקורי) =====
PAGE_W, PAGE_H = A4              # 21.0 x 29.7 ס"מ
COLS = 3
ROWS = 5
# מידות שחולצו מקובץ הייחוס: מדבקה 5.91x5.05 ס"מ עם רווחים בין המדבקות
LABEL_W = 5.91 * cm
LABEL_H = 5.05 * cm
H_GAP = 0.92 * cm               # רווח אופקי בין מדבקות
V_GAP = 0.59 * cm               # רווח אנכי בין מדבקות
PER_PAGE = COLS * ROWS

# כיול מדפסת: הזזה אנכית כלפי מטה לכל שורה (במ"מ). ערך חיובי = מטה.
# שורה 1=2 מ"מ, שורה 2=1.5 מ"מ, שורה 3=1 מ"מ, שורות 4-5 ללא הזזה.
ROW_SHIFT_DOWN_MM = [2.0, 1.5, 1.0, 0.0, 0.0]

# ===== פריסת תוכן המדבקה (ניתן לכיול) =====
BORDER_INSET = 0.05 * cm        # מרחק המסגרת המקווקוות מקצה התא (קטן - שהמסגרת תתאים לגודל הייחוס)
CONTENT_PAD = 0.20 * cm         # ריפוד פנימי בין המסגרת לתוכן
HEADER_H = 1.02 * cm            # אזור הלוגו (עליון)
FOOTER_H = 1.28 * cm            # אזור הברקוד + תיבת היחידות (תחתון)
IMG_W_RATIO = 0.40              # רוחב תמונת המוצר ביחס לרוחב התוכן
IMG_TEXT_GAP = 0.18 * cm        # רווח בין תמונת המוצר לטקסט
LOGO_BOX_W = 3.6 * cm           # רוחב תיבת הלוגו
UNITS_BOX_W = 2.05 * cm         # רוחב תיבת היחידות השחורה
UNITS_BOX_H = 0.60 * cm         # גובה תיבת היחידות
BC_HEIGHT = 0.88 * cm           # גובה הברקוד הסרוק

LOGO_PATH = os.path.join('assets', 'labels', 'logo.png')

# פונט עברי: Heebo-Medium (תואם לתבנית הבוטיק). נפילה ל-Arial אם חסר.
_FONT_NAME = 'LabelHe'
_FONT_BOLD = 'LabelHeBold'
_FONT_LIGHT = 'LabelHeLight'    # Heebo-Regular - לשורת הבד (משקל קל יותר)
_FONT_NUM = 'Helvetica'         # ספרות הברקוד (פונט מובנה ב-reportlab)
_FONT_REGISTERED = False

HEEBO_MEDIUM_PATH = os.path.join('assets', 'fonts', 'Heebo-Medium.ttf')
HEEBO_REGULAR_PATH = os.path.join('assets', 'fonts', 'Heebo-Regular.ttf')


def _register_fonts():
    global _FONT_REGISTERED
    if _FONT_REGISTERED:
        return
    # עדיפות ראשונה: Heebo (Medium לשם הדגם, Regular לשורת הבד)
    heebo_m = _abs_path(HEEBO_MEDIUM_PATH)
    heebo_r = _abs_path(HEEBO_REGULAR_PATH)
    if heebo_m and os.path.exists(heebo_m):
        try:
            pdfmetrics.registerFont(TTFont(_FONT_NAME, heebo_m))
            pdfmetrics.registerFont(TTFont(_FONT_BOLD, heebo_m))
            light = heebo_r if (heebo_r and os.path.exists(heebo_r)) else heebo_m
            pdfmetrics.registerFont(TTFont(_FONT_LIGHT, light))
            _FONT_REGISTERED = True
            return
        except Exception:
            pass
    # נפילה ל-Arial/Tahoma מ-Windows אם Heebo חסר
    candidates_reg = [r'C:\Windows\Fonts\arial.ttf', r'C:\Windows\Fonts\tahoma.ttf']
    candidates_bold = [r'C:\Windows\Fonts\arialbd.ttf', r'C:\Windows\Fonts\tahomabd.ttf']
    reg = next((p for p in candidates_reg if os.path.exists(p)), None)
    bold = next((p for p in candidates_bold if os.path.exists(p)), reg)
    if reg:
        pdfmetrics.registerFont(TTFont(_FONT_NAME, reg))
        pdfmetrics.registerFont(TTFont(_FONT_BOLD, bold or reg))
        pdfmetrics.registerFont(TTFont(_FONT_LIGHT, reg))
    _FONT_REGISTERED = True


def _has_hebrew(text: str) -> bool:
    return any('\u0590' <= ch <= '\u05FF' for ch in (text or ''))


def _format_fabric(fabric: str) -> str:
    """מרחיב מילת בד קצרה (פלנל/טריקו וכו') ל'<בד> 100% כותנה איכותית'.

    אם כבר הוזן טקסט מלא (מכיל % / 'כותנה' / 'איכותית') - מוחזר כפי שהוא.
    """
    f = str(fabric or '').strip()
    if not f:
        return ''
    if '%' in f or 'כותנה' in f or 'איכותית' in f:
        return f
    return f + ' 100% כותנה איכותית'


def _shape(text: str) -> str:
    """סידור RTL לעברית; טקסט לטיני/מספרי נשאר כפי שהוא."""
    text = str(text or '')
    if _has_hebrew(text):
        try:
            return get_display(text)
        except Exception:
            return text
    return text


def _make_barcode(value: str):
    """יוצר ציור ברקוד: EAN-13 אם תקין (12-13 ספרות), אחרת Code128."""
    value = str(value or '').strip()
    digits = ''.join(ch for ch in value if ch.isdigit())
    try:
        if len(digits) in (12, 13):
            return createBarcodeDrawing('EAN13', value=digits, humanReadable=False)
    except Exception:
        pass
    try:
        if value:
            return createBarcodeDrawing('Code128', value=value, humanReadable=False)
    except Exception:
        pass
    return None


_LOGO_CACHE = {'path': None, 'reader': None}

# שורש הפרויקט (assets/ נמצא כאן) - גיבוי אם ספריית העבודה שונה
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


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
    if os.path.exists(p_root):
        return p_root
    return p_cwd


def _get_logo():
    """טוען את לוגו החנות (מ-assets/labels/logo.png). מנסה מחדש עד שיצליח."""
    path = _abs_path(LOGO_PATH)
    if not path or not os.path.exists(path):
        return None
    if _LOGO_CACHE['reader'] is not None and _LOGO_CACHE['path'] == path:
        return _LOGO_CACHE['reader']
    try:
        reader = ImageReader(path)
        _LOGO_CACHE['reader'] = reader
        _LOGO_CACHE['path'] = path
        return reader
    except Exception:
        return None


def _image_reader(rel_or_abs: str):
    path = _abs_path(rel_or_abs)
    if path and os.path.exists(path):
        try:
            return ImageReader(path)
        except Exception:
            return None
    return None


def _draw_image_fit(c: canvas.Canvas, reader, x, y, w, h, anchor='c'):
    """מצייר תמונה בתוך תיבה (x,y,w,h) תוך שמירת יחס וכיוון העגינה."""
    if reader is None:
        return
    try:
        c.drawImage(reader, x, y, width=w, height=h, mask='auto',
                    preserveAspectRatio=True, anchor=anchor)
    except Exception:
        pass


def _fit_font_size(c: canvas.Canvas, text: str, font: str, max_size: float,
                   min_size: float, max_width: float) -> float:
    """מקטין את גודל הפונט עד שהטקסט נכנס לרוחב הנתון."""
    size = max_size
    while size > min_size:
        if c.stringWidth(text, font, size) <= max_width:
            break
        size -= 0.5
    return size


def _draw_label(c: canvas.Canvas, x_left: float, y_bottom: float, item: dict):
    """מצייר מדבקה אחת בעיצוב מותג בתוך תא הממוקם ב-(x_left, y_bottom)."""
    print_name = str(item.get('print_name', '')).strip()
    size = str(item.get('size', '')).strip()
    size_unit = str(item.get('size_unit', '')).strip()
    if size and size_unit:
        # "0-3 חודשים" - ב-RTL היחידה מופיעה משמאל למספר המידה
        size = f"{size} {size_unit}"
    fabric = _format_fabric(item.get('fabric', ''))
    try:
        pack_qty = int(float(str(item.get('pack_qty', 1)) or 1))
    except Exception:
        pack_qty = 1
    barcode = str(item.get('barcode', '')).strip()
    image_rel = str(item.get('image', '')).strip()

    # ----- גבולות התא והתוכן -----
    cell_x = x_left
    cell_y = y_bottom
    # מסגרת מקווקוות מעוגלת
    c.saveState()
    c.setLineWidth(0.7)
    c.setDash(3, 2)
    c.roundRect(cell_x + BORDER_INSET, cell_y + BORDER_INSET,
                LABEL_W - 2 * BORDER_INSET, LABEL_H - 2 * BORDER_INSET,
                radius=0.28 * cm, stroke=1, fill=0)
    c.restoreState()

    x0 = cell_x + BORDER_INSET + CONTENT_PAD            # שמאל התוכן
    x1 = cell_x + LABEL_W - BORDER_INSET - CONTENT_PAD  # ימין התוכן
    yt = cell_y + LABEL_H - BORDER_INSET - CONTENT_PAD  # ראש התוכן
    yb = cell_y + BORDER_INSET + CONTENT_PAD            # תחתית התוכן
    content_w = x1 - x0

    header_top = yt
    header_bottom = yt - HEADER_H
    footer_top = yb + FOOTER_H
    footer_bottom = yb
    body_top = header_bottom
    body_bottom = footer_top

    # ----- כותרת: לוגו בפינה הימנית-עליונה -----
    logo = _get_logo()
    if logo is not None:
        _draw_image_fit(c, logo, x1 - LOGO_BOX_W, header_bottom,
                        LOGO_BOX_W, HEADER_H, anchor='ne')

    # ----- גוף: תמונת מוצר משמאל, טקסט מימין -----
    img_w = content_w * IMG_W_RATIO
    img_reader = _image_reader(image_rel)
    if img_reader is not None:
        _draw_image_fit(c, img_reader, x0, body_bottom, img_w, body_top - body_bottom,
                        anchor='c')

    text_left = x0 + img_w + IMG_TEXT_GAP
    text_w = x1 - text_left

    # בלוק הטקסט ממורכז אנכית באזור הגוף
    name_size = _fit_font_size(c, _shape(print_name), _FONT_BOLD, 15, 9, text_w) if print_name else 0
    fabric_size = 8.5 if fabric else 0
    size_size = 10.5 if size else 0
    SEP_GAP = 0.14 * cm
    line_gaps = 0.12 * cm

    def _hline(y):
        c.saveState()
        c.setLineWidth(0.5)
        c.line(text_left, y, x1, y)
        c.restoreState()

    block_h = 0.0
    if print_name:
        block_h += SEP_GAP * 2 + name_size + line_gaps   # קו עליון + שם הדגם
    if fabric:
        block_h += fabric_size + line_gaps
    if size:
        block_h += SEP_GAP * 2 + size_size               # קו מפריד + מידה
    by = body_bottom + (body_top - body_bottom + block_h) / 2.0

    if print_name:
        # קו מעל שם הדגם
        by -= SEP_GAP
        _hline(by)
        by -= SEP_GAP + name_size
        # שם הדגם נמתח לכל רוחב הקווים (justified) - רק שם הדגם
        shaped_name = _shape(print_name)
        name_target_w = x1 - text_left
        name_nat_w = c.stringWidth(shaped_name, _FONT_BOLD, name_size)
        n_chars = len(shaped_name)
        name_cs = ((name_target_w - name_nat_w) / (n_chars - 1)
                   if n_chars > 1 and name_target_w > name_nat_w else 0)
        tn = c.beginText(text_left, by)
        tn.setFont(_FONT_BOLD, name_size)
        tn.setCharSpace(name_cs)
        tn.textOut(shaped_name)
        c.drawText(tn)
        by -= line_gaps
    if fabric:
        by -= fabric_size
        shaped_fabric = _shape(fabric)
        char_space = -0.6           # ריווח קטן יותר בין אות לאות בשורת הבד
        fab_w = (c.stringWidth(shaped_fabric, _FONT_LIGHT, fabric_size)
                 + char_space * max(0, len(shaped_fabric) - 1))
        t = c.beginText(x1 - fab_w, by)
        t.setFont(_FONT_LIGHT, fabric_size)
        t.setCharSpace(char_space)
        t.textOut(shaped_fabric)
        c.drawText(t)
        by -= line_gaps
    if size:
        # קו מפריד מעל המידה
        by -= SEP_GAP
        _hline(by)
        by -= SEP_GAP + size_size
        c.setFont(_FONT_BOLD, size_size)
        # מידה ממורכזת בעמודת הטקסט
        c.drawCentredString((text_left + x1) / 2.0, by, _shape(size))

    # ----- תחתית: תיבת יחידות שחורה משמאל, ברקוד מימין -----
    if pack_qty and pack_qty > 1:
        box_y = footer_bottom + (FOOTER_H - UNITS_BOX_H) / 2.0
        c.saveState()
        c.setFillColorRGB(0, 0, 0)
        c.rect(x0, box_y, UNITS_BOX_W, UNITS_BOX_H, stroke=0, fill=1)
        c.setFillColorRGB(1, 1, 1)
        c.setFont(_FONT_BOLD, 10)
        c.drawCentredString(x0 + UNITS_BOX_W / 2.0,
                            box_y + (UNITS_BOX_H - 10) / 2.0 + 1.5,
                            _shape(f"{pack_qty} יחידות"))
        c.restoreState()

    if barcode:
        bc = _make_barcode(barcode)
        target_w = min(3.6 * cm, content_w * 0.55)
        bc_xr = x1                                  # קצה ימני של הברקוד
        bc_xl = bc_xr - target_w
        num_size = 7.5
        num_y = footer_bottom + 0.06 * cm
        bc_bottom = num_y + num_size + 0.06 * cm
        c.setFont(_FONT_NUM, num_size)
        c.drawCentredString((bc_xl + bc_xr) / 2.0, num_y, barcode)
        if bc is not None and bc.width and bc.height:
            sx = target_w / float(bc.width)
            sy = BC_HEIGHT / float(bc.height)
            bc.scale(sx, sy)
            bc.drawOn(c, bc_xl, bc_bottom)


def build_label_sheet_pdf(items, file_path: str) -> int:
    """בונה דף/דפי A4 של מדבקות.

    items: רשימת dict {print_name,size,fabric,pack_qty,barcode} (כבר מורחבת לפי כמות).
    מחזיר את מספר המדבקות שנכתבו.
    """
    if not items:
        raise ValueError("אין מדבקות לייצוא")
    _register_fonts()
    os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)

    # מרכוז הגריד (כולל הרווחים) על הדף - משחזר את מיקומי המסגרות בקובץ הייחוס
    left_margin = (PAGE_W - (COLS * LABEL_W + (COLS - 1) * H_GAP)) / 2.0
    top_margin = (PAGE_H - (ROWS * LABEL_H + (ROWS - 1) * V_GAP)) / 2.0
    c = canvas.Canvas(file_path, pagesize=A4)

    for idx, item in enumerate(items):
        pos = idx % PER_PAGE
        if idx > 0 and pos == 0:
            c.showPage()
        row = pos // COLS
        col = pos % COLS
        x_left = left_margin + col * (LABEL_W + H_GAP)
        y_bottom = PAGE_H - top_margin - row * (LABEL_H + V_GAP) - LABEL_H
        # כיול מדפסת: הזזת השורה כלפי מטה (הורדת y)
        shift_mm = ROW_SHIFT_DOWN_MM[row] if row < len(ROW_SHIFT_DOWN_MM) else 0.0
        y_bottom -= shift_mm * mm
        _draw_label(c, x_left, y_bottom, item)

    c.showPage()
    c.save()
    return len(items)
