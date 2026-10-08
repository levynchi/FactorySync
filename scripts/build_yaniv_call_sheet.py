# -*- coding: utf-8 -*-
"""דף שיחה עם יניב (אקספרסליין) על דרישת תשלום 2630 — מע"מ טורקי.

Builds an RTL Hebrew DOCX under supplier_documents/6 and registers it in suppliers.json.
"""
import json
import os
from datetime import datetime

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "supplier_documents", "6")
OUT_NAME = "דף_שיחה_יניב_דרישה_2630_06.10.2026.docx"
OUT_PATH = os.path.join(OUT_DIR, OUT_NAME)

# ---------------- numbers ----------------
RATE = 3.10
S1_UNITS, S1_USD = 14750, 26071.24
S2_UNITS, S2_USD = 18278, 37383.44
GOODS = round(S1_USD + S2_USD, 2)            # 63,454.68
VAT_TR = round(GOODS * 0.10, 2)              # 6,345.47
GOODS_INCL = round(GOODS + VAT_TR, 2)        # 69,800.15
ADV_OMER, ADV_YANIV = 19250.00, 19310.00
BAL_EX = round(GOODS - ADV_OMER, 2)          # 44,204.68
BAL_INCL = round(GOODS_INCL - ADV_OMER, 2)   # 50,550.15
HALF_VAT = round(VAT_TR / 2, 2)              # 3,172.73
CARTONS, PER_CARTON = 106, 35
FREIGHT = CARTONS * PER_CARTON               # 3,710
YANIV_GOODS = 50500.00


def scenario(goods):
    sub = round(goods + FREIGHT, 2)
    vat = round(sub * 0.18, 2)
    tot = round(sub + vat, 2)
    ils = round(tot * RATE, 0)
    ils_local = round(FREIGHT * RATE + sub * RATE * 0.18, 0)  # what is paid in ₪ in Israel
    return goods, sub, vat, tot, ils, ils_local


SC_YANIV = scenario(YANIV_GOODS)
SC_A = scenario(BAL_EX)                      # my proposal
SC_B = scenario(round(BAL_EX + HALF_VAT, 2))  # if I end up paying 5%


def f(x, d=2):
    return f"{x:,.{d}f}"


# ---------------- docx helpers ----------------
def set_rtl(paragraph, align=WD_ALIGN_PARAGRAPH.RIGHT):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)
    paragraph.alignment = align
    return paragraph


def run_fmt(run, size=11, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "Arial"
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:cs"), "Arial")
    rtl = OxmlElement("w:rtl")
    rPr.append(rtl)
    if color:
        run.font.color.rgb = RGBColor(*color)
    return run


def para(doc, text="", size=11, bold=False, color=None, space_after=4, align=WD_ALIGN_PARAGRAPH.RIGHT, style=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    set_rtl(p, align)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if text:
        run_fmt(p.add_run(text), size, bold, color)
    return p


def heading(doc, text, size=13):
    p = para(doc, text, size=size, bold=True, color=(0x1F, 0x3A, 0x5F), space_after=3)
    p.paragraph_format.space_before = Pt(8)
    return p


def bullet(doc, text, bold_prefix=None, size=11):
    p = doc.add_paragraph(style="List Bullet")
    set_rtl(p)
    p.paragraph_format.space_after = Pt(2)
    if bold_prefix:
        run_fmt(p.add_run(bold_prefix + " "), size, True)
    run_fmt(p.add_run(text), size)
    return p


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def table(doc, headers, rows, widths=None, bold_last=False, highlight_rows=()):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = t._tbl.tblPr
    bidi = OxmlElement("w:bidiVisual")
    tblPr.append(bidi)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ""
        p = c.paragraphs[0]
        set_rtl(p, WD_ALIGN_PARAGRAPH.CENTER)
        run_fmt(p.add_run(h), 10, True, (0xFF, 0xFF, 0xFF))
        shade(c, "1F3A5F")
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        is_bold = (bold_last and ri == len(rows) - 1) or ri in highlight_rows
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            set_rtl(p, WD_ALIGN_PARAGRAPH.RIGHT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER)
            run_fmt(p.add_run(str(v)), 10, is_bold)
            if ri in highlight_rows:
                shade(cells[i], "FFF2CC")
            elif is_bold:
                shade(cells[i], "E8EEF5")
    if widths:
        t.autofit = False
        tblLayout = OxmlElement("w:tblLayout")
        tblLayout.set(qn("w:type"), "fixed")
        tblPr.append(tblLayout)
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
        grid = t._tbl.tblGrid
        for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
            if i < len(widths):
                gc.set(qn("w:w"), str(int(Cm(widths[i]).twips)))
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = OxmlElement("w:tblW")
            tblPr.append(tblW)
        tblW.set(qn("w:type"), "dxa")
        tblW.set(qn("w:w"), str(int(Cm(sum(widths)).twips)))
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


# ---------------- build ----------------
def build():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.6)
    sec.top_margin = sec.bottom_margin = Cm(1.3)
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(11)

    para(doc, "דף שיחה — יניב / אקספרסליין קרגו", size=16, bold=True, color=(0x1F, 0x3A, 0x5F), space_after=0)
    para(doc, "דרישת תשלום 2630 מ-04.10.2026 — מע״מ טורקי על סחורת עומר (Nibby Baby)", size=11, space_after=0)
    para(doc, "הוכן 06.10.2026 · יניב 0585001984 · משרד 0528194040 · order@expresslinetr.com", size=9, color=(0x66, 0x66, 0x66), space_after=6)

    # ---- 0. outcome
    heading(doc, "תוצאה — 06.10.2026, אחרי השיחה: נסגר")
    table(doc,
          ["", "סחורה $", "הובלה $", "לפני מע״מ $", "מע״מ 18% $", "סה״כ $", "סה״כ ₪", "משולם בארץ ₪"],
          [
              ["דרישה 2630 מקורית (04.10)", f(YANIV_GOODS, 0), f(FREIGHT, 0), f(SC_YANIV[1], 0), f(SC_YANIV[2], 0), f(SC_YANIV[3], 0), f(SC_YANIV[4], 0), f(SC_YANIV[5], 0)],
              ["גרסה 2 (06.10 בוקר)", "50,550", "700", "51,250", "9,225", "60,475", "187,473", "30,768"],
              ["גרסה 3 — סופית (06.10)", "50,550", "150", "50,700", "9,126", "59,826", "185,461", "28,756"],
          ],
          widths=[4.4, 2.1, 1.4, 2.1, 2.0, 2.1, 1.9, 1.8], highlight_rows=(2,))
    bullet(doc, "המע״מ הטורקי האמיתי הוא 4,595$ — זה מה שבחשבונית הרשמית של עומר ל-Express Line (225,316 TL, שער 49.03). על המקדמה אין חשבונית, אז הטענה על 1,750$ נוספים ירדה.", "מה התברר:")
    bullet(doc, "יניב לוקח על עצמו 2,250$ מהמע״מ, ומזכה גם את 2,409 ₪ הפרש ההובלה ממשלוח 1. הכול דרך שורת ההובלה: 150$ במקום 3,710$ (זיכוי 3,560$). הסחורה נשארה 50,550$ — עומר מקבל את מלוא היתרה.", "מה סוכם:")
    bullet(doc, "2,345$ (4,595 − 2,250). נשלחה דרישה לעומר לספוג 1,500$ (אני 845$). אם מאשר — היתרה לטורקיה 49,050$ ויניב מוציא דרישה 2630 מתוקנת.", "עדיין פתוח:")
    bullet(doc, "בוצע ב-07.10.2026 (10:00): הועברו 28,755 ₪ מחשבון דיסקונט 1093039 לחשבון פועלים 658/611610 אקספרסליין קרגו (אסמכתא 324577). חלק ההובלה והמע״מ הישראלי שולם ונסגר! חלק הסחורה (50,550$ או 49,050$) מעוכב עד לתשובת עומר.", "שולם בארץ:")
    bullet(doc, "2,250$ + 2,409 ₪ (≈777$) = 3,027$, אבל הזיכוי בפועל 3,560$ — 533$ לטובתך שלא מוסברים (אולי הוזיל הובלה ל~30$/קרטון). לא להעיר, רק לדעת.", "לבדוק בשקט:")
    bullet(doc, "עלות הסחורה למחירון: 48,300$ (50,550 − 2,250). מהמשלוח הבא — עומר לא מוציא חשבונית לפני תיאום עם יניב; לסכם מראש שהמחירים בלי מע״מ טורקי.", "לרישום:")
    para(doc, "— להלן דף ההכנה המקורי לשיחה —", size=9, color=(0x99, 0x99, 0x99), align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)

    # ---- 1. numbers
    heading(doc, "1. המספרים — החשבון שלי מול עומר (דולר, לפני מע״מ טורקי)")
    table(doc,
          ["פריט", "יחידות", "לפני מע״מ $", "עם 10% מע״מ $"],
          [
              ["משלוח 1 — 200/201/202 (חשבונית 01.09)", f(S1_UNITS, 0), f(S1_USD), f(round(S1_USD * 1.1, 2))],
              ["משלוח 2 — 025/204 (חשבונית SF01078, 30.09)", f(S2_UNITS, 0), f(S2_USD), f(round(S2_USD * 1.1, 2))],
              ["סה״כ סחורה", f(S1_UNITS + S2_UNITS, 0), f(GOODS), f(GOODS_INCL)],
              ["פחות מקדמה 03.09 (עומר רשם 19,250; דרך יניב שולם 19,310)", "", f"({f(ADV_OMER)})", f"({f(ADV_OMER)})"],
              ["יתרה לסחורה", "", f(BAL_EX), f(BAL_INCL)],
          ],
          widths=[8.5, 2.2, 3.3, 3.3], highlight_rows=(4,))
    para(doc, f"ההפרש כולו = המע״מ הטורקי: {f(VAT_TR)}$. חצי ממנו (5%) = {f(HALF_VAT)}$.", size=10, bold=True)

    # ---- 2. what Yaniv asked vs options
    heading(doc, "2. דרישה 2630 כפי שהגיעה מול שתי אפשרויות תיקון")
    rows = []
    for name, sc in (("כפי שיניב שלח (יתרת עומר כולל מע״מ)", SC_YANIV),
                     ("א. ההצעה שלי — יניב סופג 5%", SC_A),
                     ("ב. אם בסוף אני משלם 5% לעומר", SC_B)):
        g, sub, vat, tot, ils, ils_local = sc
        rows.append([name, f(g, 0), f(FREIGHT, 0), f(sub, 0), f(vat, 0), f(tot, 0), f(ils, 0), f(ils_local, 0)])
    table(doc,
          ["תרחיש", "סחורה $", "הובלה $", "לפני מע״מ $", "מע״מ 18% $", "סה״כ $", "סה״כ ₪ (3.10)", "משולם בארץ ₪"],
          rows, widths=[4.4, 2.1, 1.4, 2.1, 2.0, 2.1, 1.9, 1.8], highlight_rows=(1,))
    para(doc, f"פער בין הדרישה לאפשרות א: {f(SC_YANIV[3] - SC_A[3])}$ ≈ {f((SC_YANIV[3] - SC_A[3]) * RATE, 0)} ₪. "
              f"בין הדרישה לאפשרות ב: {f(SC_YANIV[3] - SC_B[3])}$ ≈ {f((SC_YANIV[3] - SC_B[3]) * RATE, 0)} ₪.", size=10)
    para(doc, "״משולם בארץ ₪״ = הובלה + מע״מ ישראלי (מה שחתום על הדרישה כ-41,750 ₪). הסחורה בדולר ב-SWIFT ל-Vakifbank.", size=9, color=(0x66, 0x66, 0x66))
    para(doc, "ההובלה (106 קרטונים × 35$ = 3,710$) — לא פותחים. הוא כבר ירד מ-42$. להגיד לו שזה בסדר.", size=10, bold=True)

    # ---- 3. talking points
    heading(doc, "3. מה אני אומר — לפי הסדר")
    bullet(doc, "כל המחירים של עומר (2.02$ / 2.14$ / 1.43$ / 2.35$ / 1.70$) הם לפי המחירון שלו, בלי מע״מ. על זה בניתי את התמחור ללקוחות שלי — בייבי בייסיק כבר קיבלו מחירון ותעודות משלוח לפי העלות הזו. 10% על הסחורה = 6,345$ ≈ 0.20$ ליחידה שאין לי ממי לקחת.", "1. התמחור שלי:")
    bullet(doc, "ככה עבדתי תמיד עם הטורקים, וככה אני עובד איתך על הבדים של אקדם — המע״מ הטורקי מתקזז בין החברות בטורקיה, לא אצלי. אני משלם מע״מ בישראל, לא גם בטורקיה.", "2. ככה תמיד עבדנו:")
    bullet(doc, "דיברנו ב-04.10 ואמרת שאתה מבין שזה כמו עם אקדם. באותו יום כתבתי את זה גם לעומר. דרישה 2630 פשוט העתיקה את היתרה מדף החשבון של עומר (50,550$) שכוללת את המע״מ.", "3. סיכמנו:")
    bullet(doc, "אני מבין שההחזר עולה לך — חצי שנה, לירה, רואה חשבון. לכן אני לא מבקש שתספוג הכול.", "4. אני מבין אותו:")
    bullet(doc, "אתה סופג את ה-5% שאתה כן מקבל חזרה מהמדינה. את ה-5% השני אני לוקח על עצמי מול עומר — אדבר איתו ואראה מה אני עושה. כלומר הדרישה המתוקנת: סחורה 44,204.68$ + הובלה 3,710$ + מע״מ 18% = 56,539$. אם אני סוגר עם עומר שאני משלם לו את ה-5% — נעביר את זה בנפרד.", "5. ההצעה:")
    bullet(doc, "מהמשלוח הבא (אינטרלוק טובולרי, הדפסים) עומר מוכר לך ״איהראץ׳ קאיטלי״ (ihraç kayıtlı teslim, סעיף 11/1-c) — יצרן ליצואן בלי לגבות מע״מ, יצוא תוך 3 חודשים. ככה אף אחד לא מממן כלום ואין הפסד שער. ככה זה עובד עם אקדם — למה לא עם עומר?", "6. קדימה:")

    # ---- 4. objections
    heading(doc, "4. תשובות לטענות שלו")
    table(doc,
          ["הוא אומר", "אני עונה"],
          [
              ["״אני מקבל רק 5% חזרה״", "בדיוק בגלל זה אני מציע שתספוג רק את ה-5% האלה ולא את כל ה-10%. את השאר אני מסדר מול עומר."],
              ["״יש הפרשי שער, ההחזר אחרי חצי שנה״", "מבין. זו בעיה של המבנה — אם עומר מוכר לך ihraç kayıtlı אין החזר בכלל ואין שער. בוא נסדר את זה למשלוח הבא."],
              ["״הסברתי לעומר, הוא לא מסכים״", "עומר אמר לי שהמחירים בלי מע״מ. אני מדבר איתו על ה-5% שלי. בינינו — אתה מתחשבן איתי כמו על הבדים."],
              ["״ככה זה בטורקיה, חייבים מע״מ״", "לא כשמייצאים. יצוא 0%. או שעומר מוציא לך חשבונית בלי גבייה (יצרן ליצואן), או שאתה מקבל החזר. בשני המקרים זה לא עלות שלי."],
              ["״אני לא יודע מה לעשות״", "אני אומר לך מה לעשות: דרישה מתוקנת על 44,204.68$ סחורה, ואני מעביר באותו יום."],
          ],
          widths=[5.5, 11.8])

    # ---- 5. questions
    heading(doc, "5. שאלות לשאול אותו (לרשום תשובות)")
    bullet(doc, "החשבונית של עומר יצאה על Express Line או על Leo Israel? (SF01078 רשומה על ״LEO İSRAİL״ עם 10% — ללקוח זר ביצוא זה אמור להיות 0%.)")
    bullet(doc, "ההחזר אצלך במזומן או בקיזוז (mahsup) מול מסים אחרים? אם קיזוז — אין חצי שנה ואין הפסד שער.")
    bullet(doc, "כבר שילמת לעומר את המע״מ בלירות? באיזה שער?")
    bullet(doc, "למה עם אקדם זה עובד בלי מע״מ ועם עומר לא? לעומר אין תעודת יצרן (sanayi sicil)?")
    bullet(doc, "המקדמה: אני העברתי 19,310$ (דרישה 2432), עומר רשם 19,250$. איפה ה-60$?")
    bullet(doc, "שער 3.10 בדרישה — השער ביום ההעברה בפועל?")

    # ---- 6. closing
    heading(doc, "6. סגירה וקווים אדומים")
    bullet(doc, f"יעד: דרישה מתוקנת — סחורה {f(BAL_EX)}$, הובלה {f(FREIGHT, 0)}$, סה״כ {f(SC_A[3])}$ ≈ {f(SC_A[4], 0)} ₪. אני מעביר מיד.", "סגירה טובה:")
    bullet(doc, f"אם הוא לא זז — משלמים עכשיו את מה שלא במחלוקת ({f(BAL_EX)}$ סחורה + הובלה + מע״מ ישראלי), ואת ה-{f(VAT_TR)}$ שבמחלוקת מחזיקים עד שמסדרים עם עומר. לא לעכב סחורה בגלל 6,000$.", "אם נתקעים:")
    bullet(doc, "לא לשלם 10% מע״מ טורקי מלא. מקסימום 5% ורק אם זה סגור מול עומר, לא דרך יניב.", "קו אדום:")
    bullet(doc, "לא לפתוח את ההובלה. לא להתווכח על ה-60$ של המקדמה מעבר לשאלה.", "לא להיגרר ל:")
    bullet(doc, "לסכם בוואטסאפ מיד אחרי השיחה: ״כפי שדיברנו — דרישה מתוקנת על X$, המשלוח הבא ihraç kayıtlı״.", "אחרי השיחה:")

    para(doc, "", space_after=2)
    para(doc, "תשלום ₪: בנק הפועלים סניף 658 ח-ן 611610 · תשלום $: Express Line Tekstil, Vakifbank IBAN TR730001500158048025399744", size=8.5, color=(0x66, 0x66, 0x66))

    os.makedirs(OUT_DIR, exist_ok=True)
    doc.save(OUT_PATH)
    return OUT_PATH


def register():
    sp = os.path.join(ROOT, "suppliers.json")
    data = json.load(open(sp, encoding="utf-8"))
    sup = next(s for s in data if s["id"] == 6)
    docs = sup.setdefault("documents", [])
    if any(d["filename"] == OUT_NAME for d in docs):
        return False
    docs.append({
        "id": max([d["id"] for d in docs] + [0]) + 1,
        "filename": OUT_NAME,
        "original_name": "דף שיחה עם יניב — דרישה 2630 / מע\"מ טורקי (06.10.2026)",
        "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    json.dump(data, open(sp, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return True


if __name__ == "__main__":
    p = build()
    added = register()
    print("built:", p)
    print("registered:", added)
    print("A:", SC_A)
    print("B:", SC_B)
    print("Yaniv:", SC_YANIV)
