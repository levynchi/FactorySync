# -*- coding: utf-8 -*-
"""דף הסבר: מע"מ טורקי על עסקת עומר (שני המשלוחים) מול החשבונית הרשמית ל-Express Line.

Reuses the RTL docx helpers from build_yaniv_call_sheet.py; saves under supplier_documents/6
and registers the document in suppliers.json (supplier 6 — Expressline / Yaniv).
"""
import json
import os
import sys
from datetime import datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_yaniv_call_sheet import para, heading, bullet, table, f, ROOT  # noqa: E402

OUT_DIR = os.path.join(ROOT, "supplier_documents", "6")
OUT_NAME = "הסבר_מעמ_טורקי_עסקת_עומר_06.10.2026.docx"
OUT_PATH = os.path.join(OUT_DIR, OUT_NAME)

# ---- numbers
S1, S2 = 26071.24, 37383.44
GOODS = S1 + S2                      # 63,454.68
VAT = GOODS * 0.10                   # 6,345.47
TOTAL = GOODS + VAT                  # 69,800.15
ADV = 19250.00
BAL = TOTAL - ADV                    # 50,550.15
INV_TL, INV_VAT_TL, INV_NET_TL = 2478476.0, 225316.0, 2253160.0
RATE_TL = INV_TL / BAL               # 49.03
INV_VAT_USD = INV_VAT_TL / RATE_TL   # 4,595
INV_NET_USD = INV_NET_TL / RATE_TL   # 45,955
ADV_NET, ADV_VAT = ADV / 1.1, ADV - ADV / 1.1   # 17,500 / 1,750


def build():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin = sec.bottom_margin = Cm(1.5)
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(11)

    para(doc, "מע״מ טורקי על עסקת עומר — מה חויב, מה דווח, ומה ההפרש", size=15, bold=True, color=(0x1F, 0x3A, 0x5F), space_after=0)
    para(doc, "Nibby Baby (עומר) · שני משלוחים 09/2026 · חשבונית רשמית ל-Express Line מס׳ NBY2026000000004 מ-01.10.2026", size=9.5, color=(0x66, 0x66, 0x66), space_after=8)

    heading(doc, "1. כל העסקה (דולר)")
    table(doc, ["", "סחורה לפני מע״מ", "מע״מ 10%", "סה״כ"],
          [["משלוח 1 — דגמים 200/201/202 (14,750 יח׳)", f(S1, 0), f(S1 * 0.1, 0), f(S1 * 1.1, 0)],
           ["משלוח 2 — דגמים 025/204 (18,278 יח׳)", f(S2, 0), f(S2 * 0.1, 0), f(S2 * 1.1, 0)],
           ["כל העסקה (33,028 יח׳)", f(GOODS, 0), f(VAT, 0), f(TOTAL, 0)]],
          widths=[7.4, 3.3, 3.3, 3.3], bold_last=True)
    para(doc, f"מתוך {f(TOTAL, 0)}$ כבר שולמה מקדמה של {f(ADV, 0)}$ (03.09.2026, דרך Express Line). נשאר לתשלום: {f(BAL, 0)}$.", size=10.5, bold=True)

    heading(doc, "2. מה כתוב בחשבונית הרשמית ליניב")
    table(doc, ["", "לירות (TL)", "דולר (שער 49.03)"],
          [["סחורה — ״28,000 BEBE BADİ × 80.47 TL״", f(INV_NET_TL, 0), f(INV_NET_USD, 0)],
           ["מע״מ 10% (Hesaplanan KDV)", f(INV_VAT_TL, 0), f(INV_VAT_USD, 0)],
           ["סה״כ לתשלום (Ödenecek Tutar)", f(INV_TL, 0), f(BAL, 0)]],
          widths=[7.4, 4.95, 4.95], bold_last=True)
    bullet(doc, f"הסה״כ בחשבונית = בדיוק היתרה {f(BAL, 0)}$ (2,478,476 ÷ 50,550 = 49.03 TL/$). עומר הוציא חשבונית על מה שנשאר לו לקבל — לא על הסחורה.")
    bullet(doc, "״28,000 בגדי גוף ב-80.47 TL״ הוא מספר מומצא כדי להגיע לסכום. בפועל נשלחו 33,028 יחידות ב-5 דגמים, בשני משלוחים.")
    bullet(doc, "החשבונית על שם Express Line (נכון), מסוג SATIS רגיל עם 10% מע״מ — לא ihraç kayıtlı. תאריך 01.10 — אחרי ששני המשלוחים כבר יצאו.")

    heading(doc, "3. למה המע״מ בחשבונית קטן ממה שעומר גובה")
    table(doc, ["", "סחורה", "מע״מ 10%", "סה״כ"],
          [["חשבונית רשמית ליניב (01.10)", f(INV_NET_USD, 0), f(INV_VAT_USD, 0), f(BAL, 0)],
           ["המקדמה — ללא חשבונית (״Nakit״ בדף החשבון)", f(ADV_NET, 0), f(ADV_VAT, 0), f(ADV, 0)],
           ["ביחד", f(GOODS, 0), f(VAT, 0), f(TOTAL, 0)]],
          widths=[7.4, 3.3, 3.3, 3.3], bold_last=True, highlight_rows=(1,))
    para(doc, "עומר רשם את שני המשלוחים בדף החשבון כולל מע״מ וניכה מהם את המקדמה. כך המקדמה ״בלעה״ 1,750$ מע״מ — אבל על המקדמה הוא לא הוציא חשבונית רשמית לאף אחד.", size=10.5)

    heading(doc, "4. שורה תחתונה")
    table(doc, ["", "דולר"],
          [["מע״מ שעומר מחייב אותי בדף החשבון (10% על כל הסחורה)", f(VAT, 0)],
           ["מע״מ שדווח רשמית — בחשבונית ליניב", f(INV_VAT_USD, 0)],
           ["מע״מ שנגבה ממני ולא דווח (על המקדמה)", f(ADV_VAT, 0)]],
          widths=[12.0, 5.3], highlight_rows=(2,))
    para(doc, "איפה זה עומד (06.10.2026)", size=11, bold=True, space_after=2)
    table(doc, ["", "דולר"],
          [["מע״מ טורקי רשמי (בחשבונית ל-Express Line)", f(INV_VAT_USD, 0)],
           ["יניב לקח על עצמו (בדרישה 2630 הסופית, הובלה 150$ במקום 3,710$)", "2,250"],
           ["נשאר פתוח", f(INV_VAT_USD - 2250, 0)],
           ["נדרש מעומר (הודעה 06.10)", "1,500"],
           ["עלי", f(INV_VAT_USD - 2250 - 1500, 0)],
           ["אם עומר מאשר — יתרה לטורקיה 50,550 − 1,500", "49,050"]],
          widths=[12.0, 5.3], highlight_rows=(2,), bold_last=True)
    para(doc, "סטטוס תשלומים (07.10.2026): שולמו 28,755 ₪ בארץ לאקספרסליין עבור ההובלה והמע״מ (אסמכתא 324577). 50,550$ לטורקיה מעוכבים עד לתשובת עומר (אם יאשר לספוג 1,500$ — יועברו 49,050$).", size=10.5, bold=True)
    bullet(doc, f"ה-{f(ADV_VAT, 0)}$ ״מע״מ על המקדמה״ — ירד מהשולחן. על המקדמה אין חשבונית רשמית, והמע״מ שבאמת שולם למדינה הוא רק {f(INV_VAT_USD, 0)}$ שבחשבונית ליניב.", "מה ירד:")
    bullet(doc, "חשבונית על 28,000 יח׳ מול יצוא של 33,028 יח׳ — אי-התאמה שעלולה לסבך ליניב את החזר המע״מ. עניין שלו מול עומר.", "סיכון של יניב:")
    bullet(doc, "למשלוח הבא — עומר לא מוציא חשבונית לפני שתיאם עם יניב איך לבנות אותה בלי מע״מ לישראל.", "להבא:")

    para(doc, "", space_after=2)
    para(doc, "מקורות: balance 01.10.26.pdf (דף חשבון עומר) · 025-204 invoice.pdf (SF01078) · Pdf_Faturalar-20261001160452.pdf (e-Fatura ל-Express Line) · דרישת תשלום 2630 מקורית ומתוקנת",
         size=8.5, color=(0x66, 0x66, 0x66))

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
        "original_name": "הסבר מע\"מ טורקי על עסקת עומר — דף חשבון מול חשבונית רשמית ל-Express Line (06.10.2026)",
        "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    json.dump(data, open(sp, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return True


if __name__ == "__main__":
    p = build()
    print("built:", p)
    print("registered:", register())
