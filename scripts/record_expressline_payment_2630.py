import json
import os

ROOT = r"c:\optitex excell"

# 1. Update suppliers.json
suppliers_path = os.path.join(ROOT, "suppliers.json")
with open(suppliers_path, "r", encoding="utf-8") as f:
    suppliers = json.load(f)

for s in suppliers:
    if s.get("id") == 6:
        docs = s.setdefault("documents", [])
        fn = "אישור_העברה_28755_אקספרסליין_2630_07.10.2026.jpg"
        if not any(d["filename"] == fn for d in docs):
            new_id = max([d["id"] for d in docs] + [0]) + 1
            docs.append({
                "id": new_id,
                "filename": fn,
                "original_name": "אישור העברה בנקאית 28,755 ₪ לאקספרסליין קרגו — דרישת תשלום 2630 (דיסקונט, 07.10.2026, אסמכתא 324577)",
                "uploaded_at": "2026-10-07 10:00:15"
            })
        
        payment_note = (
            " 07.10.2026 10:00: שולמו 28,755 ₪ בהעברה בנקאית מדיסקונט (1093039/93033) לפועלים 658/611610 אקספרסליין קרגו "
            '(אסמכתא 324577, מטרת העברה: "דרישת תשלום 2630 מיום 41026 ביגוד"). '
            'תשלום זה סוגר במלואו את חלק ההובלה והמע"מ הישראלי של דרישה 2630 הסופית (הובלה 150$ + מע"מ ישראלי 9,126$ = 9,276$ × 3.10 = 28,755.60 ₪). '
            'חלק הסחורה הדולרי (50,550$ או 49,050$ בכפוף לתשובת עומר לספיגת 1,500$ מע"מ טורקי — אירוע 131) עדיין מעוכב עד לתשובת עומר.'
        )
        if "324577" not in s.get("notes", ""):
            s["notes"] = s.get("notes", "") + payment_note

with open(suppliers_path, "w", encoding="utf-8") as f:
    json.dump(suppliers, f, ensure_ascii=False, indent=2)
print("suppliers.json updated successfully.")

# 2. Update import_history.json
history_path = os.path.join(ROOT, "import_history.json")
with open(history_path, "r", encoding="utf-8") as f:
    imp = json.load(f)

events = imp["events"]
ev130 = next((e for e in events if e["id"] == 130), None)
if ev130 and "324577" not in ev130.get("notes", ""):
    ev130["notes"] += ' חלק השקלים בארץ (28,755 ₪) שולם ב-07.10.2026 באסמכתא 324577 (אירוע 132). חלק הדולרים (50,550$) עדיין מעוכב עד לתשובת עומר.'

if not any(e["id"] == 132 for e in events):
    ev132 = {
        "id": 132,
        "date": "2026-10-07",
        "kind": "תשלום שילוח/עמילות",
        "party": "אקספרסליין קרגו",
        "description": 'הוראת העברה 28,755 ₪ מדיסקונט (ח-ן 1093039 / 93033) לפועלים סניף 658 ח-ן 611610 אקספרסליין קרגו בע"מ — תשלום הובלה ומע"מ ישראלי עבור דרישת תשלום 2630 (ביגוד עומר, 106 קרטונים). אסמכתא 324577',
        "amount": 28755,
        "currency": "ILS",
        "ils": 28755,
        "qty": "106 קרטונים",
        "certainty": "ודאי",
        "source": "אישור העברה בנקאית דיסקונט 07/10/2026 10:00:15 (אסמכתא 324577)",
        "notes": 'סגירת החלק השקלי של דרישה 2630 סופית (אירוע 130: הובלה 150$ + מע"מ ישראלי 9,126$ = 9,276$ × 3.10 = 28,755.60 ₪). חלקה הדולרי של הדרישה עבור הסחורה לטורקיה (50,550$ או 49,050$ בכפוף לאישור עומר לספיגת 1,500$ מע"מ טורקי — אירוע 131) טרם הועבר וממתין לתשובת עומר.',
        "attachments": [
            {
                "label": "אישור העברה 28,755 ₪ לאקספרסליין 07.10.2026 (אסמכתא 324577)",
                "path": "import_documents/132/אישור_העברה_28755_אקספרסליין_2630_07.10.2026.jpg"
            }
        ],
        "gmail_threads": []
    }
    events.append(ev132)
    print("Event 132 added to import_history.json.")

with open(history_path, "w", encoding="utf-8") as f:
    json.dump(imp, f, ensure_ascii=False, indent=2)
print("import_history.json saved successfully.")

# 3. Update shipping_companies.json
sc_path = os.path.join(ROOT, "shipping_companies.json")
with open(sc_path, "r", encoding="utf-8") as f:
    sc_comp = json.load(f)

for c in sc_comp:
    if "סאקיס" in c.get("company_name", "") or "אקספרסליין" in c.get("company_name", ""):
        note = c.get("notes", "")
        if "324577" not in note:
            c["notes"] = (
                'מוריד לנו את הסחורה בבר הופמן. הפרש הובלה 2,409 ₪ ממשלוח 53 קרטונים (09/2026, דרישה #2432) — זוכה בדרישה 2630 הסופית. '
                'דרישה 2630 (106 קרטונים, 10/2026): חלק השקלים 28,755 ₪ שולם ב-07.10.2026 (אסמכתא 324577). '
                'חלק הסחורה הדולרי מעוכב עד תשובת עומר על המע"מ הטורקי. ראה כרטיס ספק 6 (אקספרסליין / יניב).'
            )

with open(sc_path, "w", encoding="utf-8") as f:
    json.dump(sc_comp, f, ensure_ascii=False, indent=2)
print("shipping_companies.json updated.")
