# FactorySync - ממיר אקספורט של אופטיטקס לאקסל נקי

## תיאור הפרויקט
תוכנה לניתוח קבצי אקספורט של אופטיטקס והמרתם לפורמט אקסל נקי עם אפשרות העלאה ישירה לאירטייבל.

## פיצ'רים עיקריים
- 🔍 ניתוח קבצי Excel מאופטיטקס
- 📊 המרה לפורמט נקי ומסודר
- 🔄 טיפול אוטומטי ב-Layout Tubular
- 📤 העלאה ישירה לאירטייבל
- ⚙️ ממשק גרפי ידידותי למשתמש
- 💾 שמירת הגדרות חיבור לאירטייבל

## התקנה

### דרישות מערכת
- Python 3.8 ומעלה
- Windows/Mac/Linux

### הוראות התקנה

1. שכפל את הפרויקט:
```bash
git clone https://github.com/levynchi/FactorySync.git
cd FactorySync
```

2. צור סביבה וירטואלית:
```bash
python -m venv .venv
```

3. הפעל את הסביבה הוירטואלית:

**Windows:**
```bash
.venv\Scripts\activate
```

**Mac/Linux:**
```bash
source .venv/bin/activate
```

4. התקן את הספריות הנדרשות:
```bash
pip install -r requirements.txt
```

## שימוש

### הפעלת התוכנה
```bash
python optitex_gui.py
```

### הגדרת חיבור לאירטייבל
1. לחץ על כפתור "⚙️ הגדר חיבור"
2. הזן את פרטי החיבור:
   - **API Key**: מפתח ה-API מאירטייבל
   - **Base ID**: מזהה הבסיס באירטייבל
   - **Table ID**: מזהה הטבלה באירטייבל

### תהליך עבודה
1. **בחר קובץ אופטיטקס**: קובץ Excel מיוצא מאופטיטקס
2. **בחר קובץ מוצרים**: קובץ מיפוי שמות מוצרים
3. **נתח קבצים**: לחץ על "🔍 נתח קבצים"
4. **שמור או העלה**:
   - "💾 שמור כ-Excel" - שמירה לקובץ מקומי
   - "📊 העלה לאייר טייבל" - העלאה ישירה לאירטייבל

## מבנה הפרויקט
```
FactorySync/
├── optitex_gui.py          # הקובץ הראשי של התוכנה
├── config.json             # קובץ הגדרות חיבור לאירטייבל
├── requirements.txt        # רשימת ספריות נדרשות
├── קובץ מוצרים.xlsx        # דוגמה לקובץ מיפוי מוצרים
├── rib4.xlsx              # דוגמה לקובץ אופטיטקס
├── rib6.xlsx              # דוגמה נוספת לקובץ אופטיטקס
└── .venv/                 # סביבה וירטואלית
```

## טכנולוגיות
- **Python 3.x**
- **tkinter** - ממשק גרפי
- **pandas** - עיבוד נתונים
- **openpyxl** - טיפול בקובצי Excel
- **pyairtable** - חיבור לאירטייבל

## תמיכה והערות
- התוכנה תומכת בעברית ובאנגלית
- טיפול אוטומטי ב-Layout Tubular (חלוקה ב-2)
- מיון חכם של מידות (מספרים, חודשים, טקסט)
- שמירת הגדרות בקובץ JSON

## רישיון
פרויקט בקוד פתוח למטרות לימוד ושימוש מסחרי.

---
**פותח על ידי לוי**
📧 levynchi@example.com
