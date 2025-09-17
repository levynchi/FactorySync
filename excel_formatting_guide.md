# מדריך עיצוב אקסל - Optitex Analyzer

## סקירה כללית
מדריך זה מכיל את כל ההגדרות העיצוביות המשמשות לייצוא קבצי אקסל במערכת Optitex Analyzer. השתמש בהגדרות אלה כדי לשמור על עקביות עיצובית בכל הדפים עם ייצוא לאקסל.

## הגדרות בסיסיות

### 1. הגדרות דף
```python
# כיוון דף
ws.page_setup.orientation = 'portrait'

# גודל נייר A4
ws.page_setup.paperSize = 9

# התאמה לרוחב בלבד
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0

# מרכוז אופקי
ws.print_options.horizontalCentered = True
ws.print_options.verticalCentered = False

# שוליים (0.5 אינץ' מכל צד)
ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)

# כיוון RTL לעברית
ws.sheet_view.rightToLeft = True
```

### 2. הגדרות פונט

#### כותרות עסקיות
```python
Font(size=16, bold=True)
```

#### כותרות טבלה
```python
Font(bold=True, size=16)
```

#### נתוני טבלה
```python
Font(size=16)
```

#### מידע מסמך (מפתחות)
```python
Font(bold=True, size=16)
```

#### מידע מסמך (ערכים)
```python
Font(size=16)
```

### 3. הגדרות יישור

#### טקסט עברי
```python
Alignment(horizontal='right')
```

#### כמויות ומספרים
```python
Alignment(horizontal='center')
```

#### כותרות טבלה
```python
Alignment(horizontal='center')
```

### 4. הגדרות צבעים

#### כותרת טבלה אפורה
```python
from openpyxl.styles import PatternFill
gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

# יישום על כותרות טבלה
for col in range(1, num_columns + 1):
    cell = ws.cell(row=header_row, column=col)
    cell.fill = gray_fill
    cell.font = Font(bold=True, size=16)
    cell.alignment = Alignment(horizontal='center')
```

### 5. הגדרות גבולות

#### גבולות טבלה
```python
from openpyxl.styles import Border, Side
thin = Side(style='thin', color='000000')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
```

### 6. חישוב רוחב עמודות (AutoFit)

```python
def calculate_column_width(ws, col_letter):
    """
    מחשב רוחב עמודה אוטומטי כמו Excel AutoFit
    """
    max_width = 0
    col_num = ord(col_letter) - ord('A') + 1
    
    # בדוק את כל התאים בעמודה
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_num, max_col=col_num):
        for cell in row:
            try:
                val = str(cell.value) if cell.value is not None else ""
                if val.strip():
                    char_width = 0
                    for char in val:
                        # תווים עבריים רחבים יותר
                        if '\u0590' <= char <= '\u05FF':
                            char_width += 1.2
                        # מספרים ואנגלית צרים יותר
                        elif char.isdigit() or char.isalpha():
                            char_width += 0.6
                        # תווים מיוחדים
                        else:
                            char_width += 0.8
                    
                    # הוסף רווח לגבולות תאים
                    char_width += 2
                    max_width = max(max_width, char_width)
            except Exception:
                pass
    
    # הגדר רוחב עם הגבלות הגיוניות
    if max_width > 0:
        return min(max(8, max_width), 50)  # מינימום 8, מקסימום 50
    else:
        return 12  # ברירת מחדל
```

### 7. הגדרות לוגו

#### מידות לוגו
```python
# גובה: 4.66 ס"מ, רוחב: 15.73 ס"מ
img.height = 4.66 * 37.8  # המרה לפיקסלים
img.width = 15.73 * 37.8

# גובה שורה עבור לוגו
ws.row_dimensions[1].height = 4.66 * 28.35  # המרה ליחידות Excel
```

### 8. פורמט תאריכים

```python
# פורמט תאריך עברי
cell.number_format = 'd/m/yyyy'
```

## דוגמה לשימוש מלא

```python
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins

def create_formatted_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "דוגמה"
    
    # הגדרות בסיסיות
    ws.sheet_view.rightToLeft = True
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
    
    # הוסף נתונים
    headers = ["כותרת 1", "כותרת 2", "כותרת 3", "כמות"]
    ws.append(headers)
    
    # עיצוב כותרות
    gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    for col in range(1, 5):
        cell = ws.cell(row=1, column=col)
        cell.fill = gray_fill
        cell.font = Font(bold=True, size=16)
        cell.alignment = Alignment(horizontal='center')
        cell.border = border
    
    # חישוב רוחב עמודות
    for col_letter in ['A', 'B', 'C', 'D']:
        width = calculate_column_width(ws, col_letter)
        ws.column_dimensions[col_letter].width = width
    
    return wb
```

## כללי עיצוב כלליים

1. **עקביות**: השתמש תמיד באותן הגדרות פונט ויישור
2. **קריאות**: גודל פונט 16 לכל הטקסט
3. **ארגון**: כותרות אפורות עם גבולות שחורים
4. **יישור**: עברית מימין, מספרים במרכז
5. **רוחב**: AutoFit עם הגבלות הגיוניות
6. **דף**: A4 Portrait עם מרכוז אופקי

## הערות חשובות

- תמיד השתמש ב-RTL עבור טקסט עברי
- הוסף גבולות לכל הטבלות
- השתמש בגופן עבה לכותרות
- וודא שהעמודות לא רחבות מדי או צרות מדי
- בדוק שהלוגו במידות הנכונות אם קיים
