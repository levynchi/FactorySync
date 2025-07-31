@echo off
echo ========================================
echo    FactorySync - ממיר אופטיטקס לאקסל
echo ========================================
echo.

REM מעבר לתיקיית התוכנה
cd /d "%~dp0"

REM בדיקה אם קיימת סביבה וירטואלית
if not exist ".venv\Scripts\activate.bat" (
    echo ❌ סביבה וירטואלית לא נמצאה!
    echo יוצר סביבה וירטואלית חדשה...
    python -m venv .venv
    if errorlevel 1 (
        echo ❌ שגיאה ביצירת סביבה וירטואלית
        echo אנא וודא ש-Python מותקן במערכת
        pause
        exit /b 1
    )
)

REM הפעלת הסביבה הוירטואלית
echo 🔄 מפעיל סביבה וירטואלית...
call .venv\Scripts\activate.bat

REM בדיקה אם requirements.txt קיים והתקנת תלויות
if exist "requirements.txt" (
    echo 📦 מתקין תלויות...
    pip install -r requirements.txt --quiet
)

REM הפעלת התוכנה
echo 🚀 מפעיל את התוכנה...
echo.

REM בדיקה איזה קובץ להפעיל
if exist "main.py" (
    python main.py
) else if exist "optitex_gui.py" (
    python optitex_gui.py
) else (
    echo ❌ לא נמצא קובץ התוכנה!
    pause
    exit /b 1
)

REM אם הגענו לכאן, התוכנה נסגרה
echo.
echo ✅ התוכנה נסגרה
pause
