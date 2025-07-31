@echo off
echo ========================================
echo    התקנה ראשונית - FactorySync
echo ========================================
echo.

REM מעבר לתיקיית התוכנה
cd /d "%~dp0"

echo 🔧 יוצר סביבה וירטואלית...
python -m venv .venv

if errorlevel 1 (
    echo ❌ שגיאה ביצירת סביבה וירטואלית
    echo אנא וודא ש-Python מותקן במערכת
    pause
    exit /b 1
)

echo ✅ סביבה וירטואלית נוצרה בהצלחה

echo 🔄 מפעיל סביבה וירטואלית...
call .venv\Scripts\activate.bat

echo 📦 מתקין תלויות...
pip install -r requirements.txt

if errorlevel 1 (
    echo ❌ שגיאה בהתקנת תלויות
    pause
    exit /b 1
)

echo.
echo ✅ התקנה הושלמה בהצלחה!
echo.
echo כעת אתה יכול להפעיל את התוכנה על ידי:
echo 1. הרצת start.bat - הפעלה מהירה
echo 2. הפעלה ידנית: .venv\Scripts\activate.bat ואז python main.py
echo.
pause
