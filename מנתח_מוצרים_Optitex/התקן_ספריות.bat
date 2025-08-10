@echo off
setlocal

REM תיקיית הסקריפט
set "ROOT=%~dp0"

REM יצירת/שימוש בסביבה וירטואלית מקומית
if not exist "%ROOT%\.venv\Scripts\python.exe" (
  py -3 -m venv "%ROOT%\.venv"
)
call "%ROOT%\.venv\Scripts\activate"

REM עדכון pip והתקנת תלויות
python -m pip install --upgrade pip
if exist "%ROOT%requirements.txt" (
  pip install -r "%ROOT%requirements.txt"
) else (
  echo requirements.txt לא נמצא ב-%ROOT%
  echo מתקין חבילות ברירת מחדל...
  pip install pandas openpyxl tk
)

echo.
echo ✔ ההתקנה הושלמה. הסביבה הווירטואלית: %ROOT%\.venv
pause
