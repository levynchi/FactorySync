"""
FactorySync - תוכנה לניתוח ועיבוד קבצי אופטיטקס

קובץ ראשי חדש עם מבנה מודולרי
"""

import tkinter as tk
import sys
import os

# הוספת התיקייה הנוכחית ל-Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

from src.gui.main_window import MainWindow


def main():
    """פונקציה ראשית"""
    try:
        # יצירת חלון ראשי
        root = tk.Tk()
        
        # יצירת האפליקציה
        app = MainWindow(root)
        
        # הפעלת לולאת הממשק
        root.mainloop()
        
    except Exception as e:
        print(f"שגיאה בהפעלת התוכנה: {e}")
        input("לחץ Enter לסגירה...")


if __name__ == "__main__":
    main()
