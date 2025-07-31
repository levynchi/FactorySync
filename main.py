"""
FactorySync - ממיר אופטיטקס לאקסל
נקודת הכניסה הראשית לתוכנה
"""

import tkinter as tk
from tkinter import messagebox
import sys
import os

# הוספת תיקיית התוכנה ל-path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def main():
    """פונקציה ראשית"""
    try:
        print("🚀 מתחיל טעינת התוכנה...")
        
        # יבוא המודולים
        print("📦 טוען מודול הגדרות...")
        from optitex_analyzer.config.settings import SettingsManager
        print("✅ מודול הגדרות נטען בהצלחה")
        
        print("📦 טוען מודול ניתוח קבצים...")
        from optitex_analyzer.core.file_analyzer import OptitexFileAnalyzer
        print("✅ מודול ניתוח קבצים נטען בהצלחה")
        
        print("📦 טוען מודול עיבוד נתונים...")
        from optitex_analyzer.core.data_processor import DataProcessor
        print("✅ מודול עיבוד נתונים נטען בהצלחה")
        
        print("📦 טוען מודול החלון הראשי...")
        from optitex_analyzer.gui.main_window import MainWindow
        print("✅ מודול החלון הראשי נטען בהצלחה")
        
        # יצירת החלון הראשי
        print("🖼️  יוצר חלון ראשי...")
        root = tk.Tk()
        print("✅ חלון ראשי נוצר בהצלחה")
        
        # הגדרת סגנון
        print("🎨 מגדיר סגנון...")
        root.tk_setPalette(background='#f0f0f0')
        print("✅ סגנון הוגדר בהצלחה")
        
        # יצירת המנהלים
        print("⚙️  יוצר מנהלי המערכת...")
        try:
            print("   - יוצר מנהל הגדרות...")
            settings_manager = SettingsManager()
            print("   ✅ מנהל הגדרות נוצר בהצלחה")
            
            print("   - יוצר מנהל ניתוח קבצים...")
            file_analyzer = OptitexFileAnalyzer()
            print("   ✅ מנהל ניתוח קבצים נוצר בהצלחה")
            
            print("   - יוצר מנהל עיבוד נתונים...")
            data_processor = DataProcessor()
            print("   ✅ מנהל עיבוד נתונים נוצר בהצלחה")
            
        except Exception as e:
            print(f"❌ שגיאה ביצירת המנהלים: {str(e)}")
            messagebox.showerror(
                "שגיאת אתחול",
                f"שגיאה באתחול המערכת:\n{str(e)}\n\nהתוכנה תיסגר."
            )
            return
        
        # יצירת החלון הראשי
        print("🖼️  יוצר ממשק המשתמש...")
        try:
            app = MainWindow(
                root,
                settings_manager,
                file_analyzer,
                data_processor
            )
            print("✅ ממשק המשתמש נוצר בהצלחה")
            
            # הגדרת סגירת התוכנה
            print("🔗 מגדיר פונקציות סגירה...")
            def on_closing():
                try:
                    print("💾 שומר הגדרות לפני סגירה...")
                    # שמירת הגדרות לפני סגירה
                    window_geometry = root.geometry()
                    settings_manager.set("app.window_size", window_geometry)
                    settings_manager.save_config()
                    print("✅ הגדרות נשמרו בהצלחה")
                except Exception as ex:
                    print(f"⚠️  שגיאה בשמירת הגדרות: {ex}")
                finally:
                    print("👋 סוגר את התוכנה...")
                    root.destroy()
            
            root.protocol("WM_DELETE_WINDOW", on_closing)
            print("✅ פונקציות סגירה הוגדרו בהצלחה")
            
            # הרצת התוכנה
            print("🚀 מפעיל את התוכנה...")
            print("=" * 50)
            print("התוכנה מוכנה לשימוש!")
            print("=" * 50)
            root.mainloop()
            
        except Exception as e:
            print(f"❌ שגיאה ביצירת הממשק: {str(e)}")
            import traceback
            print("פרטי השגיאה המלאים:")
            traceback.print_exc()
            messagebox.showerror(
                "שגיאת GUI",
                f"שגיאה ביצירת הממשק:\n{str(e)}\n\nהתוכנה תיסגר."
            )
            return
            
    except ImportError as e:
        print(f"❌ שגיאת יבוא: {str(e)}")
        import traceback
        print("פרטי השגיאה המלאים:")
        traceback.print_exc()
        messagebox.showerror(
            "שגיאת יבוא",
            f"שגיאה בטעינת המודולים:\n{str(e)}\n\nודא שכל הקבצים נמצאים במקום הנכון."
        )
        
    except Exception as e:
        print(f"❌ שגיאה כללית: {str(e)}")
        import traceback
        print("פרטי השגיאה המלאים:")
        traceback.print_exc()
        messagebox.showerror(
            "שגיאה כללית",
            f"שגיאה לא צפויה:\n{str(e)}\n\nהתוכנה תיסגר."
        )

if __name__ == "__main__":
    main()
