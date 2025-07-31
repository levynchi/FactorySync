"""
חלון הגדרות אייר טייבל
"""

import tkinter as tk
from tkinter import messagebox

class AirtableConfigWindow:
    """חלון הגדרות אייר טייבל"""
    
    def __init__(self, parent, settings_manager, on_settings_changed=None):
        self.parent = parent
        self.settings = settings_manager
        self.on_settings_changed = on_settings_changed
        self.window = None
        
        self.api_key_var = tk.StringVar()
        self.base_id_var = tk.StringVar()
        self.table_id_var = tk.StringVar()
    
    def show(self):
        """הצגת חלון ההגדרות"""
        if self.window and self.window.winfo_exists():
            self.window.lift()
            return
        
        self.window = tk.Toplevel(self.parent)
        self.window.title("הגדרות Airtable")
        self.window.geometry("600x400")
        self.window.configure(bg='#f0f0f0')
        self.window.grab_set()  # חלון מודלי
        self.window.resizable(False, False)
        
        # מרכוז החלון
        self.window.transient(self.parent)
        
        self._create_widgets()
        self._load_current_settings()
    
    def _create_widgets(self):
        """יצירת הרכיבים"""
        # כותרת
        title_label = tk.Label(
            self.window,
            text="הגדרות חיבור לאייר טייבל",
            font=('Arial', 16, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        title_label.pack(pady=15)
        
        # מסגרת עיקרית
        main_frame = tk.Frame(self.window, bg='#f0f0f0')
        main_frame.pack(fill="both", expand=True, padx=30, pady=10)
        
        # API Key
        self._create_field(main_frame, "API Key:", self.api_key_var, show_password=True)
        
        # Base ID
        self._create_field(main_frame, "Base ID:", self.base_id_var)
        
        # Table ID
        self._create_field(main_frame, "Table ID:", self.table_id_var)
        
        # הערות עזרה
        help_frame = tk.LabelFrame(main_frame, text="עזרה", bg='#f0f0f0', fg='#666')
        help_frame.pack(fill="x", pady=(20, 10))
        
        help_text = """💡 איך למצוא את הפרטים:

• API Key: Account → API → Generate API key
• Base ID: מכתובת האתר של הבסיס (מתחיל ב-app...)
• Table ID: מכתובת האתר של הטבלה (מתחיל ב-tbl...)

🔗 מדריך מפורט: ראה קובץ 'הוראות_Airtable.md'"""
        
        help_label = tk.Label(
            help_frame,
            text=help_text,
            bg='#f0f0f0',
            fg='#666',
            font=('Arial', 9),
            justify="left",
            anchor="w"
        )
        help_label.pack(fill="both", padx=10, pady=5)
        
        # כפתורים
        buttons_frame = tk.Frame(self.window, bg='#f0f0f0')
        buttons_frame.pack(fill="x", padx=30, pady=15)
        
        # כפתור בדיקת חיבור
        tk.Button(
            buttons_frame,
            text="🔍 בדוק חיבור",
            command=self._test_connection,
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=12
        ).pack(side="left", padx=5)
        
        # כפתור שמירה
        tk.Button(
            buttons_frame,
            text="💾 שמור",
            command=self._save_settings,
            bg='#27ae60',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=12
        ).pack(side="right", padx=5)
        
        # כפתור ביטול
        tk.Button(
            buttons_frame,
            text="❌ ביטול",
            command=self._cancel,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=12
        ).pack(side="right", padx=5)
    
    def _create_field(self, parent, label_text, text_var, show_password=False):
        """יצירת שדה קלט"""
        field_frame = tk.Frame(parent, bg='#f0f0f0')
        field_frame.pack(fill="x", pady=8)
        
        label = tk.Label(
            field_frame,
            text=label_text,
            bg='#f0f0f0',
            font=('Arial', 11, 'bold'),
            width=12,
            anchor="w"
        )
        label.pack(side="left")
        
        entry = tk.Entry(
            field_frame,
            textvariable=text_var,
            font=('Arial', 10),
            width=50,
            show="*" if show_password else ""
        )
        entry.pack(side="right", padx=(10, 0))
        
        return entry
    
    def _load_current_settings(self):
        """טעינת ההגדרות הנוכחיות"""
        self.api_key_var.set(self.settings.get("airtable.api_key", ""))
        self.base_id_var.set(self.settings.get("airtable.base_id", ""))
        self.table_id_var.set(self.settings.get("airtable.table_id", "tblC0hR3gZFXxstbM"))
    
    def _test_connection(self):
        """בדיקת חיבור לאייר טייבל"""
        api_key = self.api_key_var.get().strip()
        base_id = self.base_id_var.get().strip()
        table_id = self.table_id_var.get().strip()
        
        if not api_key:
            messagebox.showerror("שגיאה", "אנא הזן API Key")
            return
        
        if not base_id:
            messagebox.showerror("שגיאה", "אנא הזן Base ID")
            return
        
        try:
            from ..core.airtable_client import AirtableClient
            
            # יצירת לקוח זמני לבדיקה
            client = AirtableClient(api_key, base_id, table_id)
            
            if client.test_connection():
                messagebox.showinfo("הצלחה", "החיבור לאייר טייבל הצליח! ✅")
            else:
                messagebox.showerror("שגיאה", "החיבור לאייר טייבל נכשל")
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בבדיקת החיבור:\n{str(e)}")
    
    def _save_settings(self):
        """שמירת ההגדרות"""
        api_key = self.api_key_var.get().strip()
        base_id = self.base_id_var.get().strip()
        table_id = self.table_id_var.get().strip()
        
        # בדיקות בסיסיות
        if api_key and not base_id:
            messagebox.showerror("שגיאה", "אנא הזן Base ID")
            return
        
        if base_id and not api_key:
            messagebox.showerror("שגיאה", "אנא הזן API Key")
            return
        
        # שמירת ההגדרות
        self.settings.set("airtable.api_key", api_key)
        self.settings.set("airtable.base_id", base_id)
        self.settings.set("airtable.table_id", table_id or "tblC0hR3gZFXxstbM")
        
        # עדכון הקורא
        if self.on_settings_changed:
            self.on_settings_changed()
        
        messagebox.showinfo("הצלחה", "ההגדרות נשמרו בהצלחה!")
        self._close()
    
    def _cancel(self):
        """ביטול השינויים"""
        self._close()
    
    def _close(self):
        """סגירת החלון"""
        if self.window:
            self.window.destroy()
            self.window = None
