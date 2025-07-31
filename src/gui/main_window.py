"""
חלון ממשק הראשי של התוכנה - גרסה מעודכנת עם מבנה מודולרי
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from threading import Thread

# ייבוא המודולים שלנו
from src.core.analyzer import OptitexAnalyzer
from src.core.drawings_manager import DrawingsManager
from src.core.airtable_manager import AirtableManager
from src.utils.config_manager import ConfigManager
from src.gui.drawings_window import DrawingsManagerWindow


class MainWindow:
    """חלון הממשק הראשי"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("FactorySync - ממיר אקספורט של אופטיטקס לאקסל נקי")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # יצירת המודולים העיקריים
        self.config_manager = ConfigManager()
        self.analyzer = OptitexAnalyzer()
        self.drawings_manager = DrawingsManager()
        self.airtable_manager = None
        
        # משתנים לממשק
        self.rib_file = ""
        self.products_file = ""
        self.output_df = None
        
        # חלון מנהל ציורים
        self.drawings_window = DrawingsManagerWindow(self.root, self.drawings_manager)
        
        # בדיקה אוטומטית לקובץ מוצרים
        auto_products_file = self.config_manager.find_products_file()
        if auto_products_file:
            self.products_file = auto_products_file
        
        # יצירת הממשק
        self.create_widgets()
        
        # עדכון סטטוס Airtable
        self.update_airtable_status()
        
        # עדכון תצוגת מוצרים
        self.update_products_display()
    
    def create_widgets(self):
        """יצירת רכיבי הממשק"""
        # כותרת
        title_label = tk.Label(
            self.root, 
            text="FactorySync - ממיר אקספורט אופטיטקס", 
            font=('Arial', 16, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        title_label.pack(pady=10)
        
        # מסגרת בחירת קבצים
        self.create_files_section()
        
        # מסגרת אפשרויות
        self.create_options_section()
        
        # כפתורי פעולה
        self.create_buttons_section()
        
        # אזור תוצאות
        self.create_results_section()
        
        # שורת סטטוס
        self.create_status_bar()
    
    def create_files_section(self):
        """יצירת מסגרת בחירת קבצים"""
        files_frame = ttk.LabelFrame(self.root, text="בחירת קבצים", padding=10)
        files_frame.pack(fill="x", padx=20, pady=10)
        
        # בחירת קובץ אופטיטקס
        rib_frame = tk.Frame(files_frame)
        rib_frame.pack(fill="x", pady=5)
        
        tk.Label(rib_frame, text="אופטיטקס אקסל אקספורט:").pack(side="left")
        self.rib_label = tk.Label(rib_frame, text="לא נבחר", bg="white", width=50, anchor="w")
        self.rib_label.pack(side="left", padx=5)
        
        tk.Button(
            rib_frame, 
            text="בחר קובץ", 
            command=self.select_rib_file,
            bg='#3498db',
            fg='white'
        ).pack(side="right")
        
        # בחירת קובץ מוצרים
        products_frame = tk.Frame(files_frame)
        products_frame.pack(fill="x", pady=5)
        
        tk.Label(products_frame, text="קובץ מוצרים:").pack(side="left")
        self.products_label = tk.Label(products_frame, text="לא נבחר", bg="white", width=50, anchor="w")
        self.products_label.pack(side="left", padx=5)
        
        tk.Button(
            products_frame, 
            text="בחר קובץ", 
            command=self.select_products_file,
            bg='#3498db',
            fg='white'
        ).pack(side="right")
        
        # עדכון תצוגת קובץ מוצרים אם נמצא
        # הערה: יקרא אחר יצירת ה-status_label ב-create_widgets
        # self.update_products_display()
    
    def create_options_section(self):
        """יצירת מסגרת אפשרויות"""
        options_frame = ttk.LabelFrame(self.root, text="אפשרויות", padding=10)
        options_frame.pack(fill="x", padx=20, pady=10)
        
        # הגדרות Airtable
        self.create_airtable_section(options_frame)
        
        # אפשרויות ניתוח
        app_config = self.config_manager.get_app_config()
        
        self.tubular_var = tk.BooleanVar(value=app_config.get("check_tubular", True))
        tk.Checkbutton(
            options_frame,
            text="טיפול אוטומטי ב-Layout Tubular (חלוקה ב-2)",
            variable=self.tubular_var
        ).pack(anchor="w")
        
        self.only_positive_var = tk.BooleanVar(value=app_config.get("only_positive_quantities", True))
        tk.Checkbutton(
            options_frame,
            text="הצג רק מידות עם כמות גדולה מ-0",
            variable=self.only_positive_var
        ).pack(anchor="w")
    
    def create_airtable_section(self, parent):
        """יצירת מסגרת הגדרות Airtable"""
        airtable_frame = ttk.LabelFrame(parent, text="הגדרות Airtable", padding=5)
        airtable_frame.pack(fill="x", pady=5)
        
        # סטטוס חיבור
        status_frame = tk.Frame(airtable_frame)
        status_frame.pack(fill="x", pady=2)
        
        tk.Label(status_frame, text="סטטוס:", font=('Arial', 9, 'bold')).pack(side="left")
        self.airtable_status = tk.Label(status_frame, text="", font=('Arial', 9))
        self.airtable_status.pack(side="left", padx=5)
        
        # כפתור הגדרות
        tk.Button(
            status_frame,
            text="⚙️ הגדר חיבור",
            command=self.configure_airtable,
            bg='#3498db',
            fg='white',
            font=('Arial', 8)
        ).pack(side="right")
    
    def create_buttons_section(self):
        """יצירת כפתורי פעולה"""
        buttons_frame = tk.Frame(self.root)
        buttons_frame.pack(fill="x", padx=20, pady=10)
        
        tk.Button(
            buttons_frame,
            text="🔍 נתח קבצים",
            command=self.analyze_files,
            bg='#27ae60',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        ).pack(side="left", padx=5)
        
        tk.Button(
            buttons_frame,
            text="💾 שמור כ-Excel",
            command=self.save_excel,
            bg='#e67e22',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        ).pack(side="left", padx=5)
        
        tk.Button(
            buttons_frame,
            text="📊 העלה לאייר טייבל",
            command=self.upload_to_airtable,
            bg='#9b59b6',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        ).pack(side="left", padx=5)
        
        tk.Button(
            buttons_frame,
            text="📋 הוסף לטבלה המקומית",
            command=self.add_to_local_table,
            bg='#16a085',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        ).pack(side="left", padx=5)
        
        tk.Button(
            buttons_frame,
            text="📁 מנהל ציורים",
            command=self.open_drawings_manager,
            bg='#2980b9',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        ).pack(side="left", padx=5)
        
        tk.Button(
            buttons_frame,
            text="🗑️ נקה הכל",
            command=self.clear_all,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        ).pack(side="right", padx=5)
    
    def create_results_section(self):
        """יצירת אזור תוצאות"""
        results_frame = ttk.LabelFrame(self.root, text="תוצאות", padding=10)
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.results_text = scrolledtext.ScrolledText(
            results_frame,
            height=15,
            font=('Courier New', 10),
            wrap=tk.WORD
        )
        self.results_text.pack(fill="both", expand=True)
    
    def create_status_bar(self):
        """יצירת שורת סטטוס"""
        self.status_label = tk.Label(
            self.root,
            text="מוכן לעבודה",
            bg='#34495e',
            fg='white',
            anchor='w',
            padx=10
        )
        self.status_label.pack(fill="x", side="bottom")
    
    def select_rib_file(self):
        """בחירת קובץ אופטיטקס"""
        file_path = filedialog.askopenfilename(
            title="בחר קובץ אופטיטקס אקסל אקספורט",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.rib_file = file_path
            self.rib_label.config(text=os.path.basename(file_path))
            self.update_status(f"נבחר קובץ אופטיטקס: {os.path.basename(file_path)}")
    
    def select_products_file(self):
        """בחירת קובץ מוצרים"""
        file_path = filedialog.askopenfilename(
            title="בחר קובץ רשימת מוצרים",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.products_file = file_path
            self.products_label.config(text=os.path.basename(file_path))
            self.update_status(f"נבחר קובץ מוצרים: {os.path.basename(file_path)}")
    
    def update_products_display(self):
        """עדכון תצוגת קובץ המוצרים"""
        if self.products_file:
            self.products_label.config(text=os.path.basename(self.products_file))
            self.update_status(f"נמצא קובץ מוצרים: {os.path.basename(self.products_file)}")
    
    def configure_airtable(self):
        """פתיחת חלון הגדרות Airtable"""
        config_window = tk.Toplevel(self.root)
        config_window.title("הגדרות Airtable")
        config_window.geometry("500x300")
        config_window.configure(bg='#f0f0f0')
        config_window.grab_set()
        
        # כותרת
        tk.Label(
            config_window,
            text="הגדרות חיבור לאייר טייבל",
            font=('Arial', 14, 'bold'),
            bg='#f0f0f0'
        ).pack(pady=10)
        
        # מסגרת עיקרית
        main_frame = tk.Frame(config_window, bg='#f0f0f0')
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # שדות הזנה
        airtable_config = self.config_manager.get_airtable_config()
        
        tk.Label(main_frame, text="API Key:", bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor="w")
        api_key_entry = tk.Entry(main_frame, width=60, show="*")
        api_key_entry.pack(fill="x", pady=(0, 10))
        api_key_entry.insert(0, airtable_config.get("api_key", ""))
        
        tk.Label(main_frame, text="Base ID:", bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor="w")
        base_id_entry = tk.Entry(main_frame, width=60)
        base_id_entry.pack(fill="x", pady=(0, 10))
        base_id_entry.insert(0, airtable_config.get("base_id", ""))
        
        tk.Label(main_frame, text="Table ID:", bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor="w")
        table_id_entry = tk.Entry(main_frame, width=60)
        table_id_entry.pack(fill="x", pady=(0, 15))
        table_id_entry.insert(0, airtable_config.get("table_id", "tblC0hR3gZFXxstbM"))
        
        # כפתורים
        buttons_frame = tk.Frame(config_window, bg='#f0f0f0')
        buttons_frame.pack(fill="x", padx=20, pady=10)
        
        def save_and_close():
            success = self.config_manager.set_airtable_config(
                api_key_entry.get(),
                base_id_entry.get(),
                table_id_entry.get()
            )
            
            if success:
                self.update_airtable_status()
                config_window.destroy()
                messagebox.showinfo("הצלחה", "ההגדרות נשמרו בהצלחה!")
            else:
                messagebox.showerror("שגיאה", "שגיאה בשמירת הגדרות")
        
        tk.Button(
            buttons_frame,
            text="💾 שמור",
            command=save_and_close,
            bg='#27ae60',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=10
        ).pack(side="left", padx=5)
        
        tk.Button(
            buttons_frame,
            text="❌ ביטול",
            command=config_window.destroy,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=10
        ).pack(side="right", padx=5)
    
    def update_airtable_status(self):
        """עדכון סטטוס Airtable"""
        status_info = self.config_manager.get_airtable_status()
        self.airtable_status.config(text=status_info['text'], fg=status_info['color'])
    
    def update_status(self, message):
        """עדכון שורת סטטוס"""
        self.status_label.config(text=message)
        self.root.update()
    
    def log_message(self, message):
        """הוספת הודעה לאזור התוצאות"""
        self.results_text.insert(tk.END, message + "\n")
        self.results_text.see(tk.END)
        self.root.update()
    
    def clear_results(self):
        """ניקוי אזור התוצאות"""
        self.results_text.delete(1.0, tk.END)
    
    def analyze_files(self):
        """ניתוח הקבצים"""
        if not self.rib_file or not self.products_file:
            messagebox.showerror("שגיאה", "אנא בחר את שני הקבצים")
            return
        
        self.clear_results()
        self.update_status("מנתח קבצים...")
        
        Thread(target=self._analyze_files_thread, daemon=True).start()
    
    def _analyze_files_thread(self):
        """ביצוע הניתוח בחוט נפרד"""
        try:
            self.log_message("=== התחלת ניתוח ===")
            
            # ביצוע הניתוח
            result = self.analyzer.analyze_files(
                self.rib_file,
                self.products_file,
                self.tubular_var.get(),
                self.only_positive_var.get()
            )
            
            if result['success']:
                self.output_df = result['data']
                
                self.log_message(f"✅ נטען מיפוי עבור {result['products_mapping_count']} מוצרים")
                
                if result['is_tubular']:
                    self.log_message("🔍 נמצא Layout: Tubular - הכמויות חולקו ב-2")
                else:
                    self.log_message("📋 לא נמצא Layout: Tubular - הכמויות יישארו כמו שהן")
                
                self.log_message(f"✅ נוצרה טבלה עם {len(self.output_df)} רשומות")
                
                # הצגת תוצאות מסודרות
                self.log_message("\n=== תוצאות הניתוח ===")
                
                current_product = None
                for _, row in self.output_df.iterrows():
                    if current_product != row['שם המוצר']:
                        current_product = row['שם המוצר']
                        self.log_message(f"\n📦 {current_product}:")
                        self.log_message("-" * 50)
                    
                    quantity_text = f"{row['כמות']}"
                    if row['כמות מקורית'] != row['כמות']:
                        quantity_text += f" (מקורי: {row['כמות מקורית']})"
                    
                    self.log_message(f"   מידה {row['מידה']:>6}: {quantity_text:>8} - {row['הערה']}")
                
                # סטטיסטיקות
                self.log_message(f"\n=== סיכום ===")
                self.log_message(f"מוצרים: {result['total_products']}")
                self.log_message(f"מידות: {result['total_sizes']}")
                self.log_message(f"סך כמויות: {result['total_quantity']}")
                
                if result['is_tubular']:
                    self.log_message("🔄 הכמויות חולקו ב-2 בגלל Layout: Tubular")
                
                self.update_status("הניתוח הושלם בהצלחה!")
            else:
                self.log_message(f"❌ {result['error']}")
                self.update_status(f"שגיאה: {result['error']}")
                
        except Exception as e:
            self.log_message(f"❌ שגיאה: {str(e)}")
            self.update_status(f"שגיאה: {str(e)}")
            messagebox.showerror("שגיאה", f"שגיאה בניתוח: {str(e)}")
    
    def save_excel(self):
        """שמירת התוצאות לקובץ Excel"""
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("אזהרה", "אין נתונים לשמירה. אנא בצע ניתוח תחילה.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="שמור כ-Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                success = self.analyzer.save_to_excel(file_path)
                if success:
                    self.log_message(f"📁 הקובץ נשמר בהצלחה: {file_path}")
                    self.update_status(f"נשמר: {os.path.basename(file_path)}")
                    messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה:\\n{file_path}")
            except Exception as e:
                messagebox.showerror("שגיאה", str(e))
    
    def upload_to_airtable(self):
        """העלאה ל-Airtable"""
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("אזהרה", "אין נתונים להעלאה. אנא בצע ניתוח תחילה.")
            return
        
        # יצירת מנהל Airtable
        airtable_config = self.config_manager.get_airtable_config()
        self.airtable_manager = AirtableManager(
            airtable_config.get("api_key", ""),
            airtable_config.get("base_id", ""),
            airtable_config.get("table_id", "")
        )
        
        try:
            self.log_message("\n=== התחלת העלאה לאייר טייבל ===")
            
            result = self.airtable_manager.upload_drawing(self.rib_file, self.output_df)
            
            if result['success']:
                self.log_message(f"✅ הנתונים הועלו בהצלחה לאייר טייבל!")
                self.log_message(f"ID של הרשומה החדשה: {result['record_id']}")
                self.log_message(f"מספר דגמים שהועלו: {result['models_count']}")
                
                # הצגת הדגמים שהועלו
                record_data = result['record_data']
                self.log_message(f"\nשם הקובץ: {record_data['שם הקובץ']}")
                for i in range(1, result['models_count'] + 1):
                    model = record_data.get(f'דגם {i}', '')
                    quantity = record_data.get(f'כמות דגם {i}', '')
                    if model:
                        self.log_message(f"  דגם {i}: {model} - כמות: {quantity}")
                
                self.update_status("העלאה לאייר טייבל הושלמה בהצלחה!")
                messagebox.showinfo("הצלחה", f"הנתונים הועלו בהצלחה לאייר טייבל!\\nID רשומה: {result['record_id']}")
            else:
                self.log_message(f"❌ {result['error']}")
                if result.get('help'):
                    self.log_message(f"💡 {result['help']}")
                messagebox.showerror("שגיאה", result['error'])
                
        except Exception as e:
            error_msg = f"שגיאה בהעלאה לאייר טייבל: {str(e)}"
            self.log_message(f"❌ {error_msg}")
            messagebox.showerror("שגיאה", error_msg)
    
    def add_to_local_table(self):
        """הוספה לטבלה המקומית"""
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("אזהרה", "אין נתונים להוספה. אנא בצע ניתוח תחילה.")
            return
        
        try:
            result = self.drawings_manager.add_drawing(self.rib_file, self.output_df)
            
            if result['success']:
                record = result['record']
                self.log_message(f"\n✅ הציור נוסף לטבלה המקומית!")
                self.log_message(f"ID רשומה: {record['id']}")
                self.log_message(f"שם הקובץ: {record['שם הקובץ']}")
                self.log_message(f"מוצרים: {len(record['מוצרים'])}")
                self.log_message(f"סך כמויות: {record['סך כמויות']}")
                
                messagebox.showinfo("הצלחה", f"הציור נוסף בהצלחה לטבלה המקומית!\\nID רשומה: {record['id']}")
            else:
                messagebox.showerror("שגיאה", result['error'])
                
        except Exception as e:
            error_msg = f"שגיאה בהוספה לטבלה המקומית: {str(e)}"
            self.log_message(f"❌ {error_msg}")
            messagebox.showerror("שגיאה", error_msg)
    
    def open_drawings_manager(self):
        """פתיחת מנהל הציורים"""
        self.drawings_window.open_window()
    
    def clear_all(self):
        """ניקוי כל הנתונים"""
        self.rib_file = ""
        self.output_df = None
        self.rib_label.config(text="לא נבחר")
        self.clear_results()
        
        # בדיקה מחדש לקובץ מוצרים
        auto_products_file = self.config_manager.find_products_file()
        if auto_products_file:
            self.products_file = auto_products_file
            self.products_label.config(text=os.path.basename(self.products_file))
            self.update_status(f"נמצא קובץ מוצרים: {os.path.basename(self.products_file)}")
        else:
            self.products_file = ""
            self.products_label.config(text="לא נבחר")
            self.update_status("מוכן לעבודה")
