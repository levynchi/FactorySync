"""
החלון הראשי של התוכנה
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from threading import Thread

class MainWindow:
    """החלון הראשי של התוכנה"""
    
    def __init__(self, root, settings_manager, file_analyzer, data_processor):
        self.root = root
        self.settings = settings_manager
        self.file_analyzer = file_analyzer
        self.data_processor = data_processor
        
        # משתנים
        self.rib_file = ""
        self.products_file = ""
        self.current_results = []
        
        # חלונות מנהל
        self.drawings_manager_window = None
        
        self._setup_window()
        self._create_widgets()
        self._load_initial_settings()
    
    def _setup_window(self):
        """הגדרת החלון הראשי"""
        self.root.title("FactorySync - ממיר אופטיטקס לאקסל")
        window_size = self.settings.get("app.window_size", "900x700")
        self.root.geometry(window_size)
        self.root.configure(bg='#f0f0f0')
        
        # מינימום גודל חלון
        self.root.minsize(800, 600)
        
        # מרכוז החלון
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (900 // 2)
        y = (self.root.winfo_screenheight() // 2) - (700 // 2)
        self.root.geometry(f"900x700+{x}+{y}")
    
    def _create_widgets(self):
        """יצירת הרכיבים"""
        # כותרת
        title_label = tk.Label(
            self.root,
            text="FactorySync - ממיר אופטיטקס לאקסל",
            font=('Arial', 18, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        title_label.pack(pady=15)
        
        # מסגרת קבצים
        self._create_files_section()
        
        # מסגרת אפשרויות
        self._create_options_section()
        
        # כפתורי פעולה
        self._create_action_buttons()
        
        # אזור תוצאות
        self._create_results_section()
        
        # שורת סטטוס
        self._create_status_bar()
    
    def _create_files_section(self):
        """יצירת מקטע בחירת קבצים"""
        files_frame = ttk.LabelFrame(self.root, text="בחירת קבצים", padding=15)
        files_frame.pack(fill="x", padx=20, pady=10)
        
        # קובץ אופטיטקס
        rib_frame = tk.Frame(files_frame)
        rib_frame.pack(fill="x", pady=8)
        
        tk.Label(
            rib_frame,
            text="קובץ אופטיטקס:",
            font=('Arial', 10, 'bold'),
            width=15,
            anchor="w"
        ).pack(side="left")
        
        self.rib_label = tk.Label(
            rib_frame,
            text="לא נבחר קובץ",
            bg="white",
            relief="sunken",
            width=60,
            anchor="w",
            padx=5
        )
        self.rib_label.pack(side="left", padx=10)
        
        tk.Button(
            rib_frame,
            text="📁 בחר קובץ",
            command=self._select_rib_file,
            bg='#3498db',
            fg='white',
            font=('Arial', 9, 'bold'),
            width=12
        ).pack(side="right")
        
        # קובץ מוצרים
        products_frame = tk.Frame(files_frame)
        products_frame.pack(fill="x", pady=8)
        
        tk.Label(
            products_frame,
            text="קובץ מוצרים:",
            font=('Arial', 10, 'bold'),
            width=15,
            anchor="w"
        ).pack(side="left")
        
        self.products_label = tk.Label(
            products_frame,
            text="לא נבחר קובץ",
            bg="white",
            relief="sunken",
            width=60,
            anchor="w",
            padx=5
        )
        self.products_label.pack(side="left", padx=10)
        
        tk.Button(
            products_frame,
            text="📁 בחר קובץ",
            command=self._select_products_file,
            bg='#3498db',
            fg='white',
            font=('Arial', 9, 'bold'),
            width=12
        ).pack(side="right")
    
    def _create_options_section(self):
        """יצירת מקטע אפשרויות"""
        options_frame = ttk.LabelFrame(self.root, text="אפשרויות", padding=15)
        options_frame.pack(fill="x", padx=20, pady=10)
        
        # אפשרויות עיבוד
        processing_frame = tk.Frame(options_frame)
        processing_frame.pack(fill="x", pady=5)
        
        self.tubular_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            processing_frame,
            text="טיפול אוטומטי ב-Layout Tubular (חלוקה ב-2)",
            variable=self.tubular_var,
            font=('Arial', 10)
        ).pack(anchor="w")
        
        self.only_positive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            processing_frame,
            text="הצג רק מידות עם כמות גדולה מ-0",
            variable=self.only_positive_var,
            font=('Arial', 10)
        ).pack(anchor="w")
    
    def _create_action_buttons(self):
        """יצירת כפתורי הפעולה"""
        buttons_frame = tk.Frame(self.root, bg='#f0f0f0')
        buttons_frame.pack(fill="x", padx=20, pady=15)
        
        # שורה ראשונה
        row1 = tk.Frame(buttons_frame, bg='#f0f0f0')
        row1.pack(fill="x", pady=5)
        
        tk.Button(
            row1,
            text="🔍 נתח קבצים",
            command=self._analyze_files,
            bg='#27ae60',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=15
        ).pack(side="left", padx=5)
        
        tk.Button(
            row1,
            text="💾 שמור כ-Excel",
            command=self._save_excel,
            bg='#e67e22',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=15
        ).pack(side="left", padx=5)
        
        tk.Button(
            row1,
            text="️ נקה הכל",
            command=self._clear_all,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=15
        ).pack(side="right", padx=5)
        
        # שורה שנייה
        row2 = tk.Frame(buttons_frame, bg='#f0f0f0')
        row2.pack(fill="x", pady=5)
        
        tk.Button(
            row2,
            text=" הוסף לטבלה מקומית",
            command=self._add_to_local_table,
            bg='#16a085',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=18
        ).pack(side="left", padx=5)
        
        tk.Button(
            row2,
            text="📁 מנהל ציורים",
            command=self._open_drawings_manager,
            bg='#2980b9',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=18
        ).pack(side="left", padx=5)
    
    def _create_results_section(self):
        """יצירת אזור התוצאות"""
        results_frame = ttk.LabelFrame(self.root, text="תוצאות וסטטוס", padding=10)
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.results_text = scrolledtext.ScrolledText(
            results_frame,
            height=15,
            font=('Consolas', 10),
            wrap=tk.WORD,
            bg='#f8f9fa',
            fg='#2c3e50'
        )
        self.results_text.pack(fill="both", expand=True)
    
    def _create_status_bar(self):
        """יצירת שורת הסטטוס"""
        self.status_label = tk.Label(
            self.root,
            text="מוכן לעבודה",
            bg='#34495e',
            fg='white',
            anchor='w',
            padx=15,
            font=('Arial', 10)
        )
        self.status_label.pack(fill="x", side="bottom")
    
    def _load_initial_settings(self):
        """טעינת הגדרות ראשוניות"""
        # טעינה אוטומטית של קובץ מוצרים
        if self.settings.get("app.auto_load_products", True):
            products_file = self.settings.get("app.products_file", "קובץ מוצרים.xlsx")
            if os.path.exists(products_file):
                self.products_file = os.path.abspath(products_file)
                self.products_label.config(text=os.path.basename(products_file))
                self._update_status(f"נטען קובץ מוצרים: {os.path.basename(products_file)}")
    
    # File Selection Methods
    def _select_rib_file(self):
        """בחירת קובץ אופטיטקס"""
        file_path = filedialog.askopenfilename(
            title="בחר קובץ אופטיטקס אקסל אקספורט",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.rib_file = file_path
            self.rib_label.config(text=os.path.basename(file_path))
            self._update_status(f"נבחר קובץ אופטיטקס: {os.path.basename(file_path)}")
    
    def _select_products_file(self):
        """בחירת קובץ מוצרים"""
        file_path = filedialog.askopenfilename(
            title="בחר קובץ רשימת מוצרים",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.products_file = file_path
            self.products_label.config(text=os.path.basename(file_path))
            self._update_status(f"נבחר קובץ מוצרים: {os.path.basename(file_path)}")
    
    # Analysis Methods
    def _analyze_files(self):
        """ניתוח הקבצים"""
        if not self.rib_file or not self.products_file:
            messagebox.showerror("שגיאה", "אנא בחר את שני הקבצים")
            return
        
        self._clear_results()
        self._update_status("מנתח קבצים...")
        
        # הרצה בחוט נפרד
        Thread(target=self._analyze_files_thread, daemon=True).start()
    
    def _analyze_files_thread(self):
        """ביצוע הניתוח בחוט נפרד"""
        try:
            self._log_message("=== התחלת ניתוח ===")
            
            # טעינת מיפוי מוצרים
            self._log_message("טוען מיפוי מוצרים...")
            if not self.file_analyzer.load_products_mapping(self.products_file):
                raise Exception("שגיאה בטעינת קובץ מוצרים")
            
            products_count = len(self.file_analyzer.product_mapping)
            self._log_message(f"✅ נטען מיפוי עבור {products_count} מוצרים")
            
            # ניתוח הקובץ
            self._log_message("מנתח קובץ אופטיטקס...")
            results = self.file_analyzer.analyze_file(
                self.rib_file,
                self.tubular_var.get(),
                self.only_positive_var.get()
            )
            
            if not results:
                self._log_message("❌ לא נמצאו נתונים מתאימים")
                self._update_status("לא נמצאו נתונים")
                return
            
            # מיון התוצאות
            self.current_results = self.file_analyzer.sort_results()
            
            # הצגת סיכום
            summary = self.file_analyzer.get_analysis_summary()
            self._log_message(f"✅ נוצרה טבלה עם {summary['total_records']} רשומות")
            
            # הצגת מוצרים שנמצאו
            found_products = self.file_analyzer.get_products_found()
            if found_products:
                self._log_message("\n📦 מוצרים שנמצאו:")
                for file_name, product_name in found_products:
                    self._log_message(f"   {file_name} → {product_name}")
            
            # הצגת תוצאות מפורטות
            self._display_detailed_results()
            
            # הצגת סטטיסטיקות
            self._display_statistics(summary)
            
            self._update_status("הניתוח הושלם בהצלחה!")
            
        except Exception as e:
            error_msg = f"❌ שגיאה בניתוח: {str(e)}"
            self._log_message(error_msg)
            self._update_status("שגיאה בניתוח")
            messagebox.showerror("שגיאה", str(e))
    
    def _display_detailed_results(self):
        """הצגת תוצאות מפורטות"""
        self._log_message("\n=== תוצאות הניתוח ===")
        
        current_product = None
        for result in self.current_results:
            if current_product != result['שם המוצר']:
                current_product = result['שם המוצר']
                self._log_message(f"\n📦 {current_product}:")
                self._log_message("-" * 60)
            
            quantity_text = f"{result['כמות']}"
            if result['כמות מקורית'] != result['כמות']:
                quantity_text += f" (מקורי: {result['כמות מקורית']})"
            
            self._log_message(f"   מידה {result['מידה']:>8}: {quantity_text:>10} - {result['הערה']}")
    
    def _display_statistics(self, summary):
        """הצגת סטטיסטיקות"""
        self._log_message("\n" + "=" * 70)
        self._log_message(f"\n=== סיכום ===")
        self._log_message(f"מוצרים: {summary['unique_products']}")
        self._log_message(f"מידות שונות: {summary['unique_sizes']}")
        self._log_message(f"סך רשומות: {summary['total_records']}")
        self._log_message(f"סך כמויות: {summary['total_quantity']:.1f}")
        
        if summary['is_tubular']:
            self._log_message("🔄 הכמויות חולקו ב-2 בגלל Layout: Tubular")
    
    # Export Methods
    def _save_excel(self):
        """שמירה כ-Excel"""
        if not self.current_results:
            messagebox.showwarning("אזהרה", "אין נתונים לשמירה. אנא בצע ניתוח תחילה.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="שמור כ-Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.data_processor.export_to_excel(self.current_results, file_path)
                self._log_message(f"📁 הקובץ נשמר בהצלחה: {file_path}")
                self._update_status(f"נשמר: {os.path.basename(file_path)}")
                messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה!")
            except Exception as e:
                messagebox.showerror("שגיאה", str(e))
    
    def _add_to_local_table(self):
        """הוספה לטבלה המקומית"""
        if not self.current_results:
            messagebox.showwarning("אזהרה", "אין נתונים להוספה. אנא בצע ניתוח תחילה.")
            return
        
        try:
            record_id = self.data_processor.add_to_local_table(self.current_results, self.rib_file)
            
            self._log_message(f"\n✅ הציור נוסף לטבלה המקומית!")
            self._log_message(f"ID רשומה חדשה: {record_id}")
            
            file_name = os.path.splitext(os.path.basename(self.rib_file))[0] if self.rib_file else 'לא ידוע'
            total_quantity = sum(r['כמות'] for r in self.current_results)
            
            self._log_message(f"שם הקובץ: {file_name}")
            self._log_message(f"סך כמויות: {total_quantity}")
            
            self._update_status("נוסף לטבלה המקומית")
            messagebox.showinfo("הצלחה", f"הציור נוסף בהצלחה לטבלה המקומית!\nID: {record_id}")
            
        except Exception as e:
            error_msg = str(e)
            self._log_message(f"❌ שגיאה בהוספה: {error_msg}")
            messagebox.showerror("שגיאה", error_msg)
    
    # Window Management
    def _open_drawings_manager(self):
        """פתיחת מנהל הציורים"""
        if not self.drawings_manager_window:
            from .drawings_manager import DrawingsManagerWindow
            self.drawings_manager_window = DrawingsManagerWindow(
                self.root,
                self.data_processor
            )
        
        self.drawings_manager_window.show()
    
    # Utility Methods
    def _update_status(self, message):
        """עדכון שורת הסטטוס"""
        self.status_label.config(text=message)
        self.root.update()
    
    def _log_message(self, message):
        """הוספת הודעה לאזור התוצאות"""
        self.results_text.insert(tk.END, message + "\n")
        self.results_text.see(tk.END)
        self.root.update()
    
    def _clear_results(self):
        """ניקוי אזור התוצאות"""
        self.results_text.delete(1.0, tk.END)
    
    def _clear_all(self):
        """ניקוי כל הנתונים"""
        self.rib_file = ""
        self.current_results = []
        self.rib_label.config(text="לא נבחר קובץ")
        self._clear_results()
        
        # טעינה מחדש של קובץ מוצרים אם קיים
        if self.settings.get("app.auto_load_products", True):
            products_file = self.settings.get("app.products_file", "קובץ מוצרים.xlsx")
            if os.path.exists(products_file):
                self.products_file = os.path.abspath(products_file)
                self.products_label.config(text=os.path.basename(products_file))
                self._update_status(f"נטען קובץ מוצרים: {os.path.basename(products_file)}")
            else:
                self.products_file = ""
                self.products_label.config(text="לא נבחר קובץ")
                self._update_status("מוכן לעבודה")
        else:
            self._update_status("מוכן לעבודה")
