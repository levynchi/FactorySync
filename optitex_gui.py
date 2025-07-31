"""
ממשק גרפי (GUI) לתוכנת ממיר אקספ        self.config = {
            "airtable": {
                "api_key": "",
                "base_id": "",
                "table_id": "tblC0hR3gZFXxstbM"
            }
        }אופטיטקס לאקסל נקי
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import json
from threading import Thread
from pyairtable import Api

class OptitexAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ממיר אקספורט של אופטיטקס לאקסל נקי - ממשק גרפי")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # משתנים
        self.rib_file = ""
        self.products_file = ""
        self.output_df = None
        
        # טעינת הגדרות Airtable מקובץ config
        self.config = self.load_config()
        
        # בדיקה אם קובץ המוצרים קיים בתיקייה ובחירתו כברירת מחדל
        products_file_path = "קובץ מוצרים.xlsx"
        if os.path.exists(products_file_path):
            self.products_file = os.path.abspath(products_file_path)
        
        self.create_widgets()
    
    def load_config(self):
        """טעינת הגדרות מקובץ config.json"""
        config_file = "config.json"
        default_config = {
            "airtable": {
                "api_key": "",
                "base_id": "",
                "table_name": "ציורים-כולל דגם וכמות מדגם"
            }
        }
        
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                # יצירת קובץ הגדרות ברירת מחדל
                self.save_config(default_config)
                return default_config
        except Exception as e:
            print(f"שגיאה בטעינת הגדרות: {e}")
            return default_config
    
    def save_config(self, config=None):
        """שמירת הגדרות לקובץ config.json"""
        if config is None:
            config = self.config
        
        try:
            with open("config.json", 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"שגיאה בשמירת הגדרות: {e}")
    
    def create_widgets(self):
        # כותרת
        title_label = tk.Label(
            self.root, 
            text="ממיר אקספורט של אופטיטקס לאקסל נקי", 
            font=('Arial', 16, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        title_label.pack(pady=10)
        
        # מסגרת לקבצים
        files_frame = ttk.LabelFrame(self.root, text="בחירת קבצים", padding=10)
        files_frame.pack(fill="x", padx=20, pady=10)
        
        # בחירת קובץ RIB
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
        
        # מסגרת לאפשרויות
        options_frame = ttk.LabelFrame(self.root, text="אפשרויות", padding=10)
        options_frame.pack(fill="x", padx=20, pady=10)
        
        # הגדרות Airtable
        airtable_frame = ttk.LabelFrame(options_frame, text="הגדרות Airtable", padding=5)
        airtable_frame.pack(fill="x", pady=5)
        
        # סטטוס חיבור
        status_frame = tk.Frame(airtable_frame)
        status_frame.pack(fill="x", pady=2)
        
        api_key = self.config.get("airtable", {}).get("api_key", "")
        base_id = self.config.get("airtable", {}).get("base_id", "")
        
        if api_key and base_id:
            status_text = "🟢 מחובר לאייר טייבל"
            status_color = "green"
        elif api_key or base_id:
            status_text = "🟡 הגדרות חלקיות"
            status_color = "orange"
        else:
            status_text = "🔴 לא מוגדר"
            status_color = "red"
        
        tk.Label(status_frame, text="סטטוס:", font=('Arial', 9, 'bold')).pack(side="left")
        self.airtable_status = tk.Label(status_frame, text=status_text, fg=status_color, font=('Arial', 9))
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
        
        # תיבת סימון לטיפול ב-Tubular
        self.tubular_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame,
            text="טיפול אוטומטי ב-Layout Tubular (חלוקה ב-2)",
            variable=self.tubular_var
        ).pack(anchor="w")
        
        # תיבת סימון להצגת רק כמויות > 0
        self.only_positive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            options_frame,
            text="הצג רק מידות עם כמות גדולה מ-0",
            variable=self.only_positive_var
        ).pack(anchor="w")
        
        # כפתורי פעולה
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
            text="🗑️ נקה הכל",
            command=self.clear_all,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 12, 'bold'),
            height=2
        ).pack(side="right", padx=5)
        
        # אזור תוצאות
        results_frame = ttk.LabelFrame(self.root, text="תוצאות", padding=10)
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # טקסט תוצאות עם גלילה
        self.results_text = scrolledtext.ScrolledText(
            results_frame,
            height=15,
            font=('Courier New', 10),
            wrap=tk.WORD
        )
        self.results_text.pack(fill="both", expand=True)
        
        # שורת סטטוס
        self.status_label = tk.Label(
            self.root,
            text="מוכן לעבודה",
            bg='#34495e',
            fg='white',
            anchor='w',
            padx=10
        )
        self.status_label.pack(fill="x", side="bottom")
        
        # עדכון תצוגת קובץ המוצרים אם נמצא
        self.update_products_display()
    
    def select_rib_file(self):
        """בחירת קובץ אופטיטקס אקסל אקספורט"""
        file_path = filedialog.askopenfilename(
            title="בחר קובץ אופטיטקס אקסל אקספורט",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.rib_file = file_path
            self.rib_label.config(text=os.path.basename(file_path))
            self.update_status(f"נבחר קובץ אופטיטקס אקסל אקספורט: {os.path.basename(file_path)}")
    
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
    
    def update_status(self, message):
        """עדכון שורת הסטטוס"""
        self.status_label.config(text=message)
        self.root.update()
    
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
        config_window.grab_set()  # חלון מודלי
        
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
        
        # API Key
        tk.Label(main_frame, text="API Key:", bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor="w")
        api_key_entry = tk.Entry(main_frame, width=60, show="*")
        api_key_entry.pack(fill="x", pady=(0, 10))
        api_key_entry.insert(0, self.config.get("airtable", {}).get("api_key", ""))
        
        # Base ID
        tk.Label(main_frame, text="Base ID:", bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor="w")
        base_id_entry = tk.Entry(main_frame, width=60)
        base_id_entry.pack(fill="x", pady=(0, 10))
        base_id_entry.insert(0, self.config.get("airtable", {}).get("base_id", ""))
        
        # Table ID
        tk.Label(main_frame, text="Table ID:", bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor="w")
        table_id_entry = tk.Entry(main_frame, width=60)
        table_id_entry.pack(fill="x", pady=(0, 15))
        table_id_entry.insert(0, self.config.get("airtable", {}).get("table_id", "tblC0hR3gZFXxstbM"))
        
        # הערה
        note_text = "💡 עזרה: ראה את הקובץ 'הוראות_Airtable.md' להוראות מפורטות"
        tk.Label(main_frame, text=note_text, bg='#f0f0f0', fg='#666', font=('Arial', 9)).pack(pady=5)
        
        # כפתורים
        buttons_frame = tk.Frame(config_window, bg='#f0f0f0')
        buttons_frame.pack(fill="x", padx=20, pady=10)
        
        def save_and_close():
            # עדכון ההגדרות
            self.config["airtable"]["api_key"] = api_key_entry.get().strip()
            self.config["airtable"]["base_id"] = base_id_entry.get().strip()
            self.config["airtable"]["table_id"] = table_id_entry.get().strip()
            
            # שמירה לקובץ
            self.save_config()
            
            # עדכון הסטטוס
            self.update_airtable_status()
            
            # סגירת החלון
            config_window.destroy()
            
            messagebox.showinfo("הצלחה", "ההגדרות נשמרו בהצלחה!")
        
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
        """עדכון סטטוס חיבור האייר טייבל"""
        api_key = self.config.get("airtable", {}).get("api_key", "")
        base_id = self.config.get("airtable", {}).get("base_id", "")
        
        if api_key and base_id:
            status_text = "🟢 מחובר לאייר טייבל"
            status_color = "green"
        elif api_key or base_id:
            status_text = "🟡 הגדרות חלקיות"
            status_color = "orange"
        else:
            status_text = "🔴 לא מוגדר"
            status_color = "red"
        
        self.airtable_status.config(text=status_text, fg=status_color)
    
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
        
        # ניקוי תוצאות קודמות
        self.clear_results()
        self.update_status("מנתח קבצים...")
        
        # הרצה בחוט נפרד כדי לא לחסום את הממשק
        Thread(target=self._analyze_files_thread, daemon=True).start()
    
    def _analyze_files_thread(self):
        """ביצוע הניתוח בחוט נפרד"""
        try:
            self.log_message("=== התחלת ניתוח ===")
            
            # טעינת מיפוי מוצרים
            self.log_message("טוען מיפוי מוצרים...")
            product_mapping = {}
            df_products = pd.read_excel(self.products_file)
            for _, row in df_products.iterrows():
                if pd.notna(row['product name']):
                    product_mapping[row['file name']] = row['product name']
            
            self.log_message(f"✅ נטען מיפוי עבור {len(product_mapping)} מוצרים")
            
            # קריאת קובץ RIB
            self.log_message("קורא קובץ אופטיטקס אקסל אקספורט...")
            df = pd.read_excel(self.rib_file, header=None)
            
            # בדיקת Tubular
            is_tubular = False
            if self.tubular_var.get():
                for i, row in df.iterrows():
                    if (pd.notna(row.iloc[0]) and row.iloc[0] == 'Layout' and 
                        pd.notna(row.iloc[2]) and row.iloc[2] == 'Tubular'):
                        is_tubular = True
                        self.log_message("🔍 נמצא Layout: Tubular - הכמויות יחולקו ב-2")
                        break
                
                if not is_tubular:
                    self.log_message("📋 לא נמצא Layout: Tubular - הכמויות יישארו כמו שהן")
            
            # חיפוש נתונים
            results = []
            current_file_name = None
            product_name = None
            
            for i, row in df.iterrows():
                # חיפוש שם קובץ
                if pd.notna(row.iloc[0]) and row.iloc[0] == 'Style File Name:':
                    if pd.notna(row.iloc[2]):
                        current_file_name = os.path.basename(row.iloc[2])
                        product_name = product_mapping.get(current_file_name)
                        if product_name:
                            self.log_message(f"📦 נמצא מוצר: {current_file_name} -> {product_name}")
                
                # חיפוש טבלת מידות
                elif (pd.notna(row.iloc[0]) and row.iloc[0] == 'Size name' and 
                      pd.notna(row.iloc[1]) and row.iloc[1] == 'Order' and product_name):
                    
                    j = i + 1
                    while j < len(df) and pd.notna(df.iloc[j, 0]) and pd.notna(df.iloc[j, 1]):
                        size_name = df.iloc[j, 0]
                        quantity = int(df.iloc[j, 1])
                        
                        if size_name not in ['Style File Name:', 'Size name']:
                            # טיפול ב-Tubular
                            original_quantity = quantity
                            if is_tubular and quantity > 0:
                                quantity = quantity / 2
                                quantity = int(quantity) if quantity == int(quantity) else round(quantity, 1)
                            
                            # הוספה לתוצאות
                            if not self.only_positive_var.get() or quantity > 0:
                                results.append({
                                    'שם המוצר': product_name,
                                    'מידה': size_name,
                                    'כמות': quantity,
                                    'כמות מקורית': original_quantity if is_tubular else quantity,
                                    'הערה': 'חולק ב-2 (Tubular)' if is_tubular and original_quantity > 0 else 'רגיל'
                                })
                        elif size_name in ['Style File Name:', 'Size name']:
                            break
                        j += 1
            
            # יצירת טבלה
            if results:
                self.output_df = pd.DataFrame(results)
                
                # סידור לפי קבוצת שם מוצר ובתוך כל קבוצה לפי מידה בסדר עולה
                def sort_size(size):
                    """פונקציה למיון מידות בסדר נומרי/אלפביתי נכון"""
                    size_str = str(size)
                    
                    # טיפול במידות חודשים (0-3, 3-6, 6-12, 12-18, 18-24, 24-30)
                    if '-' in size_str and any(char.isdigit() for char in size_str):
                        try:
                            # חילוץ המספר הראשון מהמידה
                            first_num = size_str.split('-')[0]
                            if first_num.isdigit():
                                return (0, int(first_num))  # מידות חודשים - קבוצה 0
                        except:
                            pass
                    
                    try:
                        # אם המידה היא מספר, נמיין לפי ערך נומרי
                        return (1, float(size))  # מידות נומריות רגילות - קבוצה 1
                    except:
                        # אם המידה היא טקסט, נמיין לפי אלפבית
                        return (2, str(size))  # מידות טקסטואליות - קבוצה 2
                
                # יצירת עמודה זמנית למיון
                self.output_df['_sort_size'] = self.output_df['מידה'].apply(sort_size)
                
                # מיון לפי שם מוצר ואז לפי מידה
                self.output_df = self.output_df.sort_values(['שם המוצר', '_sort_size'])
                
                # הסרת עמודת המיון הזמנית
                self.output_df = self.output_df.drop('_sort_size', axis=1)
                
                self.log_message(f"✅ נוצרה טבלה עם {len(self.output_df)} רשומות (מסודרת לפי מוצר ומידה)")
                
                # הצגת תוצאות מסודרות לפי קבוצות
                self.log_message("\n=== תוצאות הניתוח ===")
                
                # הצגה מקובצת לפי מוצר
                current_product = None
                for _, row in self.output_df.iterrows():
                    if current_product != row['שם המוצר']:
                        current_product = row['שם המוצר']
                        self.log_message(f"\n📦 {current_product}:")
                        self.log_message("-" * 50)
                    
                    # הצגת השורה
                    quantity_text = f"{row['כמות']}"
                    if row['כמות מקורית'] != row['כמות']:
                        quantity_text += f" (מקורי: {row['כמות מקורית']})"
                    
                    self.log_message(f"   מידה {row['מידה']:>6}: {quantity_text:>8} - {row['הערה']}")
                
                self.log_message("\n" + "=" * 60)
                
                # סטטיסטיקות
                self.log_message(f"\n=== סיכום ===")
                self.log_message(f"מוצרים: {self.output_df['שם המוצר'].nunique()}")
                self.log_message(f"מידות: {self.output_df['מידה'].nunique()}")
                self.log_message(f"סך כמויות: {self.output_df['כמות'].sum()}")
                if is_tubular:
                    self.log_message("🔄 הכמויות חולקו ב-2 בגלל Layout: Tubular")
                
                self.update_status("הניתוח הושלם בהצלחה!")
            else:
                self.log_message("❌ לא נמצאו נתונים")
                self.update_status("לא נמצאו נתונים")
                
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
                self.output_df.to_excel(file_path, index=False)
                self.log_message(f"📁 הקובץ נשמר בהצלחה: {file_path}")
                self.update_status(f"נשמר: {os.path.basename(file_path)}")
                messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה:\n{file_path}")
            except Exception as e:
                messagebox.showerror("שגיאה", f"שגיאה בשמירה: {str(e)}")
    
    def upload_to_airtable(self):
        """העלאת הנתונים ישירות לאייר טייבל"""
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("אזהרה", "אין נתונים להעלאה. אנא בצע ניתוח תחילה.")
            return
        
        # בדיקת הגדרות API מהקובץ
        api_key = self.config.get("airtable", {}).get("api_key", "").strip()
        base_id = self.config.get("airtable", {}).get("base_id", "").strip()
        table_id = self.config.get("airtable", {}).get("table_id", "tblC0hR3gZFXxstbM")
        
        if not api_key:
            messagebox.showerror("שגיאה", "API Key לא מוגדר.\nאנא לחץ על 'הגדר חיבור' כדי להגדיר את הפרטים.")
            return
        
        if not base_id:
            messagebox.showerror("שגיאה", "Base ID לא מוגדר.\nאנא לחץ על 'הגדר חיבור' כדי להגדיר את הפרטים.")
            return
        
        try:
            self.log_message("\n=== התחלת העלאה לאייר טייבל ===")
            
            # חיבור לאייר טייבל
            api = Api(api_key)
            table = api.table(base_id, table_id)
            
            # הכנת הנתונים לאייר טייבל
            airtable_record = {
                'שם הקובץ': os.path.splitext(os.path.basename(self.rib_file))[0] if self.rib_file else '',
                'אורך הגיזרה': 0.0,  # ערך ברירת מחדל
                'אורך גיזרה כולל קיפול': 0.0,  # ערך ברירת מחדל
            }
            
            # הוספת הדגמים והכמויות
            model_index = 1
            for _, row in self.output_df.iterrows():
                if model_index > 21:  # מקסימום 21 דגמים באייר טייבל
                    break
                    
                product_name = row['שם המוצר']
                size = row['מידה']
                quantity = row['כמות']
                
                # יצירת מחרוזת דגם משולבת
                model_name = f"{product_name} {size}"
                
                airtable_record[f'דגם {model_index}'] = model_name
                airtable_record[f'כמות דגם {model_index}'] = int(quantity) if quantity == int(quantity) else quantity
                
                model_index += 1
            
            # העלאת הרשומה לאייר טייבל
            self.log_message("מעלה נתונים לאייר טייבל...")
            created_record = table.create(airtable_record)
            
            self.log_message(f"✅ הנתונים הועלו בהצלחה לאייר טייבל!")
            self.log_message(f"ID של הרשומה החדשה: {created_record['id']}")
            
            # הצגת פרטי ההעלאה
            self.log_message(f"\nשם הקובץ: {airtable_record['שם הקובץ']}")
            self.log_message(f"מספר דגמים שהועלו: {model_index-1}")
            
            for i in range(1, model_index):
                model = airtable_record.get(f'דגם {i}', '')
                quantity = airtable_record.get(f'כמות דגם {i}', '')
                if model:
                    self.log_message(f"  דגם {i}: {model} - כמות: {quantity}")
            
            self.update_status("העלאה לאייר טייבל הושלמה בהצלחה!")
            messagebox.showinfo("הצלחה", f"הנתונים הועלו בהצלחה לאייר טייבל!\nID רשומה: {created_record['id']}")
            
        except Exception as e:
            error_msg = f"שגיאה בהעלאה לאייר טייבל: {str(e)}"
            self.log_message(f"❌ {error_msg}")
            messagebox.showerror("שגיאה", error_msg)
            
            # הצעת פתרונות לבעיות נפוצות
            if "Invalid API key" in str(e):
                messagebox.showinfo("עזרה", "API Key לא תקין. אנא בדוק שהמפתח נכון ופעיל.")
            elif "Base not found" in str(e):
                messagebox.showinfo("עזרה", "Base ID לא נמצא. אנא בדוק שה-Base ID נכון.")
            elif "Table not found" in str(e):
                messagebox.showinfo("עזרה", "הטבלה לא נמצאה. אנא בדוק ששם הטבלה נכון.")
    
    def convert_to_airtable(self):
        """המרת התוצאות לפורמט אייר טייבל"""
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("אזהרה", "אין נתונים להמרה. אנא בצע ניתוח תחילה.")
            return
        
        try:
            self.log_message("\n=== התחלת המרה לפורמט אייר טייבל ===")
            
            # יצירת מבנה נתונים עבור אייר טייבל
            airtable_data = {}
            
            # מעבר על כל השורות ויצירת רשימה של דגמים עבור כל קובץ
            for _, row in self.output_df.iterrows():
                product_name = row['שם המוצר']
                size = row['מידה']
                quantity = row['כמות']
                
                # יצירת מחרוזת דגם משולבת (שם מוצר + מידה)
                model_name = f"{product_name} {size}"
                
                # אם זה הרשומה הראשונה, יוצרים מבנה ריק
                if not airtable_data:
                    airtable_data = {
                        'שם הקובץ': '',  # ייקבע לפי קובץ האקספורט
                        'אורך הגיזרה': '',
                        'אורך גיזרה כולל קיפול': '',
                    }
                    # יצירת עמודות דגמים (עד 21 כמו באייר טייבל)
                    for i in range(1, 22):
                        airtable_data[f'דגם {i}'] = ''
                        airtable_data[f'כמות דגם {i}'] = ''
                
            # מילוי הדגמים והכמויות
            model_index = 1
            for _, row in self.output_df.iterrows():
                if model_index > 21:  # מקסימום 21 דגמים באייר טייבל
                    break
                    
                product_name = row['שם המוצר']
                size = row['מידה']
                quantity = row['כמות']
                
                # יצירת מחרוזת דגם משולבת
                model_name = f"{product_name} {size}"
                
                airtable_data[f'דגם {model_index}'] = model_name
                airtable_data[f'כמות דגם {model_index}'] = quantity
                
                model_index += 1
            
            # קביעת שם הקובץ מהקובץ המקורי
            if self.rib_file:
                airtable_data['שם הקובץ'] = os.path.splitext(os.path.basename(self.rib_file))[0]
            
            # יצירת DataFrame חדש עבור אייר טייבל
            airtable_df = pd.DataFrame([airtable_data])
            
            self.log_message(f"✅ נוצרה טבלת אייר טייבל עם {model_index-1} דגמים")
            
            # הצגת התוצאות
            self.log_message("\n=== תוצאות ההמרה לאייר טייבל ===")
            self.log_message(f"שם הקובץ: {airtable_data['שם הקובץ']}")
            self.log_message("דגמים וכמויות:")
            
            for i in range(1, model_index):
                model = airtable_data[f'דגם {i}']
                quantity = airtable_data[f'כמות דגם {i}']
                if model:
                    self.log_message(f"  דגם {i}: {model} - כמות: {quantity}")
            
            # שמירת הקובץ
            file_path = filedialog.asksaveasfilename(
                title="שמור כקובץ אייר טייבל",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if file_path:
                if file_path.endswith('.csv'):
                    airtable_df.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    airtable_df.to_excel(file_path, index=False)
                
                self.log_message(f"📁 קובץ אייר טייבל נשמר בהצלחה: {file_path}")
                self.update_status(f"נשמר קובץ אייר טייבל: {os.path.basename(file_path)}")
                messagebox.showinfo("הצלחה", f"קובץ אייר טייבל נשמר בהצלחה:\n{file_path}")
            
        except Exception as e:
            self.log_message(f"❌ שגיאה בהמרה לאייר טייבל: {str(e)}")
            messagebox.showerror("שגיאה", f"שגיאה בהמרה לאייר טייבל: {str(e)}")
    
    def clear_all(self):
        """ניקוי כל הנתונים"""
        self.rib_file = ""
        self.output_df = None
        self.rib_label.config(text="לא נבחר")
        self.clear_results()
        
        # בדיקה מחדש אם קובץ המוצרים קיים בתיקייה
        products_file_path = "קובץ מוצרים.xlsx"
        if os.path.exists(products_file_path):
            self.products_file = os.path.abspath(products_file_path)
            self.products_label.config(text=os.path.basename(self.products_file))
            self.update_status(f"נמצא קובץ מוצרים: {os.path.basename(self.products_file)}")
        else:
            self.products_file = ""
            self.products_label.config(text="לא נבחר")
            self.update_status("מוכן לעבודה")

def main():
    root = tk.Tk()
    app = OptitexAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
