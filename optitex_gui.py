"""
ממשק גרפי (GUI) לתוכנת ממיר אקספורט של אופטיטקס לאקסל נקי
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
from threading import Thread

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
        
        # בדיקה אם קובץ המוצרים קיים בתיקייה ובחירתו כברירת מחדל
        products_file_path = "קובץ מוצרים.xlsx"
        if os.path.exists(products_file_path):
            self.products_file = os.path.abspath(products_file_path)
        
        self.create_widgets()
    
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
            text="📊 המר לאייר טייבל",
            command=self.convert_to_airtable,
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
    
    def add_to_local_table(self):
        """הוספת הנתונים הנוכחיים לטבלה המקומית"""
        if self.output_df is None or self.output_df.empty:
            messagebox.showwarning("אזהרה", "אין נתונים להוספה. אנא בצע ניתוח תחילה.")
            return
        
        try:
            # יצירת רשומה חדשה
            from datetime import datetime
            
            record = {
                'id': len(self.drawings_data) + 1,
                'שם הקובץ': os.path.splitext(os.path.basename(self.rib_file))[0] if self.rib_file else 'לא ידוע',
                'תאריך יצירה': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'מוצרים': [],
                'סך כמויות': 0
            }
            
            # הוספת מידע על המוצרים
            for product_name in self.output_df['שם המוצר'].unique():
                product_data = self.output_df[self.output_df['שם המוצר'] == product_name]
                
                product_info = {
                    'שם המוצר': product_name,
                    'מידות': []
                }
                
                for _, row in product_data.iterrows():
                    size_info = {
                        'מידה': row['מידה'],
                        'כמות': row['כמות'],
                        'הערה': row['הערה']
                    }
                    product_info['מידות'].append(size_info)
                    record['סך כמויות'] += row['כמות']
                
                record['מוצרים'].append(product_info)
            
            # הוספה לרשימה ושמירה
            self.drawings_data.append(record)
            self.save_drawings_data()
            
            self.log_message(f"\n✅ הציור נוסף לטבלה המקומית!")
            self.log_message(f"ID רשומה: {record['id']}")
            self.log_message(f"שם הקובץ: {record['שם הקובץ']}")
            self.log_message(f"מוצרים: {len(record['מוצרים'])}")
            self.log_message(f"סך כמויות: {record['סך כמויות']}")
            
            messagebox.showinfo("הצלחה", f"הציור נוסף בהצלחה לטבלה המקומית!\nID רשומה: {record['id']}")
            
        except Exception as e:
            error_msg = f"שגיאה בהוספה לטבלה המקומית: {str(e)}"
            self.log_message(f"❌ {error_msg}")
            messagebox.showerror("שגיאה", error_msg)
    
    def open_drawings_manager(self):
        """פתיחת חלון מנהל הציורים"""
        manager_window = tk.Toplevel(self.root)
        manager_window.title("מנהל ציורים")
        manager_window.geometry("1000x700")
        manager_window.configure(bg='#f0f0f0')
        
        # כותרת
        tk.Label(
            manager_window,
            text="מנהל ציורים - טבלה מקומית",
            font=('Arial', 16, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50'
        ).pack(pady=10)
        
        # מסגרת לכפתורי פעולה
        actions_frame = tk.Frame(manager_window, bg='#f0f0f0')
        actions_frame.pack(fill="x", padx=20, pady=5)
        
        tk.Button(
            actions_frame,
            text="🔄 רענן",
            command=lambda: self.refresh_drawings_table(table_frame),
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold')
        ).pack(side="left", padx=5)
        
        tk.Button(
            actions_frame,
            text="📊 ייצא לאקסל",
            command=self.export_drawings_to_excel,
            bg='#27ae60',
            fg='white',
            font=('Arial', 10, 'bold')
        ).pack(side="left", padx=5)
        
        # כפתור מחיקת רשומה נבחרת
        tk.Button(
            actions_frame,
            text="🗑️ מחק נבחר",
            command=lambda: self.delete_selected_drawing(table_frame),
            bg='#e67e22',
            fg='white',
            font=('Arial', 10, 'bold')
        ).pack(side="left", padx=5)
        
        tk.Button(
            actions_frame,
            text="🗑️ מחק הכל",
            command=lambda: self.clear_all_drawings(table_frame),
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold')
        ).pack(side="right", padx=5)
        
        # הודעת עזרה
        help_frame = tk.Frame(manager_window, bg='#f0f0f0')
        help_frame.pack(fill="x", padx=20, pady=5)
        
        tk.Label(
            help_frame,
            text="💡 עזרה: בחר ציור בטבלה ולחץ 'מחק נבחר' למחיקה, או לחץ קליק ימני לתפריט נוסף",
            bg='#f0f0f0',
            fg='#666',
            font=('Arial', 9)
        ).pack()
        
        # מסגרת לטבלה
        table_frame = tk.Frame(manager_window, bg='#f0f0f0')
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # יצירת הטבלה
        self.create_drawings_table(table_frame)
    
    def create_drawings_table(self, parent_frame):
        """יצירת טבלת הציורים"""
        # ניקוי המסגרת
        for widget in parent_frame.winfo_children():
            widget.destroy()
        
        # יצירת Treeview עם גלילה
        tree_frame = tk.Frame(parent_frame)
        tree_frame.pack(fill="both", expand=True)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Treeview
        columns = ("ID", "שם הקובץ", "תאריך יצירה", "מוצרים", "סך כמויות", "פעולות")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", 
                           yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # הגדרת כותרות
        tree.heading("ID", text="ID")
        tree.heading("שם הקובץ", text="שם הקובץ")
        tree.heading("תאריך יצירה", text="תאריך יצירה")
        tree.heading("מוצרים", text="מוצרים")
        tree.heading("סך כמויות", text="סך כמויות")
        tree.heading("פעולות", text="פעולות")
        
        # הגדרת רוחב עמודות
        tree.column("ID", width=50, anchor="center")
        tree.column("שם הקובץ", width=200, anchor="w")
        tree.column("תאריך יצירה", width=150, anchor="center")
        tree.column("מוצרים", width=100, anchor="center")
        tree.column("סך כמויות", width=100, anchor="center")
        tree.column("פעולות", width=150, anchor="center")
        
        # הוספת נתונים
        for record in self.drawings_data:
            products_count = len(record.get('מוצרים', []))
            total_quantity = record.get('סך כמויות', 0)
            
            tree.insert("", "end", values=(
                record.get('id', ''),
                record.get('שם הקובץ', ''),
                record.get('תאריך יצירה', ''),
                products_count,
                total_quantity,
                "לחץ כפול לפרטים"
            ))
        
        # קישור אירועים
        tree.bind("<Double-1>", lambda e: self.show_record_details(tree))
        tree.bind("<Button-3>", lambda e: self.show_context_menu(e, tree))
        
        # סידור הרכיבים
        tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        v_scrollbar.config(command=tree.yview)
        h_scrollbar.config(command=tree.xview)
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # שמירת הפניה לטבלה
        self.current_tree = tree
    
    def delete_selected_drawing(self, table_frame):
        """מחיקת ציור נבחר מהטבלה"""
        if not hasattr(self, 'current_tree') or not self.current_tree:
            messagebox.showwarning("אזהרה", "לא נמצאה טבלה פעילה")
            return
        
        selection = self.current_tree.selection()
        if not selection:
            messagebox.showwarning("אזהרה", "אנא בחר ציור למחיקה")
            return
        
        # קבלת פרטי הרשומה הנבחרת
        item = self.current_tree.item(selection[0])
        record_id = int(item['values'][0])
        file_name = item['values'][1]
        
        # בקשת אישור מהמשתמש
        if messagebox.askyesno("אישור מחיקה", 
                              f"האם אתה בטוח שברצונך למחוק את הציור?\n\n"
                              f"שם הקובץ: {file_name}\n"
                              f"ID: {record_id}\n\n"
                              f"פעולה זו לא ניתנת לביטול!"):
            try:
                # מחיקה מהרשימה
                original_count = len(self.drawings_data)
                self.drawings_data = [r for r in self.drawings_data if r.get('id') != record_id]
                
                if len(self.drawings_data) < original_count:
                    # שמירה
                    self.save_drawings_data()
                    
                    # רענון הטבלה
                    self.refresh_drawings_table(table_frame)
                    
                    messagebox.showinfo("הצלחה", f"הציור '{file_name}' נמחק בהצלחה!")
                else:
                    messagebox.showerror("שגיאה", "הציור לא נמצא או לא נמחק")
                    
            except Exception as e:
                messagebox.showerror("שגיאה", f"שגיאה במחיקת הציור: {str(e)}")
    
    def show_record_details(self, tree):
        """הצגת פרטי רשומה"""
        selection = tree.selection()
        if not selection:
            return
        
        item = tree.item(selection[0])
        record_id = int(item['values'][0])
        
        # חיפוש הרשומה
        record = None
        for r in self.drawings_data:
            if r.get('id') == record_id:
                record = r
                break
        
        if not record:
            messagebox.showerror("שגיאה", "רשומה לא נמצאה")
            return
        
        # חלון פרטים
        details_window = tk.Toplevel(self.root)
        details_window.title(f"פרטי ציור - {record.get('שם הקובץ', '')}")
        details_window.geometry("800x600")
        details_window.configure(bg='#f0f0f0')
        
        # כותרת
        tk.Label(
            details_window,
            text=f"פרטי ציור: {record.get('שם הקובץ', '')}",
            font=('Arial', 14, 'bold'),
            bg='#f0f0f0'
        ).pack(pady=10)
        
        # מידע כללי
        info_frame = tk.Frame(details_window, bg='#f0f0f0')
        info_frame.pack(fill="x", padx=20, pady=5)
        
        tk.Label(info_frame, text=f"ID: {record.get('id', '')}", bg='#f0f0f0', font=('Arial', 10, 'bold')).pack(anchor="w")
        tk.Label(info_frame, text=f"תאריך יצירה: {record.get('תאריך יצירה', '')}", bg='#f0f0f0').pack(anchor="w")
        tk.Label(info_frame, text=f"סך כמויות: {record.get('סך כמויות', 0)}", bg='#f0f0f0').pack(anchor="w")
        
        # טבלת מוצרים
        tk.Label(details_window, text="פירוט מוצרים:", font=('Arial', 12, 'bold'), bg='#f0f0f0').pack(anchor="w", padx=20, pady=(10, 5))
        
        products_text = scrolledtext.ScrolledText(details_window, height=20, font=('Courier New', 10))
        products_text.pack(fill="both", expand=True, padx=20, pady=5)
        
        # הוספת פירוט המוצרים
        for product in record.get('מוצרים', []):
            products_text.insert(tk.END, f"\n📦 {product.get('שם המוצר', '')}\n")
            products_text.insert(tk.END, "-" * 50 + "\n")
            
            for size_info in product.get('מידות', []):
                size = size_info.get('מידה', '')
                quantity = size_info.get('כמות', 0)
                note = size_info.get('הערה', '')
                products_text.insert(tk.END, f"   מידה {size:>6}: {quantity:>8} - {note}\n")
        
        products_text.config(state="disabled")
        
        # כפתור סגירה
        tk.Button(
            details_window,
            text="סגור",
            command=details_window.destroy,
            bg='#95a5a6',
            fg='white',
            font=('Arial', 10, 'bold')
        ).pack(pady=10)
    
    def show_context_menu(self, event, tree):
        """הצגת תפריט הקשר"""
        selection = tree.selection()
        if not selection:
            return
        
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="📋 הצג פרטים", command=lambda: self.show_record_details(tree))
        context_menu.add_separator()
        context_menu.add_command(label="🗑️ מחק רשומה", command=lambda: self.delete_record(tree))
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def delete_record(self, tree):
        """מחיקת רשומה"""
        selection = tree.selection()
        if not selection:
            return
        
        item = tree.item(selection[0])
        record_id = int(item['values'][0])
        file_name = item['values'][1]
        
        if messagebox.askyesno("אישור מחיקה", f"האם אתה בטוח שברצונך למחוק את הרשומה:\n{file_name}?"):
            # מחיקה מהרשימה
            self.drawings_data = [r for r in self.drawings_data if r.get('id') != record_id]
            
            # שמירה
            self.save_drawings_data()
            
            # רענון הטבלה
            self.refresh_drawings_table(tree.master.master)
            
            messagebox.showinfo("הצלחה", "הרשומה נמחקה בהצלחה")
    
    def refresh_drawings_table(self, table_frame):
        """רענון טבלת הציורים"""
        self.load_drawings_data()
        self.create_drawings_table(table_frame)
    
    def clear_all_drawings(self, table_frame):
        """מחיקת כל הציורים"""
        if messagebox.askyesno("אישור מחיקה", "האם אתה בטוח שברצונך למחוק את כל הציורים?\nפעולה זו לא ניתנת לביטול!"):
            self.drawings_data = []
            self.save_drawings_data()
            self.refresh_drawings_table(table_frame)
            messagebox.showinfo("הצלחה", "כל הציורים נמחקו בהצלחה")
    
    def export_drawings_to_excel(self):
        """ייצוא כל הציורים לקובץ Excel"""
        if not self.drawings_data:
            messagebox.showwarning("אזהרה", "אין ציורים לייצוא")
            return
        
        try:
            # יצירת רשימה של כל השורות
            rows = []
            
            for record in self.drawings_data:
                for product in record.get('מוצרים', []):
                    for size_info in product.get('מידות', []):
                        rows.append({
                            'ID רשומה': record.get('id', ''),
                            'שם הקובץ': record.get('שם הקובץ', ''),
                            'תאריך יצירה': record.get('תאריך יצירה', ''),
                            'שם המוצר': product.get('שם המוצר', ''),
                            'מידה': size_info.get('מידה', ''),
                            'כמות': size_info.get('כמות', 0),
                            'הערה': size_info.get('הערה', '')
                        })
            
            if not rows:
                messagebox.showwarning("אזהרה", "אין נתונים לייצוא")
                return
            
            # יצירת DataFrame
            df = pd.DataFrame(rows)
            
            # בחירת מיקום שמירה
            file_path = filedialog.asksaveasfilename(
                title="ייצא ציורים לאקסל",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if file_path:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("הצלחה", f"הציורים יוצאו בהצלחה לקובץ:\n{file_path}")
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בייצוא: {str(e)}")

def main():
    root = tk.Tk()
    app = OptitexAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
