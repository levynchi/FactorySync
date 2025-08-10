"""
מנהל ציורים - ממשק גרפי
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog

class DrawingsManagerWindow:
    """חלון מנהל הציורים"""
    
    def __init__(self, parent, data_processor):
        self.parent = parent
        self.data_processor = data_processor
        self.window = None
        self.current_tree = None
    
    def show(self):
        """הצגת חלון מנהל הציורים"""
        if self.window and self.window.winfo_exists():
            self.window.lift()
            self._refresh_table()
            return
        
        self.window = tk.Toplevel(self.parent)
        self.window.title("מנהל ציורים")
        self.window.geometry("1200x800")
        self.window.configure(bg='#f0f0f0')
        
        self._create_widgets()
        self._create_table()
    
    def _create_widgets(self):
        """יצירת הרכיבים"""
        # כותרת
        title_label = tk.Label(
            self.window,
            text="מנהל ציורים - טבלה מקומית",
            font=('Arial', 16, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        title_label.pack(pady=15)
        
        # מסגרת כפתורי פעולה
        actions_frame = tk.Frame(self.window, bg='#f0f0f0')
        actions_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        # כפתורים בצד שמאל
        left_buttons = tk.Frame(actions_frame, bg='#f0f0f0')
        left_buttons.pack(side="left")
        
        tk.Button(
            left_buttons,
            text="🔄 רענן",
            command=self._refresh_table,
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=10
        ).pack(side="left", padx=5)
        
        tk.Button(
            left_buttons,
            text="📊 ייצא לאקסל",
            command=self._export_to_excel,
            bg='#27ae60',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=12
        ).pack(side="left", padx=5)
        
        # כפתורים בצד ימין
        right_buttons = tk.Frame(actions_frame, bg='#f0f0f0')
        right_buttons.pack(side="right")
        
        tk.Button(
            right_buttons,
            text="🗑️ מחק הכל",
            command=self._clear_all_drawings,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=10
        ).pack(side="right", padx=5)
        
        tk.Button(
            right_buttons,
            text="❌ מחק נבחר",
            command=self.delete_selected_drawing,
            bg='#e67e22',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=10
        ).pack(side="right", padx=5)
        
        # מסגרת הטבלה
        self.table_frame = tk.Frame(self.window, bg='#f0f0f0')
        self.table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # מידע סטטיסטי
        self.stats_label = tk.Label(
            self.window,
            text="",
            bg='#34495e',
            fg='white',
            font=('Arial', 10),
            anchor='w',
            padx=10
        )
        self.stats_label.pack(fill="x", side="bottom")
    
    def _create_table(self):
        """יצירת הטבלה"""
        # ניקוי המסגרת
        for widget in self.table_frame.winfo_children():
            widget.destroy()
        
        # מסגרת עם גלילה
        tree_frame = tk.Frame(self.table_frame)
        tree_frame.pack(fill="both", expand=True)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # עמודות הטבלה
        columns = ("ID", "שם הקובץ", "תאריך יצירה", "מוצרים", "סך כמויות")
        
        # Treeview
        self.current_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set
        )
        
        # הגדרת כותרות ורוחב עמודות
        column_widths = {
            "ID": 80,
            "שם הקובץ": 300,
            "תאריך יצירה": 180,
            "מוצרים": 120,
            "סך כמויות": 120
        }
        
        for col in columns:
            self.current_tree.heading(col, text=col)
            self.current_tree.column(col, width=column_widths[col], anchor="center")
        
        # הוספת נתונים
        self._populate_table()
        
        # קישור אירועים
        self.current_tree.bind("<Double-1>", self._on_double_click)
        self.current_tree.bind("<Button-3>", self._on_right_click)
        
        # סידור הרכיבים
        self.current_tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        # קישור הגלילה
        v_scrollbar.config(command=self.current_tree.yview)
        h_scrollbar.config(command=self.current_tree.xview)
        
        # הגדרת משקולות
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # עדכון סטטיסטיקות
        self._update_stats()
    
    def _populate_table(self):
        """מילוי הטבלה בנתונים"""
        # ניקוי הטבלה
        for item in self.current_tree.get_children():
            self.current_tree.delete(item)
        
        # הוספת נתונים
        for record in self.data_processor.drawings_data:
            products_count = len(record.get('מוצרים', []))
            total_quantity = record.get('סך כמויות', 0)
            
            self.current_tree.insert("", "end", values=(
                record.get('id', ''),
                record.get('שם הקובץ', ''),
                record.get('תאריך יצירה', ''),
                products_count,
                f"{total_quantity:.1f}" if isinstance(total_quantity, float) else str(total_quantity)
            ))
    
    def _update_stats(self):
        """עדכון הסטטיסטיקות"""
        total_drawings = len(self.data_processor.drawings_data)
        total_quantity = sum(record.get('סך כמויות', 0) for record in self.data_processor.drawings_data)
        
        stats_text = f"סך הכל: {total_drawings} ציורים | סך כמויות: {total_quantity:.1f}"
        self.stats_label.config(text=stats_text)
    
    def _refresh_table(self):
        """רענון הטבלה"""
        self.data_processor.refresh_drawings_data()
        self._populate_table()
        self._update_stats()
    
    def _on_double_click(self, event):
        """טיפול בלחיצה כפולה"""
        self._show_record_details()
    
    def _on_right_click(self, event):
        """טיפול בלחיצה ימנית"""
        selection = self.current_tree.selection()
        if not selection:
            return
        
        # יצירת תפריט הקשר
        context_menu = tk.Menu(self.window, tearoff=0)
        context_menu.add_command(label="📋 הצג פרטים", command=self._show_record_details)
        context_menu.add_separator()
        context_menu.add_command(label="🗑️ מחק רשומה", command=self._delete_selected_record)
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def _show_record_details(self):
        """הצגת פרטי רשומה"""
        selection = self.current_tree.selection()
        if not selection:
            return
        
        item = self.current_tree.item(selection[0])
        record_id = int(item['values'][0])
        
        # חיפוש הרשומה
        record = self.data_processor.get_drawing_by_id(record_id)
        if not record:
            messagebox.showerror("שגיאה", "רשומה לא נמצאה")
            return
        
        self._create_details_window(record)
    
    def _create_details_window(self, record):
        """יצירת חלון פרטי רשומה"""
        details_window = tk.Toplevel(self.window)
        details_window.title(f"פרטי ציור - {record.get('שם הקובץ', '')}")
        details_window.geometry("900x700")
        details_window.configure(bg='#f0f0f0')
        details_window.grab_set()
        
        # כותרת
        tk.Label(
            details_window,
            text=f"פרטי ציור: {record.get('שם הקובץ', '')}",
            font=('Arial', 14, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50'
        ).pack(pady=15)
        
        # מידע כללי
        info_frame = tk.LabelFrame(details_window, text="מידע כללי", bg='#f0f0f0', font=('Arial', 11, 'bold'))
        info_frame.pack(fill="x", padx=20, pady=10)
        
        info_text = f"""ID: {record.get('id', '')}
תאריך יצירה: {record.get('תאריך יצירה', '')}
מספר מוצרים: {len(record.get('מוצרים', []))}
סך הכמויות: {record.get('סך כמויות', 0)}"""
        
        tk.Label(
            info_frame,
            text=info_text,
            bg='#f0f0f0',
            font=('Arial', 10),
            justify="left",
            anchor="w"
        ).pack(fill="x", padx=10, pady=10)
        
        # פירוט מוצרים
        tk.Label(
            details_window,
            text="פירוט מוצרים ומידות:",
            font=('Arial', 12, 'bold'),
            bg='#f0f0f0'
        ).pack(anchor="w", padx=20, pady=(10, 5))
        
        # טקסט מגלל עם פירוט
        products_text = scrolledtext.ScrolledText(
            details_window,
            height=20,
            font=('Courier New', 10),
            wrap=tk.WORD
        )
        products_text.pack(fill="both", expand=True, padx=20, pady=5)
        
        # הוספת פירוט המוצרים
        for product in record.get('מוצרים', []):
            products_text.insert(tk.END, f"\n📦 {product.get('שם המוצר', '')}\n")
            products_text.insert(tk.END, "=" * 60 + "\n")
            
            total_product_quantity = 0
            for size_info in product.get('מידות', []):
                size = size_info.get('מידה', '')
                quantity = size_info.get('כמות', 0)
                note = size_info.get('הערה', '')
                total_product_quantity += quantity
                
                products_text.insert(tk.END, f"   מידה {size:>8}: {quantity:>8} - {note}\n")
            
            products_text.insert(tk.END, f"\nסך עבור מוצר זה: {total_product_quantity}\n")
            products_text.insert(tk.END, "-" * 60 + "\n")
        
        products_text.config(state="disabled")
        
        # כפתור סגירה
        tk.Button(
            details_window,
            text="סגור",
            command=details_window.destroy,
            bg='#95a5a6',
            fg='white',
            font=('Arial', 11, 'bold'),
            width=15
        ).pack(pady=15)
    
    def _delete_selected_record(self):
        """מחיקת הרשומה הנבחרת"""
        selection = self.current_tree.selection()
        if not selection:
            return
        
        item = self.current_tree.item(selection[0])
        record_id = int(item['values'][0])
        file_name = item['values'][1]
        
        if messagebox.askyesno(
            "אישור מחיקה",
            f"האם אתה בטוח שברצונך למחוק את הרשומה:\n{file_name}?\n\nפעולה זו לא ניתנת לביטול!"
        ):
            if self.data_processor.delete_drawing(record_id):
                self._refresh_table()
                messagebox.showinfo("הצלחה", "הרשומה נמחקה בהצלחה")
            else:
                messagebox.showerror("שגיאה", "שגיאה במחיקת הרשומה")
    
    def _clear_all_drawings(self):
        """מחיקת כל הציורים"""
        if not self.data_processor.drawings_data:
            messagebox.showinfo("מידע", "אין ציורים למחיקה")
            return
        
        if messagebox.askyesno(
            "אישור מחיקה",
            f"האם אתה בטוח שברצונך למחוק את כל {len(self.data_processor.drawings_data)} הציורים?\n\nפעולה זו לא ניתנת לביטול!"
        ):
            if self.data_processor.clear_all_drawings():
                self._refresh_table()
                messagebox.showinfo("הצלחה", "כל הציורים נמחקו בהצלחה")
            else:
                messagebox.showerror("שגיאה", "שגיאה במחיקת הציורים")
    
    def _export_to_excel(self):
        """ייצוא הציורים לאקסל"""
        if not self.data_processor.drawings_data:
            messagebox.showwarning("אזהרה", "אין ציורים לייצוא")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="ייצא ציורים לאקסל",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.data_processor.export_drawings_to_excel(file_path)
                messagebox.showinfo("הצלחה", f"הציורים יוצאו בהצלחה לקובץ:\n{file_path}")
            except Exception as e:
                messagebox.showerror("שגיאה", f"שגיאה בייצוא:\n{str(e)}")
    
    def close(self):
        """סגירת החלון"""
        if self.window:
            self.window.destroy()
            self.window = None
    
    def delete_record(self):
        """(legacy method wrapper) קורא למחיקה הפעילה אם יש בחירה"""
        try:
            # מפה ישנה שקראה delete_record(tree) - כאן אין tree אז נשתמש ב-current_tree
            if hasattr(self, 'current_tree') and self.current_tree:
                self.delete_selected_drawing()
        except Exception:
            pass
    
    def delete_selected_drawing(self):
        """מחיקת ציור נבחר מהטבלה"""
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
                original_count = len(self.data_processor.drawings_data)
                self.data_processor.drawings_data = [
                    r for r in self.data_processor.drawings_data 
                    if r.get('id') != record_id
                ]
                
                if len(self.data_processor.drawings_data) < original_count:
                    # שמירה
                    self.data_processor.save_drawings_data()
                    
                    # רענון הטבלה
                    self._refresh_table()
                    
                    messagebox.showinfo("הצלחה", f"הציור '{file_name}' נמחק בהצלחה!")
                else:
                    messagebox.showerror("שגיאה", "הציור לא נמצא או לא נמחק")
                    
            except Exception as e:
                messagebox.showerror("שגיאה", f"שגיאה במחיקת הציור: {str(e)}")
