"""
חלון מנהל הציורים - עמוד נפרד לניהול הטבלה המקומית
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog


class DrawingsManagerWindow:
    """חלון מנהל הציורים"""
    
    def __init__(self, parent, drawings_manager):
        self.parent = parent
        self.drawings_manager = drawings_manager
        self.window = None
        self.current_tree = None
        
    def open_window(self):
        """פתיחת חלון מנהל הציורים"""
        if self.window:
            self.window.lift()
            return
            
        self.window = tk.Toplevel(self.parent)
        self.window.title("מנהל ציורים")
        self.window.geometry("1000x700")
        self.window.configure(bg='#f0f0f0')
        
        # קישור לסגירת החלון
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)
        
        self.create_widgets()
    
    def close_window(self):
        """סגירת החלון"""
        if self.window:
            self.window.destroy()
            self.window = None
    
    def create_widgets(self):
        """יצירת רכיבי הממשק"""
        # כותרת
        tk.Label(
            self.window,
            text="מנהל ציורים - טבלה מקומית",
            font=('Arial', 16, 'bold'),
            bg='#f0f0f0',
            fg='#2c3e50'
        ).pack(pady=10)
        
        # סטטיסטיקות
        self.create_statistics_panel()
        
        # כפתורי פעולה
        self.create_actions_panel()
        
        # טבלת הציורים
        self.create_table_panel()
    
    def create_statistics_panel(self):
        """יצירת פאנל סטטיסטיקות"""
        stats_frame = tk.Frame(self.window, bg='#f0f0f0')
        stats_frame.pack(fill="x", padx=20, pady=5)
        
        stats = self.drawings_manager.get_statistics()
        
        tk.Label(
            stats_frame,
            text=f"📊 סטטיסטיקות: {stats['total_drawings']} ציורים | {stats['total_products']} מוצרים | {stats['total_quantity']} יחידות",
            font=('Arial', 10),
            bg='#f0f0f0',
            fg='#666'
        ).pack(side="left")
    
    def create_actions_panel(self):
        """יצירת פאנל כפתורי פעולה"""
        actions_frame = tk.Frame(self.window, bg='#f0f0f0')
        actions_frame.pack(fill="x", padx=20, pady=5)
        
        tk.Button(
            actions_frame,
            text="🔄 רענן",
            command=self.refresh_table,
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold')
        ).pack(side="left", padx=5)
        
        tk.Button(
            actions_frame,
            text="📊 ייצא לאקסל",
            command=self.export_to_excel,
            bg='#27ae60',
            fg='white',
            font=('Arial', 10, 'bold')
        ).pack(side="left", padx=5)
        
        tk.Button(
            actions_frame,
            text="🗑️ מחק הכל",
            command=self.clear_all_drawings,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold')
        ).pack(side="right", padx=5)
    
    def create_table_panel(self):
        """יצירת פאנל הטבלה"""
        table_frame = tk.Frame(self.window, bg='#f0f0f0')
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.create_table(table_frame)
    
    def create_table(self, parent_frame):
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
        drawings = self.drawings_manager.get_all_drawings()
        for record in drawings:
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
        tree.bind("<Double-1>", self.show_record_details)
        tree.bind("<Button-3>", self.show_context_menu)
        
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
    
    def show_record_details(self, event):
        """הצגת פרטי רשומה"""
        tree = event.widget
        selection = tree.selection()
        if not selection:
            return
        
        item = tree.item(selection[0])
        record_id = int(item['values'][0])
        
        # חיפוש הרשומה
        record = self.drawings_manager.get_drawing_by_id(record_id)
        
        if not record:
            messagebox.showerror("שגיאה", "רשומה לא נמצאה")
            return
        
        # פתיחת חלון פרטים
        self.open_details_window(record)
    
    def open_details_window(self, record):
        """פתיחת חלון פרטי רשומה"""
        details_window = tk.Toplevel(self.window)
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
    
    def show_context_menu(self, event):
        """הצגת תפריט הקשר"""
        tree = event.widget
        selection = tree.selection()
        if not selection:
            return
        
        context_menu = tk.Menu(self.window, tearoff=0)
        context_menu.add_command(label="📋 הצג פרטים", command=lambda: self.show_record_details(event))
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
            # מחיקה מהמנהל
            result = self.drawings_manager.delete_drawing(record_id)
            
            if result['success']:
                self.refresh_table()
                messagebox.showinfo("הצלחה", "הרשומה נמחקה בהצלחה")
            else:
                messagebox.showerror("שגיאה", f"שגיאה במחיקה: {result['error']}")
    
    def refresh_table(self):
        """רענון הטבלה"""
        self.drawings_manager.load_data()
        
        # עדכון סטטיסטיקות
        for widget in self.window.winfo_children():
            if isinstance(widget, tk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, tk.Label) and "סטטיסטיקות" in child.cget("text"):
                        stats = self.drawings_manager.get_statistics()
                        child.config(text=f"📊 סטטיסטיקות: {stats['total_drawings']} ציורים | {stats['total_products']} מוצרים | {stats['total_quantity']} יחידות")
                        break
                break
        
        # רענון הטבלה
        for widget in self.window.winfo_children():
            if isinstance(widget, tk.Frame) and widget.winfo_children():
                for child in widget.winfo_children():
                    if isinstance(child, tk.Frame):
                        self.create_table(child)
                        break
                break
    
    def clear_all_drawings(self):
        """מחיקת כל הציורים"""
        if messagebox.askyesno("אישור מחיקה", "האם אתה בטוח שברצונך למחוק את כל הציורים?\nפעולה זו לא ניתנת לביטול!"):
            result = self.drawings_manager.clear_all_drawings()
            
            if result['success']:
                self.refresh_table()
                messagebox.showinfo("הצלחה", "כל הציורים נמחקו בהצלחה")
            else:
                messagebox.showerror("שגיאה", f"שגיאה במחיקה: {result['error']}")
    
    def export_to_excel(self):
        """ייצוא לאקסל"""
        drawings = self.drawings_manager.get_all_drawings()
        if not drawings:
            messagebox.showwarning("אזהרה", "אין ציורים לייצוא")
            return
        
        # בחירת מיקום שמירה
        file_path = filedialog.asksaveasfilename(
            title="ייצא ציורים לאקסל",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            result = self.drawings_manager.export_to_excel(file_path)
            
            if result['success']:
                messagebox.showinfo("הצלחה", f"הציורים יוצאו בהצלחה לקובץ:\n{file_path}\n\nמספר שורות: {result['rows_count']}")
            else:
                messagebox.showerror("שגיאה", f"שגיאה בייצוא: {result['error']}")
