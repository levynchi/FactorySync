import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import re  # עבור פיצול מידות מרובות (וריאנטים)

class ProductsCatalogTabMixin:
    """Mixin לטאב ניהול קטלוג מוצרים (הוספה / מחיקה / ייצוא)."""
    def _create_products_catalog_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="קטלוג מוצרים")
        tk.Label(tab, text="ניהול קטלוג מוצרים", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)
        form = ttk.LabelFrame(tab, text="הוספת מוצר", padding=10)
        form.pack(fill='x', padx=10, pady=6)
        self.prod_name_var = tk.StringVar(); self.prod_size_var = tk.StringVar(); self.prod_fabric_type_var = tk.StringVar(); self.prod_fabric_color_var = tk.StringVar(); self.prod_print_name_var = tk.StringVar()
        # עדכון: שדה המידה תומך במספר וריאנטים בבת אחת מופרדים בפסיק / רווח (למשל: "0-3,3-6,6-12")
        labels = [("שם מוצר", self.prod_name_var, 25), ("מידות (פסיק)", self.prod_size_var, 18), ("סוג בד", self.prod_fabric_type_var, 15), ("צבע בד", self.prod_fabric_color_var, 15), ("שם פרינט", self.prod_print_name_var, 15)]
        for i,(lbl,var,width) in enumerate(labels):
            tk.Label(form, text=f"{lbl}:", font=('Arial',10,'bold')).grid(row=0, column=i*2, sticky='w', padx=4, pady=4)
            tk.Entry(form, textvariable=var, width=width).grid(row=0, column=i*2+1, sticky='w', padx=2, pady=4)
        tk.Button(form, text="➕ הוסף", command=self._add_product_catalog_entry, bg='#27ae60', fg='white').grid(row=0, column=len(labels)*2, padx=8)
        tk.Button(form, text="🗑️ מחק נבחר", command=self._delete_selected_product_entry, bg='#e67e22', fg='white').grid(row=0, column=len(labels)*2+1, padx=4)
        tk.Button(form, text="💾 ייצוא ל-Excel", command=self._export_products_catalog, bg='#2c3e50', fg='white').grid(row=0, column=len(labels)*2+2, padx=4)
        # Treeview
        tree_frame = ttk.LabelFrame(tab, text="מוצרים", padding=6)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=6)
        cols = ('id','name','size','fabric_type','fabric_color','print_name','created_at')
        self.products_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=12)
        headers = {'id':'ID','name':'שם מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','print_name':'שם פרינט','created_at':'נוצר'}
        widths = {'id':50,'name':180,'size':80,'fabric_type':120,'fabric_color':120,'print_name':140,'created_at':140}
        for c in cols:
            self.products_tree.heading(c, text=headers[c])
            self.products_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(tree_frame, orient='vertical', command=self.products_tree.yview)
        self.products_tree.configure(yscroll=vs.set)
        self.products_tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self._load_products_catalog_into_tree()

    # Loading
    def _load_products_catalog_into_tree(self):
        if not hasattr(self, 'products_tree'): return
        for item in self.products_tree.get_children(): self.products_tree.delete(item)
        try:
            for rec in getattr(self.data_processor, 'products_catalog', []):
                self.products_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('size'), rec.get('fabric_type'), rec.get('fabric_color'), rec.get('print_name'), rec.get('created_at')))
        except Exception:
            pass

    def _add_product_catalog_entry(self):
        """הוספת מוצר לקטלוג. תומך בווריאנטים (מספר מידות בשדה אחד)."""
        name = self.prod_name_var.get().strip(); sizes_raw = self.prod_size_var.get().strip(); ft = self.prod_fabric_type_var.get().strip(); fc = self.prod_fabric_color_var.get().strip(); pn = self.prod_print_name_var.get().strip()
        if not name:
            messagebox.showerror("שגיאה", "חובה להזין שם מוצר")
            return
        # פיצול וריאנטים: מפרידים בפסיק או רווחים
        if sizes_raw:
            size_tokens = [s.strip() for s in re.split(r'[;,\s]+', sizes_raw) if s.strip()]
        else:
            size_tokens = ['']  # אפשרות למוצר בלי מידה
        if not size_tokens:
            size_tokens = ['']
        added_ids = []
        try:
            for sz in size_tokens:
                new_id = self.data_processor.add_product_catalog_entry(name, sz, ft, fc, pn)
                added_ids.append((new_id, sz))
                self.products_tree.insert('', 'end', values=(new_id, name, sz, ft, fc, pn, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            # ניקוי שדות – נשאיר שם מוצר אם המשתמש רוצה להמשיך להזין וריאנטים נוספים לאותו מוצר? לפי נוחות – נשאיר ריק.
            self.prod_name_var.set(''); self.prod_size_var.set(''); self.prod_fabric_type_var.set(''); self.prod_fabric_color_var.set(''); self.prod_print_name_var.set('')
            if len(added_ids) > 1:
                messagebox.showinfo("הצלחה", f"נוספו {len(added_ids)} וריאנטים למוצר '{name}'")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _delete_selected_product_entry(self):
        sel = self.products_tree.selection()
        if not sel: return
        ids = []
        for item in sel:
            vals = self.products_tree.item(item, 'values')
            if vals:
                ids.append(int(vals[0]))
        if not ids: return
        deleted_any = False
        for _id in ids:
            if self.data_processor.delete_product_catalog_entry(_id):
                deleted_any = True
        if deleted_any:
            self._load_products_catalog_into_tree()

    def _export_products_catalog(self):
        if not getattr(self.data_processor, 'products_catalog', []):
            messagebox.showerror("שגיאה", "אין מוצרים לייצוא")
            return
        file_path = filedialog.asksaveasfilename(title="ייצוא קטלוג מוצרים", defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')])
        if not file_path: return
        try:
            self.data_processor.export_products_catalog_to_excel(file_path)
            messagebox.showinfo("הצלחה", "הקטלוג יוצא בהצלחה")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
