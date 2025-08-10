import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class SupplierIntakeTabMixin:
    """Mixin לטאב קליטת ספק."""
    def _create_supplier_intake_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa'); self.notebook.add(tab, text="קליטת ספק")
        tk.Label(tab, text="קליטת מוצרים מספק (הזנה ידנית)", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)
        form = ttk.LabelFrame(tab, text="פרטי קליטה", padding=10); form.pack(fill='x', padx=10, pady=6)
        tk.Label(form, text="שם ספק:", font=('Arial',10,'bold')).grid(row=0,column=0,sticky='w',padx=4,pady=4)
        self.supplier_name_var = tk.StringVar(); tk.Entry(form, textvariable=self.supplier_name_var, width=30).grid(row=0,column=1,sticky='w',padx=4,pady=4)
        tk.Label(form, text="תאריך:", font=('Arial',10,'bold')).grid(row=0,column=2,sticky='w',padx=4,pady=4)
        self.supplier_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d')); tk.Entry(form, textvariable=self.supplier_date_var, width=15).grid(row=0,column=3,sticky='w',padx=4,pady=4)
        lines_frame = ttk.LabelFrame(tab, text="שורות קליטה", padding=8); lines_frame.pack(fill='both', expand=True, padx=10, pady=4)
        entry_bar = tk.Frame(lines_frame, bg='#f7f9fa'); entry_bar.pack(fill='x', pady=(0,6))
        self.sup_product_var = tk.StringVar(); self.sup_size_var = tk.StringVar(); self.sup_qty_var = tk.StringVar(); self.sup_note_var = tk.StringVar()
        self._supplier_products_allowed = []
        try:
            if getattr(self.file_analyzer, 'product_mapping', None):
                self._supplier_products_allowed = sorted(set(self.file_analyzer.product_mapping.values()))
            elif self.products_file and os.path.exists(self.products_file):
                try:
                    self.file_analyzer.load_products_mapping(self.products_file)
                    self._supplier_products_allowed = sorted(set(self.file_analyzer.product_mapping.values()))
                except Exception:
                    pass
        except Exception:
            self._supplier_products_allowed = []
        self.sup_product_combo = ttk.Combobox(entry_bar, textvariable=self.sup_product_var, width=16, state='readonly', values=self._supplier_products_allowed)
        def _product_chosen(event=None):
            try:
                widgets_after = [w for w in entry_bar.grid_slaves(row=1) if isinstance(w, tk.Entry)]
            except Exception:
                widgets_after = []
            for w in widgets_after:
                if hasattr(w, 'cget') and w.cget('textvariable') == str(self.sup_size_var):
                    w.focus_set(); break
        self.sup_product_combo.bind('<<ComboboxSelected>>', _product_chosen)
        lbls = ["מוצר","מידה","כמות","הערה"]
        for i,lbl in enumerate(lbls): tk.Label(entry_bar, text=lbl, bg='#f7f9fa').grid(row=0,column=i*2,sticky='w',padx=2)
        widgets = [self.sup_product_combo, tk.Entry(entry_bar, textvariable=self.sup_size_var, width=10), tk.Entry(entry_bar, textvariable=self.sup_qty_var, width=7), tk.Entry(entry_bar, textvariable=self.sup_note_var, width=18)]
        for i,w in enumerate(widgets): w.grid(row=1,column=i*2,sticky='w',padx=2)
        tk.Button(entry_bar, text="➕ הוסף", command=self._add_supplier_line, bg='#27ae60', fg='white').grid(row=1,column=8,padx=6)
        tk.Button(entry_bar, text="🗑️ מחק נבחר", command=self._delete_supplier_selected, bg='#e67e22', fg='white').grid(row=1,column=9,padx=4)
        tk.Button(entry_bar, text="❌ נקה הכל", command=self._clear_supplier_lines, bg='#e74c3c', fg='white').grid(row=1,column=10,padx=4)
        cols = ('product','size','quantity','note')
        self.supplier_tree = ttk.Treeview(lines_frame, columns=cols, show='headings', height=10)
        headers = {'product':'מוצר','size':'מידה','quantity':'כמות','note':'הערה'}; widths = {'product':180,'size':80,'quantity':80,'note':240}
        for c in cols: self.supplier_tree.heading(c, text=headers[c]); self.supplier_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(lines_frame, orient='vertical', command=self.supplier_tree.yview); self.supplier_tree.configure(yscroll=vs.set)
        self.supplier_tree.pack(side='left', fill='both', expand=True, padx=(4,0), pady=4); vs.pack(side='right', fill='y', pady=4)
        bottom_actions = tk.Frame(tab, bg='#f7f9fa'); bottom_actions.pack(fill='x', padx=10, pady=6)
        tk.Button(bottom_actions, text="💾 שמור קליטה", command=self._save_supplier_receipt, bg='#2c3e50', fg='white', font=('Arial',11,'bold')).pack(side='right', padx=4)
        self.supplier_summary_var = tk.StringVar(value="0 שורות | 0 כמות")
        tk.Label(tab, textvariable=self.supplier_summary_var, bg='#34495e', fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')
        self._supplier_lines = []

    def _add_supplier_line(self):
        product = self.sup_product_var.get().strip(); size = self.sup_size_var.get().strip(); qty_raw = self.sup_qty_var.get().strip(); note = self.sup_note_var.get().strip()
        if not product or not qty_raw: messagebox.showerror("שגיאה", "חובה לבחור מוצר ולהזין כמות"); return
        if hasattr(self, '_supplier_products_allowed') and self._supplier_products_allowed and product not in self._supplier_products_allowed:
            messagebox.showerror("שגיאה", "יש לבחור מוצר מהרשימה בלבד"); return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי"); return
        line = {'product': product, 'size': size, 'quantity': qty, 'note': note}
        self._supplier_lines.append(line); self.supplier_tree.insert('', 'end', values=(product,size,qty,note))
        self.sup_product_var.set(''); self.sup_size_var.set(''); self.sup_qty_var.set(''); self.sup_note_var.set('')
        self._update_supplier_summary()

    def _delete_supplier_selected(self):
        sel = self.supplier_tree.selection();
        if not sel: return
        all_items = self.supplier_tree.get_children(); indices = [all_items.index(i) for i in sel]
        for item in sel: self.supplier_tree.delete(item)
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._supplier_lines): del self._supplier_lines[idx]
        self._update_supplier_summary()

    def _clear_supplier_lines(self):
        self._supplier_lines = []
        for item in self.supplier_tree.get_children(): self.supplier_tree.delete(item)
        self._update_supplier_summary()

    def _update_supplier_summary(self):
        total_rows = len(self._supplier_lines); total_qty = sum(l.get('quantity',0) for l in self._supplier_lines)
        self.supplier_summary_var.set(f"{total_rows} שורות | {total_qty} כמות")

    def _save_supplier_receipt(self):
        supplier = self.supplier_name_var.get().strip(); date_str = self.supplier_date_var.get().strip()
        if not supplier: messagebox.showerror("שגיאה", "חובה למלא שם ספק"); return
        if not self._supplier_lines: messagebox.showerror("שגיאה", "אין שורות לשמירה"); return
        try:
            new_id = self.data_processor.add_supplier_receipt(supplier, date_str, self._supplier_lines)
            messagebox.showinfo("הצלחה", f"קליטה נשמרה (ID: {new_id})")
            self._clear_supplier_lines()
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
