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
        # רשימת המוצרים תגיע מעכשיו ישירות מקטלוג המוצרים (data_processor.products_catalog)
        # ולא ממיפוי קובץ ה-RIB.
        self._supplier_products_allowed = []
        self._refresh_supplier_products_allowed(initial=True)
        # Autocomplete using popup Listbox (more reliable than mutating Combobox values each keystroke)
        self._supplier_products_allowed_full = list(self._supplier_products_allowed)
        self.sup_product_combo = ttk.Combobox(entry_bar, textvariable=self.sup_product_var, width=16, state='normal')
        self.sup_product_combo['values'] = self._supplier_products_allowed_full

        # Popup elements
        self._sup_ac_popup = None
        self._sup_ac_list = None

        def _ensure_popup():
            if self._sup_ac_popup and self._sup_ac_popup.winfo_exists():
                return
            popup = tk.Toplevel(self.sup_product_combo)
            popup.overrideredirect(True)
            popup.attributes('-topmost', True)
            lb = tk.Listbox(popup, height=8, activestyle='dotbox')
            lb.pack(fill='both', expand=True)
            self._sup_ac_popup = popup
            self._sup_ac_list = lb

            def _choose(event=None):
                sel = lb.curselection()
                if not sel:
                    _hide_popup(); return
                val = lb.get(sel[0])
                self.sup_product_var.set(val)
                _hide_popup()
                # Move focus to size field automatically for flow
                try:
                    size_entry = [w for w in entry_bar.grid_slaves(row=1) if isinstance(w, tk.Entry)][0]
                    size_entry.focus_set()
                except Exception:
                    pass
            lb.bind('<Return>', _choose)
            lb.bind('<Double-Button-1>', _choose)
            lb.bind('<Escape>', lambda e: _hide_popup())

        def _hide_popup():
            if self._sup_ac_popup and self._sup_ac_popup.winfo_exists():
                self._sup_ac_popup.destroy()

        def _position_popup():
            if not (self._sup_ac_popup and self._sup_ac_popup.winfo_exists()):
                return
            try:
                x = self.sup_product_combo.winfo_rootx()
                y = self.sup_product_combo.winfo_rooty() + self.sup_product_combo.winfo_height()
                w = self.sup_product_combo.winfo_width()
                self._sup_ac_popup.geometry(f"{w}x180+{x}+{y}")
            except Exception:
                pass

        def _filter_products(event=None):
            if event and event.keysym in ('Escape',):
                _hide_popup(); return
            text = self.sup_product_var.get().strip()
            base = self._supplier_products_allowed_full
            if not text:
                matches = base[:50]
            else:
                tokens = [t for t in text.lower().replace('-', ' ').split() if t]
                def match(prod):
                    pl = prod.lower()
                    words = pl.replace('-', ' ').split()
                    for tok in tokens:
                        if not any(w.startswith(tok) or tok in w for w in words):
                            return None
                    prefix_hits = sum(any(w.startswith(tok) for w in words) for tok in tokens)
                    first_idx_sum = sum(min((i for i,w in enumerate(words) if (w.startswith(tok) or tok in w)), default=99) for tok in tokens)
                    return (-prefix_hits, first_idx_sum, len(prod))
                scored = []
                for p in base:
                    sc = match(p)
                    if sc is not None:
                        scored.append((sc,p))
                scored.sort(key=lambda x: x[0])
                matches = [p for _,p in scored][:50]
            if not matches:
                _hide_popup(); return
            _ensure_popup(); _position_popup()
            lb = self._sup_ac_list
            lb.delete(0, tk.END)
            for m in matches:
                lb.insert(tk.END, m)
            lb.selection_clear(0, tk.END)
            lb.selection_set(0)
            lb.activate(0)

        def _on_key(event):
            if event.keysym in ('Down','Up') and self._sup_ac_popup and self._sup_ac_popup.winfo_exists():
                lb = self._sup_ac_list; size = lb.size()
                if size == 0: return
                cur = lb.curselection()
                idx = cur[0] if cur else 0
                if event.keysym == 'Down': idx = (idx + 1) % size
                else: idx = (idx - 1) % size
                lb.selection_clear(0, tk.END); lb.selection_set(idx); lb.activate(idx)
                return 'break'
            if event.keysym == 'Return' and self._sup_ac_popup and self._sup_ac_popup.winfo_exists():
                lb = self._sup_ac_list; cur = lb.curselection()
                if cur:
                    self.sup_product_var.set(lb.get(cur[0])); _hide_popup(); return 'break'
            _filter_products()

        self.sup_product_combo.bind('<KeyRelease>', _on_key)
        self.sup_product_combo.bind('<FocusOut>', lambda e: self.root.after(150, _hide_popup))
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
        # קומבו של מידות יתעדכן לפי המוצר הנבחר (וריאנטים מהקטלוג)
        self.sup_size_combo = ttk.Combobox(entry_bar, textvariable=self.sup_size_var, width=10, state='readonly')
        widgets = [self.sup_product_combo, self.sup_size_combo, tk.Entry(entry_bar, textvariable=self.sup_qty_var, width=7), tk.Entry(entry_bar, textvariable=self.sup_note_var, width=18)]
        for i,w in enumerate(widgets):
            w.grid(row=1,column=i*2,sticky='w',padx=2)
        # טרייס לעדכון מידות עם שינוי מוצר
        def _on_product_change(*_a):
            try:
                self._update_supplier_size_options()
            except Exception:
                pass
        try:
            self.sup_product_var.trace_add('write', _on_product_change)
        except Exception:
            pass
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

    def _refresh_supplier_products_allowed(self, initial: bool = False):
        """רענון רשימת המוצרים האפשריים מתוך קטלוג המוצרים.

        אם אין קטלוג נטען -> רשימה ריקה. נשמר סדר אלפביתי וייחודיות.
        מופעל בטעינה ראשונית ובעדכונים מהטאב של הקטלוג.
        """
        try:
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
            names = sorted({ (rec.get('name') or '').strip() for rec in catalog if rec.get('name') })
            self._supplier_products_allowed = names
            self._supplier_products_allowed_full = list(names)
            # עדכון קומבובוקס אם כבר נוצרה
            if hasattr(self, 'sup_product_combo'):
                try:
                    current = self.sup_product_var.get()
                    self.sup_product_combo['values'] = self._supplier_products_allowed_full
                    # אם המוצר הנוכחי כבר לא קיים ננקה
                    if current and current not in self._supplier_products_allowed_full:
                        self.sup_product_var.set('')
                except Exception:
                    pass
        except Exception:
            self._supplier_products_allowed = []
            self._supplier_products_allowed_full = []

    def _update_supplier_size_options(self):
        """עדכון רשימת המידות (וריאנטים) עבור המוצר שנבחר מקטלוג המוצרים.

        לוגיקה: אוסף כל הרשומות בקטלוג עם אותו name ומוציא את השדה size (לא ריק) בצורה ממויינת לוגית.
        אם אין מידות -> מנקה.
        """
        product = (self.sup_product_var.get() or '').strip()
        sizes = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        sz = (rec.get('size') or '').strip()
                        if sz:
                            sizes.append(sz)
                # מיון לוגי (מספר תחילי אם קיים)
                import re
                def _size_key(s):
                    m = re.match(r"(\d+)", s)
                    base = int(m.group(1)) if m else 0
                    return (base, s)
                sizes = sorted({s for s in sizes}, key=_size_key)
            except Exception:
                sizes = []
        # עדכון קומבובוקס
        if hasattr(self, 'sup_size_combo'):
            try:
                self.sup_size_combo['values'] = sizes
                if sizes:
                    self.sup_size_combo.state(['!disabled','readonly'])
                    # אם המידה הנוכחית לא ברשימה – ננקה
                    if self.sup_size_var.get() not in sizes:
                        self.sup_size_var.set('')
                else:
                    self.sup_size_var.set('')
                    # ללא מידות – נעבור למצב מושבת כדי לא להטעות
                    self.sup_size_combo.set('')
                    try:
                        self.sup_size_combo.state(['disabled'])
                    except Exception:
                        pass
            except Exception:
                pass

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
        # Restore full suggestions list after adding a line
        if hasattr(self, 'sup_product_combo'):
            try:
                self.sup_product_combo['values'] = self._supplier_products_allowed_full
            except Exception:
                pass
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
