import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class DeliveryNoteTabMixin:
    """Mixin לטאב 'תעודת משלוח' (העתק של קליטת סחורה מספק עם שמות מבודדים)."""

    def _create_delivery_note_tab(self):
        # Tab container
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="תעודת משלוח")
        tk.Label(tab, text="תעודת משלוח (הזנה ידנית)", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=4)

        # Inner notebook (entry + list)
        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=4, pady=4)
        entry_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        list_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(entry_wrapper, text="קליטה")
        inner_nb.add(list_wrapper, text="תעודות שמורות")

        container = entry_wrapper

        # Header form
        form = ttk.LabelFrame(container, text="פרטי תעודה", padding=10)
        form.pack(fill='x', padx=10, pady=6)
        tk.Label(form, text="שם ספק:", font=('Arial',10,'bold')).grid(row=0,column=0,sticky='w',padx=4,pady=4)
        self.dn_supplier_name_var = tk.StringVar()
        self.dn_supplier_name_combo = ttk.Combobox(form, textvariable=self.dn_supplier_name_var, width=28, state='readonly')
        try:
            names = self._get_supplier_names() if hasattr(self,'_get_supplier_names') else []
            self.dn_supplier_name_combo['values'] = names
        except Exception:
            pass
        self.dn_supplier_name_combo.grid(row=0,column=1,sticky='w',padx=4,pady=4)
        tk.Label(form, text="תאריך:", font=('Arial',10,'bold')).grid(row=0,column=2,sticky='w',padx=4,pady=4)
        self.dn_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(form, textvariable=self.dn_date_var, width=15).grid(row=0,column=3,sticky='w',padx=4,pady=4)

        # Lines frame
        lines_frame = ttk.LabelFrame(container, text="שורות תעודה", padding=8)
        lines_frame.pack(fill='both', expand=False, padx=10, pady=4)
        entry_bar = tk.Frame(lines_frame, bg='#f7f9fa')
        entry_bar.pack(fill='x', pady=(0,6))

        # Variables
        self.dn_product_var = tk.StringVar()
        self.dn_size_var = tk.StringVar()
        self.dn_qty_var = tk.StringVar()
        self.dn_note_var = tk.StringVar()
        self.dn_fabric_type_var = tk.StringVar()
        self.dn_fabric_color_var = tk.StringVar(value='לבן')
        self.dn_print_name_var = tk.StringVar(value='חלק')

        # Product list
        self._delivery_products_allowed = []
        self._refresh_delivery_products_allowed(initial=True)
        self._delivery_products_allowed_full = list(self._delivery_products_allowed)
        self.dn_product_combo = ttk.Combobox(entry_bar, textvariable=self.dn_product_var, width=16, state='normal')
        self.dn_product_combo['values'] = self._delivery_products_allowed_full

        # Popup state
        self._dn_ac_popup = None
        self._dn_ac_list = None

        def _ensure_popup():
            if self._dn_ac_popup and self._dn_ac_popup.winfo_exists():
                return
            popup = tk.Toplevel(self.dn_product_combo)
            popup.overrideredirect(True)
            popup.attributes('-topmost', True)
            lb = tk.Listbox(popup, height=8, activestyle='dotbox')
            lb.pack(fill='both', expand=True)
            self._dn_ac_popup = popup
            self._dn_ac_list = lb

            def _choose(event=None):
                sel = lb.curselection()
                if not sel:
                    _hide_popup(); return
                val = lb.get(sel[0])
                self.dn_product_var.set(val)
                _hide_popup()
                try:
                    size_entry = [w for w in entry_bar.grid_slaves(row=1) if isinstance(w, tk.Entry)][0]
                    size_entry.focus_set()
                except Exception:
                    pass
            lb.bind('<Return>', _choose)
            lb.bind('<Double-Button-1>', _choose)
            lb.bind('<Escape>', lambda e: _hide_popup())

        def _hide_popup():
            if self._dn_ac_popup and self._dn_ac_popup.winfo_exists():
                self._dn_ac_popup.destroy()

        def _position_popup():
            if not (self._dn_ac_popup and self._dn_ac_popup.winfo_exists()):
                return
            try:
                x = self.dn_product_combo.winfo_rootx()
                y = self.dn_product_combo.winfo_rooty() + self.dn_product_combo.winfo_height()
                w = self.dn_product_combo.winfo_width()
                self._dn_ac_popup.geometry(f"{w}x180+{x}+{y}")
            except Exception:
                pass

        def _filter_products(event=None):
            if event and event.keysym in ('Escape',):
                _hide_popup(); return
            text = self.dn_product_var.get().strip()
            base = self._delivery_products_allowed_full
            if not text:
                matches = base[:50]
            else:
                tokens = [t for t in text.lower().replace('-', ' ').split() if t]
                def match(prod):
                    pl = prod.lower(); words = pl.replace('-', ' ').split()
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
            lb = self._dn_ac_list; lb.delete(0, tk.END)
            for m in matches: lb.insert(tk.END, m)
            lb.selection_clear(0, tk.END); lb.selection_set(0); lb.activate(0)

        def _on_key(event):
            if event.keysym in ('Down','Up') and self._dn_ac_popup and self._dn_ac_popup.winfo_exists():
                lb = self._dn_ac_list; size = lb.size()
                if size == 0: return
                cur = lb.curselection(); idx = cur[0] if cur else 0
                if event.keysym == 'Down': idx = (idx + 1) % size
                else: idx = (idx - 1) % size
                lb.selection_clear(0, tk.END); lb.selection_set(idx); lb.activate(idx)
                return 'break'
            if event.keysym == 'Return' and self._dn_ac_popup and self._dn_ac_popup.winfo_exists():
                lb = self._dn_ac_list; cur = lb.curselection()
                if cur:
                    self.dn_product_var.set(lb.get(cur[0])); _hide_popup(); return 'break'
            _filter_products()

        self.dn_product_combo.bind('<KeyRelease>', _on_key)
        self.dn_product_combo.bind('<FocusOut>', lambda e: self.root.after(150, _hide_popup))

        def _product_chosen(event=None):
            try:
                widgets_after = [w for w in entry_bar.grid_slaves(row=1) if isinstance(w, tk.Entry)]
            except Exception:
                widgets_after = []
            for w in widgets_after:
                if hasattr(w,'cget') and w.cget('textvariable') == str(self.dn_size_var):
                    w.focus_set(); break
        self.dn_product_combo.bind('<<ComboboxSelected>>', _product_chosen)

        lbls = ["מוצר","מידה","סוג בד","צבע בד","שם פרינט","כמות","הערה"]
        for i,lbl in enumerate(lbls):
            tk.Label(entry_bar, text=lbl, bg='#f7f9fa').grid(row=0,column=i*2,sticky='w',padx=2)

        self.dn_size_combo = ttk.Combobox(entry_bar, textvariable=self.dn_size_var, width=10, state='readonly')
        self.dn_fabric_type_combo = ttk.Combobox(entry_bar, textvariable=self.dn_fabric_type_var, width=12, state='readonly')
        self.dn_fabric_color_combo = ttk.Combobox(entry_bar, textvariable=self.dn_fabric_color_var, width=10, state='readonly')
        dn_print_entry = tk.Entry(entry_bar, textvariable=self.dn_print_name_var, width=12)

        widgets = [
            self.dn_product_combo,
            self.dn_size_combo,
            self.dn_fabric_type_combo,
            self.dn_fabric_color_combo,
            dn_print_entry,
            tk.Entry(entry_bar, textvariable=self.dn_qty_var, width=7),
            tk.Entry(entry_bar, textvariable=self.dn_note_var, width=18)
        ]
        for i,w in enumerate(widgets):
            w.grid(row=1,column=i*2,sticky='w',padx=2)

        def _on_product_change(*_a):
            try:
                self._update_delivery_size_options()
                self._update_delivery_fabric_type_options()
                self._update_delivery_fabric_color_options()
            except Exception:
                pass
        try:
            self.dn_product_var.trace_add('write', _on_product_change)
        except Exception:
            pass

        try:
            def _on_fabric_type_change(*_a):
                try:
                    self._update_delivery_fabric_color_options()
                except Exception:
                    pass
            self.dn_fabric_type_var.trace_add('write', _on_fabric_type_change)
        except Exception:
            pass

        for combo in (self.dn_size_combo, self.dn_fabric_type_combo, self.dn_fabric_color_combo):
            try: combo.state(['disabled'])
            except Exception: pass

        tk.Button(entry_bar, text="➕ הוסף", command=self._add_delivery_line, bg='#27ae60', fg='white').grid(row=1,column=14,padx=6)
        tk.Button(entry_bar, text="🗑️ מחק נבחר", command=self._delete_delivery_selected, bg='#e67e22', fg='white').grid(row=1,column=15,padx=4)
        tk.Button(entry_bar, text="❌ נקה הכל", command=self._clear_delivery_lines, bg='#e74c3c', fg='white').grid(row=1,column=16,padx=4)

        cols = ('product','size','fabric_type','fabric_color','print_name','quantity','note')
        self.delivery_tree = ttk.Treeview(lines_frame, columns=cols, show='headings', height=10)
        headers = {'product':'מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','print_name':'שם פרינט','quantity':'כמות','note':'הערה'}
        widths = {'product':160,'size':80,'fabric_type':110,'fabric_color':90,'print_name':110,'quantity':70,'note':220}
        for c in cols:
            self.delivery_tree.heading(c, text=headers[c])
            self.delivery_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(lines_frame, orient='vertical', command=self.delivery_tree.yview)
        self.delivery_tree.configure(yscroll=vs.set)
        self.delivery_tree.pack(side='left', fill='both', expand=True, padx=(4,0), pady=4)
        vs.pack(side='right', fill='y')

        # Transportation section (replaces packaging)
        pkg_frame = ttk.LabelFrame(container, text="הובלה", padding=8)
        pkg_frame.pack(fill='x', padx=10, pady=(4,4))
        self.pkg_type_var = tk.StringVar(value='שקית קטנה')
        self.pkg_qty_var = tk.StringVar()
        self.pkg_driver_var = tk.StringVar()
        tk.Label(pkg_frame, text="פריט הובלה:").grid(row=0,column=0,sticky='w',padx=4,pady=2)
        self.pkg_type_combo = ttk.Combobox(pkg_frame, textvariable=self.pkg_type_var, state='readonly', width=14, values=['שקית קטנה','שק','בד'])
        self.pkg_type_combo.grid(row=0,column=1,sticky='w',padx=4,pady=2)
        tk.Label(pkg_frame, text="כמות:").grid(row=0,column=2,sticky='w',padx=4,pady=2)
        tk.Entry(pkg_frame, textvariable=self.pkg_qty_var, width=8).grid(row=0,column=3,sticky='w',padx=4,pady=2)
        tk.Label(pkg_frame, text="מי מוביל:").grid(row=0,column=4,sticky='w',padx=4,pady=2)
        tk.Entry(pkg_frame, textvariable=self.pkg_driver_var, width=14).grid(row=0,column=5,sticky='w',padx=4,pady=2)
        tk.Button(pkg_frame, text="➕ הוסף", command=self._add_package_line, bg='#27ae60', fg='white').grid(row=0,column=6,padx=8)
        tk.Button(pkg_frame, text="🗑️ מחק נבחר", command=self._delete_selected_package, bg='#e67e22', fg='white').grid(row=0,column=7,padx=4)
        tk.Button(pkg_frame, text="❌ נקה", command=self._clear_packages, bg='#e74c3c', fg='white').grid(row=0,column=8,padx=4)
        self.packages_tree = ttk.Treeview(pkg_frame, columns=('type','quantity','driver'), show='headings', height=4)
        self.packages_tree.heading('type', text='פריט הובלה')
        self.packages_tree.heading('quantity', text='כמות')
        self.packages_tree.heading('driver', text='מי מוביל')
        self.packages_tree.column('type', width=120, anchor='center')
        self.packages_tree.column('quantity', width=70, anchor='center')
        self.packages_tree.column('driver', width=110, anchor='center')
        self.packages_tree.grid(row=1,column=0,columnspan=9, sticky='ew', padx=2, pady=(6,2))

        bottom_actions = tk.Frame(container, bg='#f7f9fa')
        bottom_actions.pack(fill='x', padx=10, pady=6)
        tk.Button(bottom_actions, text="💾 שמור תעודה", command=self._save_delivery_note, bg='#2c3e50', fg='white', font=('Arial',11,'bold')).pack(side='right', padx=4)
        self.delivery_summary_var = tk.StringVar(value="0 שורות | 0 כמות")
        tk.Label(container, textvariable=self.delivery_summary_var, bg='#34495e', fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')

        # Saved delivery notes list tab
        self.delivery_notes_tree = ttk.Treeview(list_wrapper, columns=('id','date','supplier','total','packages'), show='headings')
        for col, txt, w in (
            ('id','ID',60),('date','תאריך',110),('supplier','ספק',180),('total','סה"כ כמות',90),('packages','הובלה',140)
        ):
            self.delivery_notes_tree.heading(col, text=txt)
            self.delivery_notes_tree.column(col, width=w, anchor='center')
        vs2 = ttk.Scrollbar(list_wrapper, orient='vertical', command=self.delivery_notes_tree.yview)
        self.delivery_notes_tree.configure(yscroll=vs2.set)
        self.delivery_notes_tree.grid(row=0,column=0,sticky='nsew', padx=6, pady=6)
        vs2.grid(row=0,column=1,sticky='ns', pady=6)
        list_wrapper.grid_columnconfigure(0, weight=1)
        list_wrapper.grid_rowconfigure(0, weight=1)
        refresh_btn = tk.Button(list_wrapper, text="🔄 רענן", command=self._refresh_delivery_notes_list, bg='#3498db', fg='white')
        refresh_btn.grid(row=1,column=0,sticky='e', padx=6, pady=(0,6))
        self._refresh_delivery_notes_list()

        # Internal state lists
        self._delivery_lines = []
        self._packages = []

    # ---- Helpers (products / variants) ----
    def _refresh_delivery_products_allowed(self, initial: bool = False):
        try:
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
            names = sorted({ (rec.get('name') or '').strip() for rec in catalog if rec.get('name') })
            self._delivery_products_allowed = names
            self._delivery_products_allowed_full = list(names)
            if hasattr(self, 'dn_product_combo'):
                try:
                    cur = self.dn_product_var.get()
                    self.dn_product_combo['values'] = self._delivery_products_allowed_full
                    if cur and cur not in self._delivery_products_allowed_full:
                        self.dn_product_var.set('')
                except Exception: pass
        except Exception:
            self._delivery_products_allowed = []
            self._delivery_products_allowed_full = []

    def _update_delivery_size_options(self):
        product = (self.dn_product_var.get() or '').strip(); sizes = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        sz = (rec.get('size') or '').strip()
                        if sz: sizes.append(sz)
                import re
                def _size_key(s):
                    m = re.match(r"(\d+)", s); base = int(m.group(1)) if m else 0
                    return (base, s)
                sizes = sorted({s for s in sizes}, key=_size_key)
            except Exception: sizes = []
        if hasattr(self, 'dn_size_combo'):
            try:
                self.dn_size_combo['values'] = sizes
                if sizes:
                    self.dn_size_combo.state(['!disabled','readonly'])
                    if self.dn_size_var.get() not in sizes: self.dn_size_var.set('')
                else:
                    self.dn_size_var.set(''); self.dn_size_combo.set('')
                    try: self.dn_size_combo.state(['disabled'])
                    except Exception: pass
            except Exception: pass

    def _update_delivery_fabric_type_options(self):
        product = (self.dn_product_var.get() or '').strip(); fabric_types = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        ft = (rec.get('fabric_type') or '').strip()
                        if ft: fabric_types.append(ft)
                fabric_types = sorted({f for f in fabric_types})
            except Exception: fabric_types = []
        if hasattr(self, 'dn_fabric_type_combo'):
            try:
                self.dn_fabric_type_combo['values'] = fabric_types
                if fabric_types:
                    self.dn_fabric_type_combo.state(['!disabled','readonly'])
                    if self.dn_fabric_type_var.get() not in fabric_types: self.dn_fabric_type_var.set('')
                else:
                    self.dn_fabric_type_var.set(''); self.dn_fabric_type_combo.set('')
                    try: self.dn_fabric_type_combo.state(['disabled'])
                    except Exception: pass
            except Exception: pass

    def _update_delivery_fabric_color_options(self):
        product = (self.dn_product_var.get() or '').strip(); chosen_ft = (self.dn_fabric_type_var.get() or '').strip(); colors = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        if chosen_ft and (rec.get('fabric_type') or '').strip() != chosen_ft:
                            continue
                        col = (rec.get('fabric_color') or '').strip()
                        if col: colors.append(col)
                colors = sorted({c for c in colors})
            except Exception: colors = []
        if hasattr(self, 'dn_fabric_color_combo'):
            try:
                self.dn_fabric_color_combo['values'] = colors
                if colors:
                    self.dn_fabric_color_combo.state(['!disabled','readonly'])
                    if self.dn_fabric_color_var.get() not in colors:
                        self.dn_fabric_color_var.set('לבן' if 'לבן' in colors else '')
                else:
                    self.dn_fabric_color_var.set(''); self.dn_fabric_color_combo.set('')
                    try: self.dn_fabric_color_combo.state(['disabled'])
                    except Exception: pass
            except Exception: pass

    # ---- Lines ops ----
    def _add_delivery_line(self):
        product = self.dn_product_var.get().strip(); size = self.dn_size_var.get().strip(); qty_raw = self.dn_qty_var.get().strip(); note = self.dn_note_var.get().strip()
        fabric_type = self.dn_fabric_type_var.get().strip(); fabric_color = self.dn_fabric_color_var.get().strip(); print_name = self.dn_print_name_var.get().strip() or 'חלק'
        if not product or not qty_raw:
            messagebox.showerror("שגיאה", "חובה לבחור מוצר ולהזין כמות"); return
        if self._delivery_products_allowed and product not in self._delivery_products_allowed:
            messagebox.showerror("שגיאה", "יש לבחור מוצר מהרשימה בלבד"); return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי"); return
        line = {'product': product, 'size': size, 'fabric_type': fabric_type, 'fabric_color': fabric_color, 'print_name': print_name, 'quantity': qty, 'note': note}
        self._delivery_lines.append(line)
        self.delivery_tree.insert('', 'end', values=(product,size,fabric_type,fabric_color,print_name,qty,note))
        self.dn_size_var.set(''); self.dn_qty_var.set(''); self.dn_note_var.set('')
        try: self.dn_product_combo['values'] = self._delivery_products_allowed_full
        except Exception: pass
        self._update_delivery_summary()

    def _delete_delivery_selected(self):
        sel = self.delivery_tree.selection()
        if not sel: return
        all_items = self.delivery_tree.get_children(); indices = [all_items.index(i) for i in sel]
        for item in sel: self.delivery_tree.delete(item)
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._delivery_lines): del self._delivery_lines[idx]
        self._update_delivery_summary()

    def _clear_delivery_lines(self):
        self._delivery_lines = []
        for item in self.delivery_tree.get_children(): self.delivery_tree.delete(item)
        self._update_delivery_summary()

    def _update_delivery_summary(self):
        total_rows = len(self._delivery_lines); total_qty = sum(l.get('quantity',0) for l in self._delivery_lines)
        self.delivery_summary_var.set(f"{total_rows} שורות | {total_qty} כמות")

    def _save_delivery_note(self):
        supplier = self.dn_supplier_name_var.get().strip(); date_str = self.dn_date_var.get().strip()
        valid_names = set(self._get_supplier_names()) if hasattr(self,'_get_supplier_names') else set()
        if not supplier or (valid_names and supplier not in valid_names):
            messagebox.showerror("שגיאה", "יש לבחור שם ספק מהרשימה"); return
        if not self._delivery_lines:
            messagebox.showerror("שגיאה", "אין שורות לשמירה"); return
        # Confirm saving without any packages defined (parallel to supplier intake tab)
        try:
            if not self._packages:
                proceed = messagebox.askyesno(
                    "אישור",
                    "לא הוזנו צורות אריזה (שקיות / שקים / בדים).\nהאם לשמור את התעודה ללא כמות שקים?"
                )
                if not proceed:
                    return
        except Exception:
            pass
        try:
            # שימוש בשיטה החדשה לאחר פיצול הקבצים (עם נפילה אחורה)
            if hasattr(self.data_processor, 'add_delivery_note'):
                new_id = self.data_processor.add_delivery_note(supplier, date_str, self._delivery_lines, packages=self._packages)
            else:
                new_id = self.data_processor.add_supplier_receipt(supplier, date_str, self._delivery_lines, packages=self._packages, receipt_kind='delivery_note')
            messagebox.showinfo("הצלחה", f"תעודה נשמרה (ID: {new_id})")
            self._clear_delivery_lines()
            self._clear_packages()
            try:
                self._refresh_delivery_notes_list()
            except Exception:
                pass
            # עדכון טאב הובלות אם קיים
            try:
                self._notify_new_receipt_saved()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    # ---- Packages ops ----
    def _add_package_line(self):
        pkg_type = (self.pkg_type_var.get() or '').strip()
        qty_raw = (self.pkg_qty_var.get() or '').strip()
        if not pkg_type or not qty_raw:
            messagebox.showerror("שגיאה", "חובה לבחור פריט הובלה ולהזין כמות")
            return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי")
            return
        driver = (getattr(self, 'pkg_driver_var', tk.StringVar()).get() or '').strip()
        record = {'package_type': pkg_type, 'quantity': qty, 'driver': driver}
        self._packages.append(record)
        self.packages_tree.insert('', 'end', values=(pkg_type, qty, driver))
        self.pkg_qty_var.set('')

    def _delete_selected_package(self):
        sel = self.packages_tree.selection()
        if not sel: return
        all_items = self.packages_tree.get_children(); indices = [all_items.index(i) for i in sel]
        for item in sel: self.packages_tree.delete(item)
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._packages): del self._packages[idx]

    def _clear_packages(self):
        self._packages = []
        for item in self.packages_tree.get_children():
            self.packages_tree.delete(item)

    # ---- Delivery notes list ----
    def _refresh_delivery_notes_list(self):
        try:
            # Always refresh from disk to include external additions
            self.data_processor.refresh_supplier_receipts()
            notes = self.data_processor.get_delivery_notes()
        except Exception:
            notes = []
        # Clear
        if hasattr(self, 'delivery_notes_tree'):
            for iid in self.delivery_notes_tree.get_children():
                self.delivery_notes_tree.delete(iid)
            for rec in notes:
                pkg_summary = ', '.join(f"{p.get('package_type')}:{p.get('quantity')}" for p in rec.get('packages', [])[:4])
                if len(rec.get('packages', [])) > 4:
                    pkg_summary += ' ...'
                self.delivery_notes_tree.insert('', 'end', values=(rec.get('id'), rec.get('date'), rec.get('supplier'), rec.get('total_quantity'), pkg_summary))
