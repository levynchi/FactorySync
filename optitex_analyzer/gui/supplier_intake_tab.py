import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class SupplierIntakeTabMixin:
    """Mixin לטאב תעודת קליטה."""
    def _create_supplier_intake_tab(self):
        # Outer tab + title
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="תעודת קליטה")
        tk.Label(tab, text="תעודת קליטה (הזנה ידנית)", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)

        # Inner notebook (entry + saved receipts)
        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=4, pady=4)
        entry_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        list_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(entry_wrapper, text="קליטה")
        inner_nb.add(list_wrapper, text="קליטות שמורות")

        container = entry_wrapper

        # Receipt header form
        form = ttk.LabelFrame(container, text="פרטי קליטה", padding=10)
        form.pack(fill='x', padx=10, pady=6)
        tk.Label(form, text="שם ספק:", font=('Arial',10,'bold')).grid(row=0,column=0,sticky='w',padx=4,pady=4)
        self.supplier_name_var = tk.StringVar()
        # Combobox עם שמות ספקים קיימים בלבד
        self.supplier_name_combo = ttk.Combobox(form, textvariable=self.supplier_name_var, width=28, state='readonly')
        try:
            names = self._get_supplier_names() if hasattr(self, '_get_supplier_names') else []
            self.supplier_name_combo['values'] = names
        except Exception:
            pass
        self.supplier_name_combo.grid(row=0,column=1,sticky='w',padx=4,pady=4)
        tk.Label(form, text="תאריך:", font=('Arial',10,'bold')).grid(row=0,column=2,sticky='w',padx=4,pady=4)
        self.supplier_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(form, textvariable=self.supplier_date_var, width=15).grid(row=0,column=3,sticky='w',padx=4,pady=4)

        # Lines frame
        lines_frame = ttk.LabelFrame(container, text="שורות קליטה", padding=8)
        lines_frame.pack(fill='both', expand=True, padx=10, pady=4)

        # Entry bar
        entry_bar = tk.Frame(lines_frame, bg='#f7f9fa')
        entry_bar.pack(fill='x', pady=(0,6))

        # Variables (include new variant fields)
        self.sup_product_var = tk.StringVar()
        self.sup_size_var = tk.StringVar()
        self.sup_qty_var = tk.StringVar()
        self.sup_note_var = tk.StringVar()
        self.sup_fabric_type_var = tk.StringVar()
        self.sup_fabric_color_var = tk.StringVar(value='לבן')  # default
        self.sup_print_name_var = tk.StringVar(value='חלק')    # default

        # Load product list from catalog
        self._supplier_products_allowed = []
        self._refresh_supplier_products_allowed(initial=True)
        self._supplier_products_allowed_full = list(self._supplier_products_allowed)
        self.sup_product_combo = ttk.Combobox(entry_bar, textvariable=self.sup_product_var, width=16, state='normal')
        self.sup_product_combo['values'] = self._supplier_products_allowed_full

        # Popup state
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

        # Labels row
        lbls = ["מוצר","מידה","סוג בד","צבע בד","שם פרינט","כמות","הערה"]
        for i,lbl in enumerate(lbls):
            tk.Label(entry_bar, text=lbl, bg='#f7f9fa').grid(row=0,column=i*2,sticky='w',padx=2)

        # Size / Fabric Type / Fabric Color comboboxes (readonly, dynamic) + print name free text
        self.sup_size_combo = ttk.Combobox(entry_bar, textvariable=self.sup_size_var, width=10, state='readonly')
        self.sup_fabric_type_combo = ttk.Combobox(entry_bar, textvariable=self.sup_fabric_type_var, width=12, state='readonly')
        self.sup_fabric_color_combo = ttk.Combobox(entry_bar, textvariable=self.sup_fabric_color_var, width=10, state='readonly')
        print_name_entry = tk.Entry(entry_bar, textvariable=self.sup_print_name_var, width=12)  # נשאר חופשי כרגע

        widgets = [
            self.sup_product_combo,
            self.sup_size_combo,
            self.sup_fabric_type_combo,
            self.sup_fabric_color_combo,
            print_name_entry,
            tk.Entry(entry_bar, textvariable=self.sup_qty_var, width=7),
            tk.Entry(entry_bar, textvariable=self.sup_note_var, width=18)
        ]
        for i,w in enumerate(widgets):
            w.grid(row=1,column=i*2,sticky='w',padx=2)

        # Trace product -> update sizes
        def _on_product_change(*_a):
            try:
                self._update_supplier_size_options()
                self._update_supplier_fabric_type_options()
                self._update_supplier_fabric_color_options()
            except Exception:
                pass
        try:
            self.sup_product_var.trace_add('write', _on_product_change)
        except Exception:
            pass

        # Trace fabric type change -> update color options
        try:
            def _on_fabric_type_change(*_a):
                try:
                    self._update_supplier_fabric_color_options()
                except Exception:
                    pass
            self.sup_fabric_type_var.trace_add('write', _on_fabric_type_change)
        except Exception:
            pass

        # Initialize combo disabled until product chosen
        for combo in (self.sup_size_combo, self.sup_fabric_type_combo, self.sup_fabric_color_combo):
            try:
                combo.state(['disabled'])
            except Exception:
                pass

        # Row buttons
        tk.Button(entry_bar, text="➕ הוסף", command=self._add_supplier_line, bg='#27ae60', fg='white').grid(row=1,column=14,padx=6)
        tk.Button(entry_bar, text="🗑️ מחק נבחר", command=self._delete_supplier_selected, bg='#e67e22', fg='white').grid(row=1,column=15,padx=4)
        tk.Button(entry_bar, text="❌ נקה הכל", command=self._clear_supplier_lines, bg='#e74c3c', fg='white').grid(row=1,column=16,padx=4)

        # Tree with new columns
        cols = ('product','size','fabric_type','fabric_color','print_name','quantity','note')
        self.supplier_tree = ttk.Treeview(lines_frame, columns=cols, show='headings', height=10)
        headers = {'product':'מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','print_name':'שם פרינט','quantity':'כמות','note':'הערה'}
        widths = {'product':160,'size':80,'fabric_type':110,'fabric_color':90,'print_name':110,'quantity':70,'note':220}
        for c in cols:
            self.supplier_tree.heading(c, text=headers[c])
            self.supplier_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(lines_frame, orient='vertical', command=self.supplier_tree.yview)
        self.supplier_tree.configure(yscroll=vs.set)
        self.supplier_tree.pack(side='left', fill='both', expand=True, padx=(4,0), pady=4)
        vs.pack(side='right', fill='y', pady=4)

        # Save receipt button + summary
        bottom_actions = tk.Frame(container, bg='#f7f9fa')
        bottom_actions.pack(fill='x', padx=10, pady=6)
        tk.Button(bottom_actions, text="💾 שמור קליטה", command=self._save_supplier_receipt, bg='#2c3e50', fg='white', font=('Arial',11,'bold')).pack(side='right', padx=4)
        self.supplier_summary_var = tk.StringVar(value="0 שורות | 0 כמות")
        tk.Label(container, textvariable=self.supplier_summary_var, bg='#34495e', fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')

        # Internal lines model
        self._supplier_lines = []

        # ===== Transportation section (replaces previous packaging concept) =====
        pkg_frame = ttk.LabelFrame(container, text="הובלה", padding=8)
        pkg_frame.pack(fill='x', padx=10, pady=(4,4))
        self.sup_pkg_type_var = tk.StringVar(value='שקית קטנה')
        self.sup_pkg_qty_var = tk.StringVar()
        self.sup_pkg_driver_var = tk.StringVar()
        tk.Label(pkg_frame, text="פריט הובלה:").grid(row=0,column=0,sticky='w',padx=4,pady=2)
        self.sup_pkg_type_combo = ttk.Combobox(pkg_frame, textvariable=self.sup_pkg_type_var, state='readonly', width=14, values=['שקית קטנה','שק','בד'])
        self.sup_pkg_type_combo.grid(row=0,column=1,sticky='w',padx=4,pady=2)
        tk.Label(pkg_frame, text="כמות:").grid(row=0,column=2,sticky='w',padx=4,pady=2)
        tk.Entry(pkg_frame, textvariable=self.sup_pkg_qty_var, width=8).grid(row=0,column=3,sticky='w',padx=4,pady=2)
        tk.Label(pkg_frame, text="שם המוביל:").grid(row=0,column=4,sticky='w',padx=4,pady=2)
        # Combobox למובילים
        self.sup_pkg_driver_combo = ttk.Combobox(pkg_frame, textvariable=self.sup_pkg_driver_var, width=16, state='readonly')
        self.sup_pkg_driver_combo.grid(row=0,column=5,sticky='w',padx=4,pady=2)
        try:
            self._refresh_driver_names_for_intake()
        except Exception:
            pass
        tk.Button(pkg_frame, text="➕ הוסף", command=self._add_supplier_package_line, bg='#27ae60', fg='white').grid(row=0,column=6,padx=8)
        tk.Button(pkg_frame, text="🗑️ מחק נבחר", command=self._delete_selected_supplier_package, bg='#e67e22', fg='white').grid(row=0,column=7,padx=4)
        tk.Button(pkg_frame, text="❌ נקה", command=self._clear_supplier_packages, bg='#e74c3c', fg='white').grid(row=0,column=8,padx=4)
        self.sup_packages_tree = ttk.Treeview(pkg_frame, columns=('type','quantity','driver'), show='headings', height=4)
        self.sup_packages_tree.heading('type', text='פריט הובלה')
        self.sup_packages_tree.heading('quantity', text='כמות')
        self.sup_packages_tree.heading('driver', text='שם המוביל')
        self.sup_packages_tree.column('type', width=120, anchor='center')
        self.sup_packages_tree.column('quantity', width=70, anchor='center')
        self.sup_packages_tree.column('driver', width=110, anchor='center')
        self.sup_packages_tree.grid(row=1,column=0,columnspan=9, sticky='ew', padx=2, pady=(6,2))
        # Internal packages model
        self._supplier_packages = []

        # Saved supplier receipts list (second tab)
        self.supplier_receipts_tree = ttk.Treeview(list_wrapper, columns=('id','date','supplier','total','packages'), show='headings')
        for col, txt, w in (
            ('id','ID',60), ('date','תאריך',110), ('supplier','ספק',180), ('total','סה"כ כמות',90), ('packages','הובלה',140)
        ):
            self.supplier_receipts_tree.heading(col, text=txt)
            self.supplier_receipts_tree.column(col, width=w, anchor='center')
        vs_sr = ttk.Scrollbar(list_wrapper, orient='vertical', command=self.supplier_receipts_tree.yview)
        self.supplier_receipts_tree.configure(yscroll=vs_sr.set)
        self.supplier_receipts_tree.grid(row=0,column=0,sticky='nsew', padx=6, pady=6)
        vs_sr.grid(row=0,column=1,sticky='ns', pady=6)
        list_wrapper.grid_columnconfigure(0, weight=1)
        list_wrapper.grid_rowconfigure(0, weight=1)
        # פתיחת פרטי תעודה בלחיצה כפולה/Enter
        try:
            self.supplier_receipts_tree.bind('<Double-1>', lambda e: self._open_supplier_receipt_details())
            self.supplier_receipts_tree.bind('<Return>', lambda e: self._open_supplier_receipt_details())
        except Exception:
            pass
        tk.Button(list_wrapper, text="🔄 רענן", command=self._refresh_supplier_intake_list, bg='#3498db', fg='white').grid(row=1,column=0,sticky='e', padx=6, pady=(0,6))
        self._refresh_supplier_intake_list()

    def _refresh_driver_names_for_intake(self):
        """טעינת שמות המובילים מקובץ drivers.json והחלתם בקומבובוקס."""
        try:
            import json, os
            path = os.path.join(os.getcwd(), 'drivers.json')
            names = []
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        for d in data:
                            name = (d.get('name') or '').strip()
                            if name:
                                names.append(name)
            names = sorted({n for n in names})
            if hasattr(self, 'sup_pkg_driver_combo'):
                self.sup_pkg_driver_combo['values'] = names
        except Exception:
            pass

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

    def _update_supplier_fabric_type_options(self):
        """עדכון רשימת סוגי הבד הזמינים עבור המוצר הנבחר (מהווריאנטים בקטלוג)."""
        product = (self.sup_product_var.get() or '').strip()
        fabric_types = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        ft = (rec.get('fabric_type') or '').strip()
                        if ft:
                            fabric_types.append(ft)
                fabric_types = sorted({f for f in fabric_types})
            except Exception:
                fabric_types = []
        if hasattr(self, 'sup_fabric_type_combo'):
            try:
                self.sup_fabric_type_combo['values'] = fabric_types
                if fabric_types:
                    self.sup_fabric_type_combo.state(['!disabled','readonly'])
                    if self.sup_fabric_type_var.get() not in fabric_types:
                        self.sup_fabric_type_var.set('')
                else:
                    self.sup_fabric_type_var.set('')
                    self.sup_fabric_type_combo.set('')
                    try:
                        self.sup_fabric_type_combo.state(['disabled'])
                    except Exception:
                        pass
            except Exception:
                pass

    def _update_supplier_fabric_color_options(self):
        """עדכון רשימת צבעי הבד הזמינים עבור המוצר / סוג בד שנבחרו."""
        product = (self.sup_product_var.get() or '').strip()
        chosen_ft = (self.sup_fabric_type_var.get() or '').strip()
        colors = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        # אם בחר סוג בד, נסנן לפי סוג בד. אם לא – ניקח את כל הצבעים של המוצר.
                        if chosen_ft:
                            if (rec.get('fabric_type') or '').strip() != chosen_ft:
                                continue
                        col = (rec.get('fabric_color') or '').strip()
                        if col:
                            colors.append(col)
                colors = sorted({c for c in colors})
            except Exception:
                colors = []
        if hasattr(self, 'sup_fabric_color_combo'):
            try:
                self.sup_fabric_color_combo['values'] = colors
                if colors:
                    self.sup_fabric_color_combo.state(['!disabled','readonly'])
                    if self.sup_fabric_color_var.get() not in colors:
                        # שמור ברירת מחדל 'לבן' אם קיימת באפשרויות אחרת ננקה
                        self.sup_fabric_color_var.set('לבן' if 'לבן' in colors else '')
                else:
                    self.sup_fabric_color_var.set('')
                    self.sup_fabric_color_combo.set('')
                    try:
                        self.sup_fabric_color_combo.state(['disabled'])
                    except Exception:
                        pass
            except Exception:
                pass

    def _add_supplier_line(self):
        product = self.sup_product_var.get().strip(); size = self.sup_size_var.get().strip(); qty_raw = self.sup_qty_var.get().strip(); note = self.sup_note_var.get().strip()
        fabric_type = self.sup_fabric_type_var.get().strip(); fabric_color = self.sup_fabric_color_var.get().strip(); print_name = self.sup_print_name_var.get().strip() or 'חלק'
        if not product or not qty_raw:
            messagebox.showerror("שגיאה", "חובה לבחור מוצר ולהזין כמות")
            return
        if hasattr(self, '_supplier_products_allowed') and self._supplier_products_allowed and product not in self._supplier_products_allowed:
            messagebox.showerror("שגיאה", "יש לבחור מוצר מהרשימה בלבד"); return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי"); return
        line = {'product': product, 'size': size, 'fabric_type': fabric_type, 'fabric_color': fabric_color, 'print_name': print_name, 'quantity': qty, 'note': note}
        self._supplier_lines.append(line)
        self.supplier_tree.insert('', 'end', values=(product,size,fabric_type,fabric_color,print_name,qty,note))
        # reset only size/qty/note (שומרים סוג/צבע/פרינט להמשך הזנה מהירה)
        self.sup_size_var.set(''); self.sup_qty_var.set(''); self.sup_note_var.set('')
        # Restore full suggestions list after adding a line
        if hasattr(self, 'sup_product_combo'):
            try:
                self.sup_product_combo['values'] = self._supplier_products_allowed_full
            except Exception:
                pass
        self._update_supplier_summary()

    def _delete_supplier_selected(self):
        sel = self.supplier_tree.selection()
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
        total_rows = len(self._supplier_lines)
        total_qty = sum(l.get('quantity',0) for l in self._supplier_lines)
        self.supplier_summary_var.set(f"{total_rows} שורות | {total_qty} כמות")

    def _save_supplier_receipt(self):
        supplier = self.supplier_name_var.get().strip(); date_str = self.supplier_date_var.get().strip()
        # אימות נגד רשימת הספקים
        valid_names = set(self._get_supplier_names()) if hasattr(self,'_get_supplier_names') else set()
        if not supplier or (valid_names and supplier not in valid_names):
            messagebox.showerror("שגיאה", "יש לבחור שם ספק מהרשימה"); return
        if not self._supplier_lines: messagebox.showerror("שגיאה", "אין שורות לשמירה"); return
        # אם אין כלל פריטי הובלה – בקשת אישור מהמשתמש לפני המשך שמירה
        try:
            if not self._supplier_packages:
                proceed = messagebox.askyesno(
                    "אישור",
                    "לא הוזנו פריטי הובלה (שקיות / שקים / בדים).\nהאם לשמור את הקליטה ללא פריטי הובלה?"
                )
                if not proceed:
                    return
        except Exception:
            pass
        try:
            # שימוש בשיטה החדשה לאחר פיצול הקבצים
            if hasattr(self.data_processor, 'add_supplier_intake'):
                new_id = self.data_processor.add_supplier_intake(supplier, date_str, self._supplier_lines, packages=self._supplier_packages)
            else:
                # תאימות לאחור
                new_id = self.data_processor.add_supplier_receipt(supplier, date_str, self._supplier_lines, packages=self._supplier_packages, receipt_kind='supplier_intake')
            messagebox.showinfo("הצלחה", f"קליטה נשמרה (ID: {new_id})")
            self._clear_supplier_lines()
            self._clear_supplier_packages()
            try:
                self._refresh_supplier_intake_list()
            except Exception:
                pass
            # עדכון טאב הובלות אם קיים
            try:
                self._notify_new_receipt_saved()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
            if hasattr(self.data_processor, 'add_supplier_intake'):
                new_id = self.data_processor.add_supplier_intake(supplier, date_str, self._supplier_lines, packages=self._supplier_packages)
            else:
                # תאימות לאחור
                new_id = self.data_processor.add_supplier_receipt(supplier, date_str, self._supplier_lines, packages=self._supplier_packages, receipt_kind='supplier_intake')
            messagebox.showinfo("הצלחה", f"קליטה נשמרה (ID: {new_id})")
            self._clear_supplier_lines()
            self._clear_supplier_packages()
            try:
                self._refresh_supplier_intake_list()
            except Exception:
                pass
            # עדכון טאב הובלות אם קיים
            try:
                self._notify_new_receipt_saved()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    # ===== Packaging methods (parallel to delivery note) =====
    def _add_supplier_package_line(self):
        pkg_type = (self.sup_pkg_type_var.get() or '').strip()
        qty_raw = (self.sup_pkg_qty_var.get() or '').strip()
        if not pkg_type or not qty_raw:
            messagebox.showerror("שגיאה", "חובה לבחור פריט הובלה ולהזין כמות")
            return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי")
            return
        driver = (getattr(self, 'sup_pkg_driver_var', tk.StringVar()).get() or '').strip()
        record = {'package_type': pkg_type, 'quantity': qty, 'driver': driver}
        self._supplier_packages.append(record)
        self.sup_packages_tree.insert('', 'end', values=(pkg_type, qty, driver))
        self.sup_pkg_qty_var.set('')

    def _delete_selected_supplier_package(self):
        sel = self.sup_packages_tree.selection()
        if not sel: return
        all_items = self.sup_packages_tree.get_children(); indices = [all_items.index(i) for i in sel]
        for item in sel: self.sup_packages_tree.delete(item)
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._supplier_packages): del self._supplier_packages[idx]

    def _clear_supplier_packages(self):
        self._supplier_packages = []
        for item in self.sup_packages_tree.get_children():
            self.sup_packages_tree.delete(item)
    # ---- Saved supplier receipts list ----
    def _refresh_supplier_intake_list(self):
        try:
            self.data_processor.refresh_supplier_receipts()
            receipts = [r for r in self.data_processor.supplier_receipts if r.get('receipt_kind') == 'supplier_intake']
        except Exception:
            receipts = []
        if hasattr(self, 'supplier_receipts_tree'):
            for iid in self.supplier_receipts_tree.get_children():
                self.supplier_receipts_tree.delete(iid)
            for rec in receipts:
                pkg_summary = ', '.join(f"{p.get('package_type')}:{p.get('quantity')}" for p in rec.get('packages', [])[:4])
                if len(rec.get('packages', [])) > 4:
                    pkg_summary += ' ...'
                self.supplier_receipts_tree.insert('', 'end', values=(rec.get('id'), rec.get('date'), rec.get('supplier'), rec.get('total_quantity'), pkg_summary))

    def _open_supplier_receipt_details(self):
        """פתיחת חלון פרטים עבור תעודת קליטה נבחרת מתוך 'קליטות שמורות'."""
        try:
            sel = self.supplier_receipts_tree.selection()
            if not sel:
                return
            values = self.supplier_receipts_tree.item(sel[0], 'values') or []
            if not values:
                return
            rec_id = int(values[0])
        except Exception:
            return
        # מציאת הרשומה
        try:
            self.data_processor.refresh_supplier_receipts()
        except Exception:
            pass
        rec = None
        try:
            for r in getattr(self.data_processor, 'supplier_receipts', []) or []:
                if r.get('receipt_kind') == 'supplier_intake' and int(r.get('id', -1)) == rec_id:
                    rec = r; break
        except Exception:
            rec = None
        if not rec:
            try:
                messagebox.showwarning("שגיאה", "הרשומה לא נמצאה")
            except Exception:
                pass
            return
        # חלון פרטים
        win = tk.Toplevel(self.root)
        win.title(f"תעודת קליטה #{rec_id}")
        win.transient(self.root)
        try:
            win.grab_set()
        except Exception:
            pass
        header = tk.Frame(win, bg='#f7f9fa')
        header.pack(fill='x', padx=10, pady=8)
        tk.Label(header, text=f"ספק: {rec.get('supplier','')}", bg='#f7f9fa', font=('Arial',11,'bold')).pack(side='right', padx=8)
        tk.Label(header, text=f"תאריך: {rec.get('date','')}", bg='#f7f9fa').pack(side='right', padx=8)
        tk.Label(header, text=f"ID: {rec.get('id','')}", bg='#f7f9fa').pack(side='right', padx=8)
        tk.Label(header, text=f"סה\"כ כמות: {rec.get('total_quantity',0)}", bg='#f7f9fa').pack(side='right', padx=8)

        body = tk.Frame(win, bg='#f7f9fa')
        body.pack(fill='both', expand=True, padx=10, pady=(0,10))

        # שורות מוצר
        lines_frame = tk.LabelFrame(body, text='שורות תעודה', bg='#f7f9fa')
        lines_frame.pack(fill='both', expand=True, pady=6)
        cols = ('product','size','fabric_type','fabric_color','print_name','quantity','note')
        tree = ttk.Treeview(lines_frame, columns=cols, show='headings', height=8)
        headers = {'product':'מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','print_name':'שם פרינט','quantity':'כמות','note':'הערה'}
        widths = {'product':180,'size':80,'fabric_type':100,'fabric_color':90,'print_name':110,'quantity':70,'note':220}
        for c in cols:
            tree.heading(c, text=headers[c])
            tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(lines_frame, orient='vertical', command=tree.yview)
        tree.configure(yscroll=vs.set)
        tree.pack(side='left', fill='both', expand=True, padx=(6,0), pady=6)
        vs.pack(side='left', fill='y', pady=6)
        for ln in rec.get('lines', []) or []:
            tree.insert('', 'end', values=(
                ln.get('product',''), ln.get('size',''), ln.get('fabric_type',''), ln.get('fabric_color',''),
                ln.get('print_name',''), ln.get('quantity',''), ln.get('note','')
            ))

        # פריטי הובלה
        pk_frame = tk.LabelFrame(body, text='פריטי הובלה', bg='#f7f9fa')
        pk_frame.pack(fill='x', pady=6)
        pk_cols = ('type','quantity','driver')
        pk_tree = ttk.Treeview(pk_frame, columns=pk_cols, show='headings', height=4)
        pk_tree.heading('type', text='פריט הובלה')
        pk_tree.heading('quantity', text='כמות')
        pk_tree.heading('driver', text='שם המוביל')
        pk_tree.column('type', width=120, anchor='center')
        pk_tree.column('quantity', width=70, anchor='center')
        pk_tree.column('driver', width=110, anchor='center')
        pk_tree.pack(fill='x', padx=6, pady=6)
        for p in rec.get('packages', []) or []:
            pk_tree.insert('', 'end', values=(p.get('package_type',''), p.get('quantity',''), p.get('driver','')))

        btns = tk.Frame(win, bg='#f7f9fa')
        btns.pack(fill='x', padx=10, pady=(0,10))
        tk.Button(btns, text='סגור', command=win.destroy).pack(side='left')
