import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class ProductsBalanceTabMixin:
    """Mixin לטאב 'מאזן מוצרים ופריטים'.

    כולל בחירת ספק וטאב פנימי 'מאזן מוצרים' המציג סיכום שנשלח מול שנתקבל לפי מוצר.
    לעת עתה מאזן לפי שם מוצר בלבד (ללא פירוט וריאנטים). ניתן להרחיב בעתיד.
    """

    def _create_products_balance_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="מאזן מוצרים ופריטים")

        # סרגל מסננים
        toolbar = tk.Frame(tab, bg='#f7f9fa')
        toolbar.pack(fill='x', padx=8, pady=(8,4))
        tk.Label(toolbar, text='ספק:', bg='#f7f9fa', font=('Arial',10,'bold')).pack(side='right', padx=(6,2))
        self.balance_supplier_var = tk.StringVar()
        self.balance_supplier_combo = ttk.Combobox(toolbar, textvariable=self.balance_supplier_var, width=28, state='readonly')
        try:
            if hasattr(self, '_get_supplier_names'):
                self.balance_supplier_combo['values'] = self._get_supplier_names()
        except Exception:
            pass
        self.balance_supplier_combo.pack(side='right')
        tk.Button(toolbar, text='🔄 רענן', command=self._refresh_products_balance_table, bg='#3498db', fg='white').pack(side='right', padx=6)

        # מסנן רק-חוסר נשאר בסרגל העליון
        self.balance_only_pending_var = tk.BooleanVar(value=False)
        tk.Checkbutton(toolbar, text='רק חוסר', variable=self.balance_only_pending_var, bg='#f7f9fa', command=self._refresh_products_balance_table).pack(side='right', padx=(10,0))

        # פנימי: נוטבוק לעתיד (כרגע עמוד אחד – מאזן מוצרים)
        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=8, pady=8)

        balance_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(balance_page, text='מאזן מוצרים')
        tk.Label(balance_page, text='מאזן מוצרים לפי ספק', font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=6)

        # סרגל פנימי: חיפוש + כפתור פירוט לפי מידות
        inner_bar = tk.Frame(balance_page, bg='#f7f9fa')
        inner_bar.pack(fill='x', padx=10, pady=(0,6))
        tk.Label(inner_bar, text='חיפוש:', bg='#f7f9fa').pack(side='right', padx=(8,4))
        self.balance_search_var = tk.StringVar()
        search_entry = tk.Entry(inner_bar, textvariable=self.balance_search_var, width=24)
        search_entry.pack(side='right', padx=(0,6))
        try:
            search_entry.bind('<KeyRelease>', lambda e: self._refresh_products_balance_table())
        except Exception:
            pass
        # מצב פירוט לפי מידות (כפתור טוגול)
        self._balance_detail_by_size = False
        self._balance_toggle_btn = tk.Button(inner_bar, text='פירוט לפי מידות', command=self._toggle_balance_detail_mode, bg='#8e44ad', fg='white')
        self._balance_toggle_btn.pack(side='left')

        columns = ('product','shipped','received','diff','status')
        self.products_balance_tree = ttk.Treeview(balance_page, columns=columns, show='headings', height=18)
        headers = {
            'product': 'מוצר',
            'shipped': 'נשלח',
            'received': 'נתקבל',
            'diff': 'הפרש (נותר לקבל)',
            'status': 'סטטוס'
        }
        widths = {'product':260, 'shipped':90, 'received':90, 'diff':130, 'status':140}
        for c in columns:
            self.products_balance_tree.heading(c, text=headers[c])
            self.products_balance_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(balance_page, orient='vertical', command=self.products_balance_tree.yview)
        self.products_balance_tree.configure(yscroll=vs.set)
        self.products_balance_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=6)
        vs.pack(side='left', fill='y', pady=6)

        # ריענון אוטומטי בעת שינוי ספק
        try:
            self.balance_supplier_var.trace_add('write', lambda *_: self._refresh_products_balance_table())
        except Exception:
            pass

        # טעינה ראשונית – ריק עד בחירת ספק
        self._refresh_products_balance_table()

    def _refresh_products_balance_table(self):
        supplier = (getattr(self, 'balance_supplier_var', None).get() if hasattr(self, 'balance_supplier_var') else '') or ''
        # ניקוי טבלה
        if hasattr(self, 'products_balance_tree'):
            for iid in self.products_balance_tree.get_children():
                self.products_balance_tree.delete(iid)
        # אם אין ספק נבחר – לא מציגים כלום
        if not supplier:
            return
        try:
            # לוודא נתונים עדכניים
            if hasattr(self.data_processor, 'refresh_supplier_receipts'):
                self.data_processor.refresh_supplier_receipts()
        except Exception:
            pass
        # איסוף כמויות לפי מוצר או לפי (מוצר, מידה)
        by_size = getattr(self, '_balance_detail_by_size', False)
        shipped = {}
        received = {}
        try:
            notes = getattr(self.data_processor, 'delivery_notes', []) or []
            for rec in notes:
                if (rec.get('supplier') or '').strip() != supplier:
                    continue
                for ln in rec.get('lines', []) or []:
                    name = (ln.get('product') or '').strip()
                    size = (ln.get('size') or '').strip()
                    qty = int(ln.get('quantity', 0) or 0)
                    if not name or qty <= 0:
                        continue
                    key = (name, size if by_size else '') if by_size else name
                    shipped[key] = shipped.get(key, 0) + qty
        except Exception:
            pass
        try:
            intakes = getattr(self.data_processor, 'supplier_intakes', []) or []
            for rec in intakes:
                if (rec.get('supplier') or '').strip() != supplier:
                    continue
                for ln in rec.get('lines', []) or []:
                    name = (ln.get('product') or '').strip()
                    size = (ln.get('size') or '').strip()
                    qty = int(ln.get('quantity', 0) or 0)
                    if not name or qty <= 0:
                        continue
                    key = (name, size if by_size else '') if by_size else name
                    received[key] = received.get(key, 0) + qty
        except Exception:
            pass
        # בניית שורות מאוחדות + החלת מסננים
        keys = sorted(set(list(shipped.keys()) + list(received.keys())))
        search_txt = (getattr(self, 'balance_search_var', tk.StringVar()).get() or '').strip().lower()
        only_pending = bool(getattr(self, 'balance_only_pending_var', tk.BooleanVar()).get())
        for key in keys:
            if by_size:
                prod, size = key if isinstance(key, tuple) else (str(key), '')
                label = f"{prod} – {size or 'ללא מידה'}"
            else:
                label = key if isinstance(key, str) else key[0]
            s = shipped.get(key, 0)
            r = received.get(key, 0)
            diff = s - r
            # סינון טקסטואלי
            if search_txt:
                if search_txt not in (label or '').lower():
                    continue
            # סינון רק חוסר
            if only_pending and diff <= 0:
                continue
            status = 'הושלם' if diff <= 0 else f"נותרו {diff} לקבל"
            self.products_balance_tree.insert('', 'end', values=(label, s, r, max(diff, 0), status))

    def _toggle_balance_detail_mode(self):
        """טוגול תצוגה בין מאוחד לפי מוצר לבין פירוט לפי מידות."""
        self._balance_detail_by_size = not getattr(self, '_balance_detail_by_size', False)
        try:
            if self._balance_detail_by_size:
                self._balance_toggle_btn.config(text='תצוגת מוצר בלבד')
            else:
                self._balance_toggle_btn.config(text='פירוט לפי מידות')
        except Exception:
            pass
        self._refresh_products_balance_table()
