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
        tk.Button(toolbar, text='🔄 רענן', command=self._refresh_balance_views, bg='#3498db', fg='white').pack(side='right', padx=6)

        # מסנן רק-חוסר נשאר בסרגל העליון
        self.balance_only_pending_var = tk.BooleanVar(value=False)
        tk.Checkbutton(toolbar, text='רק חוסר', variable=self.balance_only_pending_var, bg='#f7f9fa', command=self._refresh_products_balance_table).pack(side='right', padx=(10,0))

        # פנימי: נוטבוק עם 2 עמודים – מאזן מוצרים + מה נגזר אצל הספק
        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=8, pady=8)

        balance_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(balance_page, text='מאזן מוצרים')
        tk.Label(balance_page, text='מאזן מוצרים לפי ספק', font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=6)

        cut_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(cut_page, text='מה נגזר אצל הספק')
        tk.Label(cut_page, text='מה נגזר אצל הספק (מחושב מציורים "נחתך" × שכבות)', font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=6)
        # סרגל פנימי לעמוד הגזירה: חיפוש + פירוט לפי מידות
        cut_bar = tk.Frame(cut_page, bg='#f7f9fa'); cut_bar.pack(fill='x', padx=10, pady=(0,6))
        tk.Label(cut_bar, text='חיפוש:', bg='#f7f9fa').pack(side='right', padx=(8,4))
        self.cut_search_var = tk.StringVar(); cut_search = tk.Entry(cut_bar, textvariable=self.cut_search_var, width=24); cut_search.pack(side='right', padx=(0,6))
        try:
            cut_search.bind('<KeyRelease>', lambda e: self._refresh_supplier_cut_table())
        except Exception:
            pass
        self._cut_detail_by_size = True
        self._cut_toggle_btn = tk.Button(cut_bar, text='תצוגת מוצר בלבד', command=self._toggle_cut_detail_mode, bg='#8e44ad', fg='white')
        self._cut_toggle_btn.pack(side='left')
        cols_cut = ('product','size','fabric','cut_qty')
        self.supplier_cut_tree = ttk.Treeview(cut_page, columns=cols_cut, show='headings', height=18)
        headers_cut = {'product':'מוצר','size':'מידה','fabric':'סוג בד','cut_qty':'נגזר (יח׳)'}
        widths_cut = {'product':250,'size':90,'fabric':200,'cut_qty':100}
        for c in cols_cut:
            self.supplier_cut_tree.heading(c, text=headers_cut[c])
            self.supplier_cut_tree.column(c, width=widths_cut[c], anchor='center')
        vs2 = ttk.Scrollbar(cut_page, orient='vertical', command=self.supplier_cut_tree.yview)
        self.supplier_cut_tree.configure(yscroll=vs2.set)
        self.supplier_cut_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=6)
        vs2.pack(side='left', fill='y', pady=6)

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
        # הוספת אפשרות לכלול "מה נגזר אצל הספק" לתוך העמודה "נשלח"
        self.include_cuts_in_shipped_var = tk.BooleanVar(value=False)
        tk.Checkbutton(inner_bar, text="הוסף נגזר ל'נשלח'", variable=self.include_cuts_in_shipped_var, bg='#f7f9fa', command=self._refresh_products_balance_table).pack(side='left', padx=(8,0))

        columns = ('product','fabric_type','fabric_color','fabric_category','print_name','shipped','received','diff','status')
        self.products_balance_tree = ttk.Treeview(balance_page, columns=columns, show='headings', height=18)
        headers = {
            'product': 'מוצר',
            'fabric_type': 'סוג בד',
            'fabric_color': 'צבע בד',
            'fabric_category': 'קטגורית בד',
            'print_name': 'שם פרינט',
            'shipped': 'נשלח',
            'received': 'נתקבל',
            'diff': 'הפרש (נותר לקבל)',
            'status': 'סטטוס'
        }
        widths = {'product':220, 'fabric_type':120, 'fabric_color':120, 'fabric_category':120, 'print_name':120, 'shipped':90, 'received':90, 'diff':130, 'status':140}
        for c in columns:
            self.products_balance_tree.heading(c, text=headers[c])
            self.products_balance_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(balance_page, orient='vertical', command=self.products_balance_tree.yview)
        self.products_balance_tree.configure(yscroll=vs.set)
        self.products_balance_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=6)
        vs.pack(side='left', fill='y', pady=6)

        # ריענון אוטומטי בעת שינוי ספק
        try:
            self.balance_supplier_var.trace_add('write', lambda *_: self._refresh_balance_views())
        except Exception:
            pass

        # טעינה ראשונית – ריק עד בחירת ספק
        self._refresh_balance_views()

    def _refresh_balance_views(self):
        """רענון שני המצבים: מאזן מוצרים ומה נגזר אצל הספק."""
        try:
            self._refresh_products_balance_table()
        except Exception:
            pass
        try:
            self._refresh_supplier_cut_table()
        except Exception:
            pass

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
        include_cuts = bool(getattr(self, 'include_cuts_in_shipped_var', tk.BooleanVar(value=False)).get())
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

        # במידת הצורך, הוספת "מה נגזר אצל הספק" לעמודת "נשלח"
        cut_totals = {}
        if include_cuts:
            try:
                if hasattr(self.data_processor, 'refresh_drawings_data'):
                    self.data_processor.refresh_drawings_data()
            except Exception:
                pass
            try:
                for rec in getattr(self.data_processor, 'drawings_data', []) or []:
                    if rec.get('status') != 'נחתך':
                        continue
                    if (rec.get('נמען') or '').strip() != supplier:
                        continue
                    layers = rec.get('שכבות')
                    try:
                        layers = int(layers)
                    except Exception:
                        layers = None
                    if not layers or layers <= 0:
                        continue
                    for prod in rec.get('מוצרים', []) or []:
                        pname = (prod.get('שם המוצר') or '').strip()
                        for sz in prod.get('מידות', []) or []:
                            size = (sz.get('מידה') or '').strip()
                            qty = int(sz.get('כמות', 0) or 0)
                            if not pname or qty <= 0:
                                continue
                            key = (pname, size if by_size else '') if by_size else pname
                            cut_totals[key] = cut_totals.get(key, 0) + qty * layers
                # מיזוג לספירת הנשלח
                for key, add_qty in cut_totals.items():
                    shipped[key] = shipped.get(key, 0) + add_qty
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
                p_name = prod
                p_size = size
            else:
                label = key if isinstance(key, str) else key[0]
                p_name = label
                p_size = ''
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
            # איסוף מאפייני מוצר (סוג/צבע/פרינט) מהקטלוג
            f_type, f_color, p_print, f_cat = self._get_product_attrs(p_name, p_size, by_size)
            # אם המשתמש ביקש להוסיף נגזר ל"נשלח" ויש נגזר עבור המוצר הזה – נשתמש בקטגוריית בד קבועה "טריקו לבן"
            if include_cuts:
                key_for_cut = (p_name, p_size) if by_size else p_name
                if cut_totals.get(key_for_cut, 0) > 0:
                    f_cat = 'טריקו לבן'
            self.products_balance_tree.insert('', 'end', values=(label, f_type, f_color, f_cat, p_print, s, r, max(diff, 0), status))

    def _get_product_attrs(self, product_name: str, size: str = '', by_size: bool = False):
        """החזרת (סוג בד, צבע בד, שם פרינט, קטגורית בד) מתוך קטלוג המוצרים עבור מוצר (ולפי מידה אם נדרש).

        אם יש כמה ערכים שונים בקטלוג – נבחר את הערך הנפוץ ביותר (majority) כדי להציג ערך עקבי.
        אם אין נתונים – יוחזרו מחרוזות ריקות.
        """
        try:
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
            if not product_name:
                return ('', '', '', '')
            def _norm(s):
                return (s or '').strip()
            if by_size and size:
                items = [r for r in catalog if _norm(r.get('name')) == _norm(product_name) and _norm(r.get('size')) == _norm(size)]
            else:
                items = [r for r in catalog if _norm(r.get('name')) == _norm(product_name)]
            if not items:
                return ('', '', '', '')
            def _majority(values: list[str]) -> str:
                counts = {}
                for v in values:
                    v2 = _norm(v)
                    if not v2:
                        continue
                    counts[v2] = counts.get(v2, 0) + 1
                if not counts:
                    return ''
                # בחר את הערך עם המספר הגבוה ביותר; אם תיקו – הראשון בסדר הופעה טבעי
                return sorted(counts.items(), key=lambda x: (-x[1], x[0]))[0][0]

            f_type = _majority([r.get('fabric_type','') for r in items])
            f_color = _majority([r.get('fabric_color','') for r in items])
            p_print = _majority([r.get('print_name','') for r in items])
            f_cat = _majority([r.get('fabric_category','') for r in items])
            return (f_type, f_color, p_print, f_cat)
        except Exception:
            return ('', '', '', '')

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

    # === מה נגזר אצל הספק ===
    def _toggle_cut_detail_mode(self):
        self._cut_detail_by_size = not getattr(self, '_cut_detail_by_size', True)
        try:
            self._cut_toggle_btn.config(text='פירוט לפי מידות' if not self._cut_detail_by_size else 'תצוגת מוצר בלבד')
        except Exception:
            pass
        self._refresh_supplier_cut_table()

    def _refresh_supplier_cut_table(self):
        supplier = (getattr(self, 'balance_supplier_var', None).get() if hasattr(self, 'balance_supplier_var') else '') or ''
        # ניקוי
        if hasattr(self, 'supplier_cut_tree'):
            for iid in self.supplier_cut_tree.get_children():
                self.supplier_cut_tree.delete(iid)
        if not supplier:
            return
        # ודא שיש נתוני ציורים עדכניים
        try:
            if hasattr(self.data_processor, 'refresh_drawings_data'):
                self.data_processor.refresh_drawings_data()
        except Exception:
            pass
        # צבירה: לכל ציור בסטטוס נחתך, עם נמען = supplier, נחשב לכל מוצר/מידה: כמות*שכבות
        by_size = getattr(self, '_cut_detail_by_size', True)
        totals = {}
        try:
            for rec in getattr(self.data_processor, 'drawings_data', []) or []:
                if rec.get('status') != 'נחתך':
                    continue
                if (rec.get('נמען') or '').strip() != supplier:
                    continue
                layers = rec.get('שכבות')
                try:
                    layers = int(layers)
                except Exception:
                    layers = None
                if not layers or layers <= 0:
                    continue
                fabric_type = (rec.get('סוג בד') or '').strip()
                for prod in rec.get('מוצרים', []) or []:
                    pname = (prod.get('שם המוצר') or '').strip()
                    for sz in prod.get('מידות', []) or []:
                        size = (sz.get('מידה') or '').strip()
                        qty = int(sz.get('כמות', 0) or 0)
                        if not pname or qty <= 0:
                            continue
                        cut_units = qty * layers
                        if by_size:
                            key = (pname, size, fabric_type)
                        else:
                            key = (pname, '', fabric_type)
                        totals[key] = totals.get(key, 0) + cut_units
        except Exception:
            pass
        # סינון טקסטואלי
        search_txt = (getattr(self, 'cut_search_var', tk.StringVar()).get() or '').strip().lower()
        for (pname, size, fabric), qty in sorted(totals.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
            label_ok = True
            if search_txt:
                hay = f"{pname} {size} {fabric}".lower()
                if search_txt not in hay:
                    label_ok = False
            if not label_ok:
                continue
            self.supplier_cut_tree.insert('', 'end', values=(pname, size or '-', fabric or '-', qty))
