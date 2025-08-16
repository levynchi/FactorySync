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
        # מצב פירוט לפי מידות (כפתור טוגול) – ברירת מחדל: פירוט מידות פעיל
        self._balance_detail_by_size = True
        self._balance_toggle_btn = tk.Button(inner_bar, text='תצוגת מוצר בלבד', command=self._toggle_balance_detail_mode, bg='#8e44ad', fg='white')
        self._balance_toggle_btn.pack(side='left')
        # הוספת אפשרות לכלול "מה נגזר אצל הספק" לתוך העמודה "נשלח"
        self.include_cuts_in_shipped_var = tk.BooleanVar(value=False)
        tk.Checkbutton(inner_bar, text="הוסף נגזר ל'נשלח'", variable=self.include_cuts_in_shipped_var, bg='#f7f9fa', command=self._refresh_products_balance_table).pack(side='left', padx=(8,0))

        # בניית טבלת המאזן עם עמודות דינמיות בהתאם למצב פירוט מידות
        self._balance_page_frame = balance_page
        self._balance_tree_scrollbar = None
        self.products_balance_tree = None
        self._build_products_balance_tree(self._balance_detail_by_size)

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

    def _build_products_balance_tree(self, by_size: bool):
        """יוצר/מחדש את עמודות הטבלה לפי מצב פירוט מידות."""
        # אם קיים עץ קודם – הורסים אותו ואת הסקרולבר
        try:
            if self.products_balance_tree is not None:
                self.products_balance_tree.destroy()
        except Exception:
            pass
        try:
            if self._balance_tree_scrollbar is not None:
                self._balance_tree_scrollbar.destroy()
        except Exception:
            pass

        if by_size:
            columns = ('product','size','fabric_category','shipped','received','diff','status')
            headers = {
                'product': 'מוצר',
                'size': 'מידה',
                'fabric_category': 'קטגורית בד',
                'shipped': 'נשלח',
                'received': 'נתקבל',
                'diff': 'הפרש (נותר לקבל)',
                'status': 'סטטוס'
            }
            widths = {'product':240, 'size':90, 'fabric_category':140, 'shipped':90, 'received':90, 'diff':140, 'status':140}
        else:
            columns = ('product','fabric_category','shipped','received','diff','status')
            headers = {
                'product': 'מוצר',
                'fabric_category': 'קטגורית בד',
                'shipped': 'נשלח',
                'received': 'נתקבל',
                'diff': 'הפרש (נותר לקבל)',
                'status': 'סטטוס'
            }
            widths = {'product':260, 'fabric_category':160, 'shipped':90, 'received':90, 'diff':150, 'status':150}

        self.products_balance_tree = ttk.Treeview(self._balance_page_frame, columns=columns, show='headings', height=18)
        for c in columns:
            self.products_balance_tree.heading(c, text=headers[c])
            self.products_balance_tree.column(c, width=widths[c], anchor='center')
        self._balance_tree_scrollbar = ttk.Scrollbar(self._balance_page_frame, orient='vertical', command=self.products_balance_tree.yview)
        self.products_balance_tree.configure(yscroll=self._balance_tree_scrollbar.set)
        self.products_balance_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=6)
        self._balance_tree_scrollbar.pack(side='left', fill='y', pady=6)
        # דאבל-קליק על שורה לפתיחת פירוט תנועות
        try:
            self.products_balance_tree.bind('<Double-1>', self._on_balance_row_double_click)
        except Exception:
            pass

    def _on_balance_row_double_click(self, event=None):
        try:
            sel = self.products_balance_tree.selection()
            if not sel:
                return
            item_id = sel[0]
            values = self.products_balance_tree.item(item_id, 'values') or []
            self._open_balance_row_details(values)
        except Exception:
            pass

    def _open_balance_row_details(self, row_values):
        """פותח חלון פירוט תנועות עבור שורת מאזן נבחרת."""
        try:
            supplier = (getattr(self, 'balance_supplier_var', None).get() if hasattr(self, 'balance_supplier_var') else '') or ''
            if not supplier:
                return
            by_size = getattr(self, '_balance_detail_by_size', False)
            include_cuts = bool(getattr(self, 'include_cuts_in_shipped_var', tk.BooleanVar(value=False)).get())
            # חילוץ ערכים לפי עמודות פעילות
            cols = tuple(self.products_balance_tree['columns']) if hasattr(self, 'products_balance_tree') else ()
            def _get(idx, default=''):
                try:
                    return (row_values[idx] if idx < len(row_values) else default) or default
                except Exception:
                    return default
            if by_size and cols[:2] == ('product','size'):
                p_name = _get(0, '')
                p_size = _get(1, '')
                if p_size == '-':
                    p_size = ''
                p_cat = _get(2, '')
            else:
                p_name = _get(0, '')
                p_size = ''
                p_cat = _get(1, '')

            # בניית רשימת תנועות
            movements = []
            def add_move(date, doc_type, doc_no, name, size, cat, qty, direction):
                movements.append({
                    'date': (date or ''),
                    'doc_type': (doc_type or ''),
                    'doc_no': (doc_no or ''),
                    'product': (name or ''),
                    'size': (size or ''),
                    'category': (cat or ''),
                    'qty': int(qty or 0),
                    'direction': (direction or ''),
                })

            def norm(s):
                return (s or '').strip()

            # לוודא נתונים עדכניים
            try:
                if hasattr(self.data_processor, 'refresh_supplier_receipts'):
                    self.data_processor.refresh_supplier_receipts()
            except Exception:
                pass

            # תעודות משלוח (נשלח)
            try:
                for rec in getattr(self.data_processor, 'delivery_notes', []) or []:
                    if norm(rec.get('supplier')) != norm(supplier):
                        continue
                    rec_date = rec.get('date') or rec.get('created_at') or ''
                    rec_no = rec.get('number') or rec.get('id') or ''
                    for ln in (rec.get('lines', []) or []):
                        name = norm(ln.get('product'))
                        size = norm(ln.get('size'))
                        qty = int(ln.get('quantity', 0) or 0)
                        if not name or qty <= 0:
                            continue
                        cat_line = norm(ln.get('fabric_category'))
                        if not cat_line:
                            try:
                                cat_line = self._get_product_attrs(name, size, by_size)[3]
                            except Exception:
                                cat_line = ''
                        # התאמה למפתח הנוכחי
                        if norm(name) != norm(p_name):
                            continue
                        if by_size and norm(size) != norm(p_size):
                            continue
                        if norm(cat_line) != norm(p_cat):
                            continue
                        add_move(rec_date, 'תעודת משלוח', rec_no, name, size if by_size else '', cat_line, qty, 'נשלח')
            except Exception:
                pass

            # תעודות קליטה (נתקבל)
            try:
                for rec in getattr(self.data_processor, 'supplier_intakes', []) or []:
                    if norm(rec.get('supplier')) != norm(supplier):
                        continue
                    rec_date = rec.get('date') or rec.get('created_at') or ''
                    rec_no = rec.get('number') or rec.get('id') or ''
                    for ln in (rec.get('lines', []) or []):
                        name = norm(ln.get('product'))
                        size = norm(ln.get('size'))
                        qty = int(ln.get('quantity', 0) or 0)
                        if not name or qty <= 0:
                            continue
                        cat_line = norm(ln.get('fabric_category'))
                        if not cat_line:
                            try:
                                cat_line = self._get_product_attrs(name, size, by_size)[3]
                            except Exception:
                                cat_line = ''
                        if norm(name) != norm(p_name):
                            continue
                        if by_size and norm(size) != norm(p_size):
                            continue
                        if norm(cat_line) != norm(p_cat):
                            continue
                        add_move(rec_date, 'תעודת קליטה', rec_no, name, size if by_size else '', cat_line, qty, 'נתקבל')
            except Exception:
                pass

            # נגזרות (אופציונלי)
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
                        if norm(rec.get('נמען')) != norm(supplier):
                            continue
                        layers = rec.get('שכבות')
                        try:
                            layers = int(layers)
                        except Exception:
                            layers = None
                        if not layers or layers <= 0:
                            continue
                        rec_date = rec.get('תאריך') or rec.get('date') or ''
                        rec_no = rec.get('מס׳') or rec.get('id') or ''
                        for prod in rec.get('מוצרים', []) or []:
                            name = norm(prod.get('שם המוצר'))
                            for sz in prod.get('מידות', []) or []:
                                size = norm(sz.get('מידה'))
                                qty = int(sz.get('כמות', 0) or 0)
                                if not name or qty <= 0:
                                    continue
                                cat_line = 'טריקו לבן'
                                if norm(name) != norm(p_name):
                                    continue
                                if by_size and norm(size) != norm(p_size):
                                    continue
                                if norm(cat_line) != norm(p_cat):
                                    continue
                                add_move(rec_date, 'ציור – נגזר', rec_no, name, size if by_size else '', cat_line, qty * layers, 'נגזר')
                except Exception:
                    pass

            # יצירת חלון פירוט
            win = tk.Toplevel(self._balance_page_frame)
            win.title(f"פירוט תנועות – {p_name}{(' – ' + p_size) if by_size and p_size else ''} | קטגוריה: {p_cat}")
            win.geometry('920x520')
            cols = ('date','doc_type','doc_no','product','size','category','direction','qty')
            headers = {
                'date': 'תאריך',
                'doc_type': 'סוג מסמך',
                'doc_no': 'מס׳',
                'product': 'מוצר',
                'size': 'מידה',
                'category': 'קטגורית בד',
                'direction': 'תנועה',
                'qty': 'כמות'
            }
            widths = {'date':120,'doc_type':140,'doc_no':100,'product':220,'size':80,'category':140,'direction':90,'qty':80}
            tree = ttk.Treeview(win, columns=cols, show='headings', height=18)
            for c in cols:
                tree.heading(c, text=headers[c])
                tree.column(c, width=widths[c], anchor='center')
            vs = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
            tree.configure(yscroll=vs.set)
            tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=10)
            vs.pack(side='left', fill='y', pady=10)

            # מיון לפי תאריך אם אפשר
            def parse_dt(s):
                s = (s or '').strip()
                for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
                    try:
                        return datetime.strptime(s, fmt)
                    except Exception:
                        pass
                return None
            movements_sorted = sorted(movements, key=lambda m: (parse_dt(m['date']) or datetime.min, m['doc_type'], m['doc_no']))
            for m in movements_sorted:
                tree.insert('', 'end', values=(m['date'], m['doc_type'], m['doc_no'], m['product'], m['size'] if by_size else '', m['category'], m['direction'], m['qty']))

            # סיכומים
            try:
                shipped_sum = sum(m['qty'] for m in movements if m['direction'] == 'נשלח')
                received_sum = sum(m['qty'] for m in movements if m['direction'] == 'נתקבל')
                cut_sum = sum(m['qty'] for m in movements if m['direction'] == 'נגזר')
                diff = shipped_sum + cut_sum - received_sum
                summary = tk.Label(win, text=f"נשלח: {shipped_sum} | נתקבל: {received_sum} | נגזר: {cut_sum} | הפרש: {max(diff,0)}", bg='#f7f9fa', anchor='e')
                summary.pack(fill='x', padx=10, pady=(0,10))
            except Exception:
                pass
        except Exception:
            # בליעה שקטה כדי לא להפיל את ה-UI
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
        # ודא שהעמודות תואמות למצב הנוכחי
        try:
            # מזהה האם העמודות כבר במצב הנכון – אם לא, נבנה מחדש
            current_cols = tuple(self.products_balance_tree['columns']) if hasattr(self, 'products_balance_tree') and self.products_balance_tree else ()
            expected_cols = ('product','size','fabric_category','shipped','received','diff','status') if by_size else ('product','fabric_category','shipped','received','diff','status')
            if current_cols != expected_cols:
                self._build_products_balance_tree(by_size)
        except Exception:
            # במקרה של שגיאה – ננסה לבנות מחדש
            try:
                self._build_products_balance_tree(by_size)
            except Exception:
                pass
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
                    # קביעת קטגורית בד לשיוך: לפי השורה, ואם חסר – לפי majority בקטלוג
                    f_cat_line = (ln.get('fabric_category') or '').strip()
                    if not f_cat_line:
                        try:
                            f_cat_line = self._get_product_attrs(name, size, by_size)[3]
                        except Exception:
                            f_cat_line = ''
                    key = (name, size if by_size else '', f_cat_line) if by_size else (name, '', f_cat_line)
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
                    f_cat_line = (ln.get('fabric_category') or '').strip()
                    if not f_cat_line:
                        try:
                            f_cat_line = self._get_product_attrs(name, size, by_size)[3]
                        except Exception:
                            f_cat_line = ''
                    key = (name, size if by_size else '', f_cat_line) if by_size else (name, '', f_cat_line)
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
                            # שיוך נגזרות לקטגוריה קבועה 'טריקו לבן' כפי שסוכם
                            cut_key = (pname, size if by_size else '', 'טריקו לבן') if by_size else (pname, '', 'טריקו לבן')
                            cut_totals[cut_key] = cut_totals.get(cut_key, 0) + qty * layers
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
            # key הוא טופל: (שם מוצר, מידה או '', קטגורית בד)
            if isinstance(key, tuple):
                if by_size:
                    p_name, p_size, p_cat = key
                else:
                    p_name, _, p_cat = key
                    p_size = ''
            else:
                # תאימות לאחור – לא צפוי לאחר שינוי זה
                p_name = str(key)
                p_size = ''
                p_cat = ''
            s = shipped.get(key, 0)
            r = received.get(key, 0)
            diff = s - r
            # סינון טקסטואלי
            if search_txt:
                hay = f"{p_name} {p_size} {p_cat}".lower()
                if search_txt not in hay:
                    continue
            # סינון רק חוסר
            if only_pending and diff <= 0:
                continue
            status = 'הושלם' if diff <= 0 else f"נותרו {diff} לקבל"
            # מציגים עמודות מצומצמות: מוצר (+מידה אם פעיל), קטגורית בד, נשלח/נתקבל/הפרש/סטטוס
            f_cat = p_cat
            if by_size:
                self.products_balance_tree.insert('', 'end', values=(p_name, p_size or '-', f_cat, s, r, max(diff, 0), status))
            else:
                self.products_balance_tree.insert('', 'end', values=(p_name, f_cat, s, r, max(diff, 0), status))

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
        # עדכון עמודות בהתאם למצב
        try:
            self._build_products_balance_tree(self._balance_detail_by_size)
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
