import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import calendar as _cal

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

        # עמוד חדש: מאזן סחורות שנחתכו אצל הספק
        cut_balance_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(cut_balance_page, text='מאזן סחורות שנחתכו אצל הספק')
        tk.Label(cut_balance_page, text='מאזן סחורות שנחתכו אצל הספק', font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=6)
        cb_bar = tk.Frame(cut_balance_page, bg='#f7f9fa'); cb_bar.pack(fill='x', padx=10, pady=(0,6))
        tk.Label(cb_bar, text='חיפוש:', bg='#f7f9fa').pack(side='right', padx=(8,4))
        self.cut_balance_search_var = tk.StringVar(); cb_search = tk.Entry(cb_bar, textvariable=self.cut_balance_search_var, width=24); cb_search.pack(side='right', padx=(0,6))
        try:
            cb_search.bind('<KeyRelease>', lambda e: self._refresh_cut_balance_table())
        except Exception:
            pass
        self.cut_balance_only_pending_var = tk.BooleanVar(value=False)
        tk.Checkbutton(cb_bar, text='רק חוסר', variable=self.cut_balance_only_pending_var, bg='#f7f9fa', command=self._refresh_cut_balance_table).pack(side='left', padx=(8,0))
        tk.Button(cb_bar, text='🔄 רענן', command=self._refresh_cut_balance_table, bg='#3498db', fg='white').pack(side='left', padx=6)
        # טבלה
        cb_cols = ('product','size','fabric_category','drawing_no','shipped','received','diff','status')
        self.cut_balance_tree = ttk.Treeview(cut_balance_page, columns=cb_cols, show='headings', height=18)
        cb_headers = {
            'product':'מוצר','size':'מידה','fabric_category':'קטגורית בד','drawing_no':'מספר ציור',
            'shipped':'נשלח (נגזר×שכבות)','received':'נתקבל (חזר מציור)','diff':'הפרש (נותר לקבל)','status':'סטטוס'
        }
        cb_widths = {'product':240,'size':90,'fabric_category':150,'drawing_no':120,'shipped':120,'received':120,'diff':140,'status':150}
        for c in cb_cols:
            self.cut_balance_tree.heading(c, text=cb_headers[c])
            self.cut_balance_tree.column(c, width=cb_widths[c], anchor='center')
        vs3 = ttk.Scrollbar(cut_balance_page, orient='vertical', command=self.cut_balance_tree.yview)
        self.cut_balance_tree.configure(yscroll=vs3.set)
        self.cut_balance_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=6)
        vs3.pack(side='left', fill='y', pady=6)
        try:
            self.cut_balance_tree.bind('<Double-1>', self._on_cut_balance_row_double_click)
            self.cut_balance_tree.bind('<Return>', self._on_cut_balance_row_double_click)
        except Exception:
            pass

        # עמוד חדש: מאזן אביזרי תפירה
        accessories_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(accessories_page, text='מאזן אביזרי תפירה')
        tk.Label(accessories_page, text='מאזן אביזרי תפירה לפי ספק', font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=6)
        acc_bar = tk.Frame(accessories_page, bg='#f7f9fa'); acc_bar.pack(fill='x', padx=10, pady=(0,6))
        tk.Label(acc_bar, text='חיפוש:', bg='#f7f9fa').pack(side='right', padx=(8,4))
        self.accessories_search_var = tk.StringVar(); acc_search = tk.Entry(acc_bar, textvariable=self.accessories_search_var, width=24); acc_search.pack(side='right', padx=(0,6))
        try:
            acc_search.bind('<KeyRelease>', lambda e: self._refresh_accessories_balance_table())
        except Exception:
            pass
        # טווח תאריכים לאביזרים – כמו בעמוד מאזן מוצרים
        try:
            DateEntry = None
            try:
                from tkcalendar import DateEntry  # type: ignore
            except Exception:
                DateEntry = None
            tk.Label(acc_bar, text='עד תאריך:', bg='#f7f9fa').pack(side='right', padx=(8,4))
            self.accessories_to_date_var = tk.StringVar()
            if DateEntry is not None:
                acc_to_entry = DateEntry(acc_bar, textvariable=self.accessories_to_date_var, width=12, date_pattern='yyyy-mm-dd', locale='he_IL')
                try: acc_to_entry.bind('<<DateEntrySelected>>', lambda e: self._refresh_accessories_balance_table())
                except Exception: pass
            else:
                acc_to_entry = tk.Entry(acc_bar, textvariable=self.accessories_to_date_var, width=12)
                acc_to_entry.bind('<KeyRelease>', lambda e: self._refresh_accessories_balance_table())
            acc_to_entry.pack(side='right', padx=(0,10))
            try:
                tk.Button(acc_bar, text='📅', width=2, command=lambda e=acc_to_entry,v=self.accessories_to_date_var: self._open_date_picker(e, v, self._refresh_accessories_balance_table)).pack(side='right', padx=(0,4))
            except Exception:
                pass

            tk.Label(acc_bar, text='מתאריך:', bg='#f7f9fa').pack(side='right', padx=(8,4))
            self.accessories_from_date_var = tk.StringVar()
            if DateEntry is not None:
                acc_from_entry = DateEntry(acc_bar, textvariable=self.accessories_from_date_var, width=12, date_pattern='yyyy-mm-dd', locale='he_IL')
                try: acc_from_entry.bind('<<DateEntrySelected>>', lambda e: self._refresh_accessories_balance_table())
                except Exception: pass
            else:
                acc_from_entry = tk.Entry(acc_bar, textvariable=self.accessories_from_date_var, width=12)
                acc_from_entry.bind('<KeyRelease>', lambda e: self._refresh_accessories_balance_table())
            acc_from_entry.pack(side='right', padx=(0,6))
            try:
                tk.Button(acc_bar, text='📅', width=2, command=lambda e=acc_from_entry,v=self.accessories_from_date_var: self._open_date_picker(e, v, self._refresh_accessories_balance_table)).pack(side='right', padx=(0,4))
            except Exception:
                pass
        except Exception:
            pass
        self.accessories_only_pending_var = tk.BooleanVar(value=False)
        tk.Checkbutton(acc_bar, text='רק חוסר', variable=self.accessories_only_pending_var, bg='#f7f9fa', command=self._refresh_accessories_balance_table).pack(side='left', padx=(8,0))
        tk.Button(acc_bar, text='🔄 רענן', command=self._refresh_accessories_balance_table, bg='#3498db', fg='white').pack(side='left', padx=6)
        # כפתורי סינון מהיר לאביזרים עיקריים
        self.accessories_kind_filter_var = tk.StringVar(value='')
        tk.Button(acc_bar, text='כל האביזרים', command=self._set_accessories_summary).pack(side='left', padx=(4,12))
        tk.Button(acc_bar, text='טיקטקים', command=lambda: self._set_accessories_kind_filter('טיק טק קומפלט')).pack(side='left', padx=(16,4))
        tk.Button(acc_bar, text='גומי', command=lambda: self._set_accessories_kind_filter('גומי')).pack(side='left', padx=4)
        tk.Button(acc_bar, text='סרט', command=lambda: self._set_accessories_kind_filter('סרט')).pack(side='left', padx=4)
        # טבלה (נבנית דינמית – סיכום או פירוט)
        self._accessories_detail_mode = False
        self._accessories_page_frame = accessories_page
        self._accessories_tree_scrollbar = None
        self.accessories_tree = None
        self._build_accessories_tree(detail=False)

        # סרגל פנימי: חיפוש + כפתור פירוט לפי מידות
        inner_bar = tk.Frame(balance_page, bg='#f7f9fa')
        inner_bar.pack(fill='x', padx=10, pady=(0,6))
        # טווח תאריכים: מתאריך ... עד תאריך ... עם בורר גרפי אם tkcalendar מותקן
        try:
            DateEntry = None
            try:
                from tkcalendar import DateEntry  # type: ignore
            except Exception:
                DateEntry = None
            tk.Label(inner_bar, text='עד תאריך:', bg='#f7f9fa').pack(side='right', padx=(8,4))
            self.balance_to_date_var = tk.StringVar()
            # ברירת מחדל: היום
            try:
                _today_str = datetime.now().strftime('%Y-%m-%d')
                self.balance_to_date_var.set(_today_str)
            except Exception:
                pass
            if DateEntry is not None:
                to_entry = DateEntry(inner_bar, textvariable=self.balance_to_date_var, width=12, date_pattern='yyyy-mm-dd', locale='he_IL')
                # רענון בעת בחירה מהקלנדר
                try: to_entry.bind('<<DateEntrySelected>>', lambda e: self._refresh_products_balance_table())
                except Exception: pass
                # החלת ברירת המחדל בפועל על ה-Widget (DateEntry מאפס לטודיי אם לא)
                try:
                    to_entry.set_date(datetime.now())
                except Exception:
                    pass
            else:
                to_entry = tk.Entry(inner_bar, textvariable=self.balance_to_date_var, width=12)
                to_entry.bind('<KeyRelease>', lambda e: self._refresh_products_balance_table())
            to_entry.pack(side='right', padx=(0,10))
            # כפתור פתיחת קלנדר מפורש – גם אם DateEntry קיים
            try:
                tk.Button(inner_bar, text='📅', width=2, command=lambda e=to_entry,v=self.balance_to_date_var: self._open_date_picker(e, v, self._refresh_products_balance_table)).pack(side='right', padx=(0,4))
            except Exception:
                pass

            tk.Label(inner_bar, text='מתאריך:', bg='#f7f9fa').pack(side='right', padx=(8,4))
            self.balance_from_date_var = tk.StringVar()
            # ברירת מחדל: תחילת השנה הנוכחית
            try:
                _now = datetime.now(); _start_year = datetime(_now.year, 1, 1).strftime('%Y-%m-%d')
                self.balance_from_date_var.set(_start_year)
            except Exception:
                pass
            if DateEntry is not None:
                from_entry = DateEntry(inner_bar, textvariable=self.balance_from_date_var, width=12, date_pattern='yyyy-mm-dd', locale='he_IL')
                try: from_entry.bind('<<DateEntrySelected>>', lambda e: self._refresh_products_balance_table())
                except Exception: pass
                # החלת ברירת המחדל: 1 בינואר של השנה הנוכחית
                try:
                    _now = datetime.now(); from_entry.set_date(datetime(_now.year, 1, 1))
                except Exception:
                    pass
            else:
                from_entry = tk.Entry(inner_bar, textvariable=self.balance_from_date_var, width=12)
                from_entry.bind('<KeyRelease>', lambda e: self._refresh_products_balance_table())
            from_entry.pack(side='right', padx=(0,6))
            try:
                tk.Button(inner_bar, text='📅', width=2, command=lambda e=from_entry,v=self.balance_from_date_var: self._open_date_picker(e, v, self._refresh_products_balance_table)).pack(side='right', padx=(0,4))
            except Exception:
                pass
        except Exception:
            pass
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

    def _open_date_picker(self, anchor_widget, target_var: tk.StringVar, on_change=None):
        """פותח חלון בחירת תאריך גרפי ללא תלות ב-tkcalendar; אם tkcalendar קיים – ישתמש בו.

        יעדכן את target_var במחרוזת YYYY-MM-DD וירענן את הטבלה.
        """
        # נסה תחילה להשתמש ב-tkcalendar אם קיים
        Calendar = None
        try:
            from tkcalendar import Calendar  # type: ignore
        except Exception:
            Calendar = None

        top = tk.Toplevel(self._balance_page_frame)
        top.transient(self._balance_page_frame)
        top.title('בחירת תאריך')
        top.resizable(False, False)
        # מיקום ליד השדה
        try:
            x = anchor_widget.winfo_rootx(); y = anchor_widget.winfo_rooty() + anchor_widget.winfo_height()
            top.geometry(f"320x320+{x}+{y}")
        except Exception:
            top.geometry('320x320')

        # עזר: פרש תאריך התחלתי
        def _parse_dt(s: str):
            s = (s or '').strip()
            for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
                try:
                    return datetime.strptime(s, fmt)
                except Exception:
                    pass
            return None

        now_dt = _parse_dt(target_var.get()) or datetime.now()

        # אם Calendar קיים – הצג אותו בפשטות
        if Calendar is not None:
            cal = Calendar(top, selectmode='day', year=now_dt.year, month=now_dt.month, day=now_dt.day, locale='he_IL')
            cal.pack(fill='both', expand=True, padx=6, pady=6)
            btns = tk.Frame(top); btns.pack(fill='x', padx=6, pady=(0,6))
            def _set_and_close_from_cal():
                try:
                    d = cal.selection_get()
                    if d:
                        target_var.set(d.strftime('%Y-%m-%d'))
                        try:
                            (on_change or self._refresh_products_balance_table)()
                        except Exception:
                            pass
                except Exception:
                    pass
                try: top.destroy()
                except Exception: pass
            def _clear_and_close():
                try: target_var.set('')
                except Exception: pass
                try:
                    (on_change or self._refresh_products_balance_table)()
                except Exception:
                    pass
                try: top.destroy()
                except Exception: pass
            tk.Button(btns, text='היום', command=lambda: [target_var.set(datetime.now().strftime('%Y-%m-%d')), (on_change or self._refresh_products_balance_table)(), top.destroy()]).pack(side='left')
            tk.Button(btns, text='נקה', command=_clear_and_close).pack(side='left')
            tk.Button(btns, text='בחר', command=_set_and_close_from_cal).pack(side='right')
            try:
                top.focus_set(); top.grab_set()
            except Exception:
                pass
            return

        # אחרת: בניית יומן בסיסי ב-Tk בלבד
        header = tk.Frame(top)
        header.pack(fill='x', padx=8, pady=(8,4))
        month_var = tk.IntVar(value=now_dt.month)
        year_var = tk.IntVar(value=now_dt.year)

        def _update_title():
            try:
                title_lbl.config(text=f"{year_var.get():04d}-{month_var.get():02d}")
            except Exception:
                pass

        def _change_month(delta):
            m = month_var.get() + delta
            y = year_var.get()
            if m < 1:
                m = 12; y -= 1
            elif m > 12:
                m = 1; y += 1
            month_var.set(m); year_var.set(y)
            _update_title(); _rebuild_days()

        tk.Button(header, text='«', width=3, command=lambda: _change_month(-1)).pack(side='right')
        title_lbl = tk.Label(header, text='', font=('Arial', 10, 'bold'))
        title_lbl.pack(side='right', padx=6)
        tk.Button(header, text='»', width=3, command=lambda: _change_month(+1)).pack(side='right')

        # Today / Clear
        tk.Button(header, text='היום', command=lambda: _select_date(datetime.now().year, datetime.now().month, datetime.now().day)).pack(side='left', padx=(0,6))
        tk.Button(header, text='נקה', command=lambda: (_clear_and_close())).pack(side='left')

        grid = tk.Frame(top)
        grid.pack(fill='both', expand=True, padx=8, pady=6)

        # כותרות ימים (א עד ש) – נתחיל ביום ראשון
        # יצירת סדר שבוע החל מיום ראשון
        weekdays = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ש']
        for i, w in enumerate(weekdays):
            tk.Label(grid, text=w, width=4, anchor='center', fg='#2c3e50').grid(row=0, column=i, pady=(0,4))

        day_btns = []  # נשמור כדי לעדכן/לנקות אם צריך

        def _select_date(y, m, d):
            try:
                dt = datetime(int(y), int(m), int(d))
                target_var.set(dt.strftime('%Y-%m-%d'))
                try:
                    (on_change or self._refresh_products_balance_table)()
                except Exception:
                    pass
            except Exception:
                pass
            try:
                top.destroy()
            except Exception:
                pass

        def _clear_and_close():
            try:
                target_var.set('')
            except Exception:
                pass
            try:
                (on_change or self._refresh_products_balance_table)()
            except Exception:
                pass
            try:
                top.destroy()
            except Exception:
                pass

        def _rebuild_days():
            # נקה כפתורים ישנים
            for b in day_btns:
                try: b.destroy()
                except Exception: pass
            day_btns.clear()
            y = year_var.get(); m = month_var.get()
            _update_title()
            # היום לצורך הדגשה
            today = datetime.now()
            # קבל את היום בשבוע של ה-1 לחודש (0=שני לפי calendar ברירת מחדל; נרצה ראשון בתחילת שבוע)
            first_weekday, days_in_month = _cal.monthrange(y, m)  # 0=Monday..6=Sunday
            # מיפוי כדי שעמודה 0 תהיה ראשון
            start_col = (first_weekday + 1) % 7  # העתקה: Monday->1,... Sunday->0
            row = 1; col = start_col
            for d in range(1, days_in_month + 1):
                txt = f"{d:02d}"
                is_today = (d == today.day and m == today.month and y == today.year)
                btn = tk.Button(
                    grid,
                    text=txt,
                    width=4,
                    relief='raised',
                    bg=('#eaf2f8' if is_today else None),
                    command=lambda D=d, Y=y, M=m: _select_date(Y, M, D)
                )
                btn.grid(row=row, column=col, padx=1, pady=1)
                day_btns.append(btn)
                col += 1
                if col > 6:
                    col = 0; row += 1

        _rebuild_days()
        try:
            top.focus_set(); top.grab_set()
        except Exception:
            pass

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
        try:
            self._refresh_cut_balance_table()
        except Exception:
            pass
        try:
            self._refresh_accessories_balance_table()
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
            def add_move(date, doc_type, doc_no, name, size, cat, qty, direction, fabric_color='', print_name=''):
                movements.append({
                    'date': (date or ''),
                    'doc_type': (doc_type or ''),
                    'doc_no': (doc_no or ''),
                    'product': (name or ''),
                    'size': (size or ''),
                    'category': (cat or ''),
                    'qty': int(qty or 0),
                    'direction': (direction or ''),
                    'fabric_color': (fabric_color or ''),
                    'print_name': (print_name or ''),
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
                        # מאפייני מוצר לפי השורה ואם חסר – מהקטלוג
                        f_color = norm(ln.get('fabric_color'))
                        p_print = norm(ln.get('print_name'))
                        if not f_color or not p_print:
                            try:
                                _, f_color2, p_print2, _ = self._get_product_attrs(name, size, by_size)
                                f_color = f_color or f_color2
                                p_print = p_print or p_print2
                            except Exception:
                                pass
                        # התאמה למפתח הנוכחי
                        if norm(name) != norm(p_name):
                            continue
                        if by_size and norm(size) != norm(p_size):
                            continue
                        if norm(cat_line) != norm(p_cat):
                            continue
                        add_move(rec_date, 'תעודת משלוח', rec_no, name, size if by_size else '', cat_line, qty, 'נשלח', f_color, p_print)
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
                        # מאפייני מוצר לפי השורה ואם חסר – מהקטלוג
                        f_color = norm(ln.get('fabric_color'))
                        p_print = norm(ln.get('print_name'))
                        if not f_color or not p_print:
                            try:
                                _, f_color2, p_print2, _ = self._get_product_attrs(name, size, by_size)
                                f_color = f_color or f_color2
                                p_print = p_print or p_print2
                            except Exception:
                                pass
                        if norm(name) != norm(p_name):
                            continue
                        if by_size and norm(size) != norm(p_size):
                            continue
                        if norm(cat_line) != norm(p_cat):
                            continue
                        add_move(rec_date, 'תעודת קליטה', rec_no, name, size if by_size else '', cat_line, qty, 'נתקבל', f_color, p_print)
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
            win.geometry('1100x560')
            cols = ('date','doc_type','doc_no','product','size','category','fabric_color','print_name','direction','qty')
            headers = {
                'date': 'תאריך',
                'doc_type': 'סוג מסמך',
                'doc_no': 'מס׳',
                'product': 'מוצר',
                'size': 'מידה',
                'category': 'קטגורית בד',
                'fabric_color': 'צבע בד',
                'print_name': 'שם פרינט',
                'direction': 'תנועה',
                'qty': 'כמות'
            }
            widths = {
                'date':120,
                'doc_type':140,
                'doc_no':100,
                'product':200,
                'size':80,
                'category':140,
                'fabric_color':120,
                'print_name':140,
                'direction':90,
                'qty':80
            }
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
                tree.insert(
                    '',
                    'end',
                    values=(
                        m['date'],
                        m['doc_type'],
                        m['doc_no'],
                        m['product'],
                        m['size'] if by_size else '',
                        m['category'],
                        m.get('fabric_color',''),
                        m.get('print_name',''),
                        m['direction'],
                        m['qty']
                    )
                )

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
        # פרשנות טווח תאריכים
        def _parse_dt(s):
            s = (s or '').strip()
            for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
                try:
                    return datetime.strptime(s, fmt)
                except Exception:
                    pass
            return None
        try:
            from_s = (getattr(self, 'balance_from_date_var', tk.StringVar()).get() or '').strip()
            to_s   = (getattr(self, 'balance_to_date_var', tk.StringVar()).get() or '').strip()
        except Exception:
            from_s = to_s = ''
        from_dt = _parse_dt(from_s) if from_s else None
        to_dt   = _parse_dt(to_s) if to_s else None
        def _in_range(date_str):
            if not (from_dt or to_dt):
                return True
            d = _parse_dt(date_str)
            if d is None:
                # אם לא ניתן לפרש – אל תכלול כאשר מופעלים מסנני תאריך
                return False
            if from_dt and d < from_dt:
                return False
            if to_dt and d > to_dt:
                return False
            return True
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
        # עזר: מיפוי מוצר -> קטגוריה ראשית (עם נפילה ל"בגדים" כברירת מחדל)
        try:
            model_names = getattr(self.data_processor, 'product_model_names', []) or []
        except Exception:
            model_names = []
        name_to_main = {}
        try:
            for m in model_names:
                n = (m.get('name') or '').strip()
                if not n:
                    continue
                name_to_main[n] = (m.get('main_category') or 'בגדים').strip()
        except Exception:
            pass
        def _main_cat_of(n: str) -> str:
            try:
                return (name_to_main.get((n or '').strip()) or 'בגדים').strip()
            except Exception:
                return 'בגדים'
        include_cuts = bool(getattr(self, 'include_cuts_in_shipped_var', tk.BooleanVar(value=False)).get())
        shipped = {}
        received = {}
        try:
            notes = getattr(self.data_processor, 'delivery_notes', []) or []
            for rec in notes:
                if (rec.get('supplier') or '').strip() != supplier:
                    continue
                # סינון לפי טווח תאריכים
                if not _in_range(rec.get('date') or rec.get('created_at') or ''):
                    continue
                for ln in rec.get('lines', []) or []:
                    name = (ln.get('product') or '').strip()
                    size = (ln.get('size') or '').strip()
                    qty = int(ln.get('quantity', 0) or 0)
                    if not name or qty <= 0:
                        continue
                    # סינון: נשלח רק כאשר קטגוריה ראשית=בגדים ותת קטגוריה=גזרות שלא נתפרו
                    try:
                        mc = _main_cat_of(name)
                        subc = (ln.get('category') or '').strip()
                        # accept both the standardized and the legacy label
                        sent_sub_ok = subc in ('גזרות שלא נתפרו', 'גזרות לא תפורות')
                        if not (mc == 'בגדים' and sent_sub_ok):
                            continue
                    except Exception:
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
                # סינון לפי טווח תאריכים
                if not _in_range(rec.get('date') or rec.get('created_at') or ''):
                    continue
                for ln in rec.get('lines', []) or []:
                    name = (ln.get('product') or '').strip()
                    size = (ln.get('size') or '').strip()
                    qty = int(ln.get('quantity', 0) or 0)
                    if not name or qty <= 0:
                        continue
                    # סינון: התקבל כאשר קטגוריה ראשית=בגדים ותת קטגוריה=בגדים תפורים
                    # (כעת ללא חובה ש"חזר מציור" יהיה 'כן')
                    try:
                        mc = _main_cat_of(name)
                        subc = (ln.get('category') or '').strip()
                        if not (mc == 'בגדים' and subc == 'בגדים תפורים'):
                            continue
                    except Exception:
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
                    # סינון לפי טווח תאריכים
                    if not _in_range(rec.get('תאריך') or rec.get('date') or ''):
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
        # מציגים רק מוצרים שנשלחו (כפי שנתבקש) – ולכן מפתחות רק מתוך shipped
        keys = sorted(set(list(shipped.keys())))
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

    # === מאזן סחורות שנחתכו אצל הספק ===
    def _refresh_cut_balance_table(self):
        supplier = (getattr(self, 'balance_supplier_var', None).get() if hasattr(self, 'balance_supplier_var') else '') or ''
        # ניקוי
        if hasattr(self, 'cut_balance_tree'):
            for iid in self.cut_balance_tree.get_children():
                self.cut_balance_tree.delete(iid)
        if not supplier:
            return
        # לוודא נתונים עדכניים
        try:
            if hasattr(self.data_processor, 'refresh_supplier_receipts'):
                self.data_processor.refresh_supplier_receipts()
        except Exception:
            pass
        try:
            if hasattr(self.data_processor, 'refresh_drawings_data'):
                self.data_processor.refresh_drawings_data()
        except Exception:
            pass
        # נשלח: מסך ציורים שנחתכו × שכבות
        shipped = {}
        shipped_drawings = {}
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
                # קביעת קטגורית בד מהציור (סוג בד) או ברירת מחדל
                fabric_category = (rec.get('סוג בד') or 'טריקו לבן').strip()
                rec_no = rec.get('מס׳') or rec.get('id') or rec.get('number') or ''
                rec_no = str(rec_no)
                for prod in rec.get('מוצרים', []) or []:
                    pname = (prod.get('שם המוצר') or '').strip()
                    for sz in prod.get('מידות', []) or []:
                        size = (sz.get('מידה') or '').strip()
                        qty = int(sz.get('כמות', 0) or 0)
                        if not pname or qty <= 0:
                            continue
                        key = (pname, size, fabric_category)
                        shipped[key] = shipped.get(key, 0) + qty * layers
                        # שמירת מספרי ציור עבור המפתח
                        lst = shipped_drawings.get(key)
                        if lst is None:
                            lst = []
                            shipped_drawings[key] = lst
                        if rec_no and rec_no not in lst:
                            lst.append(rec_no)
        except Exception:
            pass
        # נתקבל: מכל קליטות עם returned_from_drawing == 'כן'
        received = {}
        try:
            for rec in getattr(self.data_processor, 'supplier_intakes', []) or []:
                if (rec.get('supplier') or '').strip() != supplier:
                    continue
                for ln in rec.get('lines', []) or []:
                    if (ln.get('returned_from_drawing') or '').strip() != 'כן':
                        continue
                    pname = (ln.get('product') or '').strip()
                    size = (ln.get('size') or '').strip()
                    qty = int(ln.get('quantity', 0) or 0)
                    if not pname or qty <= 0:
                        continue
                    fcat = (ln.get('fabric_category') or '').strip()
                    # אם חסר קטגוריה בשורה, ננסה להשלים מהקטלוג לשמירה על עקביות
                    if not fcat:
                        try:
                            fcat = self._get_product_attrs(pname, size, True)[3]
                        except Exception:
                            fcat = ''
                    key = (pname, size, fcat)
                    received[key] = received.get(key, 0) + qty
        except Exception:
            pass
        # איחוד, סינון, והצגה
        keys = sorted(set(list(shipped.keys()) + list(received.keys())))
        search_txt = (getattr(self, 'cut_balance_search_var', tk.StringVar()).get() or '').strip().lower()
        only_pending = bool(getattr(self, 'cut_balance_only_pending_var', tk.BooleanVar()).get())
        for key in keys:
            pname, size, fcat = key
            s = shipped.get(key, 0)
            r = received.get(key, 0)
            diff = s - r
            draw_str = ', '.join(shipped_drawings.get(key, []))
            if search_txt:
                hay = f"{pname} {size} {fcat} {draw_str}".lower()
                if search_txt not in hay:
                    continue
            if only_pending and diff <= 0:
                continue
            status = 'הושלם' if diff <= 0 else f"נותרו {diff} לקבל"
            self.cut_balance_tree.insert('', 'end', values=(pname, size or '-', fcat, draw_str, s, r, max(diff,0), status))

    def _on_cut_balance_row_double_click(self, event=None):
        try:
            sel = self.cut_balance_tree.selection()
            if not sel:
                return
            item_id = sel[0]
            values = self.cut_balance_tree.item(item_id, 'values') or []
            self._open_cut_balance_row_details(values)
        except Exception:
            pass

    def _open_cut_balance_row_details(self, row_values):
        """פרטי מאזן לשורה בטאב 'מאזן סחורות שנחתכו אצל הספק' – מתי נשלח, כמה נשלח, ומתי התקבל."""
        try:
            supplier = (getattr(self, 'balance_supplier_var', None).get() if hasattr(self, 'balance_supplier_var') else '') or ''
            if not supplier:
                return
            # שליפת המפתחות מהשורה
            # עמודות: ('product','size','fabric_category','drawing_no','shipped','received','diff','status')
            def _get(i, d=''):
                try:
                    return (row_values[i] if i < len(row_values) else d) or d
                except Exception:
                    return d
            pname = _get(0, '')
            psize = _get(1, '')
            if psize == '-':
                psize = ''
            fcat = _get(2, '')

            def norm(s):
                return (s or '').strip()

            # ודא נתונים
            try:
                if hasattr(self.data_processor, 'refresh_supplier_receipts'):
                    self.data_processor.refresh_supplier_receipts()
            except Exception:
                pass
            try:
                if hasattr(self.data_processor, 'refresh_drawings_data'):
                    self.data_processor.refresh_drawings_data()
            except Exception:
                pass

            # בניית תנועות
            movements = []
            def add_move(date, doc_type, doc_no, qty, direction):
                movements.append({
                    'date': date or '',
                    'doc_type': doc_type or '',
                    'doc_no': str(doc_no or ''),
                    'qty': int(qty or 0),
                    'direction': direction or ''
                })

            # נשלח – מציורים
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
                    # קטגורית בד כפי שהשתמשנו בטבלה
                    cat_line = (rec.get('סוג בד') or 'טריקו לבן').strip()
                    if norm(cat_line) != norm(fcat):
                        continue
                    rec_date = rec.get('תאריך') or rec.get('date') or ''
                    rec_no = rec.get('מס׳') or rec.get('id') or rec.get('number') or ''
                    for prod in rec.get('מוצרים', []) or []:
                        name = norm(prod.get('שם המוצר'))
                        if norm(name) != norm(pname):
                            continue
                        for sz in prod.get('מידות', []) or []:
                            size = norm(sz.get('מידה'))
                            if norm(size) != norm(psize):
                                continue
                            qty = int(sz.get('כמות', 0) or 0)
                            if qty > 0:
                                add_move(rec_date, 'ציור – נגזר', rec_no, qty * layers, 'נשלח')
            except Exception:
                pass

            # נתקבל – מקליטות שסומנו "חזר מציור"
            try:
                for rec in getattr(self.data_processor, 'supplier_intakes', []) or []:
                    if norm(rec.get('supplier')) != norm(supplier):
                        continue
                    rec_date = rec.get('date') or rec.get('created_at') or ''
                    rec_no = rec.get('number') or rec.get('id') or ''
                    for ln in (rec.get('lines', []) or []):
                        if (ln.get('returned_from_drawing') or '').strip() != 'כן':
                            continue
                        name = norm(ln.get('product'))
                        size = norm(ln.get('size'))
                        qty = int(ln.get('quantity', 0) or 0)
                        if not name or qty <= 0:
                            continue
                        if norm(name) != norm(pname):
                            continue
                        if norm(size) != norm(psize):
                            continue
                        cat_line = norm(ln.get('fabric_category'))
                        if not cat_line:
                            try:
                                cat_line = self._get_product_attrs(name, size, True)[3]
                            except Exception:
                                cat_line = ''
                        if norm(cat_line) != norm(fcat):
                            continue
                        add_move(rec_date, 'תעודת קליטה', rec_no, qty, 'נתקבל')
            except Exception:
                pass

            # חלון פירוט
            win = tk.Toplevel(self._balance_page_frame)
            win.title(f"פירוט - מאזן מוצר חתוך | {pname}{(' – ' + psize) if psize else ''} | קטגוריה: {fcat}")
            win.geometry('820x520')
            cols = ('date','doc_type','doc_no','direction','qty')
            headers = {'date':'תאריך','doc_type':'סוג','doc_no':'מס׳','direction':'תנועה','qty':'כמות'}
            widths = {'date':140,'doc_type':170,'doc_no':120,'direction':90,'qty':80}
            tree = ttk.Treeview(win, columns=cols, show='headings', height=18)
            for c in cols:
                tree.heading(c, text=headers[c])
                tree.column(c, width=widths[c], anchor='center')
            vs = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
            tree.configure(yscroll=vs.set)
            tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=10)
            vs.pack(side='left', fill='y', pady=10)

            def parse_dt(s):
                s = (s or '').strip()
                for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
                    try:
                        return datetime.strptime(s, fmt)
                    except Exception:
                        pass
                return None
            moves_sorted = sorted(movements, key=lambda m: (parse_dt(m['date']) or datetime.min, m['doc_type'], m['doc_no']))
            for m in moves_sorted:
                tree.insert('', 'end', values=(m['date'], m['doc_type'], m['doc_no'], m['direction'], m['qty']))

            try:
                shipped_sum = sum(m['qty'] for m in movements if m['direction'] == 'נשלח')
                received_sum = sum(m['qty'] for m in movements if m['direction'] == 'נתקבל')
                diff = shipped_sum - received_sum
                summary = tk.Label(win, text=f"נשלח: {shipped_sum} | נתקבל: {received_sum} | הפרש: {max(diff,0)}", bg='#f7f9fa', anchor='e')
                summary.pack(fill='x', padx=10, pady=(0,10))
            except Exception:
                pass
        except Exception:
            pass

    # === מאזן אביזרי תפירה ===
    def _refresh_accessories_balance_table(self):
        supplier = (getattr(self, 'balance_supplier_var', None).get() if hasattr(self, 'balance_supplier_var') else '') or ''
        if not hasattr(self, 'accessories_tree'):
            return
        # ניקוי
        for iid in self.accessories_tree.get_children():
            self.accessories_tree.delete(iid)
        if not supplier:
            return
        # סינון לפי טווח תאריכים
        def parse_dt(s):
            s = (s or '').strip()
            for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
                try:
                    return datetime.strptime(s, fmt)
                except Exception:
                    pass
            return None
        try:
            from_s = (getattr(self, 'accessories_from_date_var', tk.StringVar()).get() or '').strip()
        except Exception:
            from_s = ''
        try:
            to_s = (getattr(self, 'accessories_to_date_var', tk.StringVar()).get() or '').strip()
        except Exception:
            to_s = ''
        start_dt = parse_dt(from_s)
        end_dt = parse_dt(to_s)
        def in_range(date_str: str) -> bool:
            if not (start_dt or end_dt):
                return True
            d = parse_dt(date_str)
            if not d:
                return True
            if start_dt and d < start_dt:
                return False
            if end_dt and d > end_dt:
                return False
            return True
        # ריענון נתונים
        try:
            if hasattr(self.data_processor, 'refresh_supplier_receipts'):
                self.data_processor.refresh_supplier_receipts()
        except Exception:
            pass

        # עזר: זיהוי האם מוצר הוא אביזר תפירה לפי הקטלוג/מיפוי שמות
        def _is_accessory(name: str) -> bool:
            n = (name or '').strip()
            if not n:
                return False
            try:
                # העדפה: product_model_names אם זמין
                for m in getattr(self.data_processor, 'product_model_names', []) or []:
                    if (m.get('name') or '').strip() == n:
                        return (m.get('main_category') or '').strip() == 'אביזרי תפירה'
            except Exception:
                pass
            # נפילה לקטלוג מוצרים: קח main_category של הרשומה הראשונה המתאימה
            try:
                for r in getattr(self.data_processor, 'products_catalog', []) or []:
                    if (r.get('name') or '').strip() == n:
                        return (r.get('main_category') or '').strip() == 'אביזרי תפירה'
            except Exception:
                pass
            return False

        # צבירה: אביזרים שנשלחו ונתקבלו
        shipped = {}   # name -> qty
        received = {}  # name -> qty

        def norm(s):
            return (s or '').strip()

        # נשלח: מתוך תעודות משלוח – לפי main_category 'אביזרי תפירה' (או fallback לרשימת אביזרים)
        try:
            acc_names = set(norm(r.get('name')) for r in getattr(self.data_processor, 'sewing_accessories', []) or [])
        except Exception:
            acc_names = set()
        try:
            for rec in getattr(self.data_processor, 'delivery_notes', []) or []:
                if norm(rec.get('supplier')) != norm(supplier):
                    continue
                rec_date = rec.get('date') or rec.get('created_at') or ''
                if not in_range(rec_date):
                    continue
                for ln in rec.get('lines', []) or []:
                    name = norm(ln.get('product'))
                    qty = int(ln.get('quantity', 0) or 0)
                    if not name or qty <= 0:
                        continue
                    # אביזרי תפירה: לפי קטגוריה ראשית בקטלוג; אם לא ידוע – fallback לשמות אביזרים
                    if _is_accessory(name) or name in acc_names:
                        shipped[name] = shipped.get(name, 0) + qty
        except Exception:
            pass

        # נתקבל: מתוך תעודות קליטה, אותו היגיון
        try:
            for rec in getattr(self.data_processor, 'supplier_intakes', []) or []:
                if norm(rec.get('supplier')) != norm(supplier):
                    continue
                rec_date = rec.get('date') or rec.get('created_at') or ''
                if not in_range(rec_date):
                    continue
                for ln in rec.get('lines', []) or []:
                    name = norm(ln.get('product'))
                    qty = int(ln.get('quantity', 0) or 0)
                    if not name or qty <= 0:
                        continue
                    if _is_accessory(name) or name in acc_names:
                        received[name] = received.get(name, 0) + qty
                    # תוספת: אביזרים שחזרו מבגדים תפורים – חישוב לפי הקטלוג
                    subc = norm(ln.get('category'))
                    if subc == 'בגדים תפורים':
                        # חפש בקטלוג את רשומת המוצר לפי שם/מידה לקבלת כמויות אביזרים ליחידה
                        psize = norm(ln.get('size'))
                        try:
                            catalog = getattr(self.data_processor, 'products_catalog', []) or []
                        except Exception:
                            catalog = []
                        def _match(r):
                            return norm(r.get('name')) == name and (not psize or norm(r.get('size')) == psize)
                        per_ticks = per_elastic = per_ribbon = 0
                        try:
                            for r in catalog:
                                if _match(r):
                                    per_ticks = int(r.get('ticks_qty', 0) or 0)
                                    per_elastic = int(r.get('elastic_qty', 0) or 0)
                                    per_ribbon = int(r.get('ribbon_qty', 0) or 0)
                                    break
                        except Exception:
                            pass
                        if per_ticks > 0:
                            received['טיק טק קומפלט'] = received.get('טיק טק קומפלט', 0) + per_ticks * qty
                        if per_elastic > 0:
                            received['גומי'] = received.get('גומי', 0) + per_elastic * qty
                        if per_ribbon > 0:
                            received['סרט'] = received.get('סרט', 0) + per_ribbon * qty
        except Exception:
            pass

        detail = bool(getattr(self, '_accessories_detail_mode', False))
        search_txt = (getattr(self, 'accessories_search_var', tk.StringVar()).get() or '').strip().lower()
        only_pending = bool(getattr(self, 'accessories_only_pending_var', tk.BooleanVar()).get())
        kind_filter = (getattr(self, 'accessories_kind_filter_var', tk.StringVar()).get() or '').strip()

        if detail and kind_filter:
            # פירוט כרונולוגי עבור אביזר נבחר
            moves = []
            def add(date, kind, doc_no, qty, direction):
                moves.append({'date': date or '', 'kind': kind or '', 'no': str(doc_no or ''), 'qty': int(qty or 0), 'direction': direction or ''})
            # משלוחים ישירים של האביזר
            try:
                for rec in getattr(self.data_processor, 'delivery_notes', []) or []:
                    if norm(rec.get('supplier')) != norm(supplier):
                        continue
                    rec_date = rec.get('date') or rec.get('created_at') or ''
                    if not in_range(rec_date):
                        continue
                    rec_no = rec.get('number') or rec.get('id') or ''
                    for ln in rec.get('lines', []) or []:
                        if norm(ln.get('product')) != norm(kind_filter):
                            continue
                        qty = int(ln.get('quantity', 0) or 0)
                        if qty > 0:
                            add(rec_date, 'תעודת משלוח', rec_no, qty, 'נשלח')
            except Exception:
                pass
            # קליטות ישירות של האביזר
            try:
                for rec in getattr(self.data_processor, 'supplier_intakes', []) or []:
                    if norm(rec.get('supplier')) != norm(supplier):
                        continue
                    rec_date = rec.get('date') or rec.get('created_at') or ''
                    if not in_range(rec_date):
                        continue
                    rec_no = rec.get('number') or rec.get('id') or ''
                    for ln in rec.get('lines', []) or []:
                        if norm(ln.get('product')) == norm(kind_filter):
                            qty = int(ln.get('quantity', 0) or 0)
                            if qty > 0:
                                add(rec_date, 'תעודת קליטה', rec_no, qty, 'נתקבל')
                        # החזרי אביזרים מבגדים תפורים
                        subc = norm(ln.get('category'))
                        if subc == 'בגדים תפורים':
                            p_name = norm(ln.get('product'))
                            p_size = norm(ln.get('size'))
                            qty_units = int(ln.get('quantity', 0) or 0)
                            if qty_units <= 0:
                                continue
                            try:
                                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                            except Exception:
                                catalog = []
                            def _match(r):
                                return norm(r.get('name')) == p_name and (not p_size or norm(r.get('size')) == p_size)
                            per_ticks = per_elastic = per_ribbon = 0
                            try:
                                for r in catalog:
                                    if _match(r):
                                        per_ticks = int(r.get('ticks_qty', 0) or 0)
                                        per_elastic = int(r.get('elastic_qty', 0) or 0)
                                        per_ribbon = int(r.get('ribbon_qty', 0) or 0)
                                        break
                            except Exception:
                                pass
                            if norm(kind_filter) == 'טיק טק קומפלט' and per_ticks > 0:
                                add(rec_date, 'קליטת בגדים תפורים – אביזרים', rec_no, per_ticks * qty_units, 'נתקבל')
                            if norm(kind_filter) == 'גומי' and per_elastic > 0:
                                add(rec_date, 'קליטת בגדים תפורים – אביזרים', rec_no, per_elastic * qty_units, 'נתקבל')
                            if norm(kind_filter) == 'סרט' and per_ribbon > 0:
                                add(rec_date, 'קליטת בגדים תפורים – אביזרים', rec_no, per_ribbon * qty_units, 'נתקבל')
            except Exception:
                pass

            # מיון והצגה
            def parse_dt(s):
                s = (s or '').strip()
                for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
                    try:
                        return datetime.strptime(s, fmt)
                    except Exception:
                        pass
                return None
            moves_sorted = sorted(moves, key=lambda m: (parse_dt(m['date']) or datetime.min, m['kind'], m['no']))
            for m in moves_sorted:
                # חיפוש טקסטואלי על שדות עיקריים
                hay = f"{m['date']} {m['kind']} {m['no']} {m['direction']}"
                if search_txt and search_txt.lower() not in hay.lower():
                    continue
                self.accessories_tree.insert('', 'end', values=(m['date'], m['kind'], m['no'], m['direction'], m['qty']))
            return

        # סיכום: כל האביזרים שנשלחו/נתקבלו (לא רק 3 העיקריים)
        # יחידות
        unit_by_name = {}
        try:
            for r in getattr(self.data_processor, 'sewing_accessories', []) or []:
                unit_by_name[norm(r.get('name'))] = norm(r.get('unit')) or "יח'"
        except Exception:
            pass
        # כל השמות שנמצאו בשניהם
        all_names = sorted({*shipped.keys(), *received.keys()})
        for name in all_names:
            s = int(shipped.get(name, 0) or 0)
            r = int(received.get(name, 0) or 0)
            diff = s - r
            if search_txt and search_txt not in name.lower():
                continue
            if only_pending and diff <= 0:
                continue
            status = 'הושלם' if diff <= 0 else f"נותרו {diff} לקבל"
            unit = unit_by_name.get(name, "יח'")
            self.accessories_tree.insert('', 'end', values=(name, unit, s, r, max(diff,0), status))
        # אם אין אף שורה – הצג שורת Placeholder כדי לסמן שאין תנועות
        if not all_names:
            unit = ""
            try:
                sel_supplier = (self.balance_supplier_var.get() if hasattr(self, 'balance_supplier_var') else '') or ''
            except Exception:
                sel_supplier = ''
            msg = 'אין תנועות אביזרים עבור ספק זה'
            self.accessories_tree.insert('', 'end', values=(msg, unit, 0, 0, 0, ''))

    def _set_accessories_kind_filter(self, value: str):
        """טוגל סינון לפי סוג אביזר עיקרי (טיק טק קומפלט/גומי/סרט). לחיצה חוזרת מנקה את הסינון."""
        try:
            cur = (self.accessories_kind_filter_var.get() or '').strip()
            if cur == value:
                # חזרה לתצוגת סיכום
                self._set_accessories_summary()
            else:
                self.accessories_kind_filter_var.set(value)
                self._accessories_detail_mode = True
                self._build_accessories_tree(detail=True)
        except Exception:
            pass
        self._refresh_accessories_balance_table()

    def _set_accessories_summary(self):
        # ביטול סינון והצגת שלוש השורות כסיכום
        try:
            self.accessories_kind_filter_var.set('')
        except Exception:
            pass
        self._accessories_detail_mode = False
        try:
            self._build_accessories_tree(detail=False)
        except Exception:
            pass
        self._refresh_accessories_balance_table()

    def _build_accessories_tree(self, detail: bool):
        # השמדת טבלה קודמת
        try:
            if self.accessories_tree is not None:
                self.accessories_tree.destroy()
        except Exception:
            pass
        try:
            if self._accessories_tree_scrollbar is not None:
                self._accessories_tree_scrollbar.destroy()
        except Exception:
            pass
        # בניה
        if detail:
            cols = ('date','doc_type','doc_no','direction','qty')
            headers = {'date':'תאריך','doc_type':'סוג מסמך','doc_no':'מס׳','direction':'תנועה','qty':'כמות'}
            widths = {'date':140,'doc_type':170,'doc_no':120,'direction':90,'qty':90}
        else:
            cols = ('name','unit','shipped','received','diff','status')
            headers = {'name':'שם אביזר','unit':'יחידה','shipped':'נשלח','received':'נתקבל','diff':'הפרש (נותר לקבל)','status':'סטטוס'}
            widths = {'name':300,'unit':80,'shipped':90,'received':90,'diff':140,'status':150}
        self.accessories_tree = ttk.Treeview(self._accessories_page_frame, columns=cols, show='headings', height=18)
        for c in cols:
            self.accessories_tree.heading(c, text=headers[c])
            self.accessories_tree.column(c, width=widths[c], anchor='center')
        self._accessories_tree_scrollbar = ttk.Scrollbar(self._accessories_page_frame, orient='vertical', command=self.accessories_tree.yview)
        self.accessories_tree.configure(yscroll=self._accessories_tree_scrollbar.set)
        self.accessories_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=6)
        self._accessories_tree_scrollbar.pack(side='left', fill='y', pady=6)
        try:
            self.accessories_tree.bind('<Double-1>', self._on_accessories_row_double_click)
            self.accessories_tree.bind('<Return>', self._on_accessories_row_double_click)
        except Exception:
            pass

    def _on_accessories_row_double_click(self, event=None):
        try:
            # אם במצב פירוט – פתח חלון פירוט עבור סוג האביזר המסונן, לא עבור השורה (שהעמודה הראשונה בה היא תאריך)
            if bool(getattr(self, '_accessories_detail_mode', False)):
                try:
                    name = (self.accessories_kind_filter_var.get() or '').strip()
                except Exception:
                    name = ''
                if name:
                    # מעביר מערך שבו התא הראשון הוא שם האביזר כדי שחלון הפירוט יעבוד נכון
                    self._open_accessory_details((name,))
                return

            # במצב סיכום – פתח פירוט לפי האביזר שנבחר מהטבלה
            sel = self.accessories_tree.selection()
            if not sel:
                return
            item_id = sel[0]
            values = self.accessories_tree.item(item_id, 'values') or []
            self._open_accessory_details(values)
        except Exception:
            pass

    def _open_accessory_details(self, row_values):
        """חלון פירוט תנועות עבור אביזר נבחר (נשלח/נתקבל לפי מסמכים)."""
        try:
            supplier = (getattr(self, 'balance_supplier_var', None).get() if hasattr(self, 'balance_supplier_var') else '') or ''
            if not supplier:
                return
            def _get(i, d=''):
                try:
                    return (row_values[i] if i < len(row_values) else d) or d
                except Exception:
                    return d
            name = _get(0, '')
            def norm(s):
                return (s or '').strip()
            try:
                if hasattr(self.data_processor, 'refresh_supplier_receipts'):
                    self.data_processor.refresh_supplier_receipts()
            except Exception:
                pass
            moves = []
            def add(date, kind, doc_no, qty, direction):
                moves.append({'date': date or '', 'kind': kind or '', 'no': str(doc_no or ''), 'qty': int(qty or 0), 'direction': direction or ''})
            # משלוחים
            try:
                for rec in getattr(self.data_processor, 'delivery_notes', []) or []:
                    if norm(rec.get('supplier')) != norm(supplier):
                        continue
                    rec_date = rec.get('date') or rec.get('created_at') or ''
                    rec_no = rec.get('number') or rec.get('id') or ''
                    for ln in rec.get('lines', []) or []:
                        if norm(ln.get('product')) != norm(name):
                            continue
                        qty = int(ln.get('quantity', 0) or 0)
                        if qty > 0:
                            add(rec_date, 'תעודת משלוח', rec_no, qty, 'נשלח')
            except Exception:
                pass
            # קליטות
            try:
                for rec in getattr(self.data_processor, 'supplier_intakes', []) or []:
                    if norm(rec.get('supplier')) != norm(supplier):
                        continue
                    rec_date = rec.get('date') or rec.get('created_at') or ''
                    rec_no = rec.get('number') or rec.get('id') or ''
                    for ln in rec.get('lines', []) or []:
                        # אם זו קליטה של אביזר ישיר
                        if norm(ln.get('product')) == norm(name):
                            qty = int(ln.get('quantity', 0) or 0)
                            if qty > 0:
                                add(rec_date, 'תעודת קליטה', rec_no, qty, 'נתקבל')
                        # אם זו קליטה של בגדים תפורים – חשב אביזרים שחזרו
                        subc = norm(ln.get('category'))
                        if subc == 'בגדים תפורים':
                            p_name = norm(ln.get('product'))
                            p_size = norm(ln.get('size'))
                            qty_units = int(ln.get('quantity', 0) or 0)
                            if qty_units <= 0:
                                continue
                            try:
                                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                            except Exception:
                                catalog = []
                            def _match(r):
                                return norm(r.get('name')) == p_name and (not p_size or norm(r.get('size')) == p_size)
                            per_ticks = per_elastic = per_ribbon = 0
                            try:
                                for r in catalog:
                                    if _match(r):
                                        per_ticks = int(r.get('ticks_qty', 0) or 0)
                                        per_elastic = int(r.get('elastic_qty', 0) or 0)
                                        per_ribbon = int(r.get('ribbon_qty', 0) or 0)
                                        break
                            except Exception:
                                pass
                            if norm(name) == 'טיק טק קומפלט' and per_ticks > 0:
                                add(rec_date, 'קליטת בגדים תפורים – אביזרים', rec_no, per_ticks * qty_units, 'נתקבל')
                            if norm(name) == 'גומי' and per_elastic > 0:
                                add(rec_date, 'קליטת בגדים תפורים – אביזרים', rec_no, per_elastic * qty_units, 'נתקבל')
                            if norm(name) == 'סרט' and per_ribbon > 0:
                                add(rec_date, 'קליטת בגדים תפורים – אביזרים', rec_no, per_ribbon * qty_units, 'נתקבל')
            except Exception:
                pass

            win = tk.Toplevel(self._balance_page_frame)
            win.title(f"פירוט תנועות – אביזר: {name}")
            win.geometry('720x520')
            cols = ('date','kind','no','direction','qty')
            headers = {'date':'תאריך','kind':'סוג מסמך','no':'מס׳','direction':'תנועה','qty':'כמות'}
            widths = {'date':140,'kind':170,'no':120,'direction':90,'qty':80}
            tree = ttk.Treeview(win, columns=cols, show='headings', height=18)
            for c in cols:
                tree.heading(c, text=headers[c])
                tree.column(c, width=widths[c], anchor='center')
            vs = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
            tree.configure(yscroll=vs.set)
            tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=10)
            vs.pack(side='left', fill='y', pady=10)

            def parse_dt(s):
                s = (s or '').strip()
                for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d'):
                    try:
                        return datetime.strptime(s, fmt)
                    except Exception:
                        pass
                return None
            moves_sorted = sorted(moves, key=lambda m: (parse_dt(m['date']) or datetime.min, m['kind'], m['no']))
            for m in moves_sorted:
                tree.insert('', 'end', values=(m['date'], m['kind'], m['no'], m['direction'], m['qty']))

            try:
                shipped_sum = sum(m['qty'] for m in moves if m['direction'] == 'נשלח')
                received_sum = sum(m['qty'] for m in moves if m['direction'] == 'נתקבל')
                diff = shipped_sum - received_sum
                summary = tk.Label(win, text=f"נשלח: {shipped_sum} | נתקבל: {received_sum} | הפרש: {max(diff,0)}", bg='#f7f9fa', anchor='e')
                summary.pack(fill='x', padx=10, pady=(0,10))
            except Exception:
                pass
        except Exception:
            pass
