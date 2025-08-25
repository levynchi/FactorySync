import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import calendar as _cal

import os

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

        # עמוד חדש: מלאי (עם טאב פנימי)
        inventory_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(inventory_page, text='מלאי')
        inv_nb = ttk.Notebook(inventory_page)
        inv_nb.pack(fill='both', expand=True, padx=6, pady=6)

        # תת-טאב: מלאי עדכני (תצוגה)
        inv_view_page = tk.Frame(inv_nb, bg='#f7f9fa')
        inv_nb.add(inv_view_page, text='מלאי עדכני')
        tk.Label(inv_view_page, text='מלאי עדכני מתוך הקטלוג (קריאה בלבד)', font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=(6,2))
        inv_bar = tk.Frame(inv_view_page, bg='#f7f9fa'); inv_bar.pack(fill='x', padx=10, pady=(0,6))
        # מסננים: שם דגם / סוג בד / קטגוריה ראשית
        tk.Label(inv_bar, text='שם דגם:', bg='#f7f9fa').pack(side='right', padx=(8,4))
        self.inv_view_name_filter_var = tk.StringVar(value='הכל')
        self.inv_view_name_filter_cb = ttk.Combobox(inv_bar, textvariable=self.inv_view_name_filter_var, width=28, state='readonly', justify='right')
        try:
            self.inv_view_name_filter_cb.bind('<<ComboboxSelected>>', lambda e: self._refresh_products_inventory_table())
        except Exception:
            pass
        self.inv_view_name_filter_cb.pack(side='right', padx=(0,10))

        tk.Label(inv_bar, text='סוג בד:', bg='#f7f9fa').pack(side='right', padx=(8,4))
        self.inv_view_fabric_filter_var = tk.StringVar(value='הכל')
        self.inv_view_fabric_filter_cb = ttk.Combobox(inv_bar, textvariable=self.inv_view_fabric_filter_var, width=18, state='readonly', justify='right')
        try:
            self.inv_view_fabric_filter_cb.bind('<<ComboboxSelected>>', lambda e: [self._inv_view_rebuild_name_filter_options(), self._refresh_products_inventory_table()])
        except Exception:
            pass
        self.inv_view_fabric_filter_cb.pack(side='right', padx=(0,10))

        tk.Label(inv_bar, text='קטגוריה ראשית:', bg='#f7f9fa').pack(side='right', padx=(8,4))
        self.inv_view_main_cat_filter_var = tk.StringVar(value='הכל')
        self.inv_view_main_cat_filter_cb = ttk.Combobox(inv_bar, textvariable=self.inv_view_main_cat_filter_var, width=18, state='readonly', justify='right')
        try:
            self.inv_view_main_cat_filter_cb.bind('<<ComboboxSelected>>', lambda e: [self._inv_view_rebuild_name_filter_options(), self._refresh_products_inventory_table()])
        except Exception:
            pass
        self.inv_view_main_cat_filter_cb.pack(side='right', padx=(0,10))
        # מיקום: בחירה מרובה או הכל
        tk.Label(inv_bar, text='מיקום:', bg='#f7f9fa').pack(side='right', padx=(8,4))
        self.inv_view_location_summary_var = tk.StringVar(value='כל המקומות')
        self.inv_view_loc_menu_btn = tk.Menubutton(inv_bar, textvariable=self.inv_view_location_summary_var, relief='raised', direction='below')
        self.inv_view_loc_menu = tk.Menu(self.inv_view_loc_menu_btn, tearoff=0)
        self.inv_view_loc_menu_btn.configure(menu=self.inv_view_loc_menu)
        self.inv_view_loc_menu_btn.pack(side='right', padx=(0,10))
        # תצוגת מלאי עדכני נמשכת מהקטלוג – אין צורך בבחירת קובץ
        tk.Button(inv_bar, text='💾 יצוא לאקסל…', command=self._export_products_inventory_to_excel, bg='#27ae60', fg='white').pack(side='right', padx=(6,0))
        tk.Button(inv_bar, text='🔄 רענן', command=self._refresh_products_inventory_table, bg='#3498db', fg='white').pack(side='right', padx=(6,0))
        self.products_inventory_status_var = tk.StringVar(value='מקור: קטלוג מוצרים')
        tk.Label(inv_bar, textvariable=self.products_inventory_status_var, bg='#f7f9fa', anchor='e').pack(side='right', expand=True, fill='x')

        # טבלת מלאי – עמודות נדרשות
        inv_cols = ('name','main_category','size','fabric_type','quantity','location','packaging')
        inv_headers = {
            'name':'שם הדגם',
            'main_category':'קטגוריה ראשית',
            'size':'מידה',
            'fabric_type':'סוג בד',
            'quantity':'כמות',
            'location':'מיקום',
            'packaging':'צורת אריזה'
        }
        inv_widths = {'name':240,'main_category':130,'size':90,'fabric_type':160,'quantity':90,'location':120,'packaging':120}
        inv_table_wrap = tk.Frame(inv_view_page, bg='#ffffff', relief='groove', bd=1)
        inv_table_wrap.pack(fill='both', expand=True, padx=10, pady=6)
        self.products_inventory_tree = ttk.Treeview(inv_table_wrap, columns=inv_cols, show='headings', height=18)
        for c in inv_cols:
            self.products_inventory_tree.heading(c, text=inv_headers[c])
            self.products_inventory_tree.column(c, width=inv_widths[c], anchor='center')
        inv_vs = ttk.Scrollbar(inv_table_wrap, orient='vertical', command=self.products_inventory_tree.yview)
        self.products_inventory_tree.configure(yscroll=inv_vs.set)
        self.products_inventory_tree.grid(row=0, column=0, sticky='nsew')
        inv_vs.grid(row=0, column=1, sticky='ns')
        inv_table_wrap.grid_rowconfigure(0, weight=1)
        inv_table_wrap.grid_columnconfigure(0, weight=1)

        # תת-טאב: יצירת עדכון מלאי
        inv_create_page = tk.Frame(inv_nb, bg='#f7f9fa')
        inv_nb.add(inv_create_page, text='יצירת עדכון')
        tk.Label(inv_create_page, text='עדכון מלאי עדכני', font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=(6,2))
        form = tk.Frame(inv_create_page, bg='#f7f9fa'); form.pack(anchor='e', padx=10, pady=(0,6))

        # שדות
        tk.Label(form, text='שם דגם:', bg='#f7f9fa').grid(row=0, column=6, padx=4, pady=2, sticky='e')
        self.inv_create_name_var = tk.StringVar()
        self.inv_create_name_cb = ttk.Combobox(form, textvariable=self.inv_create_name_var, width=32, justify='right')
        try:
            names = sorted({(r.get('name') or '').strip() for r in getattr(self.data_processor, 'products_catalog', []) if r.get('name')})
            self.inv_create_name_cb['values'] = names
        except Exception:
            pass
        self.inv_create_name_cb.grid(row=0, column=5, padx=4, pady=2, sticky='e')
        try:
            self.inv_create_name_cb.bind('<<ComboboxSelected>>', lambda e: self._inv_create_on_name_change())
        except Exception:
            pass

        tk.Label(form, text='קטגוריה ראשית:', bg='#f7f9fa').grid(row=0, column=4, padx=4, pady=2, sticky='e')
        self.inv_create_main_cat_var = tk.StringVar()
        self.inv_create_main_cat_cb = ttk.Combobox(form, textvariable=self.inv_create_main_cat_var, width=20, state='readonly', justify='right')
        # ערכי ברירת מחדל לקטגוריות ראשיות
        try:
            cats = self._get_main_category_names_for_inventory()
            self.inv_create_main_cat_cb['values'] = cats
            if 'בגדים' in cats:
                self.inv_create_main_cat_var.set('בגדים')
            elif cats:
                self.inv_create_main_cat_var.set(cats[0])
        except Exception:
            pass
        self.inv_create_main_cat_cb.grid(row=0, column=3, padx=4, pady=2, sticky='e')

        tk.Label(form, text='מידה:', bg='#f7f9fa').grid(row=0, column=2, padx=4, pady=2, sticky='e')
        self.inv_create_size_var = tk.StringVar()
        self.inv_create_size_cb = ttk.Combobox(form, textvariable=self.inv_create_size_var, width=16, justify='right')
        self.inv_create_size_cb.grid(row=0, column=1, padx=4, pady=2, sticky='e')

        tk.Label(form, text='סוג בד:', bg='#f7f9fa').grid(row=1, column=6, padx=4, pady=2, sticky='e')
        self.inv_create_fabric_var = tk.StringVar()
        self.inv_create_fabric_cb = ttk.Combobox(form, textvariable=self.inv_create_fabric_var, width=24, justify='right')
        self.inv_create_fabric_cb.grid(row=1, column=5, padx=4, pady=2, sticky='e')

        tk.Label(form, text='כמות:', bg='#f7f9fa').grid(row=1, column=4, padx=4, pady=2, sticky='e')
        self.inv_create_qty_var = tk.StringVar(value='0')
        tk.Entry(form, textvariable=self.inv_create_qty_var, width=10, justify='right').grid(row=1, column=3, padx=4, pady=2, sticky='e')

        tk.Label(form, text='מיקום:', bg='#f7f9fa').grid(row=1, column=2, padx=4, pady=2, sticky='e')
        self.inv_create_location_var = tk.StringVar()
        self.inv_create_location_cb = ttk.Combobox(form, textvariable=self.inv_create_location_var, width=16, justify='right')
        self.inv_create_location_cb.grid(row=1, column=1, padx=4, pady=2, sticky='e')

        tk.Label(form, text='צורת אריזה:', bg='#f7f9fa').grid(row=2, column=6, padx=4, pady=2, sticky='e')
        self.inv_create_packaging_var = tk.StringVar()
        self.inv_create_packaging_cb = ttk.Combobox(form, textvariable=self.inv_create_packaging_var, width=16, justify='right')
        self.inv_create_packaging_cb.grid(row=2, column=5, padx=4, pady=2, sticky='e')

        actions = tk.Frame(inv_create_page, bg='#f7f9fa'); actions.pack(fill='x', padx=10, pady=(0,6))
        tk.Button(actions, text='➕ הוסף לשורות', command=self._inv_create_add_row, bg='#27ae60', fg='white').pack(side='right', padx=4)
        tk.Button(actions, text='🗑️ הסר שורה', command=self._inv_create_delete_selected).pack(side='right', padx=4)
        tk.Button(actions, text='🧹 נקה הכל', command=self._inv_create_clear_all).pack(side='right', padx=4)
        # כפתור עדכון שיחיל את השורות על המלאי וירשום להיסטוריה
        tk.Button(actions, text='⬆️ עדכן', command=self._inv_create_apply_updates, bg='#2ecc71', fg='white').pack(side='left', padx=4)
        tk.Button(actions, text='💾 שמור לאקסל…', command=self._inv_create_export_to_excel, bg='#2c3e50', fg='white').pack(side='left', padx=4)

        # טבלת בנייה
        self.inv_create_tree = ttk.Treeview(inv_create_page, columns=inv_cols, show='headings', height=14)
        for c in inv_cols:
            self.inv_create_tree.heading(c, text=inv_headers[c])
            self.inv_create_tree.column(c, width=inv_widths[c], anchor='center')
        ivs2 = ttk.Scrollbar(inv_create_page, orient='vertical', command=self.inv_create_tree.yview)
        self.inv_create_tree.configure(yscroll=ivs2.set)
        self.inv_create_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=(0,8))
        ivs2.pack(side='left', fill='y', pady=(0,8))

        # טען אפשרויות ברירת מחדל לשדות מיקום/צורת אריזה
        try:
            self._reload_inventory_aux_options()
        except Exception:
            pass

        # תת-טאב: היסטוריית עדכונים
        inv_hist_page = tk.Frame(inv_nb, bg='#f7f9fa')
        inv_nb.add(inv_hist_page, text='היסטוריית עדכונים')
        tk.Label(inv_hist_page, text='היסטוריית עדכוני מלאי', font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=(6,2))
        hist_bar = tk.Frame(inv_hist_page, bg='#f7f9fa'); hist_bar.pack(fill='x', padx=10, pady=(0,6))
        tk.Button(hist_bar, text='🔄 רענן', command=self._inv_history_reload, bg='#3498db', fg='white').pack(side='right')
        # חלוקה לשניים: תקציר למעלה, פריטים למטה
        hist_wrap = tk.Frame(inv_hist_page, bg='#f7f9fa'); hist_wrap.pack(fill='both', expand=True, padx=10, pady=6)
        hist_wrap.grid_columnconfigure(0, weight=1)
        hist_wrap.grid_rowconfigure(1, weight=1)
        # טבלת באצ׳ים
        self.inv_updates_batches_tree = ttk.Treeview(hist_wrap, columns=(
            'id','created_at','items_count'
        ), show='headings', height=6)
        self.inv_updates_batches_tree.heading('id', text='מזהה')
        self.inv_updates_batches_tree.heading('created_at', text='נוצר בתאריך')
        self.inv_updates_batches_tree.heading('items_count', text='מס׳ פריטים')
        self.inv_updates_batches_tree.column('id', width=120, anchor='center')
        self.inv_updates_batches_tree.column('created_at', width=180, anchor='center')
        self.inv_updates_batches_tree.column('items_count', width=100, anchor='center')
        self.inv_updates_batches_tree.grid(row=0, column=0, sticky='ew')
        try:
            self.inv_updates_batches_tree.bind('<<TreeviewSelect>>', self._on_inv_history_select)
        except Exception:
            pass
        # טבלת פרטים
        self.inv_updates_details_tree = ttk.Treeview(hist_wrap, columns=(
            'name','main_category','size','fabric_type','quantity','location','packaging'
        ), show='headings', height=12)
        headers = {'name':'שם הדגם','main_category':'קטגוריה ראשית','size':'מידה','fabric_type':'סוג בד','quantity':'כמות','location':'מיקום','packaging':'צורת אריזה'}
        widths = {'name':240,'main_category':130,'size':90,'fabric_type':160,'quantity':90,'location':120,'packaging':120}
        for c in ('name','main_category','size','fabric_type','quantity','location','packaging'):
            self.inv_updates_details_tree.heading(c, text=headers[c])
            self.inv_updates_details_tree.column(c, width=widths[c], anchor='center')
        vs_hist = ttk.Scrollbar(hist_wrap, orient='vertical', command=self.inv_updates_details_tree.yview)
        self.inv_updates_details_tree.configure(yscroll=vs_hist.set)
        self.inv_updates_details_tree.grid(row=1, column=0, sticky='nsew')
        vs_hist.grid(row=1, column=1, sticky='ns')

        # תת-טאב ניהול: צורות אריזה
        pkg_page = tk.Frame(inv_nb, bg='#f7f9fa')
        inv_nb.add(pkg_page, text='צורות אריזה')
        tk.Label(pkg_page, text='ניהול צורות אריזה', font=('Arial',14,'bold'), bg='#f7f9fa').pack(pady=(6,2))
        pkg_bar = tk.Frame(pkg_page, bg='#f7f9fa'); pkg_bar.pack(fill='x', padx=10, pady=(0,6))
        self.pkg_new_var = tk.StringVar()
        tk.Entry(pkg_bar, textvariable=self.pkg_new_var, width=24).pack(side='right', padx=6)
        tk.Button(pkg_bar, text='➕ הוסף', command=self._inv_pkg_add).pack(side='right')
        tk.Button(pkg_bar, text='🗑️ מחק נבחר', command=self._inv_pkg_delete).pack(side='left', padx=6)
        tk.Button(pkg_bar, text='💾 שמור', command=self._inv_pkg_save).pack(side='left')
        self.pkg_list = tk.Listbox(pkg_page, height=12)
        self.pkg_list.pack(fill='both', expand=True, padx=10, pady=6)
        try:
            self._load_pkg_list()
        except Exception:
            pass

        # תת-טאב ניהול: מיקומים
        loc_page = tk.Frame(inv_nb, bg='#f7f9fa')
        inv_nb.add(loc_page, text='מיקומים')
        tk.Label(loc_page, text='ניהול מיקומים', font=('Arial',14,'bold'), bg='#f7f9fa').pack(pady=(6,2))
        loc_bar = tk.Frame(loc_page, bg='#f7f9fa'); loc_bar.pack(fill='x', padx=10, pady=(0,6))
        self.loc_new_var = tk.StringVar()
        tk.Entry(loc_bar, textvariable=self.loc_new_var, width=24).pack(side='right', padx=6)
        tk.Button(loc_bar, text='➕ הוסף', command=self._inv_loc_add).pack(side='right')
        tk.Button(loc_bar, text='🗑️ מחק נבחר', command=self._inv_loc_delete).pack(side='left', padx=6)
        tk.Button(loc_bar, text='💾 שמור', command=self._inv_loc_save).pack(side='left')
        self.loc_list = tk.Listbox(loc_page, height=12)
        self.loc_list.pack(fill='both', expand=True, padx=10, pady=6)
        try:
            self._load_loc_list()
        except Exception:
            pass

        # נסה לטעון אוטומטית קובץ מלאי אחרון מהגדרות
        try:
            last_inv = ''
            if hasattr(self, 'settings') and hasattr(self.settings, 'get'):
                last_inv = self.settings.get('app.last_products_inventory_file', '') or ''
            self.products_inventory_file = last_inv
        except Exception:
            self.products_inventory_file = ''
        # אתחול ערכים למסננים של "מלאי עדכני"
        try:
            self._reload_inventory_view_filters_options()
        except Exception:
            pass
        # אתחול אפשרויות מיקום
        try:
            self._inv_view_reload_location_options()
        except Exception:
            pass
        self._refresh_products_inventory_table()

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
            # ברירת מחדל: היום
            try:
                _today_str = datetime.now().strftime('%Y-%m-%d')
                self.accessories_to_date_var.set(_today_str)
            except Exception:
                pass
            if DateEntry is not None:
                acc_to_entry = DateEntry(acc_bar, textvariable=self.accessories_to_date_var, width=12, date_pattern='yyyy-mm-dd', locale='he_IL')
                try: acc_to_entry.bind('<<DateEntrySelected>>', lambda e: self._refresh_accessories_balance_table())
                except Exception: pass
                # החלת ברירת המחדל בפועל על ה-Widget
                try:
                    acc_to_entry.set_date(datetime.now())
                except Exception:
                    pass
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
            # ברירת מחדל: תחילת השנה הנוכחית
            try:
                _now = datetime.now(); _start_year = datetime(_now.year, 1, 1)
                self.accessories_from_date_var.set(_start_year.strftime('%Y-%m-%d'))
            except Exception:
                pass
            if DateEntry is not None:
                acc_from_entry = DateEntry(acc_bar, textvariable=self.accessories_from_date_var, width=12, date_pattern='yyyy-mm-dd', locale='he_IL')
                try: acc_from_entry.bind('<<DateEntrySelected>>', lambda e: self._refresh_accessories_balance_table())
                except Exception: pass
                # החלת ברירת המחדל בפועל על ה-Widget
                try:
                    _now = datetime.now(); acc_from_entry.set_date(datetime(_now.year, 1, 1))
                except Exception:
                    pass
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

    def _reload_inventory_view_filters_options(self):
        """טוען ערכי אפשרויות למסנני 'מלאי עדכני' מתוך הקטלוג."""
        try:
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
        except Exception:
            catalog = []
        def norm(s):
            return (str(s or '').strip())
        # הפקת קטגוריה ראשית גם כאשר השדה לא קיים ברשומות הקטלוג (נשתמש בנגזר/ברירת מחדל)
        def _derive_mc(rec: dict):
            try:
                mc = norm(rec.get('main_category'))
                if mc:
                    return mc
                # נסה מתוך השדה 'category' – קח הטוקן הראשון לפני פסיק
                cat = norm(rec.get('category'))
                if cat:
                    first = norm(cat.split(',')[0]) if ',' in cat else cat
                    if first:
                        return first
                # נסה למצוא לפי שם הדגם מתוך טבלת שמות-דגמים
                name = norm(rec.get('name'))
                if name:
                    model_names = getattr(self.data_processor, 'product_model_names', []) or getattr(self.data_processor, 'model_names', []) or []
                    for m in model_names:
                        if norm(m.get('name')) == name and norm(m.get('main_category')):
                            return norm(m.get('main_category'))
                # לבסוף – ברירת מחדל
                return 'בגדים'
            except Exception:
                return 'בגדים'
        main_cats = sorted({ _derive_mc(r) for r in catalog if r })
        fabrics = sorted({norm(r.get('fabric_type')) for r in catalog if r.get('fabric_type')})
        try:
            values_main = ['הכל'] + main_cats
            values_fab = ['הכל'] + fabrics
            if hasattr(self, 'inv_view_main_cat_filter_cb'):
                self.inv_view_main_cat_filter_cb['values'] = values_main
            if hasattr(self, 'inv_view_fabric_filter_cb'):
                self.inv_view_fabric_filter_cb['values'] = values_fab
            # קבע ברירות מחדל אם ריק
            if hasattr(self, 'inv_view_main_cat_filter_var') and not (self.inv_view_main_cat_filter_var.get() or '').strip():
                self.inv_view_main_cat_filter_var.set('הכל')
            if hasattr(self, 'inv_view_fabric_filter_var') and not (self.inv_view_fabric_filter_var.get() or '').strip():
                self.inv_view_fabric_filter_var.set('הכל')
            # בנה אפשרויות לשם דגם בהתאם למסננים הפעילים
            try:
                self._inv_view_rebuild_name_filter_options()
            except Exception:
                pass
        except Exception:
            pass

    def _inv_view_rebuild_name_filter_options(self):
        """בונה את רשימת 'שם דגם' בהתאם לקטגוריה ראשית ולסוג בד הנבחרים (או 'הכל')."""
        try:
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
        except Exception:
            catalog = []
        def norm(s):
            return (str(s or '').strip())
        # קריאת המסננים הפעילים
        try:
            mc_sel = norm(getattr(self, 'inv_view_main_cat_filter_var', tk.StringVar(value='הכל')).get())
        except Exception:
            mc_sel = 'הכל'
        try:
            fab_sel = norm(getattr(self, 'inv_view_fabric_filter_var', tk.StringVar(value='הכל')).get())
        except Exception:
            fab_sel = 'הכל'
        # פונקציה להפקת קטגוריה ראשית מהשורה
        def _derive_mc(rec: dict):
            mc = norm(rec.get('main_category'))
            if not mc:
                cat = norm(rec.get('category'))
                if cat:
                    mc = norm(cat.split(',')[0]) if ',' in cat else cat
            if not mc:
                name = norm(rec.get('name'))
                if name:
                    model_names = getattr(self.data_processor, 'product_model_names', []) or getattr(self.data_processor, 'model_names', []) or []
                    for m in model_names:
                        if norm(m.get('name')) == name and norm(m.get('main_category')):
                            mc = norm(m.get('main_category')); break
            return mc or 'בגדים'
        # סינון לפי מסננים נבחרים
        def _match(rec):
            if mc_sel != 'הכל' and _derive_mc(rec) != mc_sel:
                return False
            if fab_sel != 'הכל' and norm(rec.get('fabric_type')) != fab_sel:
                return False
            return True
        names = sorted({ norm(r.get('name')) for r in catalog if r.get('name') and _match(r) })
        try:
            if hasattr(self, 'inv_view_name_filter_cb'):
                self.inv_view_name_filter_cb['values'] = ['הכל'] + names
                # אם הערך הנוכחי לא קיים – אפס ל'הכל'
                cur = (self.inv_view_name_filter_var.get() or '').strip()
                if not cur or (cur != 'הכל' and cur not in names):
                    self.inv_view_name_filter_var.set('הכל')
        except Exception:
            pass

    # ===== Inventory view: Location multi-select filter =====
    def _inv_view_reload_location_options(self):
        """בונה את תפריט המיקומים לבחירה מרובה, כולל 'כל המקומות' ו'ללא'."""
        # קבל רשימת מיקומים מהגדרות
        locations = []
        try:
            if hasattr(self, 'settings') and hasattr(self.settings, 'get'):
                locations = self.settings.get('inventory.location_options', []) or []
        except Exception:
            locations = []
        # בנה פריטים עם BooleanVar לכל מיקום
        self._inv_view_loc_vars = {}
        menu = getattr(self, 'inv_view_loc_menu', None)
        if not menu:
            return
        try:
            menu.delete(0, 'end')
        except Exception:
            pass
        # פעולה: עדכון טקסט תצוגה ורענון טבלה
        def _on_change():
            self._inv_view_update_location_summary()
            try:
                self._refresh_products_inventory_table()
            except Exception:
                pass
        # פריט מיוחד: כל המקומות – מסמן/מנקה הכל
        menu.add_command(label='כל המקומות', command=lambda: [self._inv_view_set_all_locations(True), _on_change()])
        menu.add_command(label='נקה בחירה', command=lambda: [self._inv_view_set_all_locations(False), _on_change()])
        menu.add_separator()
        # פריט 'ללא' עבור ערך מיקום ריק
        none_var = tk.BooleanVar(value=False)
        self._inv_view_loc_vars['__NONE__'] = none_var
        menu.add_checkbutton(label='ללא', variable=none_var, onvalue=True, offvalue=False, command=_on_change)
        # שאר המיקומים לפי א-ב
        for loc in sorted({str(x).strip() for x in locations if str(x).strip()}):
            var = tk.BooleanVar(value=False)
            self._inv_view_loc_vars[loc] = var
            menu.add_checkbutton(label=loc, variable=var, onvalue=True, offvalue=False, command=_on_change)
        # טקסט סיכום התחלתי
        self._inv_view_update_location_summary()

    def _inv_view_set_all_locations(self, checked: bool):
        for v in getattr(self, '_inv_view_loc_vars', {}).values():
            try:
                v.set(bool(checked))
            except Exception:
                pass

    def _inv_view_get_selected_locations(self):
        """החזרת קבוצה של מיקומים שנבחרו. '__NONE__' מייצג ערך ריק."""
        selected = set()
        for key, var in (getattr(self, '_inv_view_loc_vars', {}) or {}).items():
            try:
                if bool(var.get()):
                    selected.add(key)
            except Exception:
                pass
        return selected

    def _inv_view_update_location_summary(self):
        sel = self._inv_view_get_selected_locations()
        label = 'כל המקומות'
        if sel:
            # הצג עד 2 ראשונים + "ועוד"
            names = []
            for k in sel:
                names.append('ללא' if k == '__NONE__' else k)
            names = sorted(names)
            if len(names) <= 2:
                label = ', '.join(names)
            else:
                label = f"{names[0]}, {names[1]} ועוד"
        try:
            self.inv_view_location_summary_var.set(label)
        except Exception:
            pass

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

        # === מוצרים: טאב מלאי ===
    def _browse_products_inventory_file(self):
        from tkinter import filedialog
        try:
            path = filedialog.askopenfilename(title='בחר קובץ מלאי מוצרים', filetypes=[('Excel','*.xlsx;*.xls')])
        except Exception:
            path = ''
        if not path:
            return
        self.products_inventory_file = path
        # שמירה בהגדרות עבור טעינה אוטומטית בפעם הבאה
        try:
            if hasattr(self, 'settings') and hasattr(self.settings, 'set'):
                self.settings.set('app.last_products_inventory_file', path)
        except Exception:
            pass
        self._refresh_products_inventory_table()

    def _refresh_products_inventory_table(self):
        """מציג מלאי עדכני מתוך קטלוג המוצרים (ללא קובץ חיצוני).

        העמודות: שם הדגם, קטגוריה ראשית, מידה, סוג בד, כמות, מיקום, צורת אריזה.
        מאחר והקטלוג לא כולל כמות/מיקום/אריזה – שדות אלה יוצגו ריקים כברירת מחדל.
        """
        tree = getattr(self, 'products_inventory_tree', None)
        if not tree:
            return
        # רענון אפשרויות המסננים (למקרה שהקטלוג השתנה)
        try:
            self._reload_inventory_view_filters_options()
        except Exception:
            pass
        # ניקוי טבלה
        try:
            for iid in tree.get_children():
                tree.delete(iid)
        except Exception:
            pass
    # שליפת נתונים מהקטלוג
        rows = []
        try:
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
            for rec in catalog:
                name = (rec.get('name') or '').strip()
                if not name:
                    continue
                # הפקת קטגוריה ראשית גם אם אינה שמורה כתכונה מפורשת
                mc = ''
                try:
                    mc = (rec.get('main_category') or '').strip()
                    if not mc:
                        cat = (rec.get('category') or '').strip()
                        if cat:
                            mc = (cat.split(',')[0] or '').strip()
                    if not mc:
                        # חפש לפי טבלת שמות-דגמים
                        try:
                            model_names = getattr(self.data_processor, 'product_model_names', []) or getattr(self.data_processor, 'model_names', []) or []
                        except Exception:
                            model_names = []
                        for m in model_names:
                            try:
                                if (m.get('name') or '').strip() == name and (m.get('main_category') or '').strip():
                                    mc = (m.get('main_category') or '').strip(); break
                            except Exception:
                                pass
                    if not mc:
                        mc = 'בגדים'
                except Exception:
                    mc = 'בגדים'
                rows.append({
                    'name': name,
                    'main_category': mc,
                    'size': (rec.get('size') or '').strip(),
                    'fabric_type': (rec.get('fabric_type') or '').strip(),
                    'quantity': '',
                    'location': '',
                    'packaging': '',
                })
        except Exception:
            rows = []

        # החלת עדכונים שמורים (אם קיימים) על רשומות התאמה לפי name+size+fabric_type+location
        try:
            updates_data = getattr(self, '_inventory_updates', None)
            if updates_data is None:
                updates_data = self._inv_updates_load_store()
                self._inventory_updates = updates_data
        except Exception:
            updates_data = None
        # מפה אחרונה לכל צירוף מזהה
        last_map = {}
        try:
            if updates_data and isinstance(updates_data.get('batches', []), list):
                for b in updates_data['batches']:
                    for it in (b.get('items') or []):
                        key = (
                            (it.get('name') or '').strip(),
                            (it.get('size') or '').strip(),
                            (it.get('fabric_type') or '').strip(),
                            (it.get('location') or '').strip()
                        )
                        last_map[key] = {
                            'quantity': (str(it.get('quantity') or '').strip()),
                            'packaging': (it.get('packaging') or '').strip()
                        }
        except Exception:
            last_map = {}
        # החלה על rows: אם קיים עדכון עבור מפתח התואם למיקום ריק – נשאיר loc ריק; אחר כך נביא גם לפי כל מיקום מפורש
        try:
            for r in rows:
                key_same_loc = (r.get('name',''), r.get('size',''), r.get('fabric_type',''), r.get('location',''))
                if key_same_loc in last_map:
                    up = last_map[key_same_loc]
                    r['quantity'] = up.get('quantity','')
                    if up.get('packaging',''):
                        r['packaging'] = up.get('packaging','')
            # הוספת שורות נוספות עבור מיקומים שהתעדכנו שאינם קיימים כעת בקטלוג (לפי אותו name/size/fabric)
            extra_rows = []
            for key, up in last_map.items():
                n, s, ft, loc = key
                # אם יש רשומת בסיס בקטלוג עבור n,s,ft – נייצר/נעדכן שורה עם המיקום הספציפי
                base = next((x for x in rows if x['name']==n and x['size']==s and x['fabric_type']==ft), None)
                if base is not None:
                    # אם המיקום במקור ריק – נעדכן באותו רשומה; אם לא, ניצור שכפול ייעודי למיקום
                    if (base.get('location') or '') == '':
                        base['location'] = loc
                        base['quantity'] = up.get('quantity','')
                        if up.get('packaging',''):
                            base['packaging'] = up.get('packaging','')
                    else:
                        newr = dict(base)
                        newr['location'] = loc
                        newr['quantity'] = up.get('quantity','')
                        if up.get('packaging',''):
                            newr['packaging'] = up.get('packaging','')
                        extra_rows.append(newr)
            if extra_rows:
                rows.extend(extra_rows)
        except Exception:
            pass
        # החלת מסננים
        def norm(s):
            return (str(s or '').strip())
        try:
            name_filter = (getattr(self, 'inv_view_name_filter_var', tk.StringVar(value='הכל')).get() or '').strip()
        except Exception:
            name_filter = 'הכל'
        try:
            main_cat_filter = norm(getattr(self, 'inv_view_main_cat_filter_var', tk.StringVar(value='הכל')).get())
        except Exception:
            main_cat_filter = 'הכל'
        try:
            fabric_filter = norm(getattr(self, 'inv_view_fabric_filter_var', tk.StringVar(value='הכל')).get())
        except Exception:
            fabric_filter = 'הכל'
        # מיקומים: סט של בחירות; אם ריק => התייחס כ"כל המקומות"
        try:
            selected_locations = set(self._inv_view_get_selected_locations())
        except Exception:
            selected_locations = set()
        filtered = []
        try:
            for r in rows:
                # שם דגם: בחירה מדויקת או 'הכל'
                if name_filter and name_filter != 'הכל' and (r.get('name','').strip() != name_filter):
                    continue
                if main_cat_filter and main_cat_filter != 'הכל' and norm(r.get('main_category','')) != main_cat_filter:
                    continue
                if fabric_filter and fabric_filter != 'הכל' and norm(r.get('fabric_type','')) != fabric_filter:
                    continue
                # מיקום: אם יש בחירות, חייב להתאים
                if selected_locations:
                    loc_val = norm(r.get('location',''))
                    # '__NONE__' מייצג שורה ללא מיקום
                    match = (loc_val in selected_locations) or (loc_val == '' and '__NONE__' in selected_locations)
                    if not match:
                        continue
                filtered.append(r)
        except Exception:
            filtered = rows
        # הצגה (הגבלה ל-3000 שורות להגנה על הביצועים)
        try:
            for r in filtered[:3000]:
                tree.insert('', 'end', values=(
                    r.get('name',''), r.get('main_category',''), r.get('size',''), r.get('fabric_type',''), r.get('quantity',''), r.get('location',''), r.get('packaging','')
                ))
        except Exception:
            pass
        # עדכון סטטוס
        try:
            self.products_inventory_status_var.set(f"מקור: קטלוג מוצרים | מסונן: {len(filtered)} מתוך {len(rows)} | שורות מוצגות: {min(len(filtered),3000)}")
        except Exception:
            pass

    # === עדכון מלאי: התמדה, היסטוריה ויישום ===
    def _inv_updates_store_path(self) -> str:
        try:
            # נשתמש בשורש העבודה של האפליקציה (שם שמורים שאר ה-json)
            return os.path.join(os.getcwd(), 'inventory_updates.json')
        except Exception:
            return 'inventory_updates.json'

    def _inv_updates_load_store(self) -> dict:
        import json
        try:
            path = self._inv_updates_store_path()
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, dict) and 'batches' in data:
                        return data
            return {'batches': []}
        except Exception:
            return {'batches': []}

    def _inv_updates_save_store(self, data: dict) -> bool:
        import json
        try:
            path = self._inv_updates_store_path()
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data or {'batches': []}, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False

    def _inv_create_apply_updates(self):
        """אוסף את שורות 'עדכון מלאי' ושומר באצ׳ חדש לקובץ ההיסטוריה, ואז מרענן את המלאי."""
        tree = getattr(self, 'inv_create_tree', None)
        if tree is None:
            return
        # אסוף פריטים
        items = []
        try:
            for iid in tree.get_children():
                vals = list(tree.item(iid, 'values') or [])
                # צורה: name, main_category, size, fabric_type, quantity, location, packaging
                item = {
                    'name': (vals[0] if len(vals)>0 else '').strip(),
                    'main_category': (vals[1] if len(vals)>1 else '').strip(),
                    'size': (vals[2] if len(vals)>2 else '').strip(),
                    'fabric_type': (vals[3] if len(vals)>3 else '').strip(),
                    'quantity': (vals[4] if len(vals)>4 else '').strip(),
                    'location': (vals[5] if len(vals)>5 else '').strip(),
                    'packaging': (vals[6] if len(vals)>6 else '').strip(),
                }
                if item['name']:
                    items.append(item)
        except Exception:
            items = []
        # נטפל גם במקרה שאין שורות
        if not items:
            try:
                messagebox.showwarning('עדכון מלאי', 'אין שורות לעדכון')
            except Exception:
                pass
            return
        # טען/שמור היסטוריה
        data = self._inv_updates_load_store()
        from datetime import datetime as _dt
        batch_id = f"batch_{_dt.now().strftime('%Y%m%d_%H%M%S')}"
        batch = {'id': batch_id, 'created_at': _dt.now().strftime('%Y-%m-%d %H:%M:%S'), 'items': items}
        try:
            data.setdefault('batches', []).append(batch)
            self._inv_updates_save_store(data)
            # שמור בזיכרון להפחתת קריאות
            self._inventory_updates = data
        except Exception:
            pass
        # ניקוי העץ של היצירה לאחר החלה (רשות, שומרות רצף עבודה נקי)
        try:
            for iid in list(tree.get_children()):
                tree.delete(iid)
        except Exception:
            pass
        # ריענון תצוגת המלאי והיסטוריה והודעה למשתמש
        try:
            self._refresh_products_inventory_table()
        except Exception:
            pass
        try:
            self._inv_history_reload()
        except Exception:
            pass
        try:
            messagebox.showinfo('עדכון מלאי', 'המלאי עודכן בהצלחה')
        except Exception:
            pass

    def _inv_history_reload(self):
        """טוען מחדש את טבלת ההיסטוריה של עדכוני המלאי ואת פרטי הבאצ׳ הנבחר (אם יש)."""
        data = None
        try:
            data = self._inv_updates_load_store()
            self._inventory_updates = data
        except Exception:
            data = {'batches': []}
        # בנה טבלת באצ׳ים
        tree = getattr(self, 'inv_updates_batches_tree', None)
        if tree is not None:
            try:
                for iid in list(tree.get_children()):
                    tree.delete(iid)
            except Exception:
                pass
            try:
                # סדר יורד לפי תאריך יצירה
                batches = list((data or {}).get('batches', []))
                def _parse_dt(b):
                    from datetime import datetime as _dt
                    s = (b.get('created_at') or '').strip()
                    try:
                        return _dt.strptime(s, '%Y-%m-%d %H:%M:%S')
                    except Exception:
                        return _dt.min
                batches.sort(key=_parse_dt, reverse=True)
                for b in batches:
                    bid = b.get('id') or ''
                    cat = b.get('created_at') or ''
                    cnt = len(b.get('items') or [])
                    tree.insert('', 'end', values=(bid, cat, cnt), iid=bid)
            except Exception:
                pass
        # נקה פרטים
        d_tree = getattr(self, 'inv_updates_details_tree', None)
        if d_tree is not None:
            try:
                for iid in list(d_tree.get_children()):
                    d_tree.delete(iid)
            except Exception:
                pass

    def _on_inv_history_select(self, event=None):
        """בעת בחירת באצ׳ – מציג את פרטי הפריטים שלו."""
        tree = getattr(self, 'inv_updates_batches_tree', None)
        d_tree = getattr(self, 'inv_updates_details_tree', None)
        if tree is None or d_tree is None:
            return
        sel = ()
        try:
            sel = tree.selection()
        except Exception:
            sel = ()
        if not sel:
            return
        bid = sel[0]
        data = getattr(self, '_inventory_updates', None) or self._inv_updates_load_store()
        batch = None
        try:
            for b in (data or {}).get('batches', []):
                if (b.get('id') or '') == bid:
                    batch = b; break
        except Exception:
            batch = None
        # הצג פריטים
        try:
            for iid in list(d_tree.get_children()):
                d_tree.delete(iid)
        except Exception:
            pass
        if not batch:
            return
        try:
            for it in (batch.get('items') or []):
                d_tree.insert('', 'end', values=(
                    it.get('name',''), it.get('main_category',''), it.get('size',''), it.get('fabric_type',''), it.get('quantity',''), it.get('location',''), it.get('packaging','')
                ))
        except Exception:
            pass

    # === יצירת קובץ מלאי: עזרים ===
    def _inv_create_on_name_change(self):
        """בעת בחירת שם דגם – ננסה למלא קטגוריה ראשית ולהציע מידות/סוגי בד מהקטלוג."""
        name = (self.inv_create_name_var.get() or '').strip()
        sizes = set(); fabrics = set(); main_cat = ''
        try:
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
            for rec in catalog:
                if (rec.get('name') or '').strip() == name:
                    if not main_cat:
                        main_cat = (rec.get('main_category') or '').strip()
                    s = (rec.get('size') or '').strip()
                    if s:
                        sizes.add(s)
                    f = (rec.get('fabric_type') or '').strip()
                    if f:
                        fabrics.add(f)
        except Exception:
            pass
        try:
            # קבע בחירת קטגוריה והצע כל האפשרויות הקיימות עבור המוצר
            cats = set()
            for rec in getattr(self.data_processor, 'products_catalog', []) or []:
                if (rec.get('name') or '').strip() == name:
                    mc = (rec.get('main_category') or '').strip()
                    if mc:
                        cats.add(mc)
            if not cats:
                cats = set(self._get_main_category_names_for_inventory())
            vals = sorted(cats, key=lambda x: (x!='בגדים', x))
            self.inv_create_main_cat_cb['values'] = vals
            if main_cat and main_cat in vals:
                self.inv_create_main_cat_var.set(main_cat)
            else:
                # ברירת מחדל 'בגדים'
                self.inv_create_main_cat_var.set('בגדים' if 'בגדים' in vals else (vals[0] if vals else ''))
        except Exception:
            pass
        try:
            self.inv_create_size_cb['values'] = sorted(sizes)
            self.inv_create_fabric_cb['values'] = sorted(fabrics)
        except Exception:
            pass

    def _inv_create_add_row(self):
        name = (self.inv_create_name_var.get() or '').strip()
        if not name:
            try: messagebox.showwarning('יצירת מלאי', 'בחר שם דגם')
            except Exception: pass
            return
        size = (self.inv_create_size_var.get() or '').strip()
        fabric = (self.inv_create_fabric_var.get() or '').strip()
        qty = (self.inv_create_qty_var.get() or '').strip()
        location = (self.inv_create_location_var.get() or '').strip()
        packaging = (self.inv_create_packaging_var.get() or '').strip()
        main_cat = (self.inv_create_main_cat_var.get() or '').strip()
        self.inv_create_tree.insert('', 'end', values=(name, main_cat, size, fabric, qty, location, packaging))

    def _inv_create_delete_selected(self):
        tree = getattr(self, 'inv_create_tree', None)
        if not tree:
            return
        for sel in tree.selection():
            try: tree.delete(sel)
            except Exception: pass

    def _inv_create_clear_all(self):
        tree = getattr(self, 'inv_create_tree', None)
        if not tree:
            return
        for iid in tree.get_children():
            try: tree.delete(iid)
            except Exception: pass

    def _inv_create_export_to_excel(self):
        """ייצוא טבלת היצירה לאקסל."""
        tree = getattr(self, 'inv_create_tree', None)
        if tree is None:
            return
        rows = []
        try:
            for iid in tree.get_children():
                rows.append(tuple(tree.item(iid, 'values') or []))
        except Exception:
            pass
        headers = ['שם הדגם','קטגוריה ראשית','מידה','סוג בד','כמות','מיקום','צורת אריזה']
        from tkinter import filedialog
        from datetime import datetime as _dt
        default_name = f"קובץ_מלאי_חדש_{_dt.now().strftime('%Y-%m-%d')}.xlsx"
        try:
            save_path = filedialog.asksaveasfilename(title='שמירת קובץ מלאי', defaultextension='.xlsx', initialfile=default_name, filetypes=[('Excel','*.xlsx')])
        except Exception:
            save_path = ''
        if not save_path:
            return
        try:
            from openpyxl import Workbook  # type: ignore
            wb = Workbook(); ws = wb.active; ws.title = 'Inventory'
            ws.append(headers)
            def to_num(v):
                s = str(v).strip()
                try:
                    if s == '':
                        return ''
                    f = float(s.replace(',', ''))
                    return int(f) if abs(f-int(f))<1e-9 else f
                except Exception:
                    return s
            for r in rows:
                out = [
                    r[0] if len(r)>0 else '', r[1] if len(r)>1 else '', r[2] if len(r)>2 else '', r[3] if len(r)>3 else '',
                    to_num(r[4] if len(r)>4 else ''), r[5] if len(r)>5 else '', r[6] if len(r)>6 else ''
                ]
                ws.append(out)
            wb.save(save_path)
            try: messagebox.showinfo('שמירה', f'נשמר בהצלחה:\n{save_path}')
            except Exception: pass
        except Exception as e:
            try: messagebox.showerror('שמירה', f'שגיאה:\n{e}')
            except Exception: pass

    def _export_products_inventory_to_excel(self):
        """ייצוא השורות המוצגות בטבלת המלאי לקובץ Excel."""
        tree = getattr(self, 'products_inventory_tree', None)
        # איסוף נתונים מהטבלה (גם אם אין – נייצא קובץ ריק עם כותרות)
        rows = []
        try:
            if tree is not None:
                for iid in tree.get_children():
                    vals = tree.item(iid, 'values') or []
                    rows.append(tuple(vals))
        except Exception:
            pass
        # כותרות מתוך הטבלה
        cols = ('name','main_category','size','fabric_type','quantity','location','packaging')
        headers = ['שם הדגם','קטגוריה ראשית','מידה','סוג בד','כמות','מיקום','צורת אריזה']
        # דיאלוג שמירה
        from tkinter import filedialog
        from datetime import datetime as _dt
        default_name = f"מלאי_מוצרים_{_dt.now().strftime('%Y-%m-%d')}.xlsx"
        try:
            save_path = filedialog.asksaveasfilename(title='שמירת מלאי לאקסל', defaultextension='.xlsx', initialfile=default_name, filetypes=[('Excel','*.xlsx')])
        except Exception:
            save_path = ''
        if not save_path:
            return
        # כתיבה עם openpyxl
        try:
            from openpyxl import Workbook  # type: ignore
            wb = Workbook()
            ws = wb.active
            ws.title = 'Inventory'
            # כתיבת כותרות
            ws.append(headers)
            # המרה עדינה של כמויות למספרים אם אפשר
            def to_num(v):
                s = str(v).strip()
                try:
                    if s == '':
                        return ''
                    # עדיף int אם אפשר
                    f = float(s.replace(',', ''))
                    if abs(f - int(f)) < 1e-9:
                        return int(f)
                    return f
                except Exception:
                    return s
            for r in rows:
                out = [r[0] if len(r)>0 else '', r[1] if len(r)>1 else '', r[2] if len(r)>2 else '', r[3] if len(r)>3 else '', to_num(r[4] if len(r)>4 else ''), r[5] if len(r)>5 else '', r[6] if len(r)>6 else '']
                ws.append(out)
            wb.save(save_path)
            try:
                messagebox.showinfo('ייצוא מלאי', f'הייצוא הושלם בהצלחה:\n{save_path}')
            except Exception:
                pass
        except Exception as e:
            try:
                messagebox.showerror('ייצוא מלאי', f'שגיאה בשמירה לאקסל:\n{e}')
            except Exception:
                pass

    # === עזרי UI: קטגוריות/מיקומים/צורות אריזה למלאי ===
    def _get_main_category_names_for_inventory(self):
        """החזרת שמות קטגוריות ראשיות מהנתונים, עם 'בגדים' כברירת מחדל ראשונה."""
        names = []
        try:
            lst = getattr(self.data_processor, 'main_categories', []) or []
            for c in lst:
                nm = (c.get('name') or c.get('שם') or '').strip()
                if nm:
                    names.append(nm)
        except Exception:
            pass
        if 'בגדים' not in names:
            names = ['בגדים'] + [n for n in names if n != 'בגדים']
        else:
            # סדר: בגדים תחילה
            names = ['בגדים'] + [n for n in names if n != 'בגדים']
        # הסר כפולים תוך שמירה על סדר
        seen = set(); out = []
        for n in names:
            if n not in seen:
                seen.add(n); out.append(n)
        return out or ['בגדים']

    def _reload_inventory_aux_options(self):
        """טען מ'Settings' את רשימות המיקומים/צורות אריזה לתוך הקומבובוקסים."""
        pkgs = []
        locs = []
        try:
            if hasattr(self, 'settings') and hasattr(self.settings, 'get'):
                pkgs = self.settings.get('inventory.packaging_options', []) or []
                locs = self.settings.get('inventory.location_options', []) or []
        except Exception:
            pkgs = locs = []
        try:
            self.inv_create_packaging_cb['values'] = pkgs
            self.inv_create_location_cb['values'] = locs
        except Exception:
            pass

    def _load_pkg_list(self):
        try:
            self.pkg_list.delete(0, 'end')
        except Exception:
            return
        opts = []
        try:
            if hasattr(self, 'settings') and hasattr(self.settings, 'get'):
                opts = self.settings.get('inventory.packaging_options', []) or []
        except Exception:
            opts = []
        for v in opts:
            try: self.pkg_list.insert('end', v)
            except Exception: pass

    def _inv_pkg_add(self):
        v = (getattr(self, 'pkg_new_var', tk.StringVar()).get() or '').strip()
        if not v:
            return
        try:
            self.pkg_list.insert('end', v)
            self.pkg_new_var.set('')
        except Exception:
            pass

    def _inv_pkg_delete(self):
        try:
            sel = list(self.pkg_list.curselection())
            for i in reversed(sel):
                self.pkg_list.delete(i)
        except Exception:
            pass

    def _inv_pkg_save(self):
        vals = []
        try:
            vals = [self.pkg_list.get(i) for i in range(self.pkg_list.size())]
        except Exception:
            pass
        try:
            if hasattr(self, 'settings') and hasattr(self.settings, 'set'):
                self.settings.set('inventory.packaging_options', vals)
        except Exception:
            pass
        # ריענון קומבובוקס
        try:
            self.inv_create_packaging_cb['values'] = vals
        except Exception:
            pass

    def _load_loc_list(self):
        try:
            self.loc_list.delete(0, 'end')
        except Exception:
            return
        opts = []
        try:
            if hasattr(self, 'settings') and hasattr(self.settings, 'get'):
                opts = self.settings.get('inventory.location_options', []) or []
        except Exception:
            opts = []
        for v in opts:
            try: self.loc_list.insert('end', v)
            except Exception: pass

    def _inv_loc_add(self):
        v = (getattr(self, 'loc_new_var', tk.StringVar()).get() or '').strip()
        if not v:
            return
        try:
            self.loc_list.insert('end', v)
            self.loc_new_var.set('')
        except Exception:
            pass

    def _inv_loc_delete(self):
        try:
            sel = list(self.loc_list.curselection())
            for i in reversed(sel):
                self.loc_list.delete(i)
        except Exception:
            pass

    def _inv_loc_save(self):
        vals = []
        try:
            vals = [self.loc_list.get(i) for i in range(self.loc_list.size())]
        except Exception:
            pass
        try:
            if hasattr(self, 'settings') and hasattr(self.settings, 'set'):
                self.settings.set('inventory.location_options', vals)
        except Exception:
            pass
        # ריענון קומבובוקס
        try:
            self.inv_create_location_cb['values'] = vals
        except Exception:
            pass
