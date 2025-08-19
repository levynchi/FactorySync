import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import re

class ProductsCatalogMethodsMixin:
    """All event handlers, loaders, and helpers for Products Catalog tab."""
    # ===== products section builders =====
    def _build_products_section(self, parent):
        form = ttk.LabelFrame(parent, text="הוספת פריט", padding=10)
        form.pack(fill='x', padx=10, pady=6)
        self.prod_name_var = tk.StringVar(); self.prod_size_var = tk.StringVar(); self.prod_fabric_type_var = tk.StringVar(); self.prod_fabric_color_var = tk.StringVar(); self.prod_print_name_var = tk.StringVar()
        self.prod_category_var = tk.StringVar(); self.prod_ticks_var = tk.StringVar(); self.prod_elastic_var = tk.StringVar(); self.prod_ribbon_var = tk.StringVar(); self.prod_unit_type_var = tk.StringVar()
        self.prod_fabric_category_var = tk.StringVar(); self.prod_main_category_var = tk.StringVar()

        # --- Main Category selector (drives field visibility) ---
        self._products_field_widgets = {}
        tk.Label(form, text="קטגוריה ראשית:", font=('Arial',10,'bold')).grid(row=0, column=0, sticky='w', padx=4, pady=(0,6))
        main_cat_names = [c.get('name','') for c in getattr(self.data_processor, 'main_categories', [])]
        self.prod_main_category_combobox = ttk.Combobox(form, textvariable=self.prod_main_category_var, values=main_cat_names, state='readonly', width=16, justify='right')
        self.prod_main_category_combobox.grid(row=0, column=1, sticky='w', padx=2, pady=(0,6))
        self.prod_main_category_combobox.bind('<<ComboboxSelected>>', lambda e: self._apply_main_category_field_visibility())

        # shift existing rows by +1 to make space for main category row
        tk.Label(form, text="שם הדגם:", font=('Arial',10,'bold')).grid(row=1, column=0, sticky='w', padx=4, pady=4)
        model_names = [r.get('name') for r in getattr(self.data_processor, 'product_model_names', [])]
        self.model_name_combobox = ttk.Combobox(
            form,
            textvariable=self.prod_name_var,
            values=model_names,
            state='readonly',
            width=16,
            justify='right'
        )
        self.model_name_combobox.grid(row=1, column=1, sticky='w', padx=2, pady=4)
        # map widgets for field visibility: model_name
        self._products_field_widgets['model_name'] = []
        self._products_field_widgets['model_name'].append(self.model_name_combobox)

        # sub category (תת קטגוריה) - multi-select like sizes/types/colors/prints
        tk.Label(form, text="תת קטגוריה:", font=('Arial',10,'bold')).grid(row=1, column=2, sticky='w', padx=4, pady=4)
        cat_names = [c.get('name','') for c in getattr(self.data_processor, 'categories', [])]
        # Group sub-category picker + display + clear into one frame to keep them adjacent
        self.subcat_frame = ttk.Frame(form)
        self.subcat_frame.grid(row=1, column=3, columnspan=3, sticky='w', padx=2, pady=4)
        self.category_combobox = ttk.Combobox(
            self.subcat_frame,
            values=cat_names,
            state='readonly',
            width=12,
            justify='right'
        )
        self.category_combobox.pack(side='left', padx=(0, 4))
        self.category_combobox.bind('<<ComboboxSelected>>', lambda e: self._on_attr_select('sub_category'))
        # readonly display of selected sub-categories and clear button
        self.subcat_selected_entry = tk.Entry(self.subcat_frame, textvariable=self.prod_category_var, width=20, state='readonly')
        self.subcat_selected_entry.pack(side='left', padx=(0, 4))
        self.btn_clear_subcat = tk.Button(self.subcat_frame, text='נקה', command=lambda: self._clear_attr('sub_category'), width=4)
        self.btn_clear_subcat.pack(side='left')
        # sub_category mapping (label+combo will be toggled together)
        # We'll collect after creating label widgets too using winfo_children search if needed

        # fabric category
        tk.Label(form, text="קטגוריית בד:", font=('Arial',10,'bold')).grid(row=1, column=6, sticky='w', padx=4, pady=4)
        fabric_cat_names = [r.get('name') for r in getattr(self.data_processor, 'product_fabric_categories', [])]
        self.fabric_category_combobox = ttk.Combobox(
            form,
            textvariable=self.prod_fabric_category_var,
            values=fabric_cat_names,
            state='readonly',
            width=12,
            justify='right'
        )
        self.fabric_category_combobox.grid(row=1, column=7, sticky='w', padx=2, pady=4)

        # multiselect helpers state
        self.selected_sub_categories = []
        self.selected_sizes = []
        self.selected_fabric_types = []
        self.selected_fabric_colors = []
        self.selected_print_names = []

    # size (new row for more space)
        tk.Label(form, text="מידות:", font=('Arial',10,'bold')).grid(row=2, column=0, sticky='w', padx=4, pady=4)
        self.size_picker = ttk.Combobox(form, values=[r.get('name') for r in getattr(self.data_processor,'product_sizes',[])], state='readonly', width=12, justify='right')
        self.size_picker.grid(row=2, column=1, sticky='w', padx=2, pady=4)
        self.size_picker.bind('<<ComboboxSelected>>', lambda e: self._on_attr_select('size'))
        tk.Entry(form, textvariable=self.prod_size_var, width=20, state='readonly').grid(row=2, column=2, sticky='w', padx=2, pady=4)
        self.btn_clear_size = tk.Button(form, text='נקה', command=lambda: self._clear_attr('size'), width=4)
        self.btn_clear_size.grid(row=2, column=3, padx=2)

        # types (move to its own row for more space)
        tk.Label(form, text="סוגי בד:", font=('Arial',10,'bold')).grid(row=3, column=0, sticky='w', padx=4, pady=4)
        self.ftype_picker = ttk.Combobox(form, values=[r.get('name') for r in getattr(self.data_processor,'product_fabric_types',[])], state='readonly', width=12, justify='right')
        self.ftype_picker.grid(row=3, column=1, sticky='w', padx=2, pady=4)
        self.ftype_picker.bind('<<ComboboxSelected>>', lambda e: self._on_attr_select('fabric_type'))
        tk.Entry(form, textvariable=self.prod_fabric_type_var, width=20, state='readonly').grid(row=3, column=2, sticky='w', padx=2, pady=4)
        self.btn_clear_ftype = tk.Button(form, text='נקה', command=lambda: self._clear_attr('fabric_type'), width=4)
        self.btn_clear_ftype.grid(row=3, column=3, padx=2)

        # colors (keep on their own row)
        tk.Label(form, text="צבעי בד:", font=('Arial',10,'bold')).grid(row=4, column=0, sticky='w', padx=4, pady=4)
        self.fcolor_picker = ttk.Combobox(form, values=[r.get('name') for r in getattr(self.data_processor,'product_fabric_colors',[])], state='readonly', width=12, justify='right')
        self.fcolor_picker.grid(row=4, column=1, sticky='w', padx=2, pady=4)
        self.fcolor_picker.bind('<<ComboboxSelected>>', lambda e: self._on_attr_select('fabric_color'))
        tk.Entry(form, textvariable=self.prod_fabric_color_var, width=20, state='readonly').grid(row=4, column=2, sticky='w', padx=2, pady=4)
        self.btn_clear_fcolor = tk.Button(form, text='נקה', command=lambda: self._clear_attr('fabric_color'), width=4)
        self.btn_clear_fcolor.grid(row=4, column=3, padx=2)

        # prints (move to a dedicated row)
        tk.Label(form, text="שמות פרינט:", font=('Arial',10,'bold')).grid(row=5, column=0, sticky='w', padx=4, pady=4)
        self.pname_picker = ttk.Combobox(form, values=[r.get('name') for r in getattr(self.data_processor,'product_print_names',[])], state='readonly', width=12, justify='right')
        self.pname_picker.grid(row=5, column=1, sticky='w', padx=2, pady=4)
        self.pname_picker.bind('<<ComboboxSelected>>', lambda e: self._on_attr_select('print_name'))
        tk.Entry(form, textvariable=self.prod_print_name_var, width=20, state='readonly').grid(row=5, column=2, sticky='w', padx=2, pady=4)
        self.btn_clear_pname = tk.Button(form, text='נקה', command=lambda: self._clear_attr('print_name'), width=4)
        self.btn_clear_pname.grid(row=5, column=3, padx=2, pady=4)

        # accessories quantities (move down to separate row)
        tk.Label(form, text="טיקטקים:", font=('Arial',10,'bold')).grid(row=6, column=0, sticky='w', padx=4, pady=4)
        tk.Entry(form, textvariable=self.prod_ticks_var, width=10).grid(row=6, column=1, sticky='w', padx=2, pady=4)
        tk.Label(form, text="גומי:", font=('Arial',10,'bold')).grid(row=6, column=2, sticky='w', padx=4, pady=4)
        tk.Entry(form, textvariable=self.prod_elastic_var, width=10).grid(row=6, column=3, sticky='w', padx=2, pady=4)
        tk.Label(form, text="סרט:", font=('Arial',10,'bold')).grid(row=6, column=4, sticky='w', padx=4, pady=4)
        tk.Entry(form, textvariable=self.prod_ribbon_var, width=10).grid(row=6, column=5, sticky='w', padx=2, pady=4)
        # Unit Type (סוג יחידה)
        tk.Label(form, text="סוג יחידה:", font=('Arial',10,'bold')).grid(row=6, column=6, sticky='w', padx=4, pady=4)
        tk.Entry(form, textvariable=self.prod_unit_type_var, width=12).grid(row=6, column=7, sticky='w', padx=2, pady=4)

        # actions moved to their own row to avoid horizontal clipping
        tk.Button(form, text="➕ הוסף", command=self._add_product_catalog_entry, bg='#27ae60', fg='white').grid(row=7, column=0, padx=12, pady=6, sticky='w')
        tk.Button(form, text="🗑️ מחק נבחר", command=self._delete_selected_product_entry, bg='#e67e22', fg='white').grid(row=7, column=1, padx=4, pady=6, sticky='w')
        tk.Button(form, text="💾 ייצוא ל-Excel", command=self._export_products_catalog, bg='#2c3e50', fg='white').grid(row=7, column=2, padx=4, pady=6, sticky='w')
        tk.Button(form, text="⬆️ יבוא מקובץ", command=self._import_products_catalog_dialog, bg='#34495e', fg='white').grid(row=7, column=3, padx=4, pady=6, sticky='w')

        tree_frame = ttk.LabelFrame(parent, text="פריטים", padding=6)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=6)
        # Add main_category column for display
        cols = ('id','name','main_category','category','size','fabric_type','fabric_color','print_name','fabric_category','ticks_qty','elastic_qty','ribbon_qty','created_at')
        self.products_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=12)
        headers = {
            'id':'ID',
            'name':'שם הדגם',
            'main_category':'קטגוריה ראשית',
            'category':'תת קטגוריה',
            'size':'מידה',
            'fabric_type':'סוג בד',
            'fabric_color':'צבע בד',
            'print_name':'שם פרינט',
            'fabric_category':'קטגוריית בד',
            'ticks_qty':'טיקטקים',
            'elastic_qty':'גומי',
            'ribbon_qty':'סרט',
            'created_at':'נוצר'
        }
        widths = {
            'id':40,
            'name':140,
            'main_category':110,
            'category':100,
            'size':70,
            'fabric_type':110,
            'fabric_color':110,
            'print_name':110,
            'fabric_category':120,
            'ticks_qty':70,
            'elastic_qty':60,
            'ribbon_qty':60,
            'created_at':140
        }
        for c in cols:
            self.products_tree.heading(c, text=headers[c])
            self.products_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(tree_frame, orient='vertical', command=self.products_tree.yview)
        self.products_tree.configure(yscroll=vs.set)
        self.products_tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self._load_products_catalog_into_tree()

        # collect field widgets for toggling (labels + inputs/buttons)
        # Helper to find label widgets by text; if not found we still toggle inputs
        def _try_find_label(text):
            for w in form.grid_slaves():
                if isinstance(w, tk.Label) and str(w.cget('text')) == text:
                    return w
            return None

        # Helper to find a widget by grid row/column, optionally class and text
        def _get_grid_widget(row: int, col: int, cls=None, text: str | None = None):
            for w in form.grid_slaves():
                info = w.grid_info()
                try:
                    r = int(info.get('row'))
                    c = int(info.get('column'))
                except Exception:
                    continue
                if r == row and c == col:
                    if cls and not isinstance(w, cls):
                        continue
                    if text is not None and str(getattr(w, 'cget', lambda *_: '')('text')) != text:
                        continue
                    return w
            return None

        # Map labels explicitly
        lbl_model = _try_find_label("שם הדגם:")
        lbl_subcat = _try_find_label("תת קטגוריה:")
        lbl_fabric_cat = _try_find_label("קטגוריית בד:")
        lbl_sizes = _try_find_label("מידות:")
        lbl_ftype = _try_find_label("סוגי בד:")
        lbl_fcolor = _try_find_label("צבעי בד:")
        lbl_pname = _try_find_label("שמות פרינט:")
        lbl_ticks = _try_find_label("טיקטקים:")
        lbl_elastic = _try_find_label("גומי:")
        lbl_ribbon = _try_find_label("סרט:")
        lbl_unit_type = _try_find_label("סוג יחידה:")

        self._products_field_widgets['model_name'] = [x for x in [lbl_model, self.model_name_combobox] if x]
        # include readonly entry and clear for sub_category
        # Toggle the entire frame to show/hide sub-category widgets together
        subcat_widgets = [lbl_subcat]
        if hasattr(self, 'subcat_frame'):
            subcat_widgets.append(self.subcat_frame)
        self._products_field_widgets['sub_category'] = [x for x in subcat_widgets if x]
        self._products_field_widgets['fabric_category'] = [x for x in [lbl_fabric_cat, self.fabric_category_combobox] if x]

        # sizes group includes picker, readonly entry, clear button
        sizes_widgets = [lbl_sizes, self.size_picker]
        for w in form.grid_slaves():
            if isinstance(w, tk.Entry) and w.cget('state') == 'readonly' and w.cget('textvariable'):
                if str(w.cget('textvariable')) == str(self.prod_size_var):
                    sizes_widgets.append(w)
        btn_clear_size = getattr(self, 'btn_clear_size', None)
        if btn_clear_size:
            sizes_widgets.append(btn_clear_size)
        self._products_field_widgets['sizes'] = [w for w in sizes_widgets if w]

        # fabric type group
        ftype_widgets = [lbl_ftype, self.ftype_picker]
        for w in form.grid_slaves():
            if isinstance(w, tk.Entry) and w.cget('state') == 'readonly' and w.cget('textvariable'):
                if str(w.cget('textvariable')) == str(self.prod_fabric_type_var):
                    ftype_widgets.append(w)
        btn_clear_ftype = getattr(self, 'btn_clear_ftype', None)
        if btn_clear_ftype:
            ftype_widgets.append(btn_clear_ftype)
        self._products_field_widgets['fabric_type'] = [w for w in ftype_widgets if w]

        # fabric color group
        fcolor_widgets = [lbl_fcolor, self.fcolor_picker]
        for w in form.grid_slaves():
            if isinstance(w, tk.Entry) and w.cget('state') == 'readonly' and w.cget('textvariable'):
                if str(w.cget('textvariable')) == str(self.prod_fabric_color_var):
                    fcolor_widgets.append(w)
        btn_clear_fcolor = getattr(self, 'btn_clear_fcolor', None)
        if btn_clear_fcolor:
            fcolor_widgets.append(btn_clear_fcolor)
        self._products_field_widgets['fabric_color'] = [w for w in fcolor_widgets if w]

        # print name group
        pname_widgets = [lbl_pname, self.pname_picker]
        for w in form.grid_slaves():
            if isinstance(w, tk.Entry) and w.cget('state') == 'readonly' and w.cget('textvariable'):
                if str(w.cget('textvariable')) == str(self.prod_print_name_var):
                    pname_widgets.append(w)
        btn_clear_pname = getattr(self, 'btn_clear_pname', None)
        if btn_clear_pname:
            pname_widgets.append(btn_clear_pname)
        self._products_field_widgets['print_name'] = [w for w in pname_widgets if w]

        # quantities
        ticks_widgets = [x for x in [lbl_ticks] if x]
        elastic_widgets = [x for x in [lbl_elastic] if x]
        ribbon_widgets = [x for x in [lbl_ribbon] if x]
        for w in form.grid_slaves():
            if isinstance(w, tk.Entry) and w.cget('state') != 'readonly' and w.cget('textvariable'):
                tv = str(w.cget('textvariable'))
                if tv == str(self.prod_ticks_var):
                    ticks_widgets.append(w)
                elif tv == str(self.prod_elastic_var):
                    elastic_widgets.append(w)
                elif tv == str(self.prod_ribbon_var):
                    ribbon_widgets.append(w)
        self._products_field_widgets['ticks_qty'] = ticks_widgets
        self._products_field_widgets['elastic_qty'] = elastic_widgets
        self._products_field_widgets['ribbon_qty'] = ribbon_widgets
        # unit type group
        unit_widgets = [x for x in [lbl_unit_type] if x]
        for w in form.grid_slaves():
            if isinstance(w, tk.Entry) and w.cget('state') != 'readonly' and w.cget('textvariable'):
                if str(w.cget('textvariable')) == str(self.prod_unit_type_var):
                    unit_widgets.append(w)
        self._products_field_widgets['unit_type'] = unit_widgets

        # initial visibility
        self._refresh_main_categories_for_products()
        self._apply_main_category_field_visibility()

    # ===== accessories builders =====
    def _build_accessories_section(self, parent):
        self.acc_name_var = tk.StringVar(); self.acc_unit_var = tk.StringVar()
        acc_form = ttk.LabelFrame(parent, text="הוספת אביזר תפירה", padding=10)
        acc_form.pack(fill='x', padx=10, pady=6)
        tk.Label(acc_form, text="שם אביזר:", font=('Arial',10,'bold')).grid(row=0, column=0, padx=4, pady=4, sticky='w')
        tk.Entry(acc_form, textvariable=self.acc_name_var, width=20).grid(row=0, column=1, padx=4, pady=4)
        tk.Label(acc_form, text="יחידת מדידה:", font=('Arial',10,'bold')).grid(row=0, column=2, padx=4, pady=4, sticky='w')
        tk.Entry(acc_form, textvariable=self.acc_unit_var, width=12).grid(row=0, column=3, padx=4, pady=4)
        tk.Button(acc_form, text="➕ הוסף", command=self._add_sewing_accessory, bg='#27ae60', fg='white').grid(row=0, column=4, padx=8)
        tk.Button(acc_form, text="🗑️ מחק נבחר", command=self._delete_selected_accessory, bg='#e67e22', fg='white').grid(row=0, column=5, padx=4)

        acc_tree_frame = ttk.LabelFrame(parent, text="אביזרים", padding=6)
        acc_tree_frame.pack(fill='both', expand=True, padx=10, pady=6)
        acc_cols = ('id','name','unit','created_at')
        self.accessories_tree = ttk.Treeview(acc_tree_frame, columns=acc_cols, show='headings', height=10)
        acc_headers = {'id':'ID','name':'שם','unit':'יחידה','created_at':'נוצר'}
        acc_widths = {'id':50,'name':160,'unit':100,'created_at':140}
        for c in acc_cols:
            self.accessories_tree.heading(c, text=acc_headers[c])
            self.accessories_tree.column(c, width=acc_widths[c], anchor='center')
        acc_vs = ttk.Scrollbar(acc_tree_frame, orient='vertical', command=self.accessories_tree.yview)
        self.accessories_tree.configure(yscroll=acc_vs.set)
        self.accessories_tree.pack(side='left', fill='both', expand=True)
        acc_vs.pack(side='right', fill='y')
        self._load_accessories_into_tree()

    # ===== categories builders =====
    def _build_categories_section(self, parent):
        self.cat_name_var = tk.StringVar()
        cat_form = ttk.LabelFrame(parent, text="הוספת תת קטגוריה", padding=10)
        cat_form.pack(fill='x', padx=10, pady=6)
        tk.Label(cat_form, text="שם תת קטגוריה:", font=('Arial',10,'bold')).grid(row=0, column=0, padx=4, pady=4, sticky='w')
        tk.Entry(cat_form, textvariable=self.cat_name_var, width=22).grid(row=0, column=1, padx=4, pady=4)
        tk.Button(cat_form, text="➕ הוסף", command=self._add_category, bg='#27ae60', fg='white').grid(row=0, column=2, padx=8)
        tk.Button(cat_form, text="🗑️ מחק נבחר", command=self._delete_selected_category, bg='#e67e22', fg='white').grid(row=0, column=3, padx=4)

        cat_tree_frame = ttk.LabelFrame(parent, text="תת קטגוריות", padding=6)
        cat_tree_frame.pack(fill='both', expand=True, padx=10, pady=6)
        cat_cols = ('id','name','created_at')
        self.categories_tree = ttk.Treeview(cat_tree_frame, columns=cat_cols, show='headings', height=10)
        cat_headers = {'id':'ID','name':'שם תת קטגוריה','created_at':'נוצר'}
        cat_widths = {'id':60,'name':180,'created_at':140}
        for c in cat_cols:
            self.categories_tree.heading(c, text=cat_headers[c])
            self.categories_tree.column(c, width=cat_widths[c], anchor='center')
        cat_vs = ttk.Scrollbar(cat_tree_frame, orient='vertical', command=self.categories_tree.yview)
        self.categories_tree.configure(yscroll=cat_vs.set)
        self.categories_tree.pack(side='left', fill='both', expand=True)
        cat_vs.pack(side='right', fill='y')
        self._load_categories_into_tree()

    # ===== main categories builders =====
    def _build_main_categories_section(self, parent):
        # Notebook inside the Main Categories tab
        mcat_nb = ttk.Notebook(parent)
        mcat_nb.pack(fill='both', expand=True, padx=8, pady=6)

        manage_tab = tk.Frame(mcat_nb, bg='#f7f9fa')
        fields_tab = tk.Frame(mcat_nb, bg='#f7f9fa')
        mcat_nb.add(manage_tab, text='קטגוריות')
        mcat_nb.add(fields_tab, text='שדות לקטגוריה')

        # Manage tab: add/delete + list
        self.main_cat_name_var = tk.StringVar()
        mcat_form = ttk.LabelFrame(manage_tab, text="הוספת קטגוריה ראשית", padding=10)
        mcat_form.pack(fill='x', padx=10, pady=6)
        tk.Label(mcat_form, text="שם קטגוריה ראשית:", font=('Arial',10,'bold')).grid(row=0, column=0, padx=4, pady=4, sticky='w')
        tk.Entry(mcat_form, textvariable=self.main_cat_name_var, width=22).grid(row=0, column=1, padx=4, pady=4)
        tk.Button(mcat_form, text="➕ הוסף", command=self._add_main_category, bg='#27ae60', fg='white').grid(row=0, column=2, padx=8)
        tk.Button(mcat_form, text="🗑️ מחק נבחר", command=self._delete_selected_main_category, bg='#e67e22', fg='white').grid(row=0, column=3, padx=4)

        mcat_tree_frame = ttk.LabelFrame(manage_tab, text="קטגוריות ראשיות", padding=6)
        mcat_tree_frame.pack(fill='both', expand=True, padx=10, pady=6)
        mcat_cols = ('id','name','created_at')
        self.main_categories_tree = ttk.Treeview(mcat_tree_frame, columns=mcat_cols, show='headings', height=10)
        mcat_headers = {'id':'ID','name':'שם קטגוריה','created_at':'נוצר'}
        mcat_widths = {'id':60,'name':180,'created_at':140}
        for c in mcat_cols:
            self.main_categories_tree.heading(c, text=mcat_headers[c])
            self.main_categories_tree.column(c, width=mcat_widths[c], anchor='center')
        mcat_vs = ttk.Scrollbar(mcat_tree_frame, orient='vertical', command=self.main_categories_tree.yview)
        self.main_categories_tree.configure(yscroll=mcat_vs.set)
        self.main_categories_tree.pack(side='left', fill='both', expand=True)
        mcat_vs.pack(side='right', fill='y')
        self._load_main_categories_into_tree()

        # Fields tab: assign item fields to main category
        self._build_main_category_fields_tab(fields_tab)

    def _get_item_fields_defs(self) -> list[tuple[str, str]]:
        # key, display label mapping for item fields in 'פריטים > הוספת פריט'
        return [
            ('model_name', 'שם הדגם'),
            ('sub_category', 'תת קטגוריה'),
            ('fabric_category', 'קטגוריית בד'),
            ('unit_type', 'סוג יחידה'),
            ('sizes', 'מידות'),
            ('fabric_type', 'סוגי בד'),
            ('fabric_color', 'צבעי בד'),
            ('print_name', 'שמות פרינט'),
            ('ticks_qty', 'טיקטקים'),
            ('elastic_qty', 'גומי'),
            ('ribbon_qty', 'סרט'),
        ]

    def _build_main_category_fields_tab(self, parent):
        frm = ttk.LabelFrame(parent, text='שיוך שדות לקטגוריה ראשית', padding=10)
        frm.pack(fill='both', expand=True, padx=10, pady=8)

        # Category selector
        tk.Label(frm, text='קטגוריה ראשית:', font=('Arial',10,'bold')).grid(row=0, column=0, sticky='w', padx=4, pady=4)
        self.main_cat_fields_var = tk.StringVar()
        self.main_cat_fields_combo = ttk.Combobox(frm, textvariable=self.main_cat_fields_var, state='readonly', width=24, justify='right')
        self.main_cat_fields_combo.grid(row=0, column=1, sticky='w', padx=4, pady=4)
        self.main_cat_fields_combo.bind('<<ComboboxSelected>>', lambda e: self._on_select_main_category_for_fields())

        # Checkboxes
        fields_defs = self._get_item_fields_defs()
        self.main_cat_field_vars = {}
        chk_container = ttk.Frame(frm)
        chk_container.grid(row=1, column=0, columnspan=3, sticky='w', padx=2, pady=4)
        col = 0; row = 0
        for key, label in fields_defs:
            var = tk.BooleanVar(value=False)
            self.main_cat_field_vars[key] = var
            cb = ttk.Checkbutton(chk_container, text=label, variable=var)
            cb.grid(row=row, column=col, padx=8, pady=4, sticky='w')
            col += 1
            if col >= 3:
                col = 0; row += 1

        # Action buttons
        tk.Button(frm, text='💾 שמור שדות', command=self._save_main_category_fields, bg='#2c3e50', fg='white').grid(row=2, column=0, padx=4, pady=8, sticky='w')
        tk.Button(frm, text='אפס בחירה', command=self._reset_main_category_fields).grid(row=2, column=1, padx=4, pady=8, sticky='w')

        # Load categories into combo
        self._load_main_categories_for_fields()

    def _load_main_categories_for_fields(self):
        names = [c.get('name','') for c in getattr(self.data_processor, 'main_categories', [])]
        if hasattr(self, 'main_cat_fields_combo'):
            self.main_cat_fields_combo['values'] = names
            # keep selection if still valid
            cur = self.main_cat_fields_var.get()
            if cur not in names:
                self.main_cat_fields_var.set('')

    def _find_main_category_by_name(self, name: str):
        for c in getattr(self.data_processor, 'main_categories', []):
            if c.get('name','') == name:
                return c
        return None

    def _on_select_main_category_for_fields(self):
        name = self.main_cat_fields_var.get().strip()
        rec = self._find_main_category_by_name(name)
        if not rec:
            self._reset_main_category_fields()
            return
        fields = rec.get('fields') or self.data_processor.get_main_category_fields(rec.get('id'))
        fields = fields or []
        for key, var in self.main_cat_field_vars.items():
            var.set(key in fields)

    def _reset_main_category_fields(self):
        for var in self.main_cat_field_vars.values():
            var.set(False)

    def _save_main_category_fields(self):
        name = self.main_cat_fields_var.get().strip()
        rec = self._find_main_category_by_name(name)
        if not rec:
            messagebox.showerror('שגיאה', 'בחר קטגוריה ראשית לשיוך שדות')
            return
        selected_keys = [k for k,v in self.main_cat_field_vars.items() if v.get()]
        ok = self.data_processor.set_main_category_fields(int(rec.get('id')), selected_keys)
        if ok:
            messagebox.showinfo('הצלחה', 'השדות נשמרו לקטגוריה')
        else:
            messagebox.showerror('שגיאה', 'שמירת השדות נכשלה')

    # ===== attributes builders =====
    def _build_attributes_section(self, attributes_tab):
        attr_nb = ttk.Notebook(attributes_tab)
        attr_nb.pack(fill='both', expand=True, padx=8, pady=6)

        sizes_tab = tk.Frame(attr_nb, bg='#f7f9fa')
        ftypes_tab = tk.Frame(attr_nb, bg='#f7f9fa')
        fcolors_tab = tk.Frame(attr_nb, bg='#f7f9fa')
        prints_tab = tk.Frame(attr_nb, bg='#f7f9fa')
        fcats_tab = tk.Frame(attr_nb, bg='#f7f9fa')
        modelnames_tab = tk.Frame(attr_nb, bg='#f7f9fa')
        attr_nb.add(sizes_tab, text='מידות')
        attr_nb.add(ftypes_tab, text='סוגי בד')
        attr_nb.add(fcolors_tab, text='צבעי בד')
        attr_nb.add(prints_tab, text='שמות פרינט')
        attr_nb.add(fcats_tab, text='קטגוריות בדים')
        attr_nb.add(modelnames_tab, text='שם הדגם')

        # bind vars
        self.attr_size_var = tk.StringVar(); self.attr_fabric_type_var = tk.StringVar(); self.attr_fabric_color_var = tk.StringVar(); self.attr_print_name_var = tk.StringVar(); self.attr_fabric_category_var = tk.StringVar(); self.attr_model_name_var = tk.StringVar()

        # Sizes
        sz_form = ttk.LabelFrame(sizes_tab, text='הוספת מידה', padding=8)
        sz_form.pack(fill='x', padx=8, pady=6)
        tk.Label(sz_form, text='מידה:', font=('Arial',10,'bold')).grid(row=0, column=0, padx=4, pady=4)
        tk.Entry(sz_form, textvariable=self.attr_size_var, width=18).grid(row=0, column=1, padx=4, pady=4)
        tk.Button(sz_form, text='➕ הוסף', command=self._add_product_size, bg='#27ae60', fg='white').grid(row=0, column=2, padx=6)
        tk.Button(sz_form, text='🗑️ מחק נבחר', command=self._delete_selected_product_size, bg='#e67e22', fg='white').grid(row=0, column=3, padx=4)
        sz_tree_frame = ttk.LabelFrame(sizes_tab, text='מידות', padding=4)
        sz_tree_frame.pack(fill='both', expand=True, padx=8, pady=4)
        self.sizes_tree = ttk.Treeview(sz_tree_frame, columns=('id','name','created_at'), show='headings', height=10)
        for c,t,w in [('id','ID',60),('name','מידה',140),('created_at','נוצר',140)]:
            self.sizes_tree.heading(c, text=t); self.sizes_tree.column(c, width=w, anchor='center')
        sz_vs = ttk.Scrollbar(sz_tree_frame, orient='vertical', command=self.sizes_tree.yview)
        self.sizes_tree.configure(yscroll=sz_vs.set)
        self.sizes_tree.pack(side='left', fill='both', expand=True); sz_vs.pack(side='right', fill='y')
        self._load_sizes_into_tree()

        # Fabric Types
        ft_form = ttk.LabelFrame(ftypes_tab, text='הוספת סוג בד', padding=8)
        ft_form.pack(fill='x', padx=8, pady=6)
        tk.Label(ft_form, text='סוג בד:', font=('Arial',10,'bold')).grid(row=0, column=0, padx=4, pady=4)
        tk.Entry(ft_form, textvariable=self.attr_fabric_type_var, width=18).grid(row=0, column=1, padx=4, pady=4)
        tk.Button(ft_form, text='➕ הוסף', command=self._add_fabric_type_item, bg='#27ae60', fg='white').grid(row=0, column=2, padx=6)
        tk.Button(ft_form, text='🗑️ מחק נבחר', command=self._delete_selected_fabric_type_item, bg='#e67e22', fg='white').grid(row=0, column=3, padx=4)
        ft_tree_frame = ttk.LabelFrame(ftypes_tab, text='סוגי בד', padding=4)
        ft_tree_frame.pack(fill='both', expand=True, padx=8, pady=4)
        self.fabric_types_tree = ttk.Treeview(ft_tree_frame, columns=('id','name','created_at'), show='headings', height=10)
        for c,t,w in [('id','ID',60),('name','סוג בד',160),('created_at','נוצר',140)]:
            self.fabric_types_tree.heading(c, text=t); self.fabric_types_tree.column(c, width=w, anchor='center')
        ft_vs = ttk.Scrollbar(ft_tree_frame, orient='vertical', command=self.fabric_types_tree.yview)
        self.fabric_types_tree.configure(yscroll=ft_vs.set)
        self.fabric_types_tree.pack(side='left', fill='both', expand=True); ft_vs.pack(side='right', fill='y')
        self._load_fabric_types_into_tree()

        # Fabric Colors
        fc_form = ttk.LabelFrame(fcolors_tab, text='הוספת צבע בד', padding=8)
        fc_form.pack(fill='x', padx=8, pady=6)
        tk.Label(fc_form, text='צבע בד:', font=('Arial',10,'bold')).grid(row=0, column=0, padx=4, pady=4)
        tk.Entry(fc_form, textvariable=self.attr_fabric_color_var, width=18).grid(row=0, column=1, padx=4, pady=4)
        tk.Button(fc_form, text='➕ הוסף', command=self._add_fabric_color_item, bg='#27ae60', fg='white').grid(row=0, column=2, padx=6)
        tk.Button(fc_form, text='🗑️ מחק נבחר', command=self._delete_selected_fabric_color_item, bg='#e67e22', fg='white').grid(row=0, column=3, padx=4)
        fc_tree_frame = ttk.LabelFrame(fcolors_tab, text='צבעי בד', padding=4)
        fc_tree_frame.pack(fill='both', expand=True, padx=8, pady=4)
        self.fabric_colors_tree = ttk.Treeview(fc_tree_frame, columns=('id','name','created_at'), show='headings', height=10)
        for c,t,w in [('id','ID',60),('name','צבע בד',160),('created_at','נוצר',140)]:
            self.fabric_colors_tree.heading(c, text=t); self.fabric_colors_tree.column(c, width=w, anchor='center')
        fc_vs = ttk.Scrollbar(fc_tree_frame, orient='vertical', command=self.fabric_colors_tree.yview)
        self.fabric_colors_tree.configure(yscroll=fc_vs.set)
        self.fabric_colors_tree.pack(side='left', fill='both', expand=True); fc_vs.pack(side='right', fill='y')
        self._load_fabric_colors_into_tree()

        # Print Names
        pn_form = ttk.LabelFrame(prints_tab, text='הוספת שם פרינט', padding=8)
        pn_form.pack(fill='x', padx=8, pady=6)
        tk.Label(pn_form, text='שם פרינט:', font=('Arial',10,'bold')).grid(row=0, column=0, padx=4, pady=4)
        tk.Entry(pn_form, textvariable=self.attr_print_name_var, width=18).grid(row=0, column=1, padx=4, pady=4)
        tk.Button(pn_form, text='➕ הוסף', command=self._add_print_name_item, bg='#27ae60', fg='white').grid(row=0, column=2, padx=6)
        tk.Button(pn_form, text='🗑️ מחק נבחר', command=self._delete_selected_print_name_item, bg='#e67e22', fg='white').grid(row=0, column=3, padx=4)
        pn_tree_frame = ttk.LabelFrame(prints_tab, text='שמות פרינט', padding=4)
        pn_tree_frame.pack(fill='both', expand=True, padx=8, pady=4)
        self.print_names_tree = ttk.Treeview(pn_tree_frame, columns=('id','name','created_at'), show='headings', height=10)
        for c,t,w in [('id','ID',60),('name','שם פרינט',160),('created_at','נוצר',140)]:
            self.print_names_tree.heading(c, text=t); self.print_names_tree.column(c, width=w, anchor='center')
        pn_vs = ttk.Scrollbar(pn_tree_frame, orient='vertical', command=self.print_names_tree.yview)
        self.print_names_tree.configure(yscroll=pn_vs.set)
        self.print_names_tree.pack(side='left', fill='both', expand=True); pn_vs.pack(side='right', fill='y')
        self._load_print_names_into_tree()

        # Fabric Categories
        fcg_form = ttk.LabelFrame(fcats_tab, text='הוספת קטגוריית בד', padding=8)
        fcg_form.pack(fill='x', padx=8, pady=6)
        tk.Label(fcg_form, text='קטגוריית בד:', font=('Arial',10,'bold')).grid(row=0, column=0, padx=4, pady=4)
        tk.Entry(fcg_form, textvariable=self.attr_fabric_category_var, width=18).grid(row=0, column=1, padx=4, pady=4)
        tk.Button(fcg_form, text='➕ הוסף', command=self._add_fabric_category_item, bg='#27ae60', fg='white').grid(row=0, column=2, padx=6)
        tk.Button(fcg_form, text='🗑️ מחק נבחר', command=self._delete_selected_fabric_category_item, bg='#e67e22', fg='white').grid(row=0, column=3, padx=4)
        fcg_tree_frame = ttk.LabelFrame(fcats_tab, text='קטגוריות בדים', padding=4)
        fcg_tree_frame.pack(fill='both', expand=True, padx=8, pady=4)
        self.fabric_categories_tree = ttk.Treeview(fcg_tree_frame, columns=('id','name','created_at'), show='headings', height=10)
        for c,t,w in [('id','ID',60),('name','קטגוריה',160),('created_at','נוצר',140)]:
            self.fabric_categories_tree.heading(c, text=t); self.fabric_categories_tree.column(c, width=w, anchor='center')
        fcg_vs = ttk.Scrollbar(fcg_tree_frame, orient='vertical', command=self.fabric_categories_tree.yview)
        self.fabric_categories_tree.configure(yscroll=fcg_vs.set)
        self.fabric_categories_tree.pack(side='left', fill='both', expand=True); fcg_vs.pack(side='right', fill='y')
        self._load_fabric_categories_into_tree()

        # Model Names (שם הדגם)
        mn_form = ttk.LabelFrame(modelnames_tab, text='הוסף שם דגם', padding=8)
        mn_form.pack(fill='x', padx=8, pady=6)
        tk.Label(mn_form, text='שם דגם:', font=('Arial',10,'bold')).grid(row=0, column=0, padx=4, pady=4)
        tk.Entry(mn_form, textvariable=self.attr_model_name_var, width=18).grid(row=0, column=1, padx=4, pady=4)
        tk.Button(mn_form, text='➕ הוסף', command=self._add_model_name_item, bg='#27ae60', fg='white').grid(row=0, column=2, padx=6)
        tk.Button(mn_form, text='🗑️ מחק נבחר', command=self._delete_selected_model_name_item, bg='#e67e22', fg='white').grid(row=0, column=3, padx=4)
        mn_tree_frame = ttk.LabelFrame(modelnames_tab, text='שמות דגם', padding=4)
        mn_tree_frame.pack(fill='both', expand=True, padx=8, pady=4)
        self.model_names_tree = ttk.Treeview(mn_tree_frame, columns=('id','name','created_at'), show='headings', height=10)
        for c,t,w in [('id','ID',60),('name','שם דגם',160),('created_at','נוצר',140)]:
            self.model_names_tree.heading(c, text=t); self.model_names_tree.column(c, width=w, anchor='center')
        mn_vs = ttk.Scrollbar(mn_tree_frame, orient='vertical', command=self.model_names_tree.yview)
        self.model_names_tree.configure(yscroll=mn_vs.set)
        self.model_names_tree.pack(side='left', fill='both', expand=True); mn_vs.pack(side='right', fill='y')
        self._load_model_names_into_tree()

    # ===== LOADERS =====
    def _load_products_catalog_into_tree(self):
        if not hasattr(self, 'products_tree'): return
        for item in self.products_tree.get_children(): self.products_tree.delete(item)
        try:
            for rec in getattr(self.data_processor, 'products_catalog', []):
                fabric_category_value = rec.get('fabric_category') or 'בלי קטגוריה'
                main_category_value = rec.get('main_category') or 'בגדים'
                self.products_tree.insert('', 'end', values=(
                    rec.get('id'), rec.get('name'), main_category_value, rec.get('category',''), rec.get('size'), rec.get('fabric_type'),
                    rec.get('fabric_color'), rec.get('print_name'), fabric_category_value, rec.get('ticks_qty'), rec.get('elastic_qty'),
                    rec.get('ribbon_qty'), rec.get('created_at')
                ))
        except Exception:
            pass

    def _load_accessories_into_tree(self):
        if not hasattr(self, 'accessories_tree'): return
        for item in self.accessories_tree.get_children():
            self.accessories_tree.delete(item)
        try:
            for rec in getattr(self.data_processor, 'sewing_accessories', []):
                self.accessories_tree.insert('', 'end', values=(
                    rec.get('id'), rec.get('name'), rec.get('unit'), rec.get('created_at')
                ))
        except Exception:
            pass

    def _add_sewing_accessory(self):
        name = self.acc_name_var.get().strip()
        unit = self.acc_unit_var.get().strip()
        if not name:
            messagebox.showerror("שגיאה", "חובה להזין שם אביזר")
            return
        try:
            new_id = self.data_processor.add_sewing_accessory(name, unit)
            self.accessories_tree.insert('', 'end', values=(new_id, name, unit, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.acc_name_var.set(''); self.acc_unit_var.set('')
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _delete_selected_accessory(self):
        if not hasattr(self, 'accessories_tree'): return
        sel = self.accessories_tree.selection()
        if not sel: return
        deleted = False
        for item in sel:
            vals = self.accessories_tree.item(item, 'values')
            if vals:
                if self.data_processor.delete_sewing_accessory(int(vals[0])):
                    deleted = True
        if deleted:
            self._load_accessories_into_tree()

    def _load_categories_into_tree(self):
        if not hasattr(self, 'categories_tree'): return
        for item in self.categories_tree.get_children():
            self.categories_tree.delete(item)
        try:
            for rec in getattr(self.data_processor, 'categories', []):
                self.categories_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('created_at')))
        except Exception:
            pass

    def _load_main_categories_into_tree(self):
        if not hasattr(self, 'main_categories_tree'): return
        for item in self.main_categories_tree.get_children():
            self.main_categories_tree.delete(item)
        try:
            for rec in getattr(self.data_processor, 'main_categories', []):
                self.main_categories_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('created_at')))
        except Exception:
            pass

    def _refresh_main_categories_for_products(self):
        names = [c.get('name','') for c in getattr(self.data_processor, 'main_categories', [])]
        if hasattr(self, 'prod_main_category_combobox'):
            self.prod_main_category_combobox['values'] = names
            cur = self.prod_main_category_var.get()
            if cur not in names:
                self.prod_main_category_var.set('')

    def _apply_main_category_field_visibility(self):
        # When no main category is selected: hide all product fields.
        # When selected: show only fields configured for that category.
        name = self.prod_main_category_var.get().strip() if hasattr(self, 'prod_main_category_var') else ''
        fields = []
        if name:
            rec = self._find_main_category_by_name(name) if hasattr(self, '_find_main_category_by_name') else None
            if rec:
                fields = rec.get('fields') or []
                if not fields and hasattr(self, 'data_processor'):
                    try:
                        fields = self.data_processor.get_main_category_fields(rec.get('id'))
                    except Exception:
                        fields = []
        # If no name chosen -> hide everything
        for key, widgets in getattr(self, '_products_field_widgets', {}).items():
            visible = (bool(name) and (not fields or key in fields))
            for w in widgets:
                try:
                    if visible:
                        w.grid()
                    else:
                        w.grid_remove()
                except Exception:
                    pass

    def _add_category(self):
        name = self.cat_name_var.get().strip()
        if not name:
            messagebox.showerror("שגיאה", "חובה להזין שם קטגוריה")
            return
        try:
            new_id = self.data_processor.add_category(name)
            self.categories_tree.insert('', 'end', values=(new_id, name, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.cat_name_var.set('')
            self._refresh_categories_for_products()
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _delete_selected_category(self):
        if not hasattr(self, 'categories_tree'): return
        sel = self.categories_tree.selection()
        if not sel: return
        deleted = False
        for item in sel:
            vals = self.categories_tree.item(item, 'values')
            if vals:
                if self.data_processor.delete_category(int(vals[0])):
                    deleted = True
        if deleted:
            self._load_categories_into_tree()
            self._refresh_categories_for_products()

    def _add_main_category(self):
        name = self.main_cat_name_var.get().strip()
        if not name:
            messagebox.showerror("שגיאה", "חובה להזין שם קטגוריה ראשית")
            return
        try:
            new_id = self.data_processor.add_main_category(name)
            self.main_categories_tree.insert('', 'end', values=(new_id, name, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.main_cat_name_var.set('')
            self._load_main_categories_for_fields()
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _delete_selected_main_category(self):
        if not hasattr(self, 'main_categories_tree'): return
        sel = self.main_categories_tree.selection()
        if not sel: return
        deleted = False
        for item in sel:
            vals = self.main_categories_tree.item(item, 'values')
            if vals:
                if self.data_processor.delete_main_category(int(vals[0])):
                    deleted = True
        if deleted:
            self._load_main_categories_into_tree()
            self._load_main_categories_for_fields()

    def _load_sizes_into_tree(self):
        if not hasattr(self, 'sizes_tree'): return
        for item in self.sizes_tree.get_children(): self.sizes_tree.delete(item)
        for rec in getattr(self.data_processor, 'product_sizes', []):
            self.sizes_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('created_at')))

    def _load_fabric_types_into_tree(self):
        if not hasattr(self, 'fabric_types_tree'): return
        for item in self.fabric_types_tree.get_children(): self.fabric_types_tree.delete(item)
        for rec in getattr(self.data_processor, 'product_fabric_types', []):
            self.fabric_types_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('created_at')))

    def _load_fabric_colors_into_tree(self):
        if not hasattr(self, 'fabric_colors_tree'): return
        for item in self.fabric_colors_tree.get_children(): self.fabric_colors_tree.delete(item)
        for rec in getattr(self.data_processor, 'product_fabric_colors', []):
            self.fabric_colors_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('created_at')))

    def _load_print_names_into_tree(self):
        if not hasattr(self, 'print_names_tree'): return
        for item in self.print_names_tree.get_children(): self.print_names_tree.delete(item)
        for rec in getattr(self.data_processor, 'product_print_names', []):
            self.print_names_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('created_at')))

    def _load_fabric_categories_into_tree(self):
        if not hasattr(self, 'fabric_categories_tree'): return
        for item in self.fabric_categories_tree.get_children(): self.fabric_categories_tree.delete(item)
        for rec in getattr(self.data_processor, 'product_fabric_categories', []):
            self.fabric_categories_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('created_at')))

    def _load_model_names_into_tree(self):
        if not hasattr(self, 'model_names_tree'): return
        for item in self.model_names_tree.get_children(): self.model_names_tree.delete(item)
        for rec in getattr(self.data_processor, 'product_model_names', []):
            self.model_names_tree.insert('', 'end', values=(rec.get('id'), rec.get('name'), rec.get('created_at')))

    # ===== ADD =====
    def _add_product_catalog_entry(self):
        name = self.prod_name_var.get().strip()
        if not name:
            messagebox.showerror("שגיאה", "חובה להזין שם מוצר")
            return
        category_raw = self.prod_category_var.get().strip()
        valid_categories = [c.get('name','') for c in getattr(self.data_processor, 'categories', [])]
        # Subcategory is optional: if blank, use an empty token; if provided, validate values
        if category_raw:
            # category_raw may be comma-separated multi-select; validate all
            category_tokens = [s.strip() for s in category_raw.split(',') if s.strip()]
            if category_tokens:
                for ct in category_tokens:
                    if ct not in valid_categories:
                        messagebox.showerror("שגיאה", f"תת קטגוריה '{ct}' לא קיימת. הוסף בטאב 'תת קטגוריות' ובחר שוב")
                        return
            else:
                # treat whitespace-only as empty category
                category_tokens = ['']
        else:
            category_tokens = ['']
        sizes_raw = self.prod_size_var.get().strip()
        ftypes_raw = self.prod_fabric_type_var.get().strip()
        fcolors_raw = self.prod_fabric_color_var.get().strip()
        prints_raw = self.prod_print_name_var.get().strip()
        fabric_category_raw = self.prod_fabric_category_var.get().strip()
        ticks_raw = self.prod_ticks_var.get().strip()
        elastic_raw = self.prod_elastic_var.get().strip()
        ribbon_raw = self.prod_ribbon_var.get().strip()

        def _split(raw):
            if not raw:
                return ['']
            return [s.strip() for s in re.split(r'[;,.\s]+', raw) if s.strip()]

        size_tokens = _split(sizes_raw)
        ft_tokens = _split(ftypes_raw)
        fc_tokens = _split(fcolors_raw)
        pn_tokens = _split(prints_raw)

        def _normalize_size(tok: str) -> str:
            t = tok.strip()
            if not t:
                return t
            if t.startswith('-'):
                t = t[1:]
            if '-' not in t and t.isdigit() and 3 <= len(t) <= 4:
                mid = len(t)//2
                a, b = t[:mid], t[mid:]
                try:
                    ai = int(a); bi = int(b)
                    if 0 <= ai <= 60 and 0 <= bi <= 60 and ai < bi:
                        return f"{ai}-{bi}"
                except Exception:
                    pass
            return t

        size_tokens = [_normalize_size(s) for s in size_tokens]

        from itertools import product
        combos = list(product(category_tokens, size_tokens, ft_tokens, fc_tokens, pn_tokens)) or [( '', '', '', '', '' )]

        existing = set()
        try:
            for rec in getattr(self.data_processor, 'products_catalog', []):
                existing.add((rec.get('name','').strip(), rec.get('category','').strip(), rec.get('size','').strip(), rec.get('fabric_type','').strip(), rec.get('fabric_color','').strip(), rec.get('print_name','').strip()))
        except Exception:
            existing = set()

        if len(combos) == 1:
            only_cat, only_sz, only_ft, only_fc, only_pn = combos[0]
            single_key = (name, only_cat, only_sz, only_ft, only_fc, only_pn)
            if single_key in existing:
                messagebox.showinfo(
                    "כפילות",
            "המוצר עם הנתונים הללו כבר קיים במערכת:\n"
            f"שם: {name}\nתת קטגוריה: {only_cat or '-'}\nמידה: {only_sz or '-'}\nסוג בד: {only_ft or '-'}\nצבע בד: {only_fc or '-'}\nשם פרינט: {only_pn or '-'}"
                )
                return

        added = 0
        try:
            for cat, sz, ft, fc, pn in combos:
                key = (name, cat, sz, ft, fc, pn)
                if key in existing:
                    continue
                new_id = self.data_processor.add_product_catalog_entry(
                    name, sz, ft, fc, pn, cat, ticks_raw, elastic_raw, ribbon_raw, fabric_category_raw
                )
                existing.add(key)
                added += 1
                fabric_category_value = fabric_category_raw or 'בלי קטגוריה'
                main_category_value = self.prod_main_category_var.get().strip() or 'בגדים'
                self.products_tree.insert('', 'end', values=(
                    new_id, name, main_category_value, cat, sz, ft, fc, pn, fabric_category_value,
                    ticks_raw or 0, elastic_raw or 0, ribbon_raw or 0,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ))
            self.prod_name_var.set(''); self.prod_category_var.set(''); self.prod_fabric_category_var.set(''); self.prod_size_var.set(''); self.prod_fabric_type_var.set(''); self.prod_fabric_color_var.set(''); self.prod_print_name_var.set(''); self.prod_ticks_var.set(''); self.prod_elastic_var.set(''); self.prod_ribbon_var.set('')
            self.selected_sizes.clear(); self.selected_fabric_types.clear(); self.selected_fabric_colors.clear(); self.selected_print_names.clear()
            if hasattr(self, 'size_picker'): self.size_picker.set('')
            if hasattr(self, 'ftype_picker'): self.ftype_picker.set('')
            if hasattr(self, 'fcolor_picker'): self.fcolor_picker.set('')
            if hasattr(self, 'pname_picker'): self.pname_picker.set('')
            if hasattr(self, 'fabric_category_combobox'): self.fabric_category_combobox.set('')
            if hasattr(self, 'model_name_combobox'): self.model_name_combobox.set('')
            if added > 1:
                messagebox.showinfo("הצלחה", f"נוספו {added} וריאנטים למוצר '{name}'")
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

    def _import_products_catalog_dialog(self):
        top = tk.Toplevel(self.notebook)
        top.title("יבוא קטלוג מוצרים מ-Excel")
        top.grab_set(); top.resizable(False, False)
        frm = ttk.Frame(top, padding=10)
        frm.pack(fill='both', expand=True)

        path_var = tk.StringVar(); overwrite_var = tk.BooleanVar(value=False)

        ttk.Label(frm, text="בחר קובץ Excel בפורמט הייצוא:").grid(row=0, column=0, columnspan=3, sticky='w', pady=(0,6))
        ttk.Entry(frm, textvariable=path_var, width=50).grid(row=1, column=0, columnspan=2, sticky='we', padx=(0,6))
        def _browse():
            p = filedialog.askopenfilename(title="בחר קובץ קטלוג", filetypes=[('Excel','*.xlsx')])
            if p: path_var.set(p)
        ttk.Button(frm, text="עיון...", command=_browse).grid(row=1, column=2, sticky='w')

        ttk.Checkbutton(frm, text="דרוס את הטבלה (במקום להוסיף)", variable=overwrite_var).grid(row=2, column=0, columnspan=3, sticky='w', pady=8)

        btns = ttk.Frame(frm); btns.grid(row=3, column=0, columnspan=3, sticky='e')
        def _do_import():
            path = path_var.get().strip()
            if not path:
                messagebox.showerror("שגיאה", "בחר קובץ לייבוא")
                return
            mode = 'overwrite' if overwrite_var.get() else 'append'
            try:
                res = self.data_processor.import_products_catalog_from_excel(path, mode=mode)
                self._load_products_catalog_into_tree()
                msg = f"יובאו {res.get('imported',0)} רשומות."
                skipped = res.get('skipped_duplicates',0)
                if skipped: msg += f"\nדולגו {skipped} כפילויות."
                if res.get('overwritten'): msg += "\nבוצעה דריסה מלאה של הטבלה."
                messagebox.showinfo("הצלחה", msg); top.destroy()
            except Exception as e:
                messagebox.showerror("שגיאה", str(e))
        ttk.Button(btns, text="ייבוא", command=_do_import).pack(side='right', padx=4)
        ttk.Button(btns, text="ביטול", command=top.destroy).pack(side='right', padx=4)

    # ===== ATTR helpers =====
    def _on_attr_select(self, kind: str):
        picker_map = {
            'sub_category': (self.category_combobox, self.selected_sub_categories, self.prod_category_var),
            'size': (self.size_picker, self.selected_sizes, self.prod_size_var),
            'fabric_type': (self.ftype_picker, self.selected_fabric_types, self.prod_fabric_type_var),
            'fabric_color': (self.fcolor_picker, self.selected_fabric_colors, self.prod_fabric_color_var),
            'print_name': (self.pname_picker, self.selected_print_names, self.prod_print_name_var)
        }
        picker, lst, var = picker_map.get(kind)
        val = picker.get().strip()
        if val and val not in lst:
            lst.append(val)
            var.set(','.join(lst))
        picker.set('')

    def _clear_attr(self, kind: str):
        if kind == 'sub_category':
            self.selected_sub_categories.clear(); self.prod_category_var.set('')
        elif kind == 'size':
            self.selected_sizes.clear(); self.prod_size_var.set('')
        elif kind == 'fabric_type':
            self.selected_fabric_types.clear(); self.prod_fabric_type_var.set('')
        elif kind == 'fabric_color':
            self.selected_fabric_colors.clear(); self.prod_fabric_color_var.set('')
        elif kind == 'print_name':
            self.selected_print_names.clear(); self.prod_print_name_var.set('')

    def _refresh_attribute_pickers(self):
        if hasattr(self, 'size_picker'):
            self.size_picker['values'] = [r.get('name') for r in getattr(self.data_processor,'product_sizes',[])]
        if hasattr(self, 'ftype_picker'):
            self.ftype_picker['values'] = [r.get('name') for r in getattr(self.data_processor,'product_fabric_types',[])]
        if hasattr(self, 'fcolor_picker'):
            self.fcolor_picker['values'] = [r.get('name') for r in getattr(self.data_processor,'product_fabric_colors',[])]
        if hasattr(self, 'pname_picker'):
            self.pname_picker['values'] = [r.get('name') for r in getattr(self.data_processor,'product_print_names',[])]
        if hasattr(self, 'fabric_category_combobox'):
            names = [r.get('name') for r in getattr(self.data_processor, 'product_fabric_categories', [])]
            self.fabric_category_combobox['values'] = names
            if self.prod_fabric_category_var.get() not in names:
                self.prod_fabric_category_var.set('')
        if hasattr(self, 'model_name_combobox'):
            model_names = [r.get('name') for r in getattr(self.data_processor, 'product_model_names', [])]
            self.model_name_combobox['values'] = model_names
            if self.prod_name_var.get() not in model_names:
                self.prod_name_var.set('')

    def _refresh_categories_for_products(self):
        if hasattr(self, 'category_combobox'):
            names = [c.get('name','') for c in getattr(self.data_processor, 'categories', [])]
            self.category_combobox['values'] = names
            # Do not clear self.prod_category_var because it can contain a comma-separated multi-select
            # Just ensure the picker (combobox) has no stale selection
            if self.category_combobox.get() not in names:
                self.category_combobox.set('')
        self._refresh_attribute_pickers()

    # ===== ATTR add/delete =====
    def _add_product_size(self):
        name = self.attr_size_var.get().strip()
        if not name:
            messagebox.showerror('שגיאה', 'חובה להזין מידה')
            return
        try:
            new_id = self.data_processor.add_product_size(name)
            self.sizes_tree.insert('', 'end', values=(new_id, name, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.attr_size_var.set(''); self._refresh_attribute_pickers()
        except Exception as e:
            messagebox.showerror('שגיאה', str(e))

    def _add_fabric_type_item(self):
        name = self.attr_fabric_type_var.get().strip()
        if not name:
            messagebox.showerror('שגיאה', 'חובה להזין סוג בד')
            return
        try:
            new_id = self.data_processor.add_fabric_type_item(name)
            self.fabric_types_tree.insert('', 'end', values=(new_id, name, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.attr_fabric_type_var.set(''); self._refresh_attribute_pickers()
        except Exception as e:
            messagebox.showerror('שגיאה', str(e))

    def _add_fabric_color_item(self):
        name = self.attr_fabric_color_var.get().strip()
        if not name:
            messagebox.showerror('שגיאה', 'חובה להזין צבע בד')
            return
        try:
            new_id = self.data_processor.add_fabric_color_item(name)
            self.fabric_colors_tree.insert('', 'end', values=(new_id, name, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.attr_fabric_color_var.set(''); self._refresh_attribute_pickers()
        except Exception as e:
            messagebox.showerror('שגיאה', str(e))

    def _add_print_name_item(self):
        name = self.attr_print_name_var.get().strip()
        if not name:
            messagebox.showerror('שגיאה', 'חובה להזין שם פרינט')
            return
        try:
            new_id = self.data_processor.add_print_name_item(name)
            self.print_names_tree.insert('', 'end', values=(new_id, name, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.attr_print_name_var.set(''); self._refresh_attribute_pickers()
        except Exception as e:
            messagebox.showerror('שגיאה', str(e))

    def _add_fabric_category_item(self):
        name = self.attr_fabric_category_var.get().strip()
        if not name:
            messagebox.showerror('שגיאה', 'חובה להזין קטגוריית בד')
            return
        try:
            new_id = self.data_processor.add_fabric_category_item(name)
            self.fabric_categories_tree.insert('', 'end', values=(new_id, name, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.attr_fabric_category_var.set(''); self._refresh_attribute_pickers()
        except Exception as e:
            messagebox.showerror('שגיאה', str(e))

    def _add_model_name_item(self):
        name = self.attr_model_name_var.get().strip()
        if not name:
            messagebox.showerror('שגיאה', 'חובה להזין שם דגם')
            return
        try:
            new_id = self.data_processor.add_model_name_item(name)
            self.model_names_tree.insert('', 'end', values=(new_id, name, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.attr_model_name_var.set('')
        except Exception as e:
            messagebox.showerror('שגיאה', str(e))

    def _delete_selected_product_size(self):
        if not hasattr(self, 'sizes_tree'): return
        sel = self.sizes_tree.selection(); deleted = False
        for item in sel:
            vals = self.sizes_tree.item(item, 'values')
            if vals and self.data_processor.delete_product_size(int(vals[0])):
                deleted = True
        if deleted:
            self._load_sizes_into_tree(); self._refresh_attribute_pickers()

    def _delete_selected_fabric_type_item(self):
        if not hasattr(self, 'fabric_types_tree'): return
        sel = self.fabric_types_tree.selection(); deleted = False
        for item in sel:
            vals = self.fabric_types_tree.item(item, 'values')
            if vals and self.data_processor.delete_fabric_type_item(int(vals[0])):
                deleted = True
        if deleted:
            self._load_fabric_types_into_tree(); self._refresh_attribute_pickers()

    def _delete_selected_fabric_color_item(self):
        if not hasattr(self, 'fabric_colors_tree'): return
        sel = self.fabric_colors_tree.selection(); deleted = False
        for item in sel:
            vals = self.fabric_colors_tree.item(item, 'values')
            if vals and self.data_processor.delete_fabric_color_item(int(vals[0])):
                deleted = True
        if deleted:
            self._load_fabric_colors_into_tree(); self._refresh_attribute_pickers()

    def _delete_selected_print_name_item(self):
        if not hasattr(self, 'print_names_tree'): return
        sel = self.print_names_tree.selection(); deleted = False
        for item in sel:
            vals = self.print_names_tree.item(item, 'values')
            if vals and self.data_processor.delete_print_name_item(int(vals[0])):
                deleted = True
        if deleted:
            self._load_print_names_into_tree(); self._refresh_attribute_pickers()

    def _delete_selected_fabric_category_item(self):
        if not hasattr(self, 'fabric_categories_tree'): return
        sel = self.fabric_categories_tree.selection(); deleted = False
        for item in sel:
            vals = self.fabric_categories_tree.item(item, 'values')
            if vals and self.data_processor.delete_fabric_category_item(int(vals[0])):
                deleted = True
        if deleted:
            self._load_fabric_categories_into_tree(); self._refresh_attribute_pickers()

    def _delete_selected_model_name_item(self):
        if not hasattr(self, 'model_names_tree'): return
        sel = self.model_names_tree.selection(); deleted = False
        for item in sel:
            vals = self.model_names_tree.item(item, 'values')
            if vals and self.data_processor.delete_model_name_item(int(vals[0])):
                deleted = True
        if deleted:
            self._load_model_names_into_tree()
