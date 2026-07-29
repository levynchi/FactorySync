import tkinter as tk
from tkinter import ttk
from .methods import SupplierIntakeMethodsMixin
from .. import theme

class SupplierIntakeTabMixin(SupplierIntakeMethodsMixin):
    """Compose the Supplier Intake tab (entry + saved list)."""
    def _create_supplier_intake_tab(self):
        tab = tk.Frame(self.notebook, bg=theme.PAGE_BG)
        self.notebook.add(tab, text="תעודת קליטה")
        tk.Label(tab, text="תעודת קליטה (הזנה ידנית)", font=(theme.FONT_FAMILY,16,'bold'), bg=theme.PAGE_BG, fg=theme.DARK).pack(pady=8)

        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=4, pady=4)
        entry_wrapper = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        list_wrapper = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        fabrics_wrapper = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        fabrics_list_wrapper = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        inner_nb.add(entry_wrapper, text="קליטה")
        inner_nb.add(list_wrapper, text="קליטות שמורות")
        inner_nb.add(fabrics_wrapper, text="קליטת בדים")
        inner_nb.add(fabrics_list_wrapper, text="קליטות בדים שמורות")

        self._build_supplier_entry_tab(entry_wrapper)
        self._build_supplier_list_tab(list_wrapper)
        self._build_fabrics_intake_tab(fabrics_wrapper)
        self._build_fabrics_list_tab(fabrics_list_wrapper)

    def _build_supplier_entry_tab(self, container: tk.Frame):
        from .entry_tab import build_entry_tab
        build_entry_tab(self, container)

    def _build_supplier_list_tab(self, container: tk.Frame):
        from .list_tab import build_list_tab
        build_list_tab(self, container)

    def _build_fabrics_intake_tab(self, container: tk.Frame):
        """UI עבור 'קליטת בדים': סריקת ברקוד, תצוגת נתונים ושמירה לשינוי סטטוס במלאי הבדים."""
        header = tk.Frame(container, bg=theme.PAGE_BG)
        header.pack(fill='x', padx=10, pady=(8,4))
        tk.Label(header, text='קליטת בדים', font=(theme.FONT_FAMILY,12,'bold'), bg=theme.PAGE_BG).pack(side='right')

        # מצב קליטה: לפי ברקוד או ללא ברקוד
        mode_frame = tk.Frame(container, bg=theme.PAGE_BG)
        mode_frame.pack(fill='x', padx=10)
        self.fi_mode_var = tk.StringVar(value='barcode')
        ttk.Radiobutton(mode_frame, text='לפי ברקוד', variable=self.fi_mode_var, value='barcode').pack(side='right', padx=(0,10))
        ttk.Radiobutton(mode_frame, text='ללא ברקוד', variable=self.fi_mode_var, value='no_barcode').pack(side='right')

        # קלט עליון עבור מצב "ללא ברקוד" ליד מתג המצב
        self.fi_nb_type = tk.StringVar()
        self.fi_nb_manu = tk.StringVar()
        self.fi_nb_color = tk.StringVar()
        self.fi_nb_shade = tk.StringVar()
        self.fi_nb_notes = tk.StringVar()
        nb_top = tk.Frame(mode_frame, bg=theme.PAGE_BG)
        def _mk_top_input(lbl_txt, var):
            box = tk.Frame(nb_top, bg=theme.PAGE_BG)
            tk.Label(box, text=lbl_txt+':', bg=theme.PAGE_BG).pack(side='right', padx=(6,2))
            tk.Entry(box, textvariable=var, width=16).pack(side='right')
            return box
        _mk_top_input('סוג בד', self.fi_nb_type).pack(side='right', padx=8)
        _mk_top_input('יצרן הבד', self.fi_nb_manu).pack(side='right', padx=8)
        _mk_top_input('צבע', self.fi_nb_color).pack(side='right', padx=8)
        _mk_top_input('גוון', self.fi_nb_shade).pack(side='right', padx=8)
        _mk_top_input('הערות', self.fi_nb_notes).pack(side='right', padx=8)
        tk.Button(nb_top, text='➕ הוסף שורה', command=self._fi_nb_add_item, bg=theme.SUCCESS, fg='white').pack(side='left')

        # מסגרת תוכן דינמית מתחת לבורר המצבים ומעל ההובלה
        content_frame = tk.Frame(container, bg=theme.PAGE_BG)
        content_frame.pack(fill='both', expand=True)

        # ברקוד + פעולות
        bar = tk.Frame(content_frame, bg=theme.PAGE_BG)
        bar.pack(fill='x', padx=10, pady=(0,6))
        tk.Label(bar, text='בר קוד:', bg=theme.PAGE_BG).pack(side='right', padx=(8,4))
        self.fi_barcode_var = tk.StringVar()
        entry = tk.Entry(bar, textvariable=self.fi_barcode_var, width=24)
        entry.pack(side='right')
        try:
            entry.bind('<Return>', lambda e: self._fi_add_fabric_by_barcode())
        except Exception:
            pass
        tk.Button(bar, text='➕ הוסף', command=self._fi_add_fabric_by_barcode, bg=theme.SUCCESS, fg='white').pack(side='right', padx=6)
        tk.Button(bar, text='🗑️ הסר נבחר', command=self._fi_remove_selected).pack(side='left', padx=6)
        tk.Button(bar, text='🧹 נקה הכל', command=self._fi_clear_all).pack(side='left')

        # טבלת פריטי בד שנבחרו לקליטה (ברקוד)
        table_wrap = tk.Frame(content_frame, bg=theme.CARD_BG, relief='groove', bd=1)
        table_wrap.pack(fill='both', expand=True, padx=10, pady=6)
        cols = ('barcode','fabric_type','color_name','color_no','design_code','width','net_kg','meters','price','location','status')
        headers = {
            'barcode':'ברקוד','fabric_type':'סוג בד','color_name':'צבע','color_no':'מס׳ צבע',
            'design_code':'Desen Kodu','width':'רוחב','net_kg':'ק"ג נטו','meters':'מטרים','price':'מחיר','location':'מיקום','status':'סטטוס'
        }
        widths = {'barcode':140,'fabric_type':140,'color_name':110,'color_no':80,'design_code':110,'width':60,'net_kg':80,'meters':80,'price':80,'location':90,'status':110}
        self.fi_tree = ttk.Treeview(table_wrap, columns=cols, show='headings', height=12)
        for c in cols:
            self.fi_tree.heading(c, text=headers[c])
            self.fi_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(table_wrap, orient='vertical', command=self.fi_tree.yview)
        self.fi_tree.configure(yscroll=vs.set)
        self.fi_tree.grid(row=0, column=0, sticky='nsew')
        vs.grid(row=0, column=1, sticky='ns')
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)
        self.fi_bar_table_wrap = table_wrap

        # טבלה לפריטים ללא ברקוד + כפתורי פעולה
        nb_table_wrap = tk.Frame(content_frame, bg=theme.CARD_BG, relief='groove', bd=1)
        nb_cols = ('fabric_type','manufacturer','color','shade','notes')
        nb_headers = {'fabric_type':'סוג בד','manufacturer':'יצרן הבד','color':'צבע','shade':'גוון','notes':'הערות'}
        nb_widths = {'fabric_type':160,'manufacturer':140,'color':100,'shade':80,'notes':220}
        self.fi_nb_tree = ttk.Treeview(nb_table_wrap, columns=nb_cols, show='headings', height=10)
        for c in nb_cols:
            self.fi_nb_tree.heading(c, text=nb_headers[c])
            self.fi_nb_tree.column(c, width=nb_widths[c], anchor='center')
        nb_vs = ttk.Scrollbar(nb_table_wrap, orient='vertical', command=self.fi_nb_tree.yview)
        self.fi_nb_tree.configure(yscroll=nb_vs.set)
        self.fi_nb_tree.grid(row=0, column=0, sticky='nsew')
        nb_vs.grid(row=0, column=1, sticky='ns')
        nb_table_wrap.grid_rowconfigure(0, weight=1)
        nb_table_wrap.grid_columnconfigure(0, weight=1)
        self.fi_nb_table_wrap = nb_table_wrap

        nb_actions = tk.Frame(content_frame, bg=theme.PAGE_BG)
        tk.Button(nb_actions, text='🗑️ הסר נבחר', command=self._fi_nb_remove_selected).pack(side='right', padx=6)
        tk.Button(nb_actions, text='🧹 נקה הכל', command=self._fi_nb_clear_all).pack(side='right')
        self.fi_nb_actions = nb_actions

        # פונקציית הצגה/הסתרה לפי מצב
        def _toggle_mode(*_):
            m = self.fi_mode_var.get()
            # הסתר הכל
            for w in (bar, nb_top, self.fi_bar_table_wrap, self.fi_nb_table_wrap, self.fi_nb_actions):
                try:
                    w.pack_forget()
                except Exception:
                    try:
                        w.grid_forget()
                    except Exception:
                        pass
            # הצג לפי מצב
            if m == 'barcode':
                bar.pack(fill='x', padx=10, pady=(0,6))
                self.fi_bar_table_wrap.pack(fill='both', expand=True, padx=10, pady=6)
            else:
                nb_top.pack(side='right', padx=10)
                self.fi_nb_table_wrap.pack(fill='both', expand=True, padx=10, pady=6)
                self.fi_nb_actions.pack(fill='x', padx=10, pady=(0,6))
            try:
                self._fi_update_summary()
            except Exception:
                pass
        try:
            self.fi_mode_var.trace_add('write', lambda *_: _toggle_mode())
        except Exception:
            pass
        _toggle_mode()

        # מקטע הובלה (כמו בקלטת מוצרים) - תמיד בתחתית לפני כפתור שמירה
        pkg_frame = ttk.LabelFrame(container, text="הובלה", padding=8)
        pkg_frame.pack(fill='x', padx=10, pady=(0,6))
        if not hasattr(self, '_fi_packages'):
            self._fi_packages = []
        self.fi_pkg_type_var = tk.StringVar(value='בד')
        self.fi_pkg_qty_var = tk.StringVar()
        self.fi_pkg_driver_var = tk.StringVar()
        tk.Label(pkg_frame, text="פריט הובלה:").grid(row=0,column=0,sticky='w',padx=4,pady=2)
        self.fi_pkg_type_combo = ttk.Combobox(pkg_frame, textvariable=self.fi_pkg_type_var, state='readonly', width=14, values=['בד','שק','שקית קטנה'])
        self.fi_pkg_type_combo.grid(row=0,column=1,sticky='w',padx=4,pady=2)
        tk.Label(pkg_frame, text="כמות:").grid(row=0,column=2,sticky='w',padx=4,pady=2)
        tk.Entry(pkg_frame, textvariable=self.fi_pkg_qty_var, width=8).grid(row=0,column=3,sticky='w',padx=4,pady=2)
        tk.Label(pkg_frame, text="שם המוביל:").grid(row=0,column=4,sticky='w',padx=4,pady=2)
        self.fi_pkg_driver_combo = ttk.Combobox(pkg_frame, textvariable=self.fi_pkg_driver_var, width=16, state='readonly')
        self.fi_pkg_driver_combo.grid(row=0,column=5,sticky='w',padx=4,pady=2)
        try:
            self._refresh_driver_names_for_intake()
        except Exception:
            pass
        tk.Button(pkg_frame, text="➕ הוסף", command=self._fi_add_package_line, bg=theme.SUCCESS, fg='white').grid(row=0,column=6,padx=8)
        tk.Button(pkg_frame, text="🗑️ מחק נבחר", command=self._fi_delete_selected_package, bg=theme.WARNING, fg='white').grid(row=0,column=7,padx=4)
        tk.Button(pkg_frame, text="❌ נקה", command=self._fi_clear_packages, bg=theme.DANGER, fg='white').grid(row=0,column=8,padx=4)
        self.fi_packages_tree = ttk.Treeview(pkg_frame, columns=('type','quantity','driver'), show='headings', height=4)
        self.fi_packages_tree.heading('type', text='פריט הובלה')
        self.fi_packages_tree.heading('quantity', text='כמות')
        self.fi_packages_tree.heading('driver', text='שם המוביל')
        self.fi_packages_tree.column('type', width=120, anchor='center')
        self.fi_packages_tree.column('quantity', width=70, anchor='center')
        self.fi_packages_tree.column('driver', width=110, anchor='center')
        self.fi_packages_tree.grid(row=1,column=0,columnspan=9, sticky='ew', padx=2, pady=(6,2))

        # פעולת שמירה + סיכום
        actions = tk.Frame(container, bg=theme.PAGE_BG)
        actions.pack(fill='x', padx=10, pady=(0,8))
        tk.Button(actions, text='💾 שמור קליטת בדים', command=self._fi_save_receipt, bg=theme.DARK, fg='white', font=(theme.FONT_FAMILY,11,'bold')).pack(side='right')
        self.fi_summary_var = tk.StringVar(value='0 בדים')
        tk.Label(container, textvariable=self.fi_summary_var, bg=theme.DARK_2, fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')

    def _build_fabrics_list_tab(self, container: tk.Frame):
        header = tk.Frame(container, bg=theme.PAGE_BG)
        header.pack(fill='x', padx=10, pady=(8,4))
        tk.Label(header, text='קליטות בדים שמורות', font=(theme.FONT_FAMILY,12,'bold'), bg=theme.PAGE_BG).pack(side='right')

        cols = ('id','date','supplier','count','packages','delete')
        # עוטפים את הטבלה במסגרת פנימית כדי להשתמש ב-grid בפנים, בעוד שה-container משתמש ב-pack
        table_wrap = tk.Frame(container, bg=theme.CARD_BG)
        table_wrap.pack(fill='both', expand=True, padx=10, pady=6)

        self.fabrics_intakes_tree = ttk.Treeview(table_wrap, columns=cols, show='headings')
        for col, txt, w in (
            ('id','ID',60),
            ('date','תאריך',110),
            ('supplier','ספק',180),
            ('count','מס׳ ברקודים',100),
            ('packages','הובלה',160),
            ('delete','מחיקה',70),
        ):
            self.fabrics_intakes_tree.heading(col, text=txt)
            self.fabrics_intakes_tree.column(col, width=w, anchor='center')

        vs = ttk.Scrollbar(table_wrap, orient='vertical', command=self.fabrics_intakes_tree.yview)
        self.fabrics_intakes_tree.configure(yscroll=vs.set)
        self.fabrics_intakes_tree.grid(row=0, column=0, sticky='nsew')
        vs.grid(row=0, column=1, sticky='ns')
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)

        try:
            self.fabrics_intakes_tree.bind('<Button-1>', lambda e: self._on_click_fabrics_intakes(e))
        except Exception:
            pass

        btns = tk.Frame(container, bg=theme.PAGE_BG)
        btns.pack(fill='x', padx=10, pady=(0,6))
        tk.Button(btns, text='🔄 רענן', command=self._refresh_fabrics_intakes_list, bg=theme.PRIMARY, fg='white').pack(side='right')
        self._refresh_fabrics_intakes_list()
