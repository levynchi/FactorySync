import tkinter as tk
from tkinter import ttk
from datetime import datetime

from .methods import BabyBasicSalesMethodsMixin
from .. import theme


class BabyBasicSalesTabMixin(BabyBasicSalesMethodsMixin):
    """טאב בייבי בייסיק: תעודות סחורה, מחירון, וחשבון מול תשלומים."""

    def _create_baby_basic_sales_tab(self):
        tab = tk.Frame(self.notebook, bg=theme.PAGE_BG)
        self.notebook.add(tab, text="בייבי בייסיק")
        tk.Label(
            tab,
            text="בייבי בייסיק — תעודות סחורה וחשבון",
            font=(theme.FONT_FAMILY, 16, 'bold'),
            bg=theme.PAGE_BG,
            fg=theme.DARK,
        ).pack(pady=4)

        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=6, pady=4)
        self._bb_inner_nb = inner_nb

        note_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        saved_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        price_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        account_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        room_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)

        inner_nb.add(note_tab, text="תעודות סחורה")
        inner_nb.add(saved_tab, text="תעודות שמורות")
        inner_nb.add(price_tab, text="מחירון")
        inner_nb.add(account_tab, text="חשבון ותשלומים")
        inner_nb.add(room_tab, text="ספירת מלאי חדר")

        self._bb_note_lines = []
        self._bb_note_qty_by_barcode = {}
        self._bb_room_lines = []

        self._build_bb_note_tab(note_tab)
        self._build_bb_saved_notes_tab(saved_tab)
        self._build_bb_price_tab(price_tab)
        self._build_bb_account_tab(account_tab)
        self._build_bb_room_count_tab(room_tab)

        inner_nb.bind('<<NotebookTabChanged>>', self._on_bb_tab_change)

    def _on_bb_tab_change(self, event=None):
        try:
            idx = self._bb_inner_nb.index(self._bb_inner_nb.select())
        except Exception:
            return
        if idx == 0:
            self._bb_refresh_note_products()
            self._bb_refresh_customer_combos()
        elif idx == 1:
            self._bb_refresh_saved_notes()
        elif idx == 2:
            self._bb_refresh_price_table()
        elif idx == 3:
            self._bb_refresh_account()
        elif idx == 4:
            self._bb_refresh_room_combos()
            self._bb_refresh_room_lines()
            self._bb_refresh_saved_room_counts()

    def _build_bb_note_tab(self, parent):
        header = ttk.LabelFrame(parent, text="פרטי תעודה", padding=10)
        header.pack(fill='x', padx=10, pady=6)

        tk.Label(header, text="שותף:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).grid(row=0, column=0, sticky='e', padx=4, pady=4)
        self.bb_customer_var = tk.StringVar(value='בייבי בייסיק')
        self.bb_customer_combo = ttk.Combobox(header, textvariable=self.bb_customer_var, width=28)
        self.bb_customer_combo.grid(row=0, column=1, sticky='w', padx=4, pady=4)

        tk.Label(header, text="תאריך:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).grid(row=0, column=2, sticky='e', padx=4, pady=4)
        self.bb_note_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(header, textvariable=self.bb_note_date_var, width=14, justify='center').grid(row=0, column=3, sticky='w', padx=4, pady=4)

        tk.Label(header, text="הערה:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).grid(row=0, column=4, sticky='e', padx=4, pady=4)
        self.bb_note_comment_var = tk.StringVar()
        tk.Entry(header, textvariable=self.bb_note_comment_var, width=36).grid(row=0, column=5, sticky='we', padx=4, pady=4)
        header.grid_columnconfigure(5, weight=1)

        products = ttk.LabelFrame(parent, text="מוצרי בייבי בייסיק (כמו במדבקות)", padding=10)
        products.pack(fill='both', expand=True, padx=10, pady=4)

        filt = tk.Frame(products, bg=theme.PAGE_BG)
        filt.pack(fill='x', pady=(0, 6))
        tk.Label(filt, text="חיפוש:", bg=theme.PAGE_BG).pack(side='right', padx=4)
        self.bb_note_search_var = tk.StringVar()
        search = ttk.Entry(filt, textvariable=self.bb_note_search_var, width=32)
        search.pack(side='right', padx=4)
        search.bind('<KeyRelease>', lambda e: self._bb_refresh_note_products())
        theme.make_button(filt, "הוסף לתעודה", kind="success", command=self._bb_add_selected_to_note).pack(side='left', padx=4)
        tk.Label(
            filt,
            text="דאבל-קליק על «כמות» (יחידות), ואז הוסף לתעודה. המחיר נמשך מהמחירון.",
            bg=theme.PAGE_BG,
            fg=theme.SUBTEXT,
            font=theme.FONT_SMALL,
        ).pack(side='left', padx=10)

        pcols = ('name', 'size', 'fabric', 'barcode', 'price', 'qty')
        pheaders = {'name': 'מוצר', 'size': 'מידה', 'fabric': 'בד', 'barcode': 'ברקוד', 'price': 'מחיר מארז', 'qty': 'כמות'}
        prod_frame = tk.Frame(products, bg=theme.PAGE_BG)
        prod_frame.pack(fill='both', expand=True)
        self.bb_note_products_tree = theme.make_treeview(prod_frame, columns=pcols, show='headings', height=8)
        for c in pcols:
            self.bb_note_products_tree.heading(c, text=pheaders[c])
            w = 70 if c in ('size', 'qty') else (110 if c in ('fabric', 'price') else 220 if c == 'name' else 140)
            self.bb_note_products_tree.column(c, width=w, anchor='center')
        pvs = ttk.Scrollbar(prod_frame, orient='vertical', command=self.bb_note_products_tree.yview)
        self.bb_note_products_tree.configure(yscrollcommand=pvs.set)
        self.bb_note_products_tree.pack(side='left', fill='both', expand=True)
        pvs.pack(side='right', fill='y')
        self.bb_note_products_tree.bind('<Double-1>', self._bb_edit_note_qty_cell)

        lines = ttk.LabelFrame(parent, text="שורות התעודה", padding=10)
        lines.pack(fill='both', expand=True, padx=10, pady=4)
        lcols = ('name', 'size', 'fabric', 'barcode', 'qty', 'price', 'total')
        lheaders = {'name': 'מוצר', 'size': 'מידה', 'fabric': 'בד', 'barcode': 'ברקוד', 'qty': 'יחידות', 'price': 'מחיר', 'total': 'סה״כ'}
        lines_frame = tk.Frame(lines, bg=theme.PAGE_BG)
        lines_frame.pack(fill='both', expand=True)
        self.bb_note_lines_tree = theme.make_treeview(lines_frame, columns=lcols, show='headings', height=7)
        for c in lcols:
            self.bb_note_lines_tree.heading(c, text=lheaders[c])
            w = 70 if c in ('size', 'qty') else (110 if c in ('fabric', 'price', 'total') else 220 if c == 'name' else 140)
            self.bb_note_lines_tree.column(c, width=w, anchor='center')
        lvs = ttk.Scrollbar(lines_frame, orient='vertical', command=self.bb_note_lines_tree.yview)
        self.bb_note_lines_tree.configure(yscrollcommand=lvs.set)
        self.bb_note_lines_tree.pack(side='left', fill='both', expand=True)
        lvs.pack(side='right', fill='y')
        self.bb_note_lines_tree.bind('<Double-1>', self._bb_edit_line_price_cell)

        footer = tk.Frame(parent, bg=theme.PAGE_BG)
        footer.pack(fill='x', padx=10, pady=(2, 10))
        self.bb_note_total_var = tk.StringVar(value='סה״כ יחידות: 0    סה״כ לתשלום: 0.00 ₪')
        tk.Label(footer, textvariable=self.bb_note_total_var, font=theme.FONT_SUBTITLE, bg=theme.PAGE_BG, fg=theme.DARK).pack(side='right')
        theme.make_button(footer, "שמור תעודה", kind="success", command=self._bb_save_note).pack(side='left', padx=4)
        theme.make_button(footer, "הסר שורה", kind="danger", command=self._bb_remove_selected_note_line).pack(side='left', padx=4)
        theme.make_button(footer, "נקה", kind="secondary", command=self._bb_clear_current_note).pack(side='left', padx=4)

        self._bb_refresh_note_products()
        self._bb_refresh_customer_combos()

    def _build_bb_saved_notes_tab(self, parent):
        cols = ('id', 'date', 'customer', 'qty', 'amount', 'note')
        headers = {'id': 'מס׳', 'date': 'תאריך', 'customer': 'שותף', 'qty': 'יחידות', 'amount': 'סכום', 'note': 'הערה'}
        wrap = tk.Frame(parent, bg=theme.PAGE_BG)
        wrap.pack(fill='both', expand=True, padx=10, pady=8)
        self.bb_saved_notes_tree = theme.make_treeview(wrap, columns=cols, show='headings')
        for c in cols:
            self.bb_saved_notes_tree.heading(c, text=headers[c])
            w = 70 if c in ('id', 'qty') else (110 if c in ('date', 'amount') else 180 if c == 'customer' else 260)
            self.bb_saved_notes_tree.column(c, width=w, anchor='center')
        vs = ttk.Scrollbar(wrap, orient='vertical', command=self.bb_saved_notes_tree.yview)
        self.bb_saved_notes_tree.configure(yscrollcommand=vs.set)
        self.bb_saved_notes_tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self.bb_saved_notes_tree.bind('<Double-1>', self._bb_open_selected_note)

        actions = tk.Frame(parent, bg=theme.PAGE_BG)
        actions.pack(fill='x', padx=10, pady=(0, 10))
        theme.make_button(actions, "צפה / ייצוא", kind="primary", command=self._bb_open_selected_note).pack(side='right', padx=4)
        theme.make_button(actions, "PDF להדפסה", kind="success", command=self._bb_print_selected_note_pdf).pack(side='right', padx=4)
        theme.make_button(actions, "מחק", kind="danger", command=self._bb_delete_selected_note).pack(side='right', padx=4)
        theme.make_button(actions, "רענן", kind="secondary", command=self._bb_refresh_saved_notes).pack(side='right', padx=4)
        self._bb_refresh_saved_notes()

    def _build_bb_price_tab(self, parent):
        tk.Label(
            parent,
            text="מחירון הסיטונאות למוצרים שנמכרים לשותף בייבי בייסיק. דאבל-קליק על מחיר כדי לערוך.",
            bg=theme.PAGE_BG,
            fg=theme.SUBTEXT,
        ).pack(pady=(8, 4))

        filt = tk.Frame(parent, bg=theme.PAGE_BG)
        filt.pack(fill='x', padx=10, pady=4)
        tk.Label(filt, text="חיפוש:", bg=theme.PAGE_BG).pack(side='right', padx=4)
        self.bb_price_search_var = tk.StringVar()
        search = ttk.Entry(filt, textvariable=self.bb_price_search_var, width=32)
        search.pack(side='right', padx=4)
        search.bind('<KeyRelease>', lambda e: self._bb_refresh_price_table())
        theme.make_button(filt, "החל מחיר לכל המידות של הדגם", kind="primary", command=self._bb_apply_price_to_family).pack(side='left', padx=4)

        cols = ('name', 'size', 'fabric', 'color', 'pack', 'barcode', 'price')
        headers = {'name': 'מוצר', 'size': 'מידה', 'fabric': 'בד', 'color': 'צבע', 'pack': 'מארז', 'barcode': 'ברקוד', 'price': 'מחיר מארז ₪'}
        wrap = tk.Frame(parent, bg=theme.PAGE_BG)
        wrap.pack(fill='both', expand=True, padx=10, pady=8)
        self.bb_price_tree = theme.make_treeview(wrap, columns=cols, show='headings')
        for c in cols:
            self.bb_price_tree.heading(c, text=headers[c])
            w = 70 if c in ('size', 'pack') else (110 if c in ('fabric', 'color', 'price') else 220 if c == 'name' else 140)
            self.bb_price_tree.column(c, width=w, anchor='center')
        vs = ttk.Scrollbar(wrap, orient='vertical', command=self.bb_price_tree.yview)
        self.bb_price_tree.configure(yscrollcommand=vs.set)
        self.bb_price_tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self.bb_price_tree.bind('<Double-1>', self._bb_edit_price_cell)
        self._bb_refresh_price_table()

    def _build_bb_account_tab(self, parent):
        top = tk.Frame(parent, bg=theme.PAGE_BG)
        top.pack(fill='x', padx=10, pady=8)
        tk.Label(
            top,
            text="חשבון שותף עסקי — בייבי בייסיק",
            font=theme.FONT_SUBTITLE,
            bg=theme.PAGE_BG,
            fg=theme.DARK,
        ).pack(side='right', padx=4)
        tk.Label(
            top,
            text="לא לקוח. כל התעודות והתשלומים בחשבון אחד.",
            bg=theme.PAGE_BG,
            fg=theme.SUBTEXT,
            font=theme.FONT_SMALL,
        ).pack(side='right', padx=8)
        theme.make_button(top, "רענן", kind="secondary", command=self._bb_refresh_account).pack(side='left')

        cards = tk.Frame(parent, bg=theme.PAGE_BG)
        cards.pack(fill='x', padx=10, pady=4)

        def make_card(host, title, var, color):
            box = tk.Frame(host, bg=theme.CARD_BG, highlightbackground=theme.BORDER, highlightthickness=1)
            box.pack(side='right', fill='x', expand=True, padx=6)
            tk.Label(box, text=title, bg=theme.CARD_BG, fg=theme.SUBTEXT, font=theme.FONT_SMALL).pack(pady=(8, 0))
            lbl = tk.Label(box, textvariable=var, bg=theme.CARD_BG, fg=color, font=(theme.FONT_FAMILY, 18, 'bold'))
            lbl.pack(pady=(0, 10))
            return lbl

        self.bb_acc_supplied_var = tk.StringVar(value='0.00 ₪')
        self.bb_acc_paid_var = tk.StringVar(value='0.00 ₪')
        self.bb_acc_balance_var = tk.StringVar(value='0.00 ₪')
        make_card(cards, 'סופק (תעודות)', self.bb_acc_supplied_var, theme.DARK)
        make_card(cards, 'שולם', self.bb_acc_paid_var, theme.SUCCESS)
        self.bb_acc_balance_label = make_card(cards, 'יתרה לספק', self.bb_acc_balance_var, theme.DANGER)

        pay = ttk.LabelFrame(parent, text="רישום תשלום — בייבי בייסיק", padding=10)
        pay.pack(fill='x', padx=10, pady=6)
        tk.Label(pay, text="שותף:", bg=theme.PAGE_BG).grid(row=0, column=0, padx=4, pady=4, sticky='e')
        tk.Label(
            pay,
            text="בייבי בייסיק",
            font=theme.FONT_BODY_BOLD,
            bg=theme.PAGE_BG,
            fg=theme.DARK,
        ).grid(row=0, column=1, padx=4, pady=4, sticky='w')
        tk.Label(pay, text="תאריך:", bg=theme.PAGE_BG).grid(row=0, column=2, padx=4, pady=4, sticky='e')
        self.bb_pay_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        date_cell = tk.Frame(pay, bg=theme.PAGE_BG)
        date_cell.grid(row=0, column=3, padx=4, pady=4, sticky='w')
        try:
            DateEntry = None
            try:
                from tkcalendar import DateEntry  # type: ignore
            except Exception:
                DateEntry = None
            if DateEntry is not None:
                pay_date_entry = DateEntry(
                    date_cell,
                    textvariable=self.bb_pay_date_var,
                    width=12,
                    date_pattern='yyyy-mm-dd',
                    locale='he_IL',
                )
                try:
                    pay_date_entry.set_date(datetime.now())
                except Exception:
                    pass
            else:
                pay_date_entry = tk.Entry(date_cell, textvariable=self.bb_pay_date_var, width=12, justify='center')
            pay_date_entry.pack(side='left')
            tk.Button(
                date_cell,
                text='📅',
                width=2,
                command=lambda e=pay_date_entry, v=self.bb_pay_date_var: self._open_date_picker(e, v, lambda: None),
            ).pack(side='left', padx=(2, 0))
        except Exception:
            tk.Entry(date_cell, textvariable=self.bb_pay_date_var, width=12, justify='center').pack(side='left')
        tk.Label(pay, text="סכום ₪:", bg=theme.PAGE_BG).grid(row=0, column=4, padx=4, pady=4, sticky='e')
        self.bb_pay_amount_var = tk.StringVar()
        tk.Entry(pay, textvariable=self.bb_pay_amount_var, width=12, justify='center').grid(row=0, column=5, padx=4, pady=4)
        tk.Label(pay, text="הערה:", bg=theme.PAGE_BG).grid(row=0, column=6, padx=4, pady=4, sticky='e')
        self.bb_pay_note_var = tk.StringVar()
        tk.Entry(pay, textvariable=self.bb_pay_note_var, width=24).grid(row=0, column=7, padx=4, pady=4)
        theme.make_button(pay, "הוסף תשלום", kind="success", command=self._bb_add_payment).grid(row=0, column=8, padx=8, pady=4)

        body = ttk.Notebook(parent)
        body.pack(fill='both', expand=True, padx=10, pady=6)

        notes_page = tk.Frame(body, bg=theme.PAGE_BG)
        pay_page = tk.Frame(body, bg=theme.PAGE_BG)
        hist_page = tk.Frame(body, bg=theme.PAGE_BG)
        body.add(hist_page, text="תנועות")
        body.add(notes_page, text="תעודות שסופקו")
        body.add(pay_page, text="תשלומים")

        ncols = ('date', 'id', 'customer', 'qty', 'amount', 'note')
        nheaders = {'date': 'תאריך', 'id': 'תעודה', 'customer': 'שותף', 'qty': 'יחידות', 'amount': 'סכום', 'note': 'הערה'}
        self.bb_acc_notes_tree = theme.make_treeview(notes_page, columns=ncols, show='headings')
        for c in ncols:
            self.bb_acc_notes_tree.heading(c, text=nheaders[c])
            self.bb_acc_notes_tree.column(c, width=110 if c != 'note' else 220, anchor='center')
        nvs = ttk.Scrollbar(notes_page, orient='vertical', command=self.bb_acc_notes_tree.yview)
        self.bb_acc_notes_tree.configure(yscrollcommand=nvs.set)
        self.bb_acc_notes_tree.pack(side='left', fill='both', expand=True, padx=(6, 0), pady=6)
        nvs.pack(side='left', fill='y', pady=6)
        self.bb_acc_notes_tree.bind('<Double-1>', self._bb_open_account_note)

        pcols = ('date', 'id', 'partner', 'amount', 'note')
        pheaders = {'date': 'תאריך', 'id': 'מס׳', 'partner': 'שותף', 'amount': 'סכום', 'note': 'הערה'}
        pay_wrap = tk.Frame(pay_page, bg=theme.PAGE_BG)
        pay_wrap.pack(fill='both', expand=True)
        self.bb_acc_pay_tree = theme.make_treeview(pay_wrap, columns=pcols, show='headings')
        for c in pcols:
            self.bb_acc_pay_tree.heading(c, text=pheaders[c])
            self.bb_acc_pay_tree.column(c, width=110 if c != 'note' else 240, anchor='center')
        pvs = ttk.Scrollbar(pay_wrap, orient='vertical', command=self.bb_acc_pay_tree.yview)
        self.bb_acc_pay_tree.configure(yscrollcommand=pvs.set)
        self.bb_acc_pay_tree.pack(side='left', fill='both', expand=True, padx=(6, 0), pady=6)
        pvs.pack(side='left', fill='y', pady=6)
        self.bb_acc_pay_tree.bind('<Double-1>', self._bb_print_selected_payment_pdf)
        pay_actions = tk.Frame(pay_page, bg=theme.PAGE_BG)
        pay_actions.pack(fill='x', padx=10, pady=(0, 8))
        theme.make_button(pay_actions, "מחק תשלום", kind="danger", command=self._bb_delete_selected_payment).pack(side='right', padx=4)
        theme.make_button(pay_actions, "PDF לתשלום", kind="success", command=self._bb_print_selected_payment_pdf).pack(side='right', padx=4)
        theme.make_button(pay_actions, "PDF כל התשלומים", kind="primary", command=self._bb_print_payments_report_pdf).pack(side='right', padx=4)

        hcols = ('date', 'type', 'id', 'partner', 'debit', 'credit', 'balance', 'note')
        hheaders = {
            'date': 'תאריך', 'type': 'סוג', 'id': 'מס׳', 'partner': 'שותף',
            'debit': 'סופק', 'credit': 'שולם', 'balance': 'יתרה', 'note': 'הערה',
        }
        self.bb_acc_hist_tree = theme.make_treeview(hist_page, columns=hcols, show='headings')
        for c in hcols:
            self.bb_acc_hist_tree.heading(c, text=hheaders[c])
            w = 90 if c in ('date', 'id', 'debit', 'credit', 'balance') else (110 if c in ('type', 'partner') else 200)
            self.bb_acc_hist_tree.column(c, width=w, anchor='center')
        hvs = ttk.Scrollbar(hist_page, orient='vertical', command=self.bb_acc_hist_tree.yview)
        self.bb_acc_hist_tree.configure(yscrollcommand=hvs.set)
        self.bb_acc_hist_tree.pack(side='left', fill='both', expand=True, padx=(6, 0), pady=6)
        hvs.pack(side='left', fill='y', pady=6)

        self._bb_refresh_account()

    def _build_bb_room_count_tab(self, parent):
        header = ttk.LabelFrame(parent, text="ספירת מלאי חדר — בלי ברקודים", padding=10)
        header.pack(fill='x', padx=10, pady=6)
        tk.Label(header, text="תאריך:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).grid(row=0, column=0, sticky='e', padx=4, pady=4)
        self.bb_room_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(header, textvariable=self.bb_room_date_var, width=14, justify='center').grid(row=0, column=1, sticky='w', padx=4, pady=4)
        tk.Label(header, text="הערה:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).grid(row=0, column=2, sticky='e', padx=4, pady=4)
        self.bb_room_note_var = tk.StringVar()
        tk.Entry(header, textvariable=self.bb_room_note_var, width=40).grid(row=0, column=3, sticky='we', padx=4, pady=4)
        header.grid_columnconfigure(3, weight=1)

        add = ttk.LabelFrame(parent, text="הוספת שורה (דגם + סוג בד + צבע רקע + הדפס + מידה)", padding=10)
        add.pack(fill='x', padx=10, pady=4)
        tk.Label(add, text="דגם:", bg=theme.PAGE_BG).grid(row=0, column=0, sticky='e', padx=4, pady=4)
        self.bb_room_model_var = tk.StringVar()
        self.bb_room_model_combo = ttk.Combobox(add, textvariable=self.bb_room_model_var, width=28)
        self.bb_room_model_combo.grid(row=0, column=1, sticky='w', padx=4, pady=4)
        self.bb_room_model_combo.bind('<<ComboboxSelected>>', self._bb_on_room_model_change)
        self.bb_room_model_combo.bind('<KeyRelease>', self._bb_on_room_model_change)
        tk.Label(add, text="סוג בד:", bg=theme.PAGE_BG).grid(row=0, column=2, sticky='e', padx=4, pady=4)
        self.bb_room_fabric_var = tk.StringVar()
        self.bb_room_fabric_combo = ttk.Combobox(add, textvariable=self.bb_room_fabric_var, width=14)
        self.bb_room_fabric_combo.grid(row=0, column=3, sticky='w', padx=4, pady=4)
        tk.Label(add, text="צבע רקע:", bg=theme.PAGE_BG).grid(row=0, column=4, sticky='e', padx=4, pady=4)
        self.bb_room_color_var = tk.StringVar()
        self.bb_room_color_combo = ttk.Combobox(add, textvariable=self.bb_room_color_var, width=18)
        self.bb_room_color_combo.grid(row=0, column=5, sticky='w', padx=4, pady=4)
        tk.Label(add, text="הדפס:", bg=theme.PAGE_BG).grid(row=1, column=0, sticky='e', padx=4, pady=4)
        self.bb_room_print_var = tk.StringVar()
        self.bb_room_print_combo = ttk.Combobox(add, textvariable=self.bb_room_print_var, width=18)
        self.bb_room_print_combo.grid(row=1, column=1, sticky='w', padx=4, pady=4)
        tk.Label(add, text="מידה:", bg=theme.PAGE_BG).grid(row=1, column=2, sticky='e', padx=4, pady=4)
        self.bb_room_size_var = tk.StringVar()
        self.bb_room_size_combo = ttk.Combobox(add, textvariable=self.bb_room_size_var, width=12)
        self.bb_room_size_combo.grid(row=1, column=3, sticky='w', padx=4, pady=4)
        tk.Label(add, text="כמות:", bg=theme.PAGE_BG).grid(row=1, column=4, sticky='e', padx=4, pady=4)
        self.bb_room_qty_var = tk.StringVar()
        qty_entry = tk.Entry(add, textvariable=self.bb_room_qty_var, width=8, justify='center')
        qty_entry.grid(row=1, column=5, sticky='w', padx=4, pady=4)
        qty_entry.bind('<Return>', lambda e: self._bb_add_room_line())
        qty_entry.bind('<KP_Enter>', lambda e: self._bb_add_room_line())
        theme.make_button(add, "הוסף", kind="success", command=self._bb_add_room_line).grid(row=1, column=6, padx=8, pady=4, sticky='w')

        lines = ttk.LabelFrame(parent, text="שורות הספירה", padding=10)
        lines.pack(fill='both', expand=True, padx=10, pady=4)
        lcols = ('print_name', 'fabric', 'color', 'print', 'size', 'qty')
        lheaders = {
            'print_name': 'דגם',
            'fabric': 'סוג בד',
            'color': 'צבע רקע',
            'print': 'הדפס',
            'size': 'מידה',
            'qty': 'כמות',
        }
        lines_frame = tk.Frame(lines, bg=theme.PAGE_BG)
        lines_frame.pack(fill='both', expand=True)
        self.bb_room_lines_tree = theme.make_treeview(lines_frame, columns=lcols, show='headings', height=8)
        for c in lcols:
            self.bb_room_lines_tree.heading(c, text=lheaders[c])
            w = 90 if c in ('size', 'qty') else (130 if c in ('fabric', 'color', 'print') else 200)
            self.bb_room_lines_tree.column(c, width=w, anchor='center')
        lvs = ttk.Scrollbar(lines_frame, orient='vertical', command=self.bb_room_lines_tree.yview)
        self.bb_room_lines_tree.configure(yscrollcommand=lvs.set)
        self.bb_room_lines_tree.pack(side='left', fill='both', expand=True)
        lvs.pack(side='right', fill='y')
        self.bb_room_lines_tree.bind('<Double-1>', self._bb_edit_room_qty_cell)

        footer = tk.Frame(parent, bg=theme.PAGE_BG)
        footer.pack(fill='x', padx=10, pady=(2, 6))
        self.bb_room_total_var = tk.StringVar(value='סה״כ יחידות: 0')
        tk.Label(footer, textvariable=self.bb_room_total_var, font=theme.FONT_SUBTITLE, bg=theme.PAGE_BG, fg=theme.DARK).pack(side='right')
        theme.make_button(footer, "שמור ספירה", kind="success", command=self._bb_save_room_count).pack(side='left', padx=4)
        theme.make_button(footer, "ייבא Excel", kind="primary", command=self._bb_import_room_count_excel).pack(side='left', padx=4)
        theme.make_button(footer, "הסר שורה", kind="danger", command=self._bb_remove_selected_room_line).pack(side='left', padx=4)
        theme.make_button(footer, "נקה", kind="secondary", command=self._bb_clear_current_room_count).pack(side='left', padx=4)
        theme.make_button(footer, "PDF גיליון ספירה", kind="primary", command=self._bb_print_count_sheet_pdf).pack(side='left', padx=4)

        saved = ttk.LabelFrame(parent, text="ספירות שמורות", padding=10)
        saved.pack(fill='both', expand=True, padx=10, pady=(0, 8))
        scols = ('id', 'date', 'qty', 'note')
        sheaders = {'id': 'מס׳', 'date': 'תאריך', 'qty': 'יחידות', 'note': 'הערה'}
        saved_wrap = tk.Frame(saved, bg=theme.PAGE_BG)
        saved_wrap.pack(fill='both', expand=True)
        self.bb_saved_room_tree = theme.make_treeview(saved_wrap, columns=scols, show='headings', height=6)
        for c in scols:
            self.bb_saved_room_tree.heading(c, text=sheaders[c])
            w = 70 if c in ('id', 'qty') else (110 if c == 'date' else 320)
            self.bb_saved_room_tree.column(c, width=w, anchor='center')
        svs = ttk.Scrollbar(saved_wrap, orient='vertical', command=self.bb_saved_room_tree.yview)
        self.bb_saved_room_tree.configure(yscrollcommand=svs.set)
        self.bb_saved_room_tree.pack(side='left', fill='both', expand=True)
        svs.pack(side='right', fill='y')
        self.bb_saved_room_tree.bind('<Double-1>', self._bb_open_selected_room_count)

        saved_actions = tk.Frame(saved, bg=theme.PAGE_BG)
        saved_actions.pack(fill='x', pady=(6, 0))
        theme.make_button(saved_actions, "צפה", kind="primary", command=self._bb_open_selected_room_count).pack(side='right', padx=4)
        theme.make_button(saved_actions, "PDF להדפסה", kind="success", command=self._bb_print_selected_room_count_pdf).pack(side='right', padx=4)
        theme.make_button(saved_actions, "מחק", kind="danger", command=self._bb_delete_selected_room_count).pack(side='right', padx=4)
        theme.make_button(saved_actions, "רענן", kind="secondary", command=self._bb_refresh_saved_room_counts).pack(side='right', padx=4)

        self._bb_refresh_room_combos()
        self._bb_refresh_room_lines()
        self._bb_refresh_saved_room_counts()
