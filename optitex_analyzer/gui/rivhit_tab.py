"""טאב 'ריווחית' - הצגת רשימת מוצרים עדכנית מקובץ הייצוא של ריווחית."""
import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class RivhitTabMixin:
    """Mixin לטאב 'ריווחית'."""

    # מיפוי עמודות לכותרות בעברית
    _RIVHIT_COLS = ('item_num', 'item_name', 'item_part_num', 'item_cost_nis', 'item_sale_nis', 'compute_0036')
    _RIVHIT_HEADERS = {
        'item_num': 'מספר פריט',
        'item_name': 'שם הפריט',
        'item_part_num': 'מק"ט / ברקוד',
        'item_cost_nis': 'עלות (₪)',
        'item_sale_nis': 'מחיר מכירה (₪)',
        'compute_0036': 'עונה / קטגוריה',
        'digital_price': 'מחיר לצרכן דיגיטלי (₪)',
    }
    _RIVHIT_WIDTHS = {
        'item_num': 80,
        'item_name': 300,
        'item_part_num': 150,
        'item_cost_nis': 100,
        'item_sale_nis': 110,
        'compute_0036': 130,
        'digital_price': 130,
    }

    def _create_rivhit_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="ריווחית")

        # Inner notebook with two sub-tabs
        inner = ttk.Notebook(tab)
        inner.pack(fill='both', expand=True)

        list_frame = tk.Frame(inner, bg='#f7f9fa')
        add_frame = tk.Frame(inner, bg='#f7f9fa')
        data_frame = tk.Frame(inner, bg='#f7f9fa')
        inner.add(list_frame, text="רשימת מוצרים")
        inner.add(add_frame, text="הוספת מוצרים וייצוא")
        inner.add(data_frame, text="העלאת נתונים מריווחית")

        self._build_rivhit_list_subtab(list_frame)
        self._build_rivhit_add_subtab(add_frame)
        self._build_rivhit_data_subtab(data_frame)

    def _build_rivhit_list_subtab(self, tab):
        tk.Label(tab, text="רשימת מוצרים מריווחית", font=('Arial', 16, 'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)

        # Action bar
        actions = tk.Frame(tab, bg='#f7f9fa')
        actions.pack(fill='x', padx=15, pady=5)
        tk.Button(actions, text="🔄 רענן", command=self._refresh_rivhit_table, bg='#3498db', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5)
        tk.Button(actions, text="✏️ עריכת מדבקה", command=self._edit_rivhit_label_fields, bg='#8e44ad', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5)

        # Last upload info
        self.rivhit_meta_var = tk.StringVar(value='')
        tk.Label(actions, textvariable=self.rivhit_meta_var, bg='#f7f9fa', fg='#7f8c8d', font=('Arial', 9)).pack(side='right', padx=15)

        # Search bar
        search_frame = tk.Frame(tab, bg='#f7f9fa')
        search_frame.pack(fill='x', padx=15, pady=(0, 5))
        tk.Label(search_frame, text='🔍 חיפוש (שם או מק"ט):', bg='#f7f9fa').pack(side='right', padx=(6, 4))
        self.rivhit_search_var = tk.StringVar()
        self.rivhit_search_var.trace_add('write', lambda *args: self._filter_rivhit())
        search_entry = tk.Entry(search_frame, textvariable=self.rivhit_search_var, width=30)
        search_entry.pack(side='right', padx=(0, 6))
        tk.Button(search_frame, text='נקה', command=self._clear_rivhit_filters).pack(side='right', padx=4)

        # Category / season selector
        tk.Label(search_frame, text='🏷️ עונה / קטגוריה:', bg='#f7f9fa').pack(side='right', padx=(14, 4))
        self.rivhit_category_var = tk.StringVar(value='הכל')
        self.rivhit_category_combo = ttk.Combobox(search_frame, textvariable=self.rivhit_category_var, state='readonly', width=20, values=['הכל'])
        self.rivhit_category_combo.pack(side='right', padx=(0, 6))
        self.rivhit_category_combo.bind('<<ComboboxSelected>>', lambda e: self._filter_rivhit())

        # Table
        table_frame = tk.Frame(tab, bg='#ffffff')
        table_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.rivhit_tree = ttk.Treeview(table_frame, columns=self._RIVHIT_COLS, show='headings')
        for c in self._RIVHIT_COLS:
            self.rivhit_tree.heading(c, text=self._RIVHIT_HEADERS[c])
            self.rivhit_tree.column(c, width=self._RIVHIT_WIDTHS[c], anchor='center')
        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.rivhit_tree.yview)
        self.rivhit_tree.configure(yscroll=vsb.set)
        self.rivhit_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        self.rivhit_tree.bind('<Double-1>', self._edit_rivhit_label_fields)

        # Footer summary
        self.rivhit_summary_var = tk.StringVar(value="אין נתונים")
        tk.Label(tab, textvariable=self.rivhit_summary_var, bg='#2c3e50', fg='white', anchor='w', padx=12, font=('Arial', 10)).pack(fill='x', side='bottom')

        # Initial population
        self._update_rivhit_categories()
        self._populate_rivhit_table()
        self._update_rivhit_meta_label()

    def _populate_rivhit_table(self, records=None):
        if records is None:
            records = list(getattr(self.data_processor, 'rivhit_products', []) or [])
        for item in self.rivhit_tree.get_children():
            self.rivhit_tree.delete(item)
        for rec in records:
            self.rivhit_tree.insert('', 'end', values=tuple(rec.get(c, '') for c in self._RIVHIT_COLS))
        self._update_rivhit_summary(len(records))

    def _update_rivhit_summary(self, count):
        self.rivhit_summary_var.set(f"סה\"כ מוצרים: {count}")

    def _edit_rivhit_label_fields(self, event=None):
        """עריכת שדות הדפסת המדבקה של המוצר הנבחר (לפי ברקוד)."""
        sel = self.rivhit_tree.selection()
        if not sel:
            messagebox.showinfo("לא נבחר", "יש לבחור מוצר מהרשימה")
            return
        vals = self.rivhit_tree.item(sel[0], 'values')
        if not vals:
            return
        # סדר העמודות: item_num, item_name, item_part_num, ...
        name = vals[1] if len(vals) > 1 else ''
        barcode = str(vals[2]).strip() if len(vals) > 2 else ''
        if not barcode:
            messagebox.showwarning("אין ברקוד", "למוצר זה אין מק\"ט/ברקוד; לא ניתן לשמור שדות מדבקה")
            return
        fields = self.data_processor.get_rivhit_label_fields(barcode, product={'item_name': name})

        dlg = tk.Toplevel(self.notebook)
        dlg.title("עריכת שדות מדבקה")
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = tk.Frame(dlg, padx=15, pady=15)
        frm.pack(fill='both', expand=True)

        tk.Label(frm, text=f'מוצר: {name}', font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=2, sticky='e', pady=(0, 2))
        tk.Label(frm, text=f'ברקוד: {barcode}', font=('Arial', 9), fg='#7f8c8d').grid(row=1, column=0, columnspan=2, sticky='e', pady=(0, 10))

        name_var = tk.StringVar(value=fields.get('print_name', ''))
        size_var = tk.StringVar(value=fields.get('size', ''))
        size_unit_var = tk.StringVar(value=fields.get('size_unit', ''))
        fabric_var = tk.StringVar(value=fields.get('fabric', ''))
        pack_var = tk.StringVar(value=str(fields.get('pack_qty', 1)))
        image_var = tk.StringVar(value=fields.get('image', ''))

        rows = [
            ('שם להדפסה:', name_var, 32, False),
            ('מידה:', size_var, 20, False),
            ('סוג בד:', fabric_var, 20, False),
            ('כמות במארז:', pack_var, 8, True),
        ]
        for i, (lbl, var, width, is_spin) in enumerate(rows, start=2):
            tk.Label(frm, text=lbl, anchor='e').grid(row=i, column=1, sticky='e', padx=(6, 2), pady=3)
            if is_spin:
                tk.Spinbox(frm, from_=1, to=99, textvariable=var, width=width, justify='center').grid(row=i, column=0, sticky='w', pady=3)
            else:
                tk.Entry(frm, textvariable=var, width=width).grid(row=i, column=0, sticky='w', pady=3)

        # שורת יחידת מידה (חודשים/שנים) - מוצגת משמאל למידה במדבקה
        unit_row = len(rows) + 2
        tk.Label(frm, text='יחידת מידה:', anchor='e').grid(row=unit_row, column=1, sticky='e', padx=(6, 2), pady=3)
        ttk.Combobox(frm, textvariable=size_unit_var, width=18,
                     values=['', 'חודשים', 'שנים']).grid(row=unit_row, column=0, sticky='w', pady=3)

        # שורת תמונת מוצר: תצוגה מקדימה + בחירה + הסרה
        img_row = len(rows) + 3
        tk.Label(frm, text='תמונת מוצר:', anchor='e').grid(row=img_row, column=1, sticky='ne', padx=(6, 2), pady=3)
        img_frame = tk.Frame(frm)
        img_frame.grid(row=img_row, column=0, sticky='w', pady=3)
        preview_lbl = tk.Label(img_frame, text='(אין תמונה)', width=14, height=6,
                               relief='solid', bd=1, bg='#f7f7f7', compound='center')
        preview_lbl.grid(row=0, column=0, columnspan=2, pady=(0, 4))
        # שמירת רפרנס לתמונה כדי שלא תימחק ע"י garbage collector
        self._label_img_ref = None

        def _abs_image_path(rel):
            rel = (rel or '').strip()
            if not rel:
                return ''
            if os.path.isabs(rel):
                return rel
            return os.path.join(os.getcwd(), rel)

        def _update_preview():
            rel = image_var.get().strip()
            abs_path = _abs_image_path(rel)
            if rel and os.path.exists(abs_path):
                try:
                    from PIL import Image, ImageTk
                    im = Image.open(abs_path)
                    im.thumbnail((96, 96))
                    self._label_img_ref = ImageTk.PhotoImage(im)
                    preview_lbl.config(image=self._label_img_ref, text='')
                    return
                except Exception:
                    pass
            self._label_img_ref = None
            preview_lbl.config(image='', text='(אין תמונה)')

        def _choose_image():
            path = filedialog.askopenfilename(
                title='בחר תמונת מוצר',
                filetypes=[('תמונות', '*.png *.jpg *.jpeg *.gif *.bmp'), ('כל הקבצים', '*.*')],
            )
            if not path:
                return
            try:
                ext = os.path.splitext(path)[1].lower() or '.png'
                dest_dir = os.path.join(os.getcwd(), 'assets', 'labels', 'products')
                os.makedirs(dest_dir, exist_ok=True)
                safe_bc = ''.join(ch for ch in barcode if ch.isalnum()) or 'product'
                dest = os.path.join(dest_dir, f"{safe_bc}{ext}")
                shutil.copyfile(path, dest)
                rel = os.path.relpath(dest, os.getcwd())
                image_var.set(rel)
                _update_preview()
            except Exception as e:
                messagebox.showerror('שגיאה', f'שגיאה בשמירת התמונה:\n{e}')

        def _remove_image():
            image_var.set('')
            _update_preview()

        tk.Button(img_frame, text='בחר תמונה…', command=_choose_image).grid(row=1, column=0, padx=(0, 4))
        tk.Button(img_frame, text='הסר', command=_remove_image).grid(row=1, column=1)
        _update_preview()

        btns = tk.Frame(frm)
        btns.grid(row=img_row + 1, column=0, columnspan=2, pady=(12, 0))

        def save():
            self.data_processor.set_rivhit_label_fields(barcode, {
                'print_name': name_var.get(),
                'size': size_var.get(),
                'size_unit': size_unit_var.get(),
                'fabric': fabric_var.get(),
                'pack_qty': pack_var.get(),
                'image': image_var.get(),
            })
            dlg.destroy()
            messagebox.showinfo("נשמר", "שדות המדבקה נשמרו")

        tk.Button(btns, text="שמור", command=save, bg='#27ae60', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btns, text="ביטול", command=dlg.destroy).pack(side='left', padx=5)

    def _update_rivhit_meta_label(self):
        meta = getattr(self.data_processor, 'rivhit_meta', {}) or {}
        if meta.get('file_name'):
            text = f"קובץ אחרון: {meta.get('file_name', '')} | {meta.get('uploaded_at', '')} | {meta.get('count', 0)} פריטים"
        else:
            text = "טרם הועלה קובץ"
        self.rivhit_meta_var.set(text)
        if hasattr(self, 'rivhit_data_meta_var'):
            self.rivhit_data_meta_var.set(text)

    def _update_rivhit_categories(self):
        """מעדכן את רשימת הערכים בבורר העונה/קטגוריה.

        מאחד את הקטגוריות מהמוצרים שיובאו עם כל הקבוצות מקובץ הקבוצות
        (item_group.txt), כך שגם קבוצות ללא מוצרים (כמו 'שלישיות') יופיעו.
        """
        base = list(getattr(self.data_processor, 'rivhit_products', []) or [])
        cats = {str(r.get('compute_0036', '')).strip() for r in base if str(r.get('compute_0036', '')).strip()}
        groups = set((getattr(self.data_processor, 'rivhit_groups', {}) or {}).keys())
        all_cats = sorted(cats | groups)
        self.rivhit_category_combo['values'] = ['הכל'] + all_cats
        if self.rivhit_category_var.get() not in (['הכל'] + all_cats):
            self.rivhit_category_var.set('הכל')

    def _clear_rivhit_filters(self):
        self.rivhit_search_var.set('')
        self.rivhit_category_var.set('הכל')
        self._filter_rivhit()

    def _filter_rivhit(self):
        q = (self.rivhit_search_var.get() or '').strip().lower()
        category = self.rivhit_category_var.get() if hasattr(self, 'rivhit_category_var') else 'הכל'
        base = list(getattr(self.data_processor, 'rivhit_products', []) or [])
        # סינון לפי עונה/קטגוריה
        if category and category != 'הכל':
            base = [r for r in base if str(r.get('compute_0036', '')).strip() == category]
        # חיפוש לפי שם/מק"ט - תמיכה בכמה מילים (כל המילים חייבות להופיע)
        if q:
            terms = q.split()
            def matches(r):
                name = str(r.get('item_name', '')).lower()
                part = str(r.get('item_part_num', '')).lower()
                return all((t in name) or (t in part) for t in terms)
            base = [r for r in base if matches(r)]
        self._populate_rivhit_table(base)

    def _import_rivhit_file(self):
        file_path = filedialog.askopenfilename(
            title="בחר קובץ ייצוא מריווחית",
            filetypes=[("Text/CSV", "*.txt;*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            count = self.data_processor.import_rivhit_products(file_path)
            self.rivhit_search_var.set('')
            self.rivhit_category_var.set('הכל')
            self._update_rivhit_categories()
            self._populate_rivhit_table()
            self._update_rivhit_meta_label()
            messagebox.showinfo("הצלחה", f"נטענו {count} מוצרים מריווחית")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _refresh_rivhit_table(self):
        try:
            self.data_processor.refresh_rivhit_products()
        except Exception:
            pass
        self.rivhit_search_var.set('')
        self.rivhit_category_var.set('הכל')
        self._update_rivhit_categories()
        self._populate_rivhit_table()
        self._update_rivhit_meta_label()

    # ===== Add / Export sub-tab =====
    _RIVHIT_NEW_COLS = ('item_num', 'item_name', 'item_part_num', 'item_cost_nis', 'item_sale_nis', 'digital_price', 'compute_0036')

    def _build_rivhit_add_subtab(self, tab):
        tk.Label(tab, text="הוספת מוצרים וייצוא לריווחית", font=('Arial', 16, 'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)

        # Add form
        form = tk.LabelFrame(tab, text="פרטי מוצר חדש", bg='#f7f9fa', fg='#2c3e50', font=('Arial', 10, 'bold'), padx=10, pady=10)
        form.pack(fill='x', padx=15, pady=5)

        self.rivhit_new_name_var = tk.StringVar()
        self.rivhit_new_part_var = tk.StringVar()
        self.rivhit_new_cost_var = tk.StringVar()
        self.rivhit_new_sale_var = tk.StringVar()
        self.rivhit_new_digital_var = tk.StringVar()
        self.rivhit_new_cat_var = tk.StringVar()
        self.rivhit_new_last_item_var = tk.StringVar()

        # Row 1
        r1 = tk.Frame(form, bg='#f7f9fa'); r1.pack(fill='x', pady=3)
        tk.Label(r1, text='שם הפריט:', bg='#f7f9fa', width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r1, textvariable=self.rivhit_new_name_var, width=40).pack(side='right', padx=(0, 12))
        tk.Label(r1, text='מק"ט / ברקוד:', bg='#f7f9fa', width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r1, textvariable=self.rivhit_new_part_var, width=22).pack(side='right', padx=(0, 4))
        tk.Button(r1, text="חולל ברקוד", command=self._generate_rivhit_barcode, bg='#8e44ad', fg='white', font=('Arial', 9, 'bold')).pack(side='right', padx=(0, 12))

        # Row 2
        r2 = tk.Frame(form, bg='#f7f9fa'); r2.pack(fill='x', pady=3)
        tk.Label(r2, text='עלות (₪):', bg='#f7f9fa', width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r2, textvariable=self.rivhit_new_cost_var, width=12).pack(side='right', padx=(0, 12))
        tk.Label(r2, text='מחיר מכירה (₪):', bg='#f7f9fa', width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r2, textvariable=self.rivhit_new_sale_var, width=12).pack(side='right', padx=(0, 12))
        tk.Label(r2, text='מחיר לצרכן דיגיטלי (₪):', bg='#f7f9fa', width=18, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r2, textvariable=self.rivhit_new_digital_var, width=12).pack(side='right', padx=(0, 12))
        tk.Label(r2, text='עונה / קטגוריה:', bg='#f7f9fa', width=14, anchor='e').pack(side='right', padx=(6, 2))
        groups = sorted((getattr(self.data_processor, 'rivhit_groups', {}) or {}).keys())
        self.rivhit_new_cat_combo = ttk.Combobox(r2, textvariable=self.rivhit_new_cat_var, values=groups, state='readonly', width=20)
        self.rivhit_new_cat_combo.pack(side='right', padx=(0, 12))

        # Row 3 - starting item number control
        r3 = tk.Frame(form, bg='#f7f9fa'); r3.pack(fill='x', pady=3)
        tk.Label(r3, text='מספר פריט אחרון:', bg='#f7f9fa', width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r3, textvariable=self.rivhit_new_last_item_var, width=12).pack(side='right', padx=(0, 6))
        tk.Label(r3, text='(המספור יתחיל ממספר זה +1. ריק = לפי רשימת המוצרים העדכנית)', bg='#f7f9fa', fg='#7f8c8d', font=('Arial', 8)).pack(side='right', padx=(0, 12))

        # Sizes multi-select for batch creation
        sizes_box = tk.LabelFrame(form, text="יצירת מוצרים לפי מידות (מוצר לכל מידה, עם ברקוד חדש)", bg='#f7f9fa', fg='#2c3e50', font=('Arial', 9, 'bold'), padx=8, pady=6)
        sizes_box.pack(fill='x', pady=(8, 2))
        sizes_actions = tk.Frame(sizes_box, bg='#f7f9fa'); sizes_actions.pack(fill='x', anchor='e')
        tk.Button(sizes_actions, text="נקה בחירה", command=self._clear_rivhit_size_selection, bg='#95a5a6', fg='white', font=('Arial', 8)).pack(side='right', padx=3)
        tk.Button(sizes_actions, text="סמן הכל", command=self._select_all_rivhit_sizes, bg='#3498db', fg='white', font=('Arial', 8)).pack(side='right', padx=3)
        self.rivhit_sizes_grid = tk.Frame(sizes_box, bg='#f7f9fa'); self.rivhit_sizes_grid.pack(fill='x', pady=(4, 0))
        self.rivhit_size_vars = {}
        self._build_rivhit_sizes_checkboxes()

        # Print (label) details - applied to the created product(s)
        self.rivhit_print_name_var = tk.StringVar()
        self.rivhit_print_size_var = tk.StringVar()
        self.rivhit_print_unit_var = tk.StringVar(value='חודשים')
        self.rivhit_print_fabric_var = tk.StringVar()
        self.rivhit_print_pack_var = tk.StringVar(value='3')
        self.rivhit_print_image_src_var = tk.StringVar()
        self.rivhit_print_image_label_var = tk.StringVar(value='ללא תמונה')

        print_box = tk.LabelFrame(form, text="פרטי הדפסה למדבקה (יחולו על כל המוצרים שייווצרו)", bg='#f7f9fa', fg='#2c3e50', font=('Arial', 9, 'bold'), padx=8, pady=6)
        print_box.pack(fill='x', pady=(8, 2))
        pr1 = tk.Frame(print_box, bg='#f7f9fa'); pr1.pack(fill='x', pady=2)
        tk.Label(pr1, text='שם להדפסה:', bg='#f7f9fa', width=12, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(pr1, textvariable=self.rivhit_print_name_var, width=26).pack(side='right', padx=(0, 12))
        tk.Label(pr1, text='סוג בד:', bg='#f7f9fa', width=8, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(pr1, textvariable=self.rivhit_print_fabric_var, width=14).pack(side='right', padx=(0, 12))
        tk.Label(pr1, text='כמות במארז:', bg='#f7f9fa', width=10, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(pr1, textvariable=self.rivhit_print_pack_var, width=6).pack(side='right', padx=(0, 12))
        pr2 = tk.Frame(print_box, bg='#f7f9fa'); pr2.pack(fill='x', pady=2)
        tk.Label(pr2, text='מידה (למוצר בודד):', bg='#f7f9fa', width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(pr2, textvariable=self.rivhit_print_size_var, width=10).pack(side='right', padx=(0, 12))
        tk.Label(pr2, text='יחידת מידה:', bg='#f7f9fa', width=10, anchor='e').pack(side='right', padx=(6, 2))
        ttk.Combobox(pr2, textvariable=self.rivhit_print_unit_var, values=['', 'חודשים', 'שנים'], state='readonly', width=10).pack(side='right', padx=(0, 12))
        tk.Button(pr2, text='בחר תמונה…', command=self._choose_rivhit_print_image, bg='#8e44ad', fg='white', font=('Arial', 8)).pack(side='right', padx=(0, 6))
        tk.Label(pr2, textvariable=self.rivhit_print_image_label_var, bg='#f7f9fa', fg='#7f8c8d', font=('Arial', 8)).pack(side='right', padx=(0, 12))
        tk.Label(print_box, text="(ביצירה לפי מידות - המידה נלקחת מכל מידה שנבחרה; שדה 'מידה' משמש להוספת מוצר בודד)", bg='#f7f9fa', fg='#7f8c8d', font=('Arial', 8)).pack(anchor='e', pady=(2, 0))

        btns = tk.Frame(form, bg='#f7f9fa'); btns.pack(pady=(8, 2))
        tk.Button(btns, text="➕ הוסף מוצר", command=self._add_rivhit_product, bg='#27ae60', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=4)
        tk.Button(btns, text="🧩 צור מוצרים לפי מידות", command=self._create_rivhit_products_by_sizes, bg='#16a085', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=4)

        # Pending list toolbar
        toolbar = tk.Frame(tab, bg='#f7f9fa')
        toolbar.pack(fill='x', padx=15, pady=(8, 2))
        tk.Label(toolbar, text="מוצרים ממתינים לייצוא:", bg='#f7f9fa', font=('Arial', 11, 'bold')).pack(side='right')
        tk.Button(toolbar, text="⬇️ ייצא קובץ לריווחית", command=self._export_rivhit_new, bg='#27ae60', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=4)
        tk.Button(toolbar, text="🗑️ נקה הכל", command=self._clear_rivhit_new, bg='#e74c3c', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=4)
        tk.Button(toolbar, text="מחק נבחר", command=self._delete_rivhit_new_selected, bg='#95a5a6', fg='white', font=('Arial', 10, 'bold')).pack(side='left', padx=4)

        # Pending list table
        table_frame = tk.Frame(tab, bg='#ffffff')
        table_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.rivhit_new_tree = ttk.Treeview(table_frame, columns=self._RIVHIT_NEW_COLS, show='headings')
        for c in self._RIVHIT_NEW_COLS:
            self.rivhit_new_tree.heading(c, text=self._RIVHIT_HEADERS[c])
            self.rivhit_new_tree.column(c, width=self._RIVHIT_WIDTHS[c], anchor='center')
        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.rivhit_new_tree.yview)
        self.rivhit_new_tree.configure(yscroll=vsb.set)
        self.rivhit_new_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        self.rivhit_new_summary_var = tk.StringVar(value="אין מוצרים ממתינים")
        tk.Label(tab, textvariable=self.rivhit_new_summary_var, bg='#2c3e50', fg='white', anchor='w', padx=12, font=('Arial', 10)).pack(fill='x', side='bottom')

        self._refresh_rivhit_new_table()

    def _refresh_rivhit_new_table(self):
        records = list(getattr(self.data_processor, 'rivhit_new_products', []) or [])
        for item in self.rivhit_new_tree.get_children():
            self.rivhit_new_tree.delete(item)
        for rec in records:
            self.rivhit_new_tree.insert('', 'end', values=tuple(rec.get(c, '') for c in self._RIVHIT_NEW_COLS))
        self.rivhit_new_summary_var.set(f"מוצרים ממתינים לייצוא: {len(records)}")

    def _generate_rivhit_barcode(self):
        try:
            code = self.data_processor.generate_and_reserve_barcode()
            self.rivhit_new_part_var.set(code)
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _choose_rivhit_print_image(self):
        path = filedialog.askopenfilename(
            title='בחר תמונת מוצר למדבקה',
            filetypes=[('תמונות', '*.png *.jpg *.jpeg *.gif *.bmp'), ('כל הקבצים', '*.*')],
        )
        if not path:
            return
        self.rivhit_print_image_src_var.set(path)
        self.rivhit_print_image_label_var.set(os.path.basename(path))

    def _apply_rivhit_print_fields(self, barcode, size=''):
        """שומר פרטי הדפסה (מדבקה) למוצר שנוצר לפי הברקוד שלו.

        מועתק מהשדות בטופס; התמונה (אם נבחרה) מועתקת לקובץ פר-ברקוד.
        """
        barcode = str(barcode or '').strip()
        if not barcode:
            return
        print_name = (self.rivhit_print_name_var.get() or '').strip()
        fabric = (self.rivhit_print_fabric_var.get() or '').strip()
        size_unit = (self.rivhit_print_unit_var.get() or '').strip()
        pack = (self.rivhit_print_pack_var.get() or '').strip()
        img_src = (self.rivhit_print_image_src_var.get() or '').strip()
        size = str(size or '').strip()
        # החל רק אם המשתמש הזין פרט הדפסה כלשהו
        if not (print_name or fabric or img_src or size or size_unit):
            return
        image_rel = ''
        if img_src and os.path.exists(img_src):
            try:
                ext = os.path.splitext(img_src)[1].lower() or '.png'
                dest_dir = os.path.join(os.getcwd(), 'assets', 'labels', 'products')
                os.makedirs(dest_dir, exist_ok=True)
                safe_bc = ''.join(ch for ch in barcode if ch.isalnum()) or 'product'
                dest = os.path.join(dest_dir, f"{safe_bc}{ext}")
                shutil.copyfile(img_src, dest)
                image_rel = os.path.relpath(dest, os.getcwd())
            except Exception:
                image_rel = ''
        self.data_processor.set_rivhit_label_fields(barcode, {
            'print_name': print_name,
            'size': size,
            'size_unit': size_unit,
            'fabric': fabric,
            'pack_qty': pack or 1,
            'image': image_rel,
        })

    def _reset_rivhit_print_fields(self):
        self.rivhit_print_name_var.set('')
        self.rivhit_print_size_var.set('')
        self.rivhit_print_fabric_var.set('')
        self.rivhit_print_image_src_var.set('')
        self.rivhit_print_image_label_var.set('ללא תמונה')

    def _add_rivhit_product(self):
        name = (self.rivhit_new_name_var.get() or '').strip()
        if not name:
            messagebox.showwarning("שדה חסר", "יש להזין שם פריט")
            return
        try:
            record = self.data_processor.add_rivhit_new_product(
                name=name,
                part_num=self.rivhit_new_part_var.get(),
                cost_nis=self.rivhit_new_cost_var.get(),
                sale_nis=self.rivhit_new_sale_var.get(),
                category=self.rivhit_new_cat_var.get(),
                digital_price=self.rivhit_new_digital_var.get(),
                last_item_num=(self.rivhit_new_last_item_var.get() or '').strip() or None,
            )
            self._apply_rivhit_print_fields(
                record.get('item_part_num', ''),
                size=(self.rivhit_print_size_var.get() or '').strip(),
            )
            self.rivhit_new_name_var.set('')
            self.rivhit_new_part_var.set('')
            self.rivhit_new_cost_var.set('')
            self.rivhit_new_sale_var.set('')
            self.rivhit_new_digital_var.set('')
            self.rivhit_new_cat_var.set('')
            self._reset_rivhit_print_fields()
            self._refresh_rivhit_new_table()
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _build_rivhit_sizes_checkboxes(self):
        for w in self.rivhit_sizes_grid.winfo_children():
            w.destroy()
        self.rivhit_size_vars = {}
        sizes = [str(s.get('name', '')).strip() for s in (getattr(self.data_processor, 'product_sizes', []) or []) if str(s.get('name', '')).strip()]
        cols = 6
        for i, size in enumerate(sizes):
            var = tk.BooleanVar(value=False)
            self.rivhit_size_vars[size] = var
            cb = tk.Checkbutton(self.rivhit_sizes_grid, text=size, variable=var, bg='#f7f9fa', anchor='w')
            cb.grid(row=i // cols, column=i % cols, sticky='w', padx=4, pady=1)

    def _select_all_rivhit_sizes(self):
        for var in self.rivhit_size_vars.values():
            var.set(True)

    def _clear_rivhit_size_selection(self):
        for var in self.rivhit_size_vars.values():
            var.set(False)

    def _create_rivhit_products_by_sizes(self):
        name = (self.rivhit_new_name_var.get() or '').strip()
        if not name:
            messagebox.showwarning("שדה חסר", "יש להזין שם פריט")
            return
        selected = [size for size, var in self.rivhit_size_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("לא נבחרו מידות", "יש לבחור לפחות מידה אחת")
            return
        try:
            created = self.data_processor.add_rivhit_new_products_by_sizes(
                base_name=name,
                sizes=selected,
                cost_nis=self.rivhit_new_cost_var.get(),
                sale_nis=self.rivhit_new_sale_var.get(),
                category=self.rivhit_new_cat_var.get(),
                digital_price=self.rivhit_new_digital_var.get(),
                last_item_num=(self.rivhit_new_last_item_var.get() or '').strip() or None,
            )
            # החל פרטי הדפסה לכל מוצר שנוצר, עם המידה התואמת שלו
            for rec, size in zip(created, selected):
                self._apply_rivhit_print_fields(rec.get('item_part_num', ''), size=size)
            self.rivhit_new_name_var.set('')
            self.rivhit_new_part_var.set('')
            self.rivhit_new_cost_var.set('')
            self.rivhit_new_sale_var.set('')
            self.rivhit_new_digital_var.set('')
            self.rivhit_new_cat_var.set('')
            self._reset_rivhit_print_fields()
            self._clear_rivhit_size_selection()
            self._refresh_rivhit_new_table()
            messagebox.showinfo("הצלחה", f"נוצרו {len(created)} מוצרים לפי המידות שנבחרו")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _delete_rivhit_new_selected(self):
        sel = self.rivhit_new_tree.selection()
        if not sel:
            messagebox.showinfo("לא נבחר", "יש לבחור מוצר למחיקה")
            return
        index = self.rivhit_new_tree.index(sel[0])
        if self.data_processor.delete_rivhit_new_product(index):
            self._refresh_rivhit_new_table()

    def _clear_rivhit_new(self):
        if not (getattr(self.data_processor, 'rivhit_new_products', []) or []):
            return
        if messagebox.askyesno("ניקוי", "למחוק את כל המוצרים הממתינים?"):
            self.data_processor.clear_rivhit_new_products()
            self._refresh_rivhit_new_table()

    def _export_rivhit_new(self):
        from datetime import datetime
        if not (getattr(self.data_processor, 'rivhit_new_products', []) or []):
            messagebox.showinfo("אין מוצרים", "אין מוצרים חדשים לייצוא")
            return
        default_name = f"ריווחית_מוצרים_חדשים_{datetime.now().strftime('%d.%m.%y')}.txt"
        file_path = filedialog.asksaveasfilename(
            title="שמירת קובץ ייצוא לריווחית",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("Text", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            count = self.data_processor.export_rivhit_new_products(file_path)
            messagebox.showinfo("הצלחה", f"יוצאו {count} מוצרים לקובץ:\n{file_path}")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    # ===== Data upload sub-tab (יבוא נתונים מריווחית) =====
    def _build_rivhit_data_subtab(self, tab):
        tk.Label(tab, text="העלאת נתונים מריווחית", font=('Arial', 16, 'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)

        # Section 1: items list import
        items_box = tk.LabelFrame(tab, text="רשימת פריטים מריווחית", bg='#f7f9fa', fg='#2c3e50', font=('Arial', 11, 'bold'), padx=12, pady=12)
        items_box.pack(fill='x', padx=15, pady=8)
        tk.Label(items_box, text="העלאת קובץ ייצוא הפריטים מריווחית מחליפה את הרשימה המוצגת בטאב 'רשימת מוצרים'.", bg='#f7f9fa', fg='#555', font=('Arial', 9), justify='right').pack(anchor='e', pady=(0, 6))
        items_row = tk.Frame(items_box, bg='#f7f9fa'); items_row.pack(fill='x')
        tk.Button(items_row, text="⬆️ העלה קובץ פריטים מריווחית", command=self._import_rivhit_file, bg='#27ae60', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5)
        self.rivhit_data_meta_var = tk.StringVar(value='')
        tk.Label(items_box, textvariable=self.rivhit_data_meta_var, bg='#f7f9fa', fg='#7f8c8d', font=('Arial', 9)).pack(anchor='e', pady=(6, 0))

        # Section 2: groups/categories import
        groups_box = tk.LabelFrame(tab, text="קבוצות / קטגוריות", bg='#f7f9fa', fg='#2c3e50', font=('Arial', 11, 'bold'), padx=12, pady=12)
        groups_box.pack(fill='x', padx=15, pady=8)
        tk.Label(groups_box, text="קובץ הקבוצות (מספר<TAB>שם) קובע אילו עונות/קטגוריות זמינות בהוספת מוצר וכיצד הן ממופות בייצוא.", bg='#f7f9fa', fg='#555', font=('Arial', 9), justify='right').pack(anchor='e', pady=(0, 6))
        groups_row = tk.Frame(groups_box, bg='#f7f9fa'); groups_row.pack(fill='x')
        tk.Button(groups_row, text="⬆️ העלה קובץ קבוצות", command=self._import_rivhit_groups_file, bg='#8e44ad', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5)
        self.rivhit_groups_meta_var = tk.StringVar(value='')
        tk.Label(groups_box, textvariable=self.rivhit_groups_meta_var, bg='#f7f9fa', fg='#7f8c8d', font=('Arial', 9)).pack(anchor='e', pady=(6, 0))

        self._update_rivhit_meta_label()
        self._update_rivhit_groups_meta_label()

    def _update_rivhit_groups_meta_label(self):
        if not hasattr(self, 'rivhit_groups_meta_var'):
            return
        meta = getattr(self.data_processor, 'rivhit_groups_meta', {}) or {}
        count = len(getattr(self.data_processor, 'rivhit_groups', {}) or {})
        if meta.get('file_name'):
            self.rivhit_groups_meta_var.set(
                f"קובץ אחרון: {meta.get('file_name', '')} | {meta.get('uploaded_at', '')} | {count} קבוצות"
            )
        else:
            self.rivhit_groups_meta_var.set(f"{count} קבוצות טעונות")

    def _import_rivhit_groups_file(self):
        file_path = filedialog.askopenfilename(
            title="בחר קובץ קבוצות מריווחית",
            filetypes=[("Text/CSV", "*.txt;*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            count = self.data_processor.import_rivhit_groups(file_path)
            self._update_rivhit_groups_meta_label()
            self._refresh_rivhit_groups_combo()
            if hasattr(self, 'rivhit_category_combo'):
                self._update_rivhit_categories()
            messagebox.showinfo("הצלחה", f"נטענו {count} קבוצות")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _refresh_rivhit_groups_combo(self):
        if hasattr(self, 'rivhit_new_cat_combo'):
            groups = sorted((getattr(self.data_processor, 'rivhit_groups', {}) or {}).keys())
            self.rivhit_new_cat_combo['values'] = groups
