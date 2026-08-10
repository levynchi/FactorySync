"""טאב 'ריווחית' - הצגת רשימת מוצרים עדכנית מקובץ הייצוא של ריווחית."""
import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from . import theme


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
        tab = tk.Frame(self.notebook, bg=theme.PAGE_BG)
        self.notebook.add(tab, text="ריווחית")

        # Inner notebook with two sub-tabs
        inner = ttk.Notebook(tab)
        inner.pack(fill='both', expand=True)

        list_frame = tk.Frame(inner, bg=theme.PAGE_BG)
        add_frame = tk.Frame(inner, bg=theme.PAGE_BG)
        data_frame = tk.Frame(inner, bg=theme.PAGE_BG)
        inner.add(list_frame, text="רשימת מוצרים")
        inner.add(add_frame, text="הוספת מוצרים וייצוא")
        inner.add(data_frame, text="העלאת נתונים מריווחית")

        self._build_rivhit_list_subtab(list_frame)
        self._build_rivhit_add_subtab(add_frame)
        self._build_rivhit_data_subtab(data_frame)

    def _build_rivhit_list_subtab(self, tab):
        tk.Label(tab, text="רשימת מוצרים מריווחית", font=(theme.FONT_FAMILY, 16, 'bold'), bg=theme.PAGE_BG, fg=theme.DARK).pack(pady=8)

        # Action bar
        actions = tk.Frame(tab, bg=theme.PAGE_BG)
        actions.pack(fill='x', padx=15, pady=5)
        tk.Button(actions, text="🔄 רענן", command=self._refresh_rivhit_table, bg=theme.PRIMARY, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        tk.Button(actions, text="✏️ עריכת מדבקה", command=self._edit_rivhit_label_fields, bg=theme.PURPLE, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        tk.Button(actions, text="🧬 ניהול דגמים", command=self._open_rivhit_family_manager, bg=theme.TEAL, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)

        # Last upload info
        self.rivhit_meta_var = tk.StringVar(value='')
        tk.Label(actions, textvariable=self.rivhit_meta_var, bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9)).pack(side='right', padx=15)

        # Search bar
        search_frame = tk.Frame(tab, bg=theme.PAGE_BG)
        search_frame.pack(fill='x', padx=15, pady=(0, 5))
        tk.Label(search_frame, text='🔍 חיפוש (שם או מק"ט):', bg=theme.PAGE_BG).pack(side='right', padx=(6, 4))
        self.rivhit_search_var = tk.StringVar()
        self.rivhit_search_var.trace_add('write', lambda *args: self._filter_rivhit())
        search_entry = tk.Entry(search_frame, textvariable=self.rivhit_search_var, width=30)
        search_entry.pack(side='right', padx=(0, 6))
        tk.Button(search_frame, text='נקה', command=self._clear_rivhit_filters).pack(side='right', padx=4)

        # Category / season selector
        tk.Label(search_frame, text='🏷️ עונה / קטגוריה:', bg=theme.PAGE_BG).pack(side='right', padx=(14, 4))
        self.rivhit_category_var = tk.StringVar(value='הכל')
        self.rivhit_category_combo = ttk.Combobox(search_frame, textvariable=self.rivhit_category_var, state='readonly', width=20, values=['הכל'])
        self.rivhit_category_combo.pack(side='right', padx=(0, 6))
        self.rivhit_category_combo.bind('<<ComboboxSelected>>', lambda e: self._filter_rivhit())

        # Table
        table_frame = tk.Frame(tab, bg=theme.CARD_BG)
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
        self.rivhit_tree.bind('<Double-1>', self._on_rivhit_tree_double_click)

        # Footer summary
        self.rivhit_summary_var = tk.StringVar(value="אין נתונים")
        tk.Label(tab, textvariable=self.rivhit_summary_var, bg=theme.DARK, fg='white', anchor='w', padx=12, font=(theme.FONT_FAMILY, 10)).pack(fill='x', side='bottom')

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

    def _on_rivhit_tree_double_click(self, event):
        """דאבל-קליק על עמודת מק\"ט מעתיק את הברקוד; אחרת פותח עריכת מדבקה."""
        tree = self.rivhit_tree
        row = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        # item_part_num הוא העמודה השלישית ב-_RIVHIT_COLS => '#3'
        if row and col == '#3':
            vals = tree.item(row, 'values')
            barcode = str(vals[2]).strip() if len(vals) > 2 else ''
            if barcode:
                tree.clipboard_clear()
                tree.clipboard_append(barcode)
                prev = self.rivhit_summary_var.get()
                self.rivhit_summary_var.set(f'הועתק מק"ט: {barcode}')
                tree.after(2000, lambda: self.rivhit_summary_var.set(prev))
            return
        self._edit_rivhit_label_fields(event)

    def _edit_rivhit_label_fields(self, event=None):
        """עריכת שדות הדפסת המדבקה של המוצר הנבחר (לפי ברקוד)."""
        sel = self.rivhit_tree.selection()
        if not sel:
            messagebox.showinfo("לא נבחר", "יש לבחור מוצר מהרשימה")
            return
        vals = self.rivhit_tree.item(sel[0], 'values')
        if not vals:
            return
        # סדר העמודות: item_num, item_name, item_part_num, ..., compute_0036 (קטגוריה)
        name = vals[1] if len(vals) > 1 else ''
        barcode = str(vals[2]).strip() if len(vals) > 2 else ''
        category = str(vals[5]).strip() if len(vals) > 5 else ''
        brand = self.data_processor.brand_key_from_category(category)
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

        tk.Label(frm, text=f'מוצר: {name}', font=(theme.FONT_FAMILY, 10, 'bold')).grid(row=0, column=0, columnspan=2, sticky='e', pady=(0, 2))
        tk.Label(frm, text=f'ברקוד: {barcode}', font=(theme.FONT_FAMILY, 9), fg=theme.SUBTEXT).grid(row=1, column=0, columnspan=2, sticky='e', pady=(0, 10))

        name_var = tk.StringVar(value=fields.get('print_name', ''))
        size_var = tk.StringVar(value=fields.get('size', ''))
        size_unit_var = tk.StringVar(value=fields.get('size_unit', ''))
        fabric_var = tk.StringVar(value=fields.get('fabric', ''))
        pack_var = tk.StringVar(value=str(fields.get('pack_qty', 1)))
        image_var = tk.StringVar(value=fields.get('image', ''))
        model_code_var = tk.StringVar(value=fields.get('model_code', ''))

        rows = [
            ('שם להדפסה:', name_var, 32, False),
            ('קוד דגם (באנגלית):', model_code_var, 20, False),
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
                               relief='solid', bd=1, bg=theme.PAGE_BG, compound='center')
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
            products_dir = os.path.join(os.getcwd(), 'assets', 'labels', 'products')
            os.makedirs(products_dir, exist_ok=True)
            path = filedialog.askopenfilename(
                title='בחר תמונת מוצר',
                initialdir=products_dir,
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
            model_code = (model_code_var.get() or '').strip()
            self.data_processor.set_rivhit_label_fields(barcode, {
                'print_name': name_var.get(),
                'size': size_var.get(),
                'size_unit': size_unit_var.get(),
                'fabric': fabric_var.get(),
                'pack_qty': pack_var.get(),
                'image': image_var.get(),
                'model_code': model_code,
                'brand': brand,
            })
            # אם למוצר יש קוד דגם - עדכון דגם האב (מותג+קוד) והפצה לכל הוואריאנטים
            propagated = 0
            if model_code:
                self.data_processor.set_rivhit_family(model_code, {
                    'print_name': name_var.get(),
                    'fabric': fabric_var.get(),
                    'pack_qty': pack_var.get(),
                    'size_unit': size_unit_var.get(),
                    'image': image_var.get(),
                }, brand=brand)
                propagated = self.data_processor.rivhit_family_variant_count(model_code, brand)
            dlg.destroy()
            if propagated > 1:
                brand_name = self.data_processor.RIVHIT_BRAND_NAMES.get(brand, brand)
                messagebox.showinfo("נשמר", f"שדות המדבקה נשמרו והופצו ל-{propagated} וואריאנטים של דגם {model_code} ({brand_name})")
            else:
                messagebox.showinfo("נשמר", "שדות המדבקה נשמרו")

        tk.Button(btns, text="שמור", command=save, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btns, text="ביטול", command=dlg.destroy).pack(side='left', padx=5)

    # ===== ניהול דגמי אב (משפחות מוצרים) =====
    def _open_rivhit_family_manager(self):
        """דיאלוג לניהול דגמי אב: רשימה + עריכה מרכזית עם הפצה לכל הוואריאנטים."""
        dlg = tk.Toplevel(self.notebook)
        dlg.title("ניהול דגמים (משפחות מוצרים)")
        dlg.geometry("640x420")
        dlg.grab_set()

        tk.Label(dlg, text="דגמי אב - עריכה מרכזית מפיצה לכל הוואריאנטים עם אותו קוד דגם",
                 font=(theme.FONT_FAMILY, 11, 'bold')).pack(pady=(10, 4))

        table_frame = tk.Frame(dlg)
        table_frame.pack(fill='both', expand=True, padx=12, pady=6)
        cols = ('brand', 'model_code', 'print_name', 'variants', 'fabric', 'pack')
        headers = {'brand': 'מותג', 'model_code': 'קוד דגם', 'print_name': 'שם להדפסה', 'variants': 'וואריאנטים', 'fabric': 'סוג בד', 'pack': 'מארז'}
        tree = ttk.Treeview(table_frame, columns=cols, show='headings')
        for c in cols:
            tree.heading(c, text=headers[c])
            tree.column(c, width=110 if c != 'print_name' else 200, anchor='center')
        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        tree.configure(yscroll=vsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        def refresh():
            for it in tree.get_children():
                tree.delete(it)
            for brand, code, fam in self.data_processor.list_rivhit_families():
                brand_name = self.data_processor.RIVHIT_BRAND_NAMES.get(brand, brand)
                tree.insert('', 'end', iid=self.data_processor._family_key(brand, code), values=(
                    brand_name, code, fam.get('print_name', ''),
                    self.data_processor.rivhit_family_variant_count(code, brand),
                    fam.get('fabric', ''), fam.get('pack_qty', ''),
                ))

        def edit_selected(event=None):
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("לא נבחר", "יש לבחור דגם לעריכה", parent=dlg)
                return
            v = tree.item(sel[0], 'values')
            # brand הוא שם תצוגה - נמיר חזרה למזהה
            brand_key = next((k for k, nm in self.data_processor.RIVHIT_BRAND_NAMES.items() if nm == v[0]), v[0])
            self._edit_rivhit_family(v[1], brand=brand_key, on_saved=refresh, parent=dlg)

        tree.bind('<Double-1>', edit_selected)

        btns = tk.Frame(dlg)
        btns.pack(fill='x', padx=12, pady=(0, 10))
        tk.Button(btns, text="✏️ ערוך דגם נבחר", command=edit_selected, bg=theme.PURPLE, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        tk.Button(btns, text="🔄 רענן", command=refresh, bg=theme.PRIMARY, fg='white').pack(side='right', padx=5)
        tk.Button(btns, text="סגור", command=dlg.destroy).pack(side='left', padx=5)

        refresh()

    def _edit_rivhit_family(self, model_code, brand='arie', on_saved=None, parent=None):
        """עריכת פרטי דגם אב (שם להדפסה, בד, מארז, יחידה, תמונה) והפצה לכל הוואריאנטים."""
        code = str(model_code or '').strip()
        if not code:
            return
        brand = (str(brand or '').strip() or 'arie')
        brand_name = self.data_processor.RIVHIT_BRAND_NAMES.get(brand, brand)
        fam = self.data_processor.get_rivhit_family(code, brand=brand)
        dlg = tk.Toplevel(parent or self.notebook)
        dlg.title(f"עריכת דגם {code} ({brand_name})")
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = tk.Frame(dlg, padx=15, pady=15)
        frm.pack(fill='both', expand=True)

        n = self.data_processor.rivhit_family_variant_count(code, brand)
        tk.Label(frm, text=f'קוד דגם: {code} | מותג: {brand_name}', font=(theme.FONT_FAMILY, 11, 'bold')).grid(row=0, column=0, columnspan=2, sticky='e', pady=(0, 2))
        tk.Label(frm, text=f'{n} וואריאנטים ישתנו', font=(theme.FONT_FAMILY, 9), fg=theme.SUBTEXT).grid(row=1, column=0, columnspan=2, sticky='e', pady=(0, 10))

        name_var = tk.StringVar(value=fam.get('print_name', ''))
        fabric_var = tk.StringVar(value=fam.get('fabric', ''))
        pack_var = tk.StringVar(value=str(fam.get('pack_qty', 1)))
        size_unit_var = tk.StringVar(value=fam.get('size_unit', ''))
        image_var = tk.StringVar(value=fam.get('image', ''))

        rows = [('שם להדפסה:', name_var, 32, False), ('סוג בד:', fabric_var, 20, False), ('כמות במארז:', pack_var, 8, True)]
        for i, (lbl, var, width, is_spin) in enumerate(rows, start=2):
            tk.Label(frm, text=lbl, anchor='e').grid(row=i, column=1, sticky='e', padx=(6, 2), pady=3)
            if is_spin:
                tk.Spinbox(frm, from_=1, to=99, textvariable=var, width=width, justify='center').grid(row=i, column=0, sticky='w', pady=3)
            else:
                tk.Entry(frm, textvariable=var, width=width).grid(row=i, column=0, sticky='w', pady=3)

        unit_row = len(rows) + 2
        tk.Label(frm, text='יחידת מידה:', anchor='e').grid(row=unit_row, column=1, sticky='e', padx=(6, 2), pady=3)
        ttk.Combobox(frm, textvariable=size_unit_var, width=18, values=['', 'חודשים', 'שנים']).grid(row=unit_row, column=0, sticky='w', pady=3)

        img_row = unit_row + 1
        tk.Label(frm, text='תמונת דגם:', anchor='e').grid(row=img_row, column=1, sticky='ne', padx=(6, 2), pady=3)
        img_frame = tk.Frame(frm)
        img_frame.grid(row=img_row, column=0, sticky='w', pady=3)
        preview_lbl = tk.Label(img_frame, text='(אין תמונה)', width=14, height=6, relief='solid', bd=1, bg=theme.PAGE_BG, compound='center')
        preview_lbl.grid(row=0, column=0, columnspan=2, pady=(0, 4))
        self._family_img_ref = None

        def _abs_path(rel):
            rel = (rel or '').strip()
            if not rel:
                return ''
            return rel if os.path.isabs(rel) else os.path.join(os.getcwd(), rel)

        def _update_preview():
            abs_path = _abs_path(image_var.get())
            if abs_path and os.path.exists(abs_path):
                try:
                    from PIL import Image, ImageTk
                    im = Image.open(abs_path)
                    im.thumbnail((96, 96))
                    self._family_img_ref = ImageTk.PhotoImage(im)
                    preview_lbl.config(image=self._family_img_ref, text='')
                    return
                except Exception:
                    pass
            self._family_img_ref = None
            preview_lbl.config(image='', text='(אין תמונה)')

        def _choose_image():
            products_dir = os.path.join(os.getcwd(), 'assets', 'labels', 'products')
            os.makedirs(products_dir, exist_ok=True)
            path = filedialog.askopenfilename(title='בחר תמונת דגם', initialdir=products_dir,
                                              filetypes=[('תמונות', '*.png *.jpg *.jpeg *.gif *.bmp'), ('כל הקבצים', '*.*')])
            if not path:
                return
            try:
                ext = os.path.splitext(path)[1].lower() or '.png'
                safe = ''.join(ch for ch in f"{brand}_{code}" if ch.isalnum() or ch == '_') or 'family'
                dest = os.path.join(products_dir, f"family_{safe}{ext}")
                shutil.copyfile(path, dest)
                image_var.set(os.path.relpath(dest, os.getcwd()))
                _update_preview()
            except Exception as e:
                messagebox.showerror('שגיאה', f'שגיאה בשמירת התמונה:\n{e}', parent=dlg)

        def _remove_image():
            image_var.set('')
            _update_preview()

        tk.Button(img_frame, text='בחר תמונה…', command=_choose_image).grid(row=1, column=0, padx=(0, 4))
        tk.Button(img_frame, text='הסר', command=_remove_image).grid(row=1, column=1)
        _update_preview()

        def save():
            self.data_processor.set_rivhit_family(code, {
                'print_name': name_var.get(),
                'fabric': fabric_var.get(),
                'pack_qty': pack_var.get(),
                'size_unit': size_unit_var.get(),
                'image': image_var.get(),
            }, brand=brand)
            dlg.destroy()
            messagebox.showinfo("נשמר", f"הדגם עודכן והופץ ל-{self.data_processor.rivhit_family_variant_count(code, brand)} וואריאנטים",
                                parent=parent or self.notebook)
            if callable(on_saved):
                on_saved()

        btns = tk.Frame(frm)
        btns.grid(row=img_row + 1, column=0, columnspan=2, pady=(12, 0))
        tk.Button(btns, text="שמור והפץ", command=save, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=5)
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
        # עטיפה נגללת - התוכן ארוך מגובה המסך (טופס + מידות + צבעים + מדבקה + ייצוא)
        outer = tk.Frame(tab, bg=theme.PAGE_BG)
        outer.pack(fill='both', expand=True)
        canvas = tk.Canvas(outer, bg=theme.PAGE_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=theme.PAGE_BG)
        scroll_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        _win_id = canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        def _configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
            canvas.itemconfig(_win_id, width=event.width)
        canvas.bind('<Configure>', _configure_scroll_region)

        # גלגלת עכבר גוללת רק כשהסמן מעל הלשונית הזו
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        canvas.bind('<Enter>', lambda e: canvas.bind_all('<MouseWheel>', _on_mousewheel))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all('<MouseWheel>'))

        tab = scroll_frame

        tk.Label(tab, text="הוספת מוצרים וייצוא לריווחית", font=(theme.FONT_FAMILY, 16, 'bold'), bg=theme.PAGE_BG, fg=theme.DARK).pack(pady=8)

        # Add form
        form = tk.LabelFrame(tab, text="פרטי מוצר חדש", bg=theme.PAGE_BG, fg=theme.DARK, font=(theme.FONT_FAMILY, 10, 'bold'), padx=10, pady=10)
        form.pack(fill='x', padx=15, pady=5)

        self.rivhit_new_name_var = tk.StringVar()
        self.rivhit_new_part_var = tk.StringVar()
        self.rivhit_new_cost_var = tk.StringVar()
        self.rivhit_new_sale_var = tk.StringVar()
        self.rivhit_new_digital_var = tk.StringVar()
        self.rivhit_new_cat_var = tk.StringVar()
        self.rivhit_new_last_item_var = tk.StringVar()

        # Row 1
        r1 = tk.Frame(form, bg=theme.PAGE_BG); r1.pack(fill='x', pady=3)
        tk.Label(r1, text='שם הפריט:', bg=theme.PAGE_BG, width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r1, textvariable=self.rivhit_new_name_var, width=40).pack(side='right', padx=(0, 12))
        tk.Label(r1, text='מק"ט / ברקוד:', bg=theme.PAGE_BG, width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r1, textvariable=self.rivhit_new_part_var, width=22).pack(side='right', padx=(0, 4))
        tk.Button(r1, text="חולל ברקוד", command=self._generate_rivhit_barcode, bg=theme.PURPLE, fg='white', font=(theme.FONT_FAMILY, 9, 'bold')).pack(side='right', padx=(0, 12))

        # Row 2
        r2 = tk.Frame(form, bg=theme.PAGE_BG); r2.pack(fill='x', pady=3)
        tk.Label(r2, text='עלות (₪):', bg=theme.PAGE_BG, width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r2, textvariable=self.rivhit_new_cost_var, width=12).pack(side='right', padx=(0, 12))
        tk.Label(r2, text='מחיר מכירה (₪):', bg=theme.PAGE_BG, width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r2, textvariable=self.rivhit_new_sale_var, width=12).pack(side='right', padx=(0, 12))
        tk.Label(r2, text='מחיר לצרכן דיגיטלי (₪):', bg=theme.PAGE_BG, width=18, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r2, textvariable=self.rivhit_new_digital_var, width=12).pack(side='right', padx=(0, 12))
        tk.Label(r2, text='עונה / קטגוריה:', bg=theme.PAGE_BG, width=14, anchor='e').pack(side='right', padx=(6, 2))
        groups = sorted((getattr(self.data_processor, 'rivhit_groups', {}) or {}).keys())
        self.rivhit_new_cat_combo = ttk.Combobox(r2, textvariable=self.rivhit_new_cat_var, values=groups, state='readonly', width=20)
        self.rivhit_new_cat_combo.pack(side='right', padx=(0, 12))

        # Row 3 - starting item number control
        r3 = tk.Frame(form, bg=theme.PAGE_BG); r3.pack(fill='x', pady=3)
        tk.Label(r3, text='מספר פריט אחרון:', bg=theme.PAGE_BG, width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(r3, textvariable=self.rivhit_new_last_item_var, width=12).pack(side='right', padx=(0, 6))
        tk.Label(r3, text='(המספור יתחיל ממספר זה +1. ריק = לפי רשימת המוצרים העדכנית)', bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 8)).pack(side='right', padx=(0, 12))

        # Sizes multi-select for batch creation
        sizes_box = tk.LabelFrame(form, text="יצירת מוצרים לפי מידות (מוצר לכל מידה, עם ברקוד חדש)", bg=theme.PAGE_BG, fg=theme.DARK, font=(theme.FONT_FAMILY, 9, 'bold'), padx=8, pady=6)
        sizes_box.pack(fill='x', pady=(8, 2))
        sizes_actions = tk.Frame(sizes_box, bg=theme.PAGE_BG); sizes_actions.pack(fill='x', anchor='e')
        tk.Button(sizes_actions, text="נקה בחירה", command=self._clear_rivhit_size_selection, bg=theme.MUTED, fg='white', font=(theme.FONT_FAMILY, 8)).pack(side='right', padx=3)
        tk.Button(sizes_actions, text="סמן הכל", command=self._select_all_rivhit_sizes, bg=theme.PRIMARY, fg='white', font=(theme.FONT_FAMILY, 8)).pack(side='right', padx=3)
        self.rivhit_sizes_grid = tk.Frame(sizes_box, bg=theme.PAGE_BG); self.rivhit_sizes_grid.pack(fill='x', pady=(4, 0))
        self.rivhit_size_vars = {}
        self._build_rivhit_sizes_checkboxes()

        # Colors multi-select for batch creation (מוצר לכל צבע)
        colors_box = tk.LabelFrame(form, text="יצירת מוצרים לפי צבעים (מוצר לכל צבע, עם ברקוד חדש)", bg=theme.PAGE_BG, fg=theme.DARK, font=(theme.FONT_FAMILY, 9, 'bold'), padx=8, pady=6)
        colors_box.pack(fill='x', pady=(8, 2))
        colors_actions = tk.Frame(colors_box, bg=theme.PAGE_BG); colors_actions.pack(fill='x', anchor='e')
        tk.Button(colors_actions, text="נקה בחירה", command=self._clear_rivhit_color_selection, bg=theme.MUTED, fg='white', font=(theme.FONT_FAMILY, 8)).pack(side='right', padx=3)
        tk.Button(colors_actions, text="סמן הכל", command=self._select_all_rivhit_colors, bg=theme.PRIMARY, fg='white', font=(theme.FONT_FAMILY, 8)).pack(side='right', padx=3)
        tk.Button(colors_actions, text="🎨 נהל צבעים", command=self._open_fabric_color_manager, bg=theme.PURPLE, fg='white', font=(theme.FONT_FAMILY, 8, 'bold')).pack(side='right', padx=3)
        self.rivhit_colors_grid = tk.Frame(colors_box, bg=theme.PAGE_BG); self.rivhit_colors_grid.pack(fill='x', pady=(4, 0))
        self.rivhit_color_vars = {}
        self._build_rivhit_colors_checkboxes()

        # Print (label) details - applied to the created product(s)
        self.rivhit_print_name_var = tk.StringVar()
        self.rivhit_print_size_var = tk.StringVar()
        self.rivhit_print_unit_var = tk.StringVar(value='חודשים')
        self.rivhit_print_fabric_var = tk.StringVar()
        self.rivhit_print_pack_var = tk.StringVar(value='3')
        self.rivhit_print_image_src_var = tk.StringVar()
        self.rivhit_print_image_label_var = tk.StringVar(value='ללא תמונה')
        self.rivhit_print_model_code_var = tk.StringVar()

        print_box = tk.LabelFrame(form, text="פרטי הדפסה למדבקה (יחולו על כל המוצרים שייווצרו)", bg=theme.PAGE_BG, fg=theme.DARK, font=(theme.FONT_FAMILY, 9, 'bold'), padx=8, pady=6)
        print_box.pack(fill='x', pady=(8, 2))
        pr1 = tk.Frame(print_box, bg=theme.PAGE_BG); pr1.pack(fill='x', pady=2)
        tk.Label(pr1, text='שם להדפסה:', bg=theme.PAGE_BG, width=12, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(pr1, textvariable=self.rivhit_print_name_var, width=26).pack(side='right', padx=(0, 12))
        tk.Label(pr1, text='סוג בד:', bg=theme.PAGE_BG, width=8, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(pr1, textvariable=self.rivhit_print_fabric_var, width=14).pack(side='right', padx=(0, 12))
        tk.Label(pr1, text='כמות במארז:', bg=theme.PAGE_BG, width=10, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(pr1, textvariable=self.rivhit_print_pack_var, width=6).pack(side='right', padx=(0, 12))
        pr2 = tk.Frame(print_box, bg=theme.PAGE_BG); pr2.pack(fill='x', pady=2)
        tk.Label(pr2, text='מידה (למוצר בודד):', bg=theme.PAGE_BG, width=14, anchor='e').pack(side='right', padx=(6, 2))
        tk.Entry(pr2, textvariable=self.rivhit_print_size_var, width=10).pack(side='right', padx=(0, 12))
        tk.Label(pr2, text='יחידת מידה:', bg=theme.PAGE_BG, width=10, anchor='e').pack(side='right', padx=(6, 2))
        ttk.Combobox(pr2, textvariable=self.rivhit_print_unit_var, values=['', 'חודשים', 'שנים'], state='readonly', width=10).pack(side='right', padx=(0, 12))
        tk.Button(pr2, text='בחר תמונה…', command=self._choose_rivhit_print_image, bg=theme.PURPLE, fg='white', font=(theme.FONT_FAMILY, 8)).pack(side='right', padx=(0, 6))
        tk.Label(pr2, textvariable=self.rivhit_print_image_label_var, bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 8)).pack(side='right', padx=(0, 12))
        pr3 = tk.Frame(print_box, bg=theme.PAGE_BG); pr3.pack(fill='x', pady=2)
        tk.Label(pr3, text='קוד דגם (באנגלית):', bg=theme.PAGE_BG, width=14, anchor='e').pack(side='right', padx=(6, 2))
        family_codes = self.data_processor.list_rivhit_family_codes() if hasattr(self.data_processor, 'list_rivhit_family_codes') else []
        self.rivhit_print_model_code_combo = ttk.Combobox(pr3, textvariable=self.rivhit_print_model_code_var, values=family_codes, width=16)
        self.rivhit_print_model_code_combo.pack(side='right', padx=(0, 6))
        self.rivhit_print_model_code_combo.bind('<<ComboboxSelected>>', lambda e: self._on_rivhit_family_code_selected())
        tk.Label(pr3, text="(קוד דגם משותף לכל הוואריאנטים; בחירת קוד קיים תטען את פרטי הדגם)", bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 8)).pack(side='right', padx=(0, 12))
        tk.Label(print_box, text="(ביצירה לפי מידות - המידה נלקחת מכל מידה שנבחרה; שדה 'מידה' משמש להוספת מוצר בודד)", bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 8)).pack(anchor='e', pady=(2, 0))

        btns = tk.Frame(form, bg=theme.PAGE_BG); btns.pack(pady=(8, 2))
        tk.Button(btns, text="➕ הוסף מוצר", command=self._add_rivhit_product, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=4)
        tk.Button(btns, text="🧩 צור מוצרים לפי מידות", command=self._create_rivhit_products_by_sizes, bg=theme.TEAL, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=4)
        tk.Button(btns, text="🎨 צור מוצרים לפי צבעים", command=self._create_rivhit_products_by_colors, bg=theme.PURPLE, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=4)

        # Pending list toolbar
        toolbar = tk.Frame(tab, bg=theme.PAGE_BG)
        toolbar.pack(fill='x', padx=15, pady=(8, 2))
        tk.Label(toolbar, text="מוצרים ממתינים לייצוא:", bg=theme.PAGE_BG, font=(theme.FONT_FAMILY, 11, 'bold')).pack(side='right')
        tk.Button(toolbar, text="⬇️ ייצא קובץ לריווחית", command=self._export_rivhit_new, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=4)
        tk.Button(toolbar, text="🔗 עדכן ריווחית (API)", command=self._update_rivhit_via_api, bg=theme.TEAL, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=4)
        tk.Button(toolbar, text="🌐 ייצוא לאתר", command=self._export_rivhit_to_website, bg=theme.PRIMARY, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=4)
        tk.Button(toolbar, text="🗑️ נקה הכל", command=self._clear_rivhit_new, bg=theme.DANGER, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=4)
        tk.Button(toolbar, text="מחק נבחר", command=self._delete_rivhit_new_selected, bg=theme.MUTED, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=4)

        # Pending list table
        table_frame = tk.Frame(tab, bg=theme.CARD_BG)
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
        tk.Label(tab, textvariable=self.rivhit_new_summary_var, bg=theme.DARK, fg='white', anchor='w', padx=12, font=(theme.FONT_FAMILY, 10)).pack(fill='x', side='bottom')

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
        products_dir = os.path.join(os.getcwd(), 'assets', 'labels', 'products')
        os.makedirs(products_dir, exist_ok=True)
        path = filedialog.askopenfilename(
            title='בחר תמונת מוצר למדבקה',
            initialdir=products_dir,
            filetypes=[('תמונות', '*.png *.jpg *.jpeg *.gif *.bmp'), ('כל הקבצים', '*.*')],
        )
        if not path:
            return
        self.rivhit_print_image_src_var.set(path)
        self.rivhit_print_image_label_var.set(os.path.basename(path))

    def _current_add_brand(self):
        """מזהה מותג לפי הקטגוריה שנבחרה בטופס ההוספה."""
        return self.data_processor.brand_key_from_category(self.rivhit_new_cat_var.get())

    def _on_rivhit_family_code_selected(self):
        """בחירת קוד דגם קיים - טעינת פרטי ההדפסה של הדגם לשדות הטופס (לפי מותג הקטגוריה)."""
        code = (self.rivhit_print_model_code_var.get() or '').strip()
        if not code:
            return
        fam = self.data_processor.get_rivhit_family(code, brand=self._current_add_brand())
        if not fam:
            return
        self.rivhit_print_name_var.set(fam.get('print_name', ''))
        self.rivhit_print_fabric_var.set(fam.get('fabric', ''))
        self.rivhit_print_pack_var.set(str(fam.get('pack_qty', '') or ''))
        self.rivhit_print_unit_var.set(fam.get('size_unit', ''))
        img = fam.get('image', '')
        if img:
            self.rivhit_print_image_label_var.set(os.path.basename(img) + ' (מהדגם)')
        else:
            self.rivhit_print_image_label_var.set('ללא תמונה')

    def _apply_rivhit_print_fields(self, barcode, size='', color=''):
        """שומר פרטי הדפסה (מדבקה) למוצר שנוצר לפי הברקוד שלו.

        אם הוזן קוד דגם - התמונה נשמרת פעם אחת ברמת הדגם (family_<code>) והשדות
        המשותפים מופצים לכל הוואריאנטים באותו קוד דגם.
        """
        barcode = str(barcode or '').strip()
        if not barcode:
            return
        print_name = (self.rivhit_print_name_var.get() or '').strip()
        fabric = (self.rivhit_print_fabric_var.get() or '').strip()
        size_unit = (self.rivhit_print_unit_var.get() or '').strip()
        pack = (self.rivhit_print_pack_var.get() or '').strip()
        img_src = (self.rivhit_print_image_src_var.get() or '').strip()
        model_code = (self.rivhit_print_model_code_var.get() or '').strip()
        brand = self._current_add_brand()
        size = str(size or '').strip()
        color = str(color or '').strip()
        # החל רק אם המשתמש הזין פרט הדפסה כלשהו
        if not (print_name or fabric or img_src or size or size_unit or model_code or color):
            return
        # תמונה: אם יש קוד דגם - שם קובץ לפי המותג+הדגם; אחרת לפי הברקוד
        image_rel = ''
        if img_src and os.path.exists(img_src):
            try:
                ext = os.path.splitext(img_src)[1].lower() or '.png'
                dest_dir = os.path.join(os.getcwd(), 'assets', 'labels', 'products')
                os.makedirs(dest_dir, exist_ok=True)
                if model_code:
                    safe = ''.join(ch for ch in f"{brand}_{model_code}" if ch.isalnum() or ch == '_') or 'family'
                    fname = f"family_{safe}{ext}"
                else:
                    fname = (''.join(ch for ch in barcode if ch.isalnum()) or 'product') + ext
                dest = os.path.join(dest_dir, fname)
                shutil.copyfile(img_src, dest)
                image_rel = os.path.relpath(dest, os.getcwd())
            except Exception:
                image_rel = ''
        # אם נבחר קוד דגם קיים בלי תמונה חדשה - השתמש בתמונת הדגם הקיימת
        if not image_rel and model_code:
            fam = self.data_processor.get_rivhit_family(model_code, brand=brand)
            image_rel = fam.get('image', '') if fam else ''
        self.data_processor.set_rivhit_label_fields(barcode, {
            'print_name': print_name,
            'size': size,
            'size_unit': size_unit,
            'fabric': fabric,
            'pack_qty': pack or 1,
            'image': image_rel,
            'model_code': model_code,
            'brand': brand,
            'color': color,
        })
        # יצירה/עדכון של דגם האב והפצה לכל הוואריאנטים (מותג+קוד)
        if model_code:
            self.data_processor.set_rivhit_family(model_code, {
                'print_name': print_name,
                'fabric': fabric,
                'pack_qty': pack or 1,
                'size_unit': size_unit,
                'image': image_rel,
            }, brand=brand)

    def _reset_rivhit_print_fields(self):
        self.rivhit_print_name_var.set('')
        self.rivhit_print_size_var.set('')
        self.rivhit_print_fabric_var.set('')
        self.rivhit_print_model_code_var.set('')
        self.rivhit_print_image_src_var.set('')
        self.rivhit_print_image_label_var.set('ללא תמונה')
        if hasattr(self, 'rivhit_print_model_code_combo'):
            try:
                self.rivhit_print_model_code_combo['values'] = self.data_processor.list_rivhit_family_codes()
            except Exception:
                pass

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
            cb = tk.Checkbutton(self.rivhit_sizes_grid, text=size, variable=var, bg=theme.PAGE_BG, anchor='w')
            cb.grid(row=i // cols, column=i % cols, sticky='w', padx=4, pady=1)

    def _select_all_rivhit_sizes(self):
        for var in self.rivhit_size_vars.values():
            var.set(True)

    def _clear_rivhit_size_selection(self):
        for var in self.rivhit_size_vars.values():
            var.set(False)

    # ===== יצירת מוצרים לפי צבעים =====
    @staticmethod
    def _valid_hex(value):
        """מחזיר hex תקין (#rrggbb) או ריק."""
        v = str(value or '').strip().lower()
        if len(v) == 7 and v.startswith('#') and all(ch in '0123456789abcdef' for ch in v[1:]):
            return v
        return ''

    @staticmethod
    def _contrast_fg(hex_value):
        """צבע טקסט (שחור/לבן) קריא על רקע נתון."""
        try:
            r, g, b = int(hex_value[1:3], 16), int(hex_value[3:5], 16), int(hex_value[5:7], 16)
            luminance = 0.299 * r + 0.587 * g + 0.114 * b
            return '#000000' if luminance > 140 else '#ffffff'
        except Exception:
            return '#000000'

    # גודל אחיד לדוגמית בד — כמו ריבוע קוד הצבע במניפה באתר (~84px), לא תמונה גדולה
    FABRIC_SWATCH_SIZE = 96

    @classmethod
    def _normalize_fabric_image(cls, img):
        """מנרמל תמונת בד לדוגמית אחידה: ריבוע FABRIC_SWATCH_SIZE x FABRIC_SWATCH_SIZE.

        חיתוך מרכזי לריבוע; תמונה גדולה מוקטנת באיכות גבוהה (LANCZOS),
        ותמונה קטנה (למשל חיתוך מסך קטן) מרוצפת בשיקוף כדי לשמור על
        חדות הטקסטורה במקום הגדלה שמטשטשת.
        """
        from PIL import Image
        size = cls.FABRIC_SWATCH_SIZE
        img = img.convert('RGB')
        w, h = img.size
        side = min(w, h)
        left, top = (w - side) // 2, (h - side) // 2
        square = img.crop((left, top, left + side, top + side))
        if side >= size:
            return square.resize((size, size), Image.LANCZOS)
        canvas = Image.new('RGB', (size, size))
        tiles = -(-size // side)  # עיגול כלפי מעלה
        for row in range(tiles):
            for col in range(tiles):
                tile = square
                if col % 2 == 1:
                    tile = tile.transpose(Image.FLIP_LEFT_RIGHT)
                if row % 2 == 1:
                    tile = tile.transpose(Image.FLIP_TOP_BOTTOM)
                canvas.paste(tile, (col * side, row * side))
        return canvas

    def _build_rivhit_colors_checkboxes(self):
        for w in self.rivhit_colors_grid.winfo_children():
            w.destroy()
        self.rivhit_color_vars = {}
        palette = list(getattr(self.data_processor, 'fabric_colors_palette', []) or [])
        if not palette:
            tk.Label(self.rivhit_colors_grid, text="אין צבעים בפלטה - לחץ על 'נהל צבעים' כדי להוסיף צבע מתמונת בד",
                     bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9)).grid(row=0, column=0, sticky='e', padx=4, pady=2)
            return
        cols = 6
        idx = 0
        for entry in palette:
            name = str((entry or {}).get('name', '')).strip()
            if not name:
                continue
            hex_val = self._valid_hex((entry or {}).get('hex', '')) or '#cccccc'
            cell = tk.Frame(self.rivhit_colors_grid, bg=theme.PAGE_BG)
            cell.grid(row=idx // cols, column=idx % cols, sticky='w', padx=4, pady=1)
            var = tk.BooleanVar(value=False)
            self.rivhit_color_vars[name] = var
            tk.Label(cell, width=2, bg=hex_val, relief='solid', bd=1).pack(side='right', padx=(0, 3))
            tk.Checkbutton(cell, text=name, variable=var, bg=theme.PAGE_BG, anchor='w').pack(side='right')
            idx += 1

    def _select_all_rivhit_colors(self):
        for var in self.rivhit_color_vars.values():
            var.set(True)

    def _clear_rivhit_color_selection(self):
        for var in self.rivhit_color_vars.values():
            var.set(False)

    def _create_rivhit_products_by_colors(self):
        name = (self.rivhit_new_name_var.get() or '').strip()
        if not name:
            messagebox.showwarning("שדה חסר", "יש להזין שם פריט")
            return
        selected_names = [c for c, var in self.rivhit_color_vars.items() if var.get()]
        if not selected_names:
            messagebox.showwarning("לא נבחרו צבעים", "יש לבחור לפחות צבע אחד")
            return
        palette = {str((e or {}).get('name', '')).strip(): (e or {}) for e in (getattr(self.data_processor, 'fabric_colors_palette', []) or [])}
        selected = [{'name': c, 'hex': palette.get(c, {}).get('hex', '')} for c in selected_names]
        try:
            created = self.data_processor.add_rivhit_new_products_by_colors(
                base_name=name,
                colors=selected,
                cost_nis=self.rivhit_new_cost_var.get(),
                sale_nis=self.rivhit_new_sale_var.get(),
                category=self.rivhit_new_cat_var.get(),
                digital_price=self.rivhit_new_digital_var.get(),
                last_item_num=(self.rivhit_new_last_item_var.get() or '').strip() or None,
            )
            # החל פרטי הדפסה לכל מוצר שנוצר, עם הצבע התואם שלו (מידה - מהשדה הבודד אם הוזנה)
            single_size = (self.rivhit_print_size_var.get() or '').strip()
            for rec, color in zip(created, selected):
                self._apply_rivhit_print_fields(rec.get('item_part_num', ''), size=single_size, color=color['name'])
            self.rivhit_new_name_var.set('')
            self.rivhit_new_part_var.set('')
            self.rivhit_new_cost_var.set('')
            self.rivhit_new_sale_var.set('')
            self.rivhit_new_digital_var.set('')
            self.rivhit_new_cat_var.set('')
            self._reset_rivhit_print_fields()
            self._clear_rivhit_color_selection()
            self._refresh_rivhit_new_table()
            messagebox.showinfo("הצלחה", f"נוצרו {len(created)} מוצרים לפי הצבעים שנבחרו")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    # ===== ניהול פלטת צבעי בד =====
    def _open_fabric_color_manager(self):
        """דיאלוג לניהול פלטת צבעי הבד: רשימה, הוספה מתמונה/ידנית, מחיקה."""
        dlg = tk.Toplevel(self.notebook)
        dlg.title("ניהול צבעי בד")
        dlg.geometry("720x440")
        dlg.grab_set()

        tk.Label(dlg, text="פלטת צבעי בד - כל צבע משמש ליצירת מוצר עם ברקוד משלו (דאבל-קליק לעריכה)",
                 font=(theme.FONT_FAMILY, 11, 'bold')).pack(pady=(10, 4))

        table_frame = tk.Frame(dlg)
        table_frame.pack(fill='both', expand=True, padx=12, pady=6)
        tree = ttk.Treeview(table_frame, columns=('name', 'hex', 'supplier', 'sampled_at', 'image'), show='headings')
        tree.heading('name', text='שם הצבע')
        tree.heading('hex', text='קוד צבע')
        tree.heading('supplier', text='ספק')
        tree.heading('sampled_at', text='תאריך דגימה')
        tree.heading('image', text='תמונה')
        tree.column('name', width=190, anchor='center')
        tree.column('hex', width=95, anchor='center')
        tree.column('supplier', width=140, anchor='center')
        tree.column('sampled_at', width=105, anchor='center')
        tree.column('image', width=60, anchor='center')
        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        tree.configure(yscroll=vsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        def refresh():
            for it in tree.get_children():
                tree.delete(it)
            for entry in (getattr(self.data_processor, 'fabric_colors_palette', []) or []):
                name = str((entry or {}).get('name', '')).strip()
                if not name:
                    continue
                hex_val = self._valid_hex((entry or {}).get('hex', '')) or '#cccccc'
                tag = f"c_{hex_val[1:]}"
                tree.tag_configure(tag, background=hex_val, foreground=self._contrast_fg(hex_val))
                tree.insert('', 'end', values=(
                    name, hex_val,
                    str((entry or {}).get('supplier', '')),
                    str((entry or {}).get('sampled_at', '')),
                    '✓' if str((entry or {}).get('image', '')).strip() else '',
                ), tags=(tag,))
            # רענון הצ'קבוקסים בטופס ההוספה
            self._build_rivhit_colors_checkboxes()

        def add_from_image():
            self._add_fabric_color_from_image(parent=dlg, on_saved=refresh)

        def add_from_screen():
            self._add_fabric_color_from_screen(parent=dlg, on_saved=refresh)

        def add_manual():
            from tkinter import colorchooser
            picked = colorchooser.askcolor(title='בחר צבע', parent=dlg)
            if not picked or not picked[1]:
                return
            self._prompt_save_picked_color(picked[1], parent=dlg, on_saved=refresh)

        def paste_from_clipboard(event=None):
            img = self._grab_clipboard_image(dlg)
            if img is None:
                messagebox.showinfo("אין תמונה בלוח",
                                    "העתק תמונה (Print Screen או Win+Shift+S לחיתוך אזור) ואז לחץ הדבק",
                                    parent=dlg)
                return
            self._add_fabric_color_from_image(parent=dlg, on_saved=refresh, image=img)

        def edit_selected(event=None):
            sel = tree.selection()
            if not sel:
                if event is None:
                    messagebox.showinfo("לא נבחר", "יש לבחור צבע לעריכה", parent=dlg)
                return
            vals = tree.item(sel[0], 'values')
            name, hex_val = str(vals[0]), str(vals[1])
            self._prompt_save_picked_color(hex_val, parent=dlg, on_saved=refresh, edit_name=name)

        def delete_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("לא נבחר", "יש לבחור צבע למחיקה", parent=dlg)
                return
            name = str(tree.item(sel[0], 'values')[0])
            if messagebox.askyesno("מחיקה", f"למחוק את הצבע '{name}' מהפלטה?", parent=dlg):
                self.data_processor.delete_fabric_palette_color(name)
                refresh()

        tree.bind('<Double-1>', edit_selected)

        btns = tk.Frame(dlg)
        btns.pack(fill='x', padx=12, pady=(0, 10))
        tk.Button(btns, text="📷 הוסף צבע מתמונה", command=add_from_image, bg=theme.PURPLE, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        tk.Button(btns, text="📋 הדבק תמונה מהלוח", command=paste_from_clipboard, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        tk.Button(btns, text="🖥️ דגום צבע מהמסך", command=add_from_screen, bg=theme.PRIMARY, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        tk.Button(btns, text="🎨 הוסף צבע ידני", command=add_manual, bg=theme.TEAL, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        tk.Button(btns, text="🗑️ מחק נבחר", command=delete_selected, bg=theme.DANGER, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        tk.Button(btns, text="סגור", command=dlg.destroy).pack(side='left', padx=5)
        # keycode 86 = מקש V בכל פריסת מקלדת (גם בעברית)
        dlg.bind('<Control-KeyPress>', lambda e: paste_from_clipboard() if e.keycode == 86 else None)

        refresh()

    def _add_fabric_color_from_screen(self, parent=None, on_saved=None):
        """פיקר צבע מהמסך: המסך מוקפא לתצוגה מלאה, קוד הצבע מוצג ליד הסמן, ולחיצה בוחרת."""
        try:
            from PIL import ImageGrab, ImageTk
        except ImportError:
            messagebox.showerror("חסרה ספרייה", "נדרשת הספרייה Pillow (PIL) לדגימת צבע מהמסך", parent=parent)
            return
        import time

        root = self.notebook.winfo_toplevel()
        # הסתרת חלונות האפליקציה כדי לחשוף את מה שמאחוריהם, ואז צילום המסך
        hidden = []
        for w in (parent, root):
            try:
                if w is not None and w.winfo_viewable():
                    w.withdraw()
                    hidden.append(w)
            except Exception:
                pass
        try:
            root.update_idletasks()
            root.update()
            time.sleep(0.3)  # מתן זמן למערכת ההפעלה לצייר את מה שמאחורי החלונות
            shot = ImageGrab.grab().convert('RGB')
        finally:
            for w in hidden:
                try:
                    w.deiconify()
                except Exception:
                    pass

        overlay = tk.Toplevel(root)
        overlay.attributes('-fullscreen', True)
        overlay.attributes('-topmost', True)
        overlay.update_idletasks()
        sw = max(overlay.winfo_screenwidth(), 1)
        sh = max(overlay.winfo_screenheight(), 1)
        # התאמת קנה מידה בין קואורדינטות tkinter לפיקסלים בצילום (DPI scaling)
        scale_x = shot.width / sw
        scale_y = shot.height / sh
        disp = shot if shot.size == (sw, sh) else shot.resize((sw, sh))
        photo = ImageTk.PhotoImage(disp)
        canvas = tk.Canvas(overlay, width=sw, height=sh, highlightthickness=0, cursor='crosshair')
        canvas.pack(fill='both', expand=True)
        canvas.create_image(0, 0, anchor='nw', image=photo)
        canvas._photo_ref = photo

        # פס הנחיה עליון
        canvas.create_rectangle(0, 0, sw, 34, fill='#1f2933', outline='')
        canvas.create_text(sw // 2, 17, text="לחץ בכל מקום על המסך כדי לדגום צבע | Esc לביטול",
                           fill='white', font=(theme.FONT_FAMILY, 11, 'bold'))

        # תווית צפה ליד הסמן: ריבוע צבע + קוד hex
        info_rect = canvas.create_rectangle(0, 0, 0, 0, fill='#ffffff', outline='#333333', state='hidden')
        info_swatch = canvas.create_rectangle(0, 0, 0, 0, fill='#ffffff', outline='#333333', state='hidden')
        info_text = canvas.create_text(0, 0, text='', anchor='w', font=('Consolas', 11, 'bold'), state='hidden')

        def _hex_at(x, y):
            px = min(max(int(x * scale_x), 0), shot.width - 1)
            py = min(max(int(y * scale_y), 0), shot.height - 1)
            r, g, b = shot.getpixel((px, py))
            return '#{:02x}{:02x}{:02x}'.format(r, g, b)

        def on_motion(event):
            hex_val = _hex_at(event.x, event.y)
            # מיקום התווית ליד הסמן, עם היפוך ליד קצוות המסך
            bx = event.x + 18 if event.x < sw - 150 else event.x - 148
            by = event.y + 18 if event.y < sh - 60 else event.y - 46
            canvas.coords(info_rect, bx, by, bx + 130, by + 28)
            canvas.coords(info_swatch, bx + 6, by + 5, bx + 24, by + 23)
            canvas.coords(info_text, bx + 32, by + 14)
            canvas.itemconfig(info_swatch, fill=hex_val)
            canvas.itemconfig(info_text, text=hex_val)
            for item in (info_rect, info_swatch, info_text):
                canvas.itemconfig(item, state='normal')
                canvas.tag_raise(item)

        def on_click(event):
            hex_val = _hex_at(event.x, event.y)
            overlay.destroy()
            self._prompt_save_picked_color(hex_val, parent=parent, on_saved=on_saved)

        def _on_escape(e):
            overlay.destroy()
            # החזרת ה-grab לדיאלוג ניהול הצבעים (אם קיים) אחרי ביטול
            if parent is not None:
                try:
                    parent.grab_set()
                except Exception:
                    pass

        canvas.bind('<Motion>', on_motion)
        canvas.bind('<Button-1>', on_click)
        overlay.bind('<Escape>', _on_escape)
        # העברת ה-grab מהדיאלוג (המוסתר בזמן הצילום) ל-overlay - בלי זה כל הקלט
        # ממשיך לזרום לדיאלוג ניהול הצבעים וה-overlay לא מקבל עכבר/מקלדת כלל
        overlay.grab_set()
        overlay.focus_force()

    def _prompt_save_picked_color(self, hex_val, parent=None, on_saved=None, edit_name=None):
        """דיאלוג שמירה/עריכה של צבע: קוד (עם העתקה), שם, ספק, תאריך דגימה ותמונת בד.

        edit_name: שם צבע קיים לעריכה - הפרטים נטענים ממנו והשמירה מעדכנת אותו.
        תמונת בד אפשר להדביק מהלוח (Ctrl+V), לבחור מקובץ, ולדגום ממנה מחדש את קוד הצבע.
        """
        from datetime import datetime
        existing = {}
        if edit_name:
            existing = next((e for e in (getattr(self.data_processor, 'fabric_colors_palette', []) or [])
                             if str((e or {}).get('name', '')).strip() == edit_name), {}) or {}
        dlg = tk.Toplevel(parent or self.notebook)
        dlg.title("עריכת צבע" if edit_name else "צבע שנדגם")
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = tk.Frame(dlg, padx=15, pady=15)
        frm.pack(fill='both', expand=True)

        # קוד הצבע וה-swatch דינמיים - דגימה מחדש מהתמונה מעדכנת אותם
        hex_state = {'hex': self._valid_hex(hex_val) or '#cccccc'}
        swatch = tk.Label(frm, width=8, height=3, bg=hex_state['hex'], relief='solid', bd=1)
        swatch.grid(row=0, column=0, rowspan=2, padx=(0, 12))
        tk.Label(frm, text='קוד הצבע:', anchor='e').grid(row=0, column=2, sticky='e', padx=(6, 2))
        hex_entry = tk.Entry(frm, width=12, font=('Consolas', 12, 'bold'), justify='center')
        hex_entry.grid(row=0, column=1, sticky='w')

        def set_hex(new_hex):
            new_hex = self._valid_hex(new_hex)
            if not new_hex:
                return
            hex_state['hex'] = new_hex
            swatch.config(bg=new_hex)
            hex_entry.config(state='normal')
            hex_entry.delete(0, 'end')
            hex_entry.insert(0, new_hex)
            hex_entry.config(state='readonly')

        set_hex(hex_state['hex'])

        copied_var = tk.StringVar(value='')

        def copy_hex():
            dlg.clipboard_clear()
            dlg.clipboard_append(hex_state['hex'])
            copied_var.set('הועתק!')
            dlg.after(1500, lambda: copied_var.set(''))

        tk.Button(frm, text='📋 העתק', command=copy_hex, font=(theme.FONT_FAMILY, 9)).grid(row=1, column=1, sticky='w', pady=(4, 0))
        tk.Label(frm, textvariable=copied_var, fg=theme.SUCCESS, font=(theme.FONT_FAMILY, 9)).grid(row=1, column=2, sticky='e', pady=(4, 0))

        tk.Label(frm, text='שם הצבע:', anchor='e').grid(row=2, column=2, sticky='e', padx=(6, 2), pady=(12, 0))
        name_var = tk.StringVar(value=edit_name or '')
        name_entry = tk.Entry(frm, textvariable=name_var, width=20)
        name_entry.grid(row=2, column=0, columnspan=2, sticky='w', pady=(12, 0))
        name_entry.focus_set()

        tk.Label(frm, text='ספק:', anchor='e').grid(row=3, column=2, sticky='e', padx=(6, 2), pady=(6, 0))
        supplier_var = tk.StringVar(value=str(existing.get('supplier', '')))
        supplier_names = sorted({str((s or {}).get('business_name', '')).strip()
                                 for s in (getattr(self.data_processor, 'suppliers', []) or [])} - {''})
        ttk.Combobox(frm, textvariable=supplier_var, values=supplier_names, width=18).grid(row=3, column=0, columnspan=2, sticky='w', pady=(6, 0))

        tk.Label(frm, text='תאריך דגימה:', anchor='e').grid(row=4, column=2, sticky='e', padx=(6, 2), pady=(6, 0))
        date_var = tk.StringVar(value=str(existing.get('sampled_at', '')).strip() or datetime.now().strftime('%d.%m.%y'))
        tk.Entry(frm, textvariable=date_var, width=12, justify='center').grid(row=4, column=0, columnspan=2, sticky='w', pady=(6, 0))

        # --- תמונת בד מצורפת ---
        # pil: תמונה חדשה בזיכרון (מהלוח/מקובץ); path: תמונה קיימת בדיסק; changed: האם לשכתב בשמירה
        img_state = {'pil': None, 'path': str(existing.get('image', '')).strip(), 'changed': False}
        tk.Label(frm, text='תמונת בד:', anchor='e').grid(row=5, column=2, sticky='ne', padx=(6, 2), pady=(10, 0))
        img_frame = tk.Frame(frm)
        img_frame.grid(row=5, column=0, columnspan=2, sticky='w', pady=(10, 0))
        preview_lbl = tk.Label(img_frame, text='(אין תמונה)', width=8, height=3,
                               relief='solid', bd=1, bg=theme.PAGE_BG, compound='center')
        preview_lbl.grid(row=0, column=0, columnspan=4, pady=(0, 4))

        def _current_pil():
            """התמונה הנוכחית כ-PIL: מהזיכרון או מהדיסק. None אם אין."""
            if img_state['pil'] is not None:
                return img_state['pil']
            rel = img_state['path']
            if rel:
                abs_path = rel if os.path.isabs(rel) else os.path.join(os.getcwd(), rel)
                if os.path.exists(abs_path):
                    try:
                        from PIL import Image
                        return Image.open(abs_path).convert('RGB')
                    except Exception:
                        return None
            return None

        brightness_var = tk.DoubleVar(value=1.0)

        def _adjusted_pil():
            """התמונה הנוכחית אחרי התאמת בהירות - זה מה שמוצג ומה שנשמר."""
            img = _current_pil()
            if img is None:
                return None
            factor = float(brightness_var.get() or 1.0)
            if abs(factor - 1.0) > 0.01:
                from PIL import ImageEnhance
                img = ImageEnhance.Brightness(img).enhance(factor)
            return img

        def _update_preview():
            img = _adjusted_pil()
            if img is not None:
                try:
                    from PIL import Image, ImageTk
                    # תצוגה מקדימה באותו גודל כמו הקובץ שנשמר (דוגמית קטנה)
                    side = self.FABRIC_SWATCH_SIZE
                    shown = img.copy()
                    if shown.size != (side, side):
                        shown = shown.resize((side, side), Image.LANCZOS)
                    preview_lbl._img_ref = ImageTk.PhotoImage(shown)
                    # עם תמונה width/height נמדדים בפיקסלים - חובה לעדכן לגודל התמונה
                    preview_lbl.config(image=preview_lbl._img_ref, text='', width=side, height=side)
                    return
                except Exception:
                    pass
            preview_lbl._img_ref = None
            preview_lbl.config(image='', text='(אין תמונה)', width=8, height=3)

        def paste_image(event=None):
            img = self._grab_clipboard_image(dlg)
            if img is None:
                messagebox.showinfo("אין תמונה בלוח",
                                    "העתק תמונה (Print Screen או Win+Shift+S לחיתוך אזור) ואז לחץ הדבק",
                                    parent=dlg)
                return
            img_state['pil'] = self._normalize_fabric_image(img)
            img_state['changed'] = True
            brightness_var.set(1.0)
            bright_pct.config(text='100%')
            _update_preview()

        def choose_image_file():
            path = filedialog.askopenfilename(
                title='בחר תמונת בד',
                filetypes=[('תמונות', '*.png *.jpg *.jpeg *.gif *.bmp *.webp'), ('כל הקבצים', '*.*')],
                parent=dlg,
            )
            if not path:
                return
            try:
                from PIL import Image
                img_state['pil'] = self._normalize_fabric_image(Image.open(path))
            except Exception as e:
                messagebox.showerror("שגיאה", f"לא ניתן לפתוח את התמונה:\n{e}", parent=dlg)
                return
            img_state['changed'] = True
            brightness_var.set(1.0)
            bright_pct.config(text='100%')
            _update_preview()

        def remove_image():
            img_state['pil'] = None
            img_state['path'] = ''
            img_state['changed'] = True
            brightness_var.set(1.0)
            bright_pct.config(text='100%')
            _update_preview()

        def resample_from_image():
            img = _adjusted_pil()
            if img is None:
                messagebox.showinfo("אין תמונה", "הדבק או בחר תמונת בד קודם", parent=dlg)
                return
            self._add_fabric_color_from_image(parent=dlg, image=img.copy(), on_picked=set_hex)

        # --- מחוון בהירות: הבהרה/הכהיה של הבד אחרי הנירמול ---
        bright_row = tk.Frame(img_frame)
        bright_row.grid(row=1, column=0, columnspan=4, sticky='ew', pady=(0, 4))
        tk.Label(bright_row, text='בהירות:', font=(theme.FONT_FAMILY, 9)).pack(side='right', padx=(0, 4))
        bright_pct = tk.Label(bright_row, text='100%', width=5, font=(theme.FONT_FAMILY, 9, 'bold'))
        bright_pct.pack(side='right')

        def on_brightness(_=None):
            bright_pct.config(text=f"{int(round(float(brightness_var.get()) * 100))}%")
            if _current_pil() is not None:
                img_state['changed'] = True
            _update_preview()

        ttk.Scale(bright_row, from_=0.5, to=2.0, variable=brightness_var,
                  orient='horizontal', command=on_brightness).pack(side='right', fill='x', expand=True, padx=(4, 4))

        def reset_brightness():
            brightness_var.set(1.0)
            on_brightness()

        tk.Button(bright_row, text='אפס', command=reset_brightness, font=(theme.FONT_FAMILY, 8)).pack(side='left')

        tk.Button(img_frame, text='📋 הדבק מהלוח', command=paste_image, font=(theme.FONT_FAMILY, 9)).grid(row=2, column=0, padx=(0, 4))
        tk.Button(img_frame, text='בחר קובץ…', command=choose_image_file, font=(theme.FONT_FAMILY, 9)).grid(row=2, column=1, padx=(0, 4))
        tk.Button(img_frame, text='הסר', command=remove_image, font=(theme.FONT_FAMILY, 9)).grid(row=2, column=2, padx=(0, 4))
        tk.Button(img_frame, text='🎯 דגום צבע מהתמונה', command=resample_from_image, bg=theme.PURPLE, fg='white', font=(theme.FONT_FAMILY, 9, 'bold')).grid(row=2, column=3)
        _update_preview()
        # keycode 86 = מקש V בכל פריסת מקלדת (גם בעברית)
        dlg.bind('<Control-KeyPress>', lambda e: paste_image() if e.keycode == 86 else None)

        def _save_image_file(name):
            """כותב תמונה חדשה לדיסק לפי שם הצבע ומחזיר נתיב יחסי; מוחק קובץ ישן אם הוסר."""
            old_rel = str(existing.get('image', '')).strip()
            if not img_state['changed']:
                return img_state['path']
            new_rel = ''
            final_img = _adjusted_pil()  # כולל התאמת בהירות; עובד גם כשרק הבהירות שונתה על תמונה מהדיסק
            if final_img is not None:
                dest_dir = os.path.join(os.getcwd(), 'assets', 'fabric_colors')
                os.makedirs(dest_dir, exist_ok=True)
                safe = ''.join(ch for ch in name if ch.isalnum() or ch in ' _-').strip().replace(' ', '_') or 'color'
                dest = os.path.join(dest_dir, f"{safe}.png")
                final_img.save(dest, 'PNG')
                new_rel = os.path.relpath(dest, os.getcwd())
            # מחיקת הקובץ הישן אם הוחלף בנתיב אחר או הוסר
            if old_rel and old_rel != new_rel:
                old_abs = old_rel if os.path.isabs(old_rel) else os.path.join(os.getcwd(), old_rel)
                try:
                    if os.path.exists(old_abs):
                        os.remove(old_abs)
                except Exception:
                    pass
            return new_rel

        def save():
            name = (name_var.get() or '').strip()
            if not name:
                messagebox.showwarning("שדה חסר", "יש להזין שם צבע כדי לשמור לפלטה", parent=dlg)
                return
            try:
                image_rel = _save_image_file(name)
                if edit_name:
                    self.data_processor.update_fabric_palette_color(edit_name, {
                        'name': name,
                        'hex': hex_state['hex'],
                        'supplier': supplier_var.get(),
                        'sampled_at': date_var.get(),
                        'image': image_rel,
                    })
                else:
                    self.data_processor.add_fabric_palette_color(
                        name, hex_state['hex'], supplier=supplier_var.get(),
                        sampled_at=date_var.get(), image=image_rel)
            except Exception as e:
                messagebox.showerror("שגיאה", str(e), parent=dlg)
                return
            dlg.destroy()
            if callable(on_saved):
                on_saved()
            else:
                self._build_rivhit_colors_checkboxes()

        btns = tk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=3, pady=(14, 0))
        tk.Button(btns, text="💾 שמור לפלטה", command=save, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btns, text="סגור", command=dlg.destroy).pack(side='left', padx=5)
        dlg.bind('<Return>', lambda e: save())

    def _grab_clipboard_image(self, parent=None):
        """מחזיר תמונת PIL מהלוח (תמונה מודבקת או נתיב לקובץ תמונה), או None אם אין."""
        try:
            from PIL import Image, ImageGrab
        except ImportError:
            messagebox.showerror("חסרה ספרייה", "נדרשת הספרייה Pillow (PIL) להדבקת תמונה", parent=parent)
            return None
        try:
            data = ImageGrab.grabclipboard()
        except Exception:
            return None
        if data is None:
            return None
        # העתקת קבצים בסייר שמה בלוח רשימת נתיבים
        if isinstance(data, list):
            for p in data:
                try:
                    return Image.open(p).convert('RGB')
                except Exception:
                    continue
            return None
        try:
            return data.convert('RGB')
        except Exception:
            return None

    def _add_fabric_color_from_image(self, parent=None, on_saved=None, image=None, on_picked=None):
        """פיקר צבע מתמונת בד: לחיצה על התמונה דוגמת ממוצע אזור 5x5.

        image: תמונת PIL מוכנה (למשל מהדבקה מהלוח); אם לא נמסרה - נפתח דיאלוג בחירת קובץ.
        on_picked: callback שמקבל את ה-hex הנדגם; כשנמסר - הקנבס רק מחזיר את הצבע
        (לדוגמה לעדכון צבע קיים) במקום לפתוח דיאלוג שמירת צבע חדש.
        """
        try:
            from PIL import Image, ImageTk
        except ImportError:
            messagebox.showerror("חסרה ספרייה", "נדרשת הספרייה Pillow (PIL) לבחירת צבע מתמונה", parent=parent)
            return
        src_img = image
        if src_img is None:
            path = filedialog.askopenfilename(
                title='בחר תמונת בד',
                filetypes=[('תמונות', '*.png *.jpg *.jpeg *.gif *.bmp *.webp'), ('כל הקבצים', '*.*')],
                parent=parent,
            )
            if not path:
                return
            try:
                src_img = Image.open(path).convert('RGB')
            except Exception as e:
                messagebox.showerror("שגיאה", f"לא ניתן לפתוח את התמונה:\n{e}", parent=parent)
                return

        dlg = tk.Toplevel(parent or self.notebook)
        dlg.title("בחירת צבע מתמונת בד")
        dlg.grab_set()

        tk.Label(dlg, text="לחץ על התמונה כדי לדגום צבע (ממוצע אזור 5x5) | Ctrl+V מדביק תמונה מהלוח",
                 font=(theme.FONT_FAMILY, 10, 'bold')).pack(pady=(10, 4))

        max_w, max_h = 640, 440
        canvas = tk.Canvas(dlg, cursor='crosshair', highlightthickness=1, highlightbackground=theme.MUTED)
        canvas.pack(padx=12, pady=4)
        state = {'img': None, 'scale': 1.0}

        def load_image(img):
            """טעינת תמונה לקנבס (מקובץ או מהלוח) עם התאמת קנה מידה."""
            state['img'] = img
            state['scale'] = min(max_w / img.width, max_h / img.height, 1.0)
            disp_w = max(1, int(img.width * state['scale']))
            disp_h = max(1, int(img.height * state['scale']))
            photo = ImageTk.PhotoImage(img.resize((disp_w, disp_h)))
            canvas.delete('all')
            canvas.config(width=disp_w, height=disp_h)
            canvas.create_image(0, 0, anchor='nw', image=photo)
            canvas._photo_ref = photo  # שמירת רפרנס מפני garbage collection

        # שורת תוצאה: swatch + hex
        picked = {'hex': ''}
        result_row = tk.Frame(dlg)
        result_row.pack(fill='x', padx=12, pady=6)
        swatch = tk.Label(result_row, width=4, height=2, relief='solid', bd=1, bg=theme.PAGE_BG)
        swatch.pack(side='right', padx=(0, 8))
        hex_var = tk.StringVar(value='(טרם נדגם צבע)')
        tk.Label(result_row, textvariable=hex_var, font=(theme.FONT_FAMILY, 10)).pack(side='right', padx=(0, 16))

        def sample(event):
            img, scale = state['img'], state['scale']
            if img is None:
                return
            # המרת קואורדינטות תצוגה לקואורדינטות בתמונה המקורית
            ox = min(max(int(event.x / scale), 0), img.width - 1)
            oy = min(max(int(event.y / scale), 0), img.height - 1)
            # ממוצע אזור 5x5 לנטרול טקסטורת הבד
            r_sum = g_sum = b_sum = count = 0
            for dx in range(-2, 3):
                for dy in range(-2, 3):
                    px, py = ox + dx, oy + dy
                    if 0 <= px < img.width and 0 <= py < img.height:
                        r, g, b = img.getpixel((px, py))
                        r_sum += r; g_sum += g; b_sum += b; count += 1
            if not count:
                return
            hex_val = '#{:02x}{:02x}{:02x}'.format(r_sum // count, g_sum // count, b_sum // count)
            picked['hex'] = hex_val
            hex_var.set(hex_val)
            swatch.config(bg=hex_val)

        canvas.bind('<Button-1>', sample)
        canvas.bind('<B1-Motion>', sample)

        def paste_clipboard():
            img = self._grab_clipboard_image(dlg)
            if img is None:
                messagebox.showinfo("אין תמונה בלוח", "העתק תמונה (Print Screen או Win+Shift+S לחיתוך אזור) ונסה שוב", parent=dlg)
                return
            picked['hex'] = ''
            hex_var.set('(טרם נדגם צבע)')
            swatch.config(bg=theme.PAGE_BG)
            load_image(img)

        def save():
            if not picked['hex']:
                messagebox.showwarning("לא נדגם צבע", "יש ללחוץ על התמונה כדי לדגום צבע", parent=dlg)
                return
            hex_val = picked['hex']
            dlg.destroy()
            if callable(on_picked):
                on_picked(hex_val)
            else:
                self._prompt_save_picked_color(hex_val, parent=parent, on_saved=on_saved)

        btns = tk.Frame(dlg)
        btns.pack(pady=(4, 12))
        tk.Button(btns, text="✔ אשר צבע" if callable(on_picked) else "💾 שמור צבע", command=save, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btns, text="📋 הדבק מהלוח", command=paste_clipboard, bg=theme.PRIMARY, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btns, text="ביטול", command=dlg.destroy).pack(side='left', padx=5)
        # keycode 86 = מקש V בכל פריסת מקלדת (גם בעברית)
        dlg.bind('<Control-KeyPress>', lambda e: paste_clipboard() if e.keycode == 86 else None)

        load_image(src_img)

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

    # ===== ייצוא לאתר הקטלוג הלבן =====
    def _encode_fabric_image_b64(self, color_name):
        """תמונת ההדמיה של צבע מהפלטה כ-JPEG בקידוד base64 (לשליחה לאתר). '' אם אין."""
        entry = next((e for e in (getattr(self.data_processor, 'fabric_colors_palette', []) or [])
                      if str((e or {}).get('name', '')).strip() == color_name), None)
        rel = str((entry or {}).get('image', '')).strip()
        if not rel:
            return ''
        path = rel if os.path.isabs(rel) else os.path.join(os.getcwd(), rel)
        if not os.path.exists(path):
            return ''
        try:
            import base64
            import io
            from PIL import Image
            buf = io.BytesIO()
            Image.open(path).convert('RGB').save(buf, 'JPEG', quality=88)
            return base64.b64encode(buf.getvalue()).decode('ascii')
        except Exception:
            return ''

    def _export_rivhit_to_website(self):
        """ייצוא המוצרים הממתינים כווריאנטים למוצר קיים באתר הקטלוג הלבן.

        אם סנכרון ריווחית מופעל - אחרי שליחה מוצלחת לאתר אותם מוצרים נשלחים גם
        לריווחית (יצירת חדשים ועדכון קיימים), כך ששני היעדים נשארים מסונכרנים.
        """
        import threading
        from ..core.website_export import WebsiteClient, WebsiteExportError
        from ..core.rivhit_api import RivhitOnlineClient, RivhitApiError

        pending = list(getattr(self.data_processor, 'rivhit_new_products', []) or [])
        if not pending:
            messagebox.showinfo("אין מוצרים", "אין מוצרים ממתינים לייצוא לאתר")
            return

        dlg = tk.Toplevel(self.notebook)
        dlg.title("ייצוא לאתר הקטלוג הלבן")
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = tk.Frame(dlg, padx=15, pady=15)
        frm.pack(fill='both', expand=True)

        # --- יעד ---
        target_box = tk.LabelFrame(frm, text="יעד", padx=8, pady=6)
        target_box.pack(fill='x', pady=(0, 8))
        target_var = tk.StringVar(value='local')
        url_var = tk.StringVar(value=self.settings.get('website.local_url', 'http://127.0.0.1:8000'))
        token_var = tk.StringVar(value=self.settings.get('website.local_api_token', 'dev-local-token'))

        def _target_keys():
            """מפתחות ההגדרות (כתובת, טוקן) לפי היעד הנבחר — לכל יעד טוקן משלו."""
            if target_var.get() == 'local':
                return ('website.local_url', 'website.local_api_token', 'http://127.0.0.1:8000', 'dev-local-token')
            return ('website.prod_url', 'website.api_token', 'https://arye-textil.co.il', '')

        def on_target_change():
            url_key, token_key, url_default, token_default = _target_keys()
            url_var.set(self.settings.get(url_key, url_default))
            token_var.set(self.settings.get(token_key, token_default))

        tk.Radiobutton(target_box, text="אתר מקומי (פיתוח)", variable=target_var, value='local', command=on_target_change).pack(side='right', padx=6)
        tk.Radiobutton(target_box, text="שרת (arye-textil.co.il)", variable=target_var, value='prod', command=on_target_change).pack(side='right', padx=6)

        conn = tk.Frame(target_box); conn.pack(fill='x', pady=(6, 0))
        tk.Label(conn, text='כתובת:', anchor='e', width=8).pack(side='right')
        tk.Entry(conn, textvariable=url_var, width=36).pack(side='right', padx=(0, 8))
        tk.Label(conn, text='טוקן API:', anchor='e', width=8).pack(side='right')
        tk.Entry(conn, textvariable=token_var, width=24, show='*').pack(side='right')

        # --- בחירת מוצר ובד ---
        pick_box = tk.LabelFrame(frm, text="שיוך באתר", padx=8, pady=6)
        pick_box.pack(fill='x', pady=(0, 8))
        product_var = tk.StringVar()
        fabric_var = tk.StringVar()
        meta_state = {'products': []}

        pr = tk.Frame(pick_box); pr.pack(fill='x', pady=2)
        tk.Label(pr, text='מוצר באתר:', anchor='e', width=10).pack(side='right')
        product_combo = ttk.Combobox(pr, textvariable=product_var, state='disabled', width=42)
        product_combo.pack(side='right', padx=(0, 8))

        fr = tk.Frame(pick_box); fr.pack(fill='x', pady=2)
        tk.Label(fr, text='סוג בד:', anchor='e', width=10).pack(side='right')
        fabric_combo = ttk.Combobox(fr, textvariable=fabric_var, width=20)
        fabric_combo.pack(side='right', padx=(0, 8))
        tk.Label(fr, text='(יחול על כל השורות; בד חדש ייווצר באתר אוטומטית)', fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 8)).pack(side='right')

        status_var = tk.StringVar(value=f'{len(pending)} מוצרים ממתינים יישלחו כווריאנטים')
        tk.Label(pick_box, textvariable=status_var, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9)).pack(anchor='e', pady=(4, 0))

        def load_meta():
            try:
                client = WebsiteClient(url_var.get(), token_var.get())
                meta = client.fetch_meta()
            except WebsiteExportError as e:
                messagebox.showerror("שגיאת התחברות", str(e), parent=dlg)
                return
            products = meta.get('products') or []
            if not products:
                messagebox.showwarning("אין מוצרים", "לא נמצאו מוצרים באתר. יש ליצור קודם מוצר באדמין של האתר.", parent=dlg)
                return
            meta_state['products'] = products
            labels = []
            for p in products:
                label = p['name'] if not p.get('category') else f"{p['name']} ({p['category']})"
                labels.append(label)
            product_combo['values'] = labels
            product_combo.config(state='readonly')
            if labels:
                product_combo.current(0)
            fabric_combo['values'] = meta.get('fabric_types') or []
            status_var.set(f"נטענו {len(products)} מוצרים מהאתר | {len(pending)} מוצרים ממתינים יישלחו")
            # שמירת ההגדרות כבר עכשיו (החיבור הצליח), לא רק אחרי שליחה
            url_key, token_key, _, _ = _target_keys()
            self.settings.set(url_key, (url_var.get() or '').strip())
            self.settings.set(token_key, (token_var.get() or '').strip())

        tk.Button(target_box, text="🔄 טען מוצרים מהאתר", command=load_meta, bg=theme.TEAL, fg='white', font=(theme.FONT_FAMILY, 9, 'bold')).pack(anchor='w', pady=(6, 0))

        # --- מקור מחיר ---
        price_box = tk.LabelFrame(frm, text="מקור המחיר ליחידה", padx=8, pady=6)
        price_box.pack(fill='x', pady=(0, 8))
        price_src_var = tk.StringVar(value='digital')
        tk.Radiobutton(price_box, text="מחיר לצרכן דיגיטלי", variable=price_src_var, value='digital').pack(side='right', padx=6)
        tk.Radiobutton(price_box, text="מחיר מכירה", variable=price_src_var, value='sale').pack(side='right', padx=6)

        # --- סנכרון ריווחית ---
        riv_box = tk.LabelFrame(frm, text="סנכרון ריווחית", padx=8, pady=6)
        riv_box.pack(fill='x', pady=(0, 8))
        riv_auto_var = tk.BooleanVar(value=bool(self.settings.get('rivhit.auto_sync', True)))
        riv_state = {'groups': [], 'load_result': None}
        riv_check = tk.Checkbutton(riv_box, text="עדכן גם את ריווחית אוטומטית (יצירת פריטים חדשים ועדכון קיימים)", variable=riv_auto_var)
        riv_check.pack(anchor='e')
        rg = tk.Frame(riv_box); rg.pack(fill='x', pady=2)
        tk.Label(rg, text='קבוצה:', anchor='e', width=8).pack(side='right')
        riv_group_combo = ttk.Combobox(rg, state='disabled', width=36)
        riv_group_combo.pack(side='right', padx=(0, 8))
        riv_status_var = tk.StringVar(value='טוען קבוצות פריטים מריווחית...')
        tk.Label(riv_box, textvariable=riv_status_var, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9)).pack(anchor='e', pady=(2, 0))

        def _rivhit_unavailable(msg):
            if not dlg.winfo_exists():
                return
            riv_auto_var.set(False)
            riv_check.config(state='disabled')
            riv_status_var.set(msg)

        def _rivhit_ready(groups):
            if not dlg.winfo_exists():
                return
            riv_state['groups'] = groups
            labels = [f"{g.get('item_group_name', '')} ({g.get('item_group_id', '')})" for g in groups]
            riv_group_combo['values'] = labels
            riv_group_combo.config(state='readonly')
            default_gid = self.settings.get('rivhit.default_group_id', 16)
            default_idx = 0
            for i, g in enumerate(groups):
                if g.get('item_group_id') == default_gid:
                    default_idx = i
                    break
            if labels:
                riv_group_combo.current(default_idx)
            riv_status_var.set(f'חיבור לריווחית תקין ✓ | {len(groups)} קבוצות')

        def _load_rivhit_groups_bg():
            # רץ ב-thread רקע: כותב את התוצאה בלבד; ה-UI קורא אותה ב-poll מה-thread הראשי
            token = (self.settings.get('rivhit.api_token', '') or '').strip()
            if not token:
                riv_state['load_result'] = ('err', 'לא הוגדר טוקן ריווחית (דרך "עדכן ריווחית (API)") - הסנכרון כבוי')
                return
            try:
                groups = RivhitOnlineClient(token).item_groups()
            except RivhitApiError as e:
                riv_state['load_result'] = ('err', f'ריווחית לא זמינה - הסנכרון כבוי ({e})')
                return
            riv_state['load_result'] = ('ok', groups)

        def _poll_rivhit_load():
            if not dlg.winfo_exists():
                return
            res = riv_state['load_result']
            if res is None:
                dlg.after(150, _poll_rivhit_load)
                return
            if res[0] == 'ok':
                _rivhit_ready(res[1])
            else:
                _rivhit_unavailable(res[1])

        threading.Thread(target=_load_rivhit_groups_bg, daemon=True).start()
        dlg.after(150, _poll_rivhit_load)

        def send():
            idx = product_combo.current()
            if idx < 0 or not meta_state['products']:
                messagebox.showwarning("לא נבחר מוצר", "יש לטעון את המוצרים מהאתר ולבחור מוצר", parent=dlg)
                return
            product = meta_state['products'][idx]
            fabric = (fabric_var.get() or '').strip()
            if not fabric:
                messagebox.showwarning("חסר סוג בד", "יש להזין סוג בד", parent=dlg)
                return
            price_key = 'digital_price' if price_src_var.get() == 'digital' else 'item_sale_nis'
            rows = []
            for rec in pending:
                barcode = str(rec.get('item_part_num', '')).strip()
                fields = self.data_processor.get_rivhit_label_fields(barcode, product=rec)
                row = {
                    'size': (fields.get('size') or '').strip(),
                    'barcode': barcode,
                    'unit_price': str(rec.get(price_key, '')).strip(),
                }
                # מוצרים שנוצרו לפי צבעים - שליחת הצבע לאתר (למוצרים התומכים בצבע)
                color = str(rec.get('color', '')).strip() or (fields.get('color') or '').strip()
                if color:
                    row['color'] = color
                    color_hex = str(rec.get('color_hex', '')).strip()
                    if color_hex:
                        row['color_hex'] = color_hex
                    # באתר מוצג רק ריבוע קוד הצבע (hex) - לא תמונת צילום הבד
                    row['clear_image'] = True
                rows.append(row)
            try:
                client = WebsiteClient(url_var.get(), token_var.get())
                result = client.send_variants(product['id'], fabric, rows)
            except WebsiteExportError as e:
                messagebox.showerror("שגיאה בשליחה", str(e), parent=dlg)
                return
            # שמירת ההגדרות לפעם הבאה
            url_key, token_key, _, _ = _target_keys()
            self.settings.set(url_key, (url_var.get() or '').strip())
            self.settings.set(token_key, (token_var.get() or '').strip())
            self.settings.set('rivhit.auto_sync', bool(riv_auto_var.get()))

            lines = [
                "--- אתר ---",
                f"מוצר: {result.get('product', '')}",
                f"נוצרו: {result.get('created', 0)} | עודכנו: {result.get('updated', 0)}",
            ]
            for w in (result.get('warnings') or []):
                lines.append(f"⚠️ {w}")
            errors = result.get('errors') or []
            if errors:
                lines.append(f"שגיאות ({len(errors)}):")
                lines.extend(errors[:15])
                if len(errors) > 15:
                    lines.append(f"...ועוד {len(errors) - 15} שגיאות")

            # סנכרון ריווחית - רק אחרי שהשליחה לאתר הצליחה
            rivhit_failed = False
            if riv_auto_var.get() and riv_state['groups']:
                gidx = riv_group_combo.current()
                group_id = riv_state['groups'][gidx].get('item_group_id') if gidx >= 0 else None
                try:
                    token = (self.settings.get('rivhit.api_token', '') or '').strip()
                    riv_result = self._sync_products_to_rivhit(
                        RivhitOnlineClient(token), pending, group_id, update_existing=True)
                    self.settings.set('rivhit.default_group_id', group_id)
                    lines.append("")
                    lines.append("--- ריווחית ---")
                    lines.extend(self._format_rivhit_sync_summary(riv_result))
                    rivhit_failed = bool(riv_result['errors'])
                except RivhitApiError as e:
                    rivhit_failed = True
                    lines.append("")
                    lines.append("⚠️ האתר עודכן, אבל עדכון ריווחית נכשל:")
                    lines.append(str(e))
                    lines.append('אפשר לנסות שוב דרך הכפתור "עדכן ריווחית (API)".')

            if errors or rivhit_failed:
                messagebox.showwarning("הסתיים עם שגיאות", "\n".join(lines), parent=dlg)
            else:
                dlg.destroy()
                messagebox.showinfo("הצלחה", "\n".join(lines).strip())

        btns = tk.Frame(frm)
        btns.pack(pady=(6, 0))
        tk.Button(btns, text="🌐 שלח לאתר", command=send, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btns, text="ביטול", command=dlg.destroy).pack(side='left', padx=5)

    def _sync_products_to_rivhit(self, client, products, group_id, update_existing=True):
        """שליחת רשימת מוצרים לריווחית: יצירת חדשים ועדכון/דילוג על קיימים לפי מק"ט/ברקוד.

        מחזיר dict עם רשימות created/updated/skipped/errors (מחרוזות לתצוגה).
        קריאת Item.List נכשלת -> RivhitApiError עולה למעלה; כשל בפריט בודד נאסף ב-errors.
        """
        from ..core.rivhit_api import RivhitApiError

        existing_items = client.item_list()
        by_key = {}
        for it in existing_items:
            for key in (str(it.get('item_part_num', '')).strip(), str(it.get('barcode', '')).strip()):
                if key:
                    by_key.setdefault(key, it)

        result = {'created': [], 'updated': [], 'skipped': [], 'errors': []}
        for rec in products:
            name = str(rec.get('item_name', '')).strip()
            part_num = str(rec.get('item_part_num', '')).strip()
            existing = by_key.get(part_num)
            try:
                if existing is not None:
                    if update_existing:
                        client.item_update(
                            existing.get('item_id'),
                            item_name=name,
                            item_part_num=part_num,
                            barcode=part_num,
                            cost_nis=rec.get('item_cost_nis', ''),
                            sale_nis=rec.get('item_sale_nis', ''),
                            item_group_id=group_id,
                        )
                        result['updated'].append(f"{name} (פריט {existing.get('item_id')})")
                    else:
                        result['skipped'].append(f"{name} - כבר קיים בריווחית (פריט {existing.get('item_id')})")
                    continue
                data = client.item_new(
                    item_name=name,
                    item_part_num=part_num,
                    barcode=part_num,
                    cost_nis=rec.get('item_cost_nis', ''),
                    sale_nis=rec.get('item_sale_nis', ''),
                    item_group_id=group_id,
                )
                item_id = data.get('item_id', '')
                result['created'].append(f"{name} (פריט {item_id})" if item_id else name)
            except RivhitApiError as e:
                result['errors'].append(f"{name}: {e}")
        return result

    @staticmethod
    def _format_rivhit_sync_summary(result):
        """שורות סיכום לתצוגה מתוצאת _sync_products_to_rivhit."""
        lines = [
            f"נוצרו: {len(result['created'])} | עודכנו: {len(result['updated'])} | "
            f"דולגו: {len(result['skipped'])} | שגיאות: {len(result['errors'])}",
            "",
        ]
        for label, items in (("נוצרו:", result['created']), ("עודכנו:", result['updated']), ("דולגו:", result['skipped'])):
            if items:
                lines.append(label)
                lines.extend(items[:15])
                if len(items) > 15:
                    lines.append(f"...ועוד {len(items) - 15}")
                lines.append("")
        if result['errors']:
            lines.append("שגיאות:")
            lines.extend(result['errors'][:15])
            if len(result['errors']) > 15:
                lines.append(f"...ועוד {len(result['errors']) - 15} שגיאות")
        return lines

    def _update_rivhit_via_api(self):
        """שליחת המוצרים הממתינים ישירות לריווחית אונליין (פריטים ומלאי) דרך ה-API."""
        from ..core.rivhit_api import RivhitOnlineClient, RivhitApiError

        pending = list(getattr(self.data_processor, 'rivhit_new_products', []) or [])
        if not pending:
            messagebox.showinfo("אין מוצרים", "אין מוצרים ממתינים לשליחה לריווחית")
            return

        dlg = tk.Toplevel(self.notebook)
        dlg.title("עדכן ריווחית (API)")
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = tk.Frame(dlg, padx=15, pady=15)
        frm.pack(fill='both', expand=True)

        # --- חיבור ---
        conn_box = tk.LabelFrame(frm, text="חיבור לריווחית אונליין", padx=8, pady=6)
        conn_box.pack(fill='x', pady=(0, 8))
        token_var = tk.StringVar(value=self.settings.get('rivhit.api_token', ''))
        tr = tk.Frame(conn_box); tr.pack(fill='x')
        tk.Label(tr, text='API TOKEN:', anchor='e', width=10).pack(side='right')
        tk.Entry(tr, textvariable=token_var, width=44, show='*').pack(side='right', padx=(0, 8))
        tk.Label(conn_box, text='הטוקן מהגדרות ריווחית אונליין ← API', fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 8)).pack(anchor='e')

        # --- קבוצת פריטים ---
        group_box = tk.LabelFrame(frm, text="קבוצת פריטים בריווחית", padx=8, pady=6)
        group_box.pack(fill='x', pady=(0, 8))
        group_var = tk.StringVar()
        group_state = {'groups': [], 'existing': None}
        gr = tk.Frame(group_box); gr.pack(fill='x', pady=2)
        tk.Label(gr, text='קבוצה:', anchor='e', width=10).pack(side='right')
        group_combo = ttk.Combobox(gr, textvariable=group_var, state='disabled', width=40)
        group_combo.pack(side='right', padx=(0, 8))

        status_var = tk.StringVar(value=f'{len(pending)} מוצרים ממתינים | יש ללחוץ "בדוק חיבור" תחילה')
        tk.Label(group_box, textvariable=status_var, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9)).pack(anchor='e', pady=(4, 0))

        update_existing_var = tk.BooleanVar(value=False)
        tk.Checkbutton(frm, text='עדכן פריטים שכבר קיימים בריווחית (לפי מק"ט/ברקוד), אחרת ידולגו', variable=update_existing_var).pack(anchor='e', pady=(0, 8))

        def _make_client():
            token = (token_var.get() or '').strip()
            if not token:
                messagebox.showwarning("חסר טוקן", "יש להזין API TOKEN של ריווחית", parent=dlg)
                return None
            return RivhitOnlineClient(token)

        def check_connection():
            client = _make_client()
            if client is None:
                return
            try:
                groups = client.item_groups()
            except RivhitApiError as e:
                messagebox.showerror("שגיאת התחברות", str(e), parent=dlg)
                return
            group_state['groups'] = groups
            labels = [f"{g.get('item_group_name', '')} ({g.get('item_group_id', '')})" for g in groups]
            group_combo['values'] = labels
            group_combo.config(state='readonly')
            # ברירת מחדל: הקבוצה שתואמת את הקטגוריה של המוצרים הממתינים
            cats = {str(r.get('compute_0036', '')).strip() for r in pending if str(r.get('compute_0036', '')).strip()}
            default_idx = 0
            for i, g in enumerate(groups):
                if str(g.get('item_group_name', '')).strip() in cats:
                    default_idx = i
                    break
            if labels:
                group_combo.current(default_idx)
            status_var.set(f"חיבור תקין ✓ | נטענו {len(groups)} קבוצות פריטים | {len(pending)} מוצרים ממתינים")
            # החיבור הצליח - שמירת הטוקן להבא
            self.settings.set('rivhit.api_token', (token_var.get() or '').strip())

        tk.Button(conn_box, text="🔌 בדוק חיבור", command=check_connection, bg=theme.TEAL, fg='white', font=(theme.FONT_FAMILY, 9, 'bold')).pack(anchor='w', pady=(6, 0))

        def send():
            client = _make_client()
            if client is None:
                return
            idx = group_combo.current()
            if idx < 0 or not group_state['groups']:
                messagebox.showwarning("לא נבחרה קבוצה", 'יש ללחוץ "בדוק חיבור" ולבחור קבוצת פריטים', parent=dlg)
                return
            group_id = group_state['groups'][idx].get('item_group_id')
            try:
                result = self._sync_products_to_rivhit(client, pending, group_id, update_existing_var.get())
            except RivhitApiError as e:
                messagebox.showerror("שגיאה במשיכת פריטים", str(e), parent=dlg)
                return

            self.settings.set('rivhit.api_token', (token_var.get() or '').strip())
            self.settings.set('rivhit.default_group_id', group_id)

            lines = self._format_rivhit_sync_summary(result)
            if result['errors']:
                messagebox.showwarning("הסתיים עם שגיאות", "\n".join(lines), parent=dlg)
            else:
                dlg.destroy()
                messagebox.showinfo("הסתיים", "\n".join(lines).strip())

        btns = tk.Frame(frm)
        btns.pack(pady=(6, 0))
        tk.Button(btns, text="🔗 שלח לריווחית", command=send, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=5)
        tk.Button(btns, text="ביטול", command=dlg.destroy).pack(side='left', padx=5)

    # ===== Data upload sub-tab (יבוא נתונים מריווחית) =====
    def _build_rivhit_data_subtab(self, tab):
        tk.Label(tab, text="העלאת נתונים מריווחית", font=(theme.FONT_FAMILY, 16, 'bold'), bg=theme.PAGE_BG, fg=theme.DARK).pack(pady=8)

        # Section 1: items list import
        items_box = tk.LabelFrame(tab, text="רשימת פריטים מריווחית", bg=theme.PAGE_BG, fg=theme.DARK, font=(theme.FONT_FAMILY, 11, 'bold'), padx=12, pady=12)
        items_box.pack(fill='x', padx=15, pady=8)
        tk.Label(items_box, text="העלאת קובץ ייצוא הפריטים מריווחית מחליפה את הרשימה המוצגת בטאב 'רשימת מוצרים'.", bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9), justify='right').pack(anchor='e', pady=(0, 6))
        items_row = tk.Frame(items_box, bg=theme.PAGE_BG); items_row.pack(fill='x')
        tk.Button(items_row, text="⬆️ העלה קובץ פריטים מריווחית", command=self._import_rivhit_file, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        self.rivhit_data_meta_var = tk.StringVar(value='')
        tk.Label(items_box, textvariable=self.rivhit_data_meta_var, bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9)).pack(anchor='e', pady=(6, 0))

        # Section 2: groups/categories import
        groups_box = tk.LabelFrame(tab, text="קבוצות / קטגוריות", bg=theme.PAGE_BG, fg=theme.DARK, font=(theme.FONT_FAMILY, 11, 'bold'), padx=12, pady=12)
        groups_box.pack(fill='x', padx=15, pady=8)
        tk.Label(groups_box, text="קובץ הקבוצות (מספר<TAB>שם) קובע אילו עונות/קטגוריות זמינות בהוספת מוצר וכיצד הן ממופות בייצוא.", bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9), justify='right').pack(anchor='e', pady=(0, 6))
        groups_row = tk.Frame(groups_box, bg=theme.PAGE_BG); groups_row.pack(fill='x')
        tk.Button(groups_row, text="⬆️ העלה קובץ קבוצות", command=self._import_rivhit_groups_file, bg=theme.PURPLE, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='right', padx=5)
        self.rivhit_groups_meta_var = tk.StringVar(value='')
        tk.Label(groups_box, textvariable=self.rivhit_groups_meta_var, bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9)).pack(anchor='e', pady=(6, 0))

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
