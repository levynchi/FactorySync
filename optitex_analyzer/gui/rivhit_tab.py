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

        # Pending list toolbar
        toolbar = tk.Frame(tab, bg=theme.PAGE_BG)
        toolbar.pack(fill='x', padx=15, pady=(8, 2))
        tk.Label(toolbar, text="מוצרים ממתינים לייצוא:", bg=theme.PAGE_BG, font=(theme.FONT_FAMILY, 11, 'bold')).pack(side='right')
        tk.Button(toolbar, text="⬇️ ייצא קובץ לריווחית", command=self._export_rivhit_new, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=4)
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

    def _apply_rivhit_print_fields(self, barcode, size=''):
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
        # החל רק אם המשתמש הזין פרט הדפסה כלשהו
        if not (print_name or fabric or img_src or size or size_unit or model_code):
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
    def _export_rivhit_to_website(self):
        """ייצוא המוצרים הממתינים כווריאנטים למוצר קיים באתר הקטלוג הלבן."""
        from ..core.website_export import WebsiteClient, WebsiteExportError

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
        token_var = tk.StringVar(value=self.settings.get('website.api_token', ''))

        def on_target_change():
            key = 'website.local_url' if target_var.get() == 'local' else 'website.prod_url'
            default = 'http://127.0.0.1:8000' if target_var.get() == 'local' else 'https://arye-textil.co.il'
            url_var.set(self.settings.get(key, default))

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

        tk.Button(target_box, text="🔄 טען מוצרים מהאתר", command=load_meta, bg=theme.TEAL, fg='white', font=(theme.FONT_FAMILY, 9, 'bold')).pack(anchor='w', pady=(6, 0))

        # --- מקור מחיר ---
        price_box = tk.LabelFrame(frm, text="מקור המחיר ליחידה", padx=8, pady=6)
        price_box.pack(fill='x', pady=(0, 8))
        price_src_var = tk.StringVar(value='digital')
        tk.Radiobutton(price_box, text="מחיר לצרכן דיגיטלי", variable=price_src_var, value='digital').pack(side='right', padx=6)
        tk.Radiobutton(price_box, text="מחיר מכירה", variable=price_src_var, value='sale').pack(side='right', padx=6)

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
                rows.append({
                    'size': (fields.get('size') or '').strip(),
                    'barcode': barcode,
                    'unit_price': str(rec.get(price_key, '')).strip(),
                })
            try:
                client = WebsiteClient(url_var.get(), token_var.get())
                result = client.send_variants(product['id'], fabric, rows)
            except WebsiteExportError as e:
                messagebox.showerror("שגיאה בשליחה", str(e), parent=dlg)
                return
            # שמירת ההגדרות לפעם הבאה
            url_key = 'website.local_url' if target_var.get() == 'local' else 'website.prod_url'
            self.settings.set(url_key, (url_var.get() or '').strip())
            self.settings.set('website.api_token', (token_var.get() or '').strip())

            lines = [
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
                messagebox.showwarning("הסתיים עם שגיאות", "\n".join(lines), parent=dlg)
            else:
                dlg.destroy()
                messagebox.showinfo("הצלחה", "\n".join(lines))

        btns = tk.Frame(frm)
        btns.pack(pady=(6, 0))
        tk.Button(btns, text="🌐 שלח לאתר", command=send, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold')).pack(side='left', padx=5)
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
