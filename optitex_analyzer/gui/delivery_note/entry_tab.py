import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

# This module defines a function that builds the Entry sub-tab UI on a given container.

def build_entry_tab(ctx, container: tk.Frame):
    """Build the delivery note entry sub-tab.

    ctx is the MainWindow (mixin host) instance. We reuse its state and methods.
    """
    # Header form
    form = ttk.LabelFrame(container, text="פרטי תעודה", padding=10)
    form.pack(fill='x', padx=10, pady=6)
    tk.Label(form, text="שם ספק:", font=('Arial',10,'bold')).grid(row=0,column=0,sticky='w',padx=4,pady=4)
    ctx.dn_supplier_name_var = tk.StringVar()
    ctx.dn_supplier_name_combo = ttk.Combobox(form, textvariable=ctx.dn_supplier_name_var, width=28, state='readonly')
    try:
        names = ctx._get_supplier_names() if hasattr(ctx,'_get_supplier_names') else []
        ctx.dn_supplier_name_combo['values'] = names
    except Exception:
        pass
    ctx.dn_supplier_name_combo.grid(row=0,column=1,sticky='w',padx=4,pady=4)
    tk.Label(form, text="תאריך:", font=('Arial',10,'bold')).grid(row=0,column=2,sticky='w',padx=4,pady=4)
    ctx.dn_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
    tk.Entry(form, textvariable=ctx.dn_date_var, width=15).grid(row=0,column=3,sticky='w',padx=4,pady=4)

    # Lines frame
    lines_frame = ttk.LabelFrame(container, text="שורות תעודה", padding=8)
    lines_frame.pack(fill='both', expand=False, padx=10, pady=4)
    entry_bar = tk.Frame(lines_frame, bg='#f7f9fa')
    entry_bar.pack(fill='x', pady=(0,6))

    # Variables
    ctx.dn_product_var = tk.StringVar()
    ctx.dn_size_var = tk.StringVar()
    ctx.dn_qty_var = tk.StringVar()
    ctx.dn_note_var = tk.StringVar()
    ctx.dn_fabric_type_var = tk.StringVar()
    ctx.dn_fabric_color_var = tk.StringVar(value='לבן')
    ctx.dn_print_name_var = tk.StringVar(value='חלק')
    ctx.dn_fabric_category_var = tk.StringVar()
    # Main category (Products / Accessories)
    ctx.dn_main_category_var = tk.StringVar(value='מוצרים')
    ctx._delivery_products_allowed = []
    ctx._refresh_delivery_products_allowed(initial=True)
    ctx._delivery_products_allowed_full = list(ctx._delivery_products_allowed)
    # Main category chooser controls
    tk.Label(entry_bar, text='קטגוריה ראשית', bg='#f7f9fa').grid(row=0, column=0, sticky='w', padx=2)
    ctx.dn_main_category_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_main_category_var, width=12, state='readonly', values=['מוצרים','אביזרי תפירה'])
    ctx.dn_main_category_combo.grid(row=1, column=0, sticky='w', padx=2)

    # Product/Accessory name field
    ctx.dn_product_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_product_var, width=16, state='normal')
    ctx.dn_product_combo['values'] = ctx._delivery_products_allowed_full

    # Popup state
    ctx._dn_ac_popup = None
    ctx._dn_ac_list = None

    def _ensure_popup():
        if ctx._dn_ac_popup and ctx._dn_ac_popup.winfo_exists():
            return
        popup = tk.Toplevel(ctx.dn_product_combo)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        lb = tk.Listbox(popup, height=8, activestyle='dotbox')
        lb.pack(fill='both', expand=True)
        ctx._dn_ac_popup = popup
        ctx._dn_ac_list = lb

        def _choose(event=None):
            sel = lb.curselection()
            if not sel:
                _hide_popup(); return
            val = lb.get(sel[0])
            ctx.dn_product_var.set(val)
            _hide_popup()
            try:
                size_entry = [w for w in entry_bar.grid_slaves(row=1) if isinstance(w, tk.Entry)][0]
                size_entry.focus_set()
            except Exception:
                pass
        lb.bind('<Return>', _choose)
        lb.bind('<Double-Button-1>', _choose)
        lb.bind('<Escape>', lambda e: _hide_popup())

    def _hide_popup():
        if ctx._dn_ac_popup and ctx._dn_ac_popup.winfo_exists():
            ctx._dn_ac_popup.destroy()

    def _position_popup():
        if not (ctx._dn_ac_popup and ctx._dn_ac_popup.winfo_exists()):
            return
        try:
            x = ctx.dn_product_combo.winfo_rootx()
            y = ctx.dn_product_combo.winfo_rooty() + ctx.dn_product_combo.winfo_height()
            w = ctx.dn_product_combo.winfo_width()
            ctx._dn_ac_popup.geometry(f"{w}x180+{x}+{y}")
        except Exception:
            pass

    def _filter_products(event=None):
        if event and event.keysym in ('Escape',):
            _hide_popup(); return
        text = ctx.dn_product_var.get().strip()
        base = ctx._delivery_products_allowed_full
        if not text:
            matches = base[:50]
        else:
            tokens = [t for t in text.lower().replace('-', ' ').split() if t]
            def match(prod):
                pl = prod.lower(); words = pl.replace('-', ' ').split()
                for tok in tokens:
                    if not any(w.startswith(tok) or tok in w for w in words):
                        return None
                prefix_hits = sum(any(w.startswith(tok) for w in words) for tok in tokens)
                first_idx_sum = sum(min((i for i,w in enumerate(words) if (w.startswith(tok) or tok in w)), default=99) for tok in tokens)
                return (-prefix_hits, first_idx_sum, len(prod))
            scored = []
            for p in base:
                sc = match(p)
                if sc is not None:
                    scored.append((sc,p))
            scored.sort(key=lambda x: x[0])
            matches = [p for _,p in scored][:50]
        if not matches:
            _hide_popup(); return
        _ensure_popup(); _position_popup()
        lb = ctx._dn_ac_list; lb.delete(0, tk.END)
        for m in matches: lb.insert(tk.END, m)
        lb.selection_clear(0, tk.END); lb.selection_set(0); lb.activate(0)

    def _on_key(event):
        if event.keysym in ('Down','Up') and ctx._dn_ac_popup and ctx._dn_ac_popup.winfo_exists():
            lb = ctx._dn_ac_list; size = lb.size()
            if size == 0: return
            cur = lb.curselection(); idx = cur[0] if cur else 0
            if event.keysym == 'Down': idx = (idx + 1) % size
            else: idx = (idx - 1) % size
            lb.selection_clear(0, tk.END); lb.selection_set(idx); lb.activate(idx)
            return 'break'
        if event.keysym == 'Return' and ctx._dn_ac_popup and ctx._dn_ac_popup.winfo_exists():
            lb = ctx._dn_ac_list; cur = lb.curselection()
            if cur:
                ctx.dn_product_var.set(lb.get(cur[0])); _hide_popup(); return 'break'
        _filter_products()

    ctx.dn_product_combo.bind('<KeyRelease>', _on_key)
    ctx.dn_product_combo.bind('<FocusOut>', lambda e: ctx.root.after(150, _hide_popup))

    def _product_chosen(event=None):
        try:
            widgets_after = [w for w in entry_bar.grid_slaves(row=1) if isinstance(w, tk.Entry)]
        except Exception:
            widgets_after = []
        for w in widgets_after:
            if hasattr(w,'cget') and w.cget('textvariable') == str(ctx.dn_size_var):
                w.focus_set(); break
    ctx.dn_product_combo.bind('<<ComboboxSelected>>', _product_chosen)

    # Accessories catalog and unit handling
    ctx.dn_unit_var = tk.StringVar()
    ctx._accessories_by_name = {}
    ctx._accessories_names = []
    try:
        import os, json
        apath = os.path.join(os.getcwd(), 'sewing_accessories.json')
        if os.path.exists(apath):
            with open(apath, 'r', encoding='utf-8') as f:
                adata = json.load(f)
            if isinstance(adata, list):
                for it in adata:
                    nm = (it.get('name') or '').strip(); un = (it.get('unit') or '').strip()
                    if nm:
                        ctx._accessories_by_name[nm] = un
                ctx._accessories_names = sorted(ctx._accessories_by_name.keys())
    except Exception:
        pass

    def _update_unit_from_accessory(*_a):
        try:
            if ctx.dn_main_category_var.get() == 'אביזרי תפירה':
                nm = (ctx.dn_product_var.get() or '').strip()
                ctx.dn_unit_var.set(ctx._accessories_by_name.get(nm, ''))
            else:
                ctx.dn_unit_var.set('')
        except Exception:
            ctx.dn_unit_var.set('')

    def _on_main_category_change(*_a):
        mode = ctx.dn_main_category_var.get()
        if mode == 'אביזרי תפירה':
            # switch product input to accessory list
            try:
                ctx.dn_product_combo['values'] = ctx._accessories_names
                ctx.dn_product_combo.set('')
            except Exception:
                pass
            # disable product-variant combos
            for combo in (ctx.dn_size_combo, ctx.dn_fabric_type_combo, ctx.dn_fabric_color_combo, ctx.dn_category_combo):
                try:
                    combo.set(''); combo.state(['disabled'])
                except Exception:
                    pass
            try:
                ctx.dn_fabric_category_var.set('')
                dn_print_entry.delete(0, tk.END)
            except Exception:
                pass
            _update_unit_from_accessory()
        else:
            # back to products
            try:
                ctx.dn_product_combo['values'] = ctx._delivery_products_allowed_full
                ctx.dn_product_combo.set('')
            except Exception:
                pass
            for combo in (ctx.dn_size_combo, ctx.dn_fabric_type_combo, ctx.dn_fabric_color_combo, ctx.dn_category_combo):
                try:
                    combo.state(['!disabled'])
                except Exception:
                    pass
            ctx.dn_unit_var.set('')

    try:
        ctx.dn_main_category_var.trace_add('write', _on_main_category_change)
    except Exception:
        pass

    lbls = ["מוצר","מידה","סוג בד","צבע בד","קטגורית בד","שם פרינט","תת קטגוריה","יחידה","כמות","הערה"]
    for i,lbl in enumerate(lbls):
        # shift by +1 column group to make space for main category controls at col 0/1
        tk.Label(entry_bar, text=lbl, bg='#f7f9fa').grid(row=0,column=(i+1)*2,sticky='w',padx=2)

    ctx.dn_size_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_size_var, width=10, state='readonly')
    ctx.dn_fabric_type_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_fabric_type_var, width=12, state='readonly')
    ctx.dn_fabric_color_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_fabric_color_var, width=10, state='readonly')
    # Fabric category is auto-filled from products_catalog match; show as read-only entry (not user-selectable)
    ctx.dn_fabric_category_entry = ttk.Entry(entry_bar, textvariable=ctx.dn_fabric_category_var, width=14, state='readonly')
    dn_print_entry = tk.Entry(entry_bar, textvariable=ctx.dn_print_name_var, width=12)

    # New: Subcategory (from categories.json)
    ctx.dn_category_var = tk.StringVar()
    ctx.dn_category_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_category_var, width=14, state='readonly')
    try:
        import os, json
        path = os.path.join(os.getcwd(), 'categories.json')
        cats = []
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                for item in data:
                    name = (item.get('name') or '').strip()
                    if name:
                        cats.append(name)
        cats = sorted({c for c in cats})
        ctx.dn_category_combo['values'] = cats
    except Exception:
        pass

    # Create unit entry (read-only)
    dn_unit_entry = ttk.Entry(entry_bar, textvariable=ctx.dn_unit_var, width=10, state='readonly')

    widgets = [
        ctx.dn_product_combo,
        ctx.dn_size_combo,
        ctx.dn_fabric_type_combo,
        ctx.dn_fabric_color_combo,
        ctx.dn_fabric_category_entry,
        dn_print_entry,
        ctx.dn_category_combo,
        dn_unit_entry,
        tk.Entry(entry_bar, textvariable=ctx.dn_qty_var, width=7),
        tk.Entry(entry_bar, textvariable=ctx.dn_note_var, width=18)
    ]
    for i,w in enumerate(widgets):
        w.grid(row=1,column=(i+1)*2,sticky='w',padx=2)

    def _on_product_change(*_a):
        try:
            ctx._update_delivery_size_options()
            ctx._update_delivery_fabric_type_options()
            ctx._update_delivery_fabric_color_options()
            # also recompute fabric category when product changes
            if hasattr(ctx, '_update_delivery_fabric_category_auto'):
                ctx._update_delivery_fabric_category_auto()
        except Exception:
            pass
    try:
        ctx.dn_product_var.trace_add('write', _on_product_change)
    except Exception:
        pass

    try:
        def _on_fabric_type_change(*_a):
            try:
                ctx._update_delivery_fabric_color_options()
                if hasattr(ctx, '_update_delivery_fabric_category_auto'):
                    ctx._update_delivery_fabric_category_auto()
            except Exception:
                pass
        ctx.dn_fabric_type_var.trace_add('write', _on_fabric_type_change)
    except Exception:
        pass

    # When size / color / print name change, try to auto-fill fabric category
    try:
        ctx.dn_size_var.trace_add('write', lambda *_: getattr(ctx, '_update_delivery_fabric_category_auto', lambda: None)())
    except Exception:
        pass
    try:
        ctx.dn_fabric_color_var.trace_add('write', lambda *_: getattr(ctx, '_update_delivery_fabric_category_auto', lambda: None)())
    except Exception:
        pass
    try:
        ctx.dn_print_name_var.trace_add('write', lambda *_: getattr(ctx, '_update_delivery_fabric_category_auto', lambda: None)())
    except Exception:
        pass

    for combo in (ctx.dn_size_combo, ctx.dn_fabric_type_combo, ctx.dn_fabric_color_combo):
        try: combo.state(['disabled'])
        except Exception: pass

    # After adding a new field, shift action buttons to the right to avoid overlap
    _btn_base_col = (len(widgets) + 1) * 2
    tk.Button(entry_bar, text="➕ הוסף", command=ctx._add_delivery_line, bg='#27ae60', fg='white').grid(row=1,column=_btn_base_col,padx=6)
    tk.Button(entry_bar, text="🗑️ מחק נבחר", command=ctx._delete_delivery_selected, bg='#e67e22', fg='white').grid(row=1,column=_btn_base_col+1,padx=4)
    tk.Button(entry_bar, text="❌ נקה הכל", command=ctx._clear_delivery_lines, bg='#e74c3c', fg='white').grid(row=1,column=_btn_base_col+2,padx=4)

    cols = ('product','size','fabric_type','fabric_color','fabric_category','print_name','category','unit','quantity','note')
    ctx.delivery_tree = ttk.Treeview(lines_frame, columns=cols, show='headings', height=10)
    headers = {'product':'מוצר/פריט','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','fabric_category':'קטגורית בד','print_name':'שם פרינט','category':'תת קטגוריה','unit':'יחידה','quantity':'כמות','note':'הערה'}
    widths = {'product':160,'size':80,'fabric_type':110,'fabric_color':90,'fabric_category':120,'print_name':110,'category':110,'unit':70,'quantity':70,'note':220}
    for c in cols:
        ctx.delivery_tree.heading(c, text=headers[c])
        ctx.delivery_tree.column(c, width=widths[c], anchor='center')
    vs = ttk.Scrollbar(lines_frame, orient='vertical', command=ctx.delivery_tree.yview)
    ctx.delivery_tree.configure(yscroll=vs.set)
    ctx.delivery_tree.pack(side='left', fill='both', expand=True, padx=(4,0), pady=4)
    vs.pack(side='right', fill='y')

    # Transportation section
    pkg_frame = ttk.LabelFrame(container, text="הובלה", padding=8)
    pkg_frame.pack(fill='x', padx=10, pady=(4,4))
    ctx.pkg_type_var = tk.StringVar(value='שקית קטנה')
    ctx.pkg_qty_var = tk.StringVar()
    ctx.pkg_driver_var = tk.StringVar()
    tk.Label(pkg_frame, text="פריט הובלה:").grid(row=0,column=0,sticky='w',padx=4,pady=2)
    ctx.pkg_type_combo = ttk.Combobox(pkg_frame, textvariable=ctx.pkg_type_var, state='readonly', width=14, values=['שקית קטנה','שק','בד'])
    ctx.pkg_type_combo.grid(row=0,column=1,sticky='w',padx=4,pady=2)
    tk.Label(pkg_frame, text="כמות:").grid(row=0,column=2,sticky='w',padx=4,pady=2)
    tk.Entry(pkg_frame, textvariable=ctx.pkg_qty_var, width=8).grid(row=0,column=3,sticky='w',padx=4,pady=2)
    tk.Label(pkg_frame, text="שם המוביל:").grid(row=0,column=4,sticky='w',padx=4,pady=2)
    ctx.pkg_driver_combo = ttk.Combobox(pkg_frame, textvariable=ctx.pkg_driver_var, width=16, state='readonly')
    ctx.pkg_driver_combo.grid(row=0,column=5,sticky='w',padx=4,pady=2)
    try:
        ctx._refresh_driver_names_for_delivery()
    except Exception:
        pass
    tk.Button(pkg_frame, text="➕ הוסף", command=ctx._add_package_line, bg='#27ae60', fg='white').grid(row=0,column=6,padx=8)
    tk.Button(pkg_frame, text="🗑️ מחק נבחר", command=ctx._delete_selected_package, bg='#e67e22', fg='white').grid(row=0,column=7,padx=4)
    tk.Button(pkg_frame, text="❌ נקה", command=ctx._clear_packages, bg='#e74c3c', fg='white').grid(row=0,column=8,padx=4)
    ctx.packages_tree = ttk.Treeview(pkg_frame, columns=('type','quantity','driver'), show='headings', height=4)
    ctx.packages_tree.heading('type', text='פריט הובלה')
    ctx.packages_tree.heading('quantity', text='כמות')
    ctx.packages_tree.heading('driver', text='שם המוביל')
    ctx.packages_tree.column('type', width=120, anchor='center')
    ctx.packages_tree.column('quantity', width=70, anchor='center')
    ctx.packages_tree.column('driver', width=110, anchor='center')
    ctx.packages_tree.grid(row=1,column=0,columnspan=9, sticky='ew', padx=2, pady=(6,2))

    bottom_actions = tk.Frame(container, bg='#f7f9fa')
    bottom_actions.pack(fill='x', padx=10, pady=6)
    tk.Button(bottom_actions, text="💾 שמור תעודה", command=ctx._save_delivery_note, bg='#2c3e50', fg='white', font=('Arial',11,'bold')).pack(side='right', padx=4)
    ctx.delivery_summary_var = tk.StringVar(value="0 שורות | 0 כמות")
    tk.Label(container, textvariable=ctx.delivery_summary_var, bg='#34495e', fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')
