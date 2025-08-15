import tkinter as tk
from tkinter import ttk

def build_entry_tab(ctx, container: tk.Frame):
    # Reuse the original UI construction from SupplierIntakeTabMixin
    # Header form
    # Ensure internal state lists exist
    if not hasattr(ctx, '_supplier_lines'):
        ctx._supplier_lines = []
    if not hasattr(ctx, '_supplier_packages'):
        ctx._supplier_packages = []

    form = ttk.LabelFrame(container, text="פרטי קליטה", padding=10)
    form.pack(fill='x', padx=10, pady=6)
    tk.Label(form, text="שם ספק:", font=('Arial',10,'bold')).grid(row=0,column=0,sticky='w',padx=4,pady=4)
    ctx.supplier_name_var = tk.StringVar()
    ctx.supplier_name_combo = ttk.Combobox(form, textvariable=ctx.supplier_name_var, width=28, state='readonly')
    try:
        names = ctx._get_supplier_names() if hasattr(ctx,'_get_supplier_names') else []
        ctx.supplier_name_combo['values'] = names
    except Exception:
        pass
    ctx.supplier_name_combo.grid(row=0,column=1,sticky='w',padx=4,pady=4)
    tk.Label(form, text="תאריך:", font=('Arial',10,'bold')).grid(row=0,column=2,sticky='w',padx=4,pady=4)
    ctx.supplier_date_var = tk.StringVar()
    try:
        from datetime import datetime
        ctx.supplier_date_var.set(datetime.now().strftime('%Y-%m-%d'))
    except Exception:
        ctx.supplier_date_var.set('')
    tk.Entry(form, textvariable=ctx.supplier_date_var, width=15).grid(row=0,column=3,sticky='w',padx=4,pady=4)

    # Lines frame and entry bar
    lines_frame = ttk.LabelFrame(container, text="שורות קליטה", padding=8)
    lines_frame.pack(fill='both', expand=True, padx=10, pady=4)
    entry_bar = tk.Frame(lines_frame, bg='#f7f9fa')
    entry_bar.pack(fill='x', pady=(0,6))

    # Variables
    ctx.sup_product_var = tk.StringVar()
    ctx.sup_size_var = tk.StringVar()
    ctx.sup_qty_var = tk.StringVar()
    ctx.sup_note_var = tk.StringVar()
    ctx.sup_fabric_type_var = tk.StringVar()
    ctx.sup_fabric_color_var = tk.StringVar(value='לבן')
    ctx.sup_print_name_var = tk.StringVar(value='חלק')
    ctx.sup_fabric_category_var = tk.StringVar()

    # Product list
    ctx._supplier_products_allowed = []
    ctx._refresh_supplier_products_allowed(initial=True)
    ctx._supplier_products_allowed_full = list(ctx._supplier_products_allowed)
    ctx.sup_product_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_product_var, width=16, state='normal')
    ctx.sup_product_combo['values'] = ctx._supplier_products_allowed_full

    # Autocomplete popup state
    ctx._sup_ac_popup = None
    ctx._sup_ac_list = None

    def _ensure_popup():
        if ctx._sup_ac_popup and ctx._sup_ac_popup.winfo_exists():
            return
        popup = tk.Toplevel(ctx.sup_product_combo)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        lb = tk.Listbox(popup, height=8, activestyle='dotbox')
        lb.pack(fill='both', expand=True)
        ctx._sup_ac_popup = popup
        ctx._sup_ac_list = lb

        def _choose(event=None):
            sel = lb.curselection()
            if not sel:
                _hide_popup(); return
            val = lb.get(sel[0])
            ctx.sup_product_var.set(val)
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
        if ctx._sup_ac_popup and ctx._sup_ac_popup.winfo_exists():
            ctx._sup_ac_popup.destroy()

    def _position_popup():
        if not (ctx._sup_ac_popup and ctx._sup_ac_popup.winfo_exists()):
            return
        try:
            x = ctx.sup_product_combo.winfo_rootx()
            y = ctx.sup_product_combo.winfo_rooty() + ctx.sup_product_combo.winfo_height()
            w = ctx.sup_product_combo.winfo_width()
            ctx._sup_ac_popup.geometry(f"{w}x180+{x}+{y}")
        except Exception:
            pass

    def _filter_products(event=None):
        if event and event.keysym in ('Escape',):
            _hide_popup(); return
        text = ctx.sup_product_var.get().strip()
        base = ctx._supplier_products_allowed_full
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
        lb = ctx._sup_ac_list; lb.delete(0, tk.END)
        for m in matches: lb.insert(tk.END, m)
        lb.selection_clear(0, tk.END); lb.selection_set(0); lb.activate(0)

    def _on_key(event):
        if event.keysym in ('Down','Up') and ctx._sup_ac_popup and ctx._sup_ac_popup.winfo_exists():
            lb = ctx._sup_ac_list; size = lb.size()
            if size == 0: return
            cur = lb.curselection(); idx = cur[0] if cur else 0
            if event.keysym == 'Down': idx = (idx + 1) % size
            else: idx = (idx - 1) % size
            lb.selection_clear(0, tk.END); lb.selection_set(idx); lb.activate(idx)
            return 'break'
        if event.keysym == 'Return' and ctx._sup_ac_popup and ctx._sup_ac_popup.winfo_exists():
            lb = ctx._sup_ac_list; cur = lb.curselection()
            if cur:
                ctx.sup_product_var.set(lb.get(cur[0])); _hide_popup(); return 'break'
        _filter_products()

    ctx.sup_product_combo.bind('<KeyRelease>', _on_key)
    ctx.sup_product_combo.bind('<FocusOut>', lambda e: ctx.root.after(150, _hide_popup))

    def _product_chosen(event=None):
        try:
            widgets_after = [w for w in entry_bar.grid_slaves(row=1) if isinstance(w, tk.Entry)]
        except Exception:
            widgets_after = []
        for w in widgets_after:
            if hasattr(w,'cget') and w.cget('textvariable') == str(ctx.sup_size_var):
                w.focus_set(); break
    ctx.sup_product_combo.bind('<<ComboboxSelected>>', _product_chosen)

    # Labels row (added 'קטגורית בד')
    for i,lbl in enumerate(["מוצר","מידה","סוג בד","צבע בד","קטגורית בד","שם פרינט","כמות","הערה"]):
        tk.Label(entry_bar, text=lbl, bg='#f7f9fa').grid(row=0,column=i*2,sticky='w',padx=2)

    # Variant controls
    ctx.sup_size_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_size_var, width=10, state='readonly')
    ctx.sup_fabric_type_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_fabric_type_var, width=12, state='readonly')
    ctx.sup_fabric_color_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_fabric_color_var, width=10, state='readonly')
    # Fabric category combobox (values from data_processor.product_fabric_categories)
    try:
        _fabric_cat_names = [r.get('name') for r in getattr(ctx.data_processor, 'product_fabric_categories', [])]
    except Exception:
        _fabric_cat_names = []
    ctx.sup_fabric_category_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_fabric_category_var, width=12, state='readonly', values=_fabric_cat_names)
    ctx.sup_print_name_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_print_name_var, width=12, state='readonly')

    widgets = [
        ctx.sup_product_combo,
        ctx.sup_size_combo,
        ctx.sup_fabric_type_combo,
    ctx.sup_fabric_color_combo,
    ctx.sup_fabric_category_combo,
        ctx.sup_print_name_combo,
        tk.Entry(entry_bar, textvariable=ctx.sup_qty_var, width=7),
        tk.Entry(entry_bar, textvariable=ctx.sup_note_var, width=18)
    ]
    for i,w in enumerate(widgets):
        w.grid(row=1,column=i*2,sticky='w',padx=2)

    def _on_product_change(*_a):
        try:
            ctx._update_supplier_size_options()
            ctx._update_supplier_fabric_type_options()
            ctx._update_supplier_fabric_color_options()
            ctx._update_supplier_print_name_options()
        except Exception:
            pass
    try: ctx.sup_product_var.trace_add('write', _on_product_change)
    except Exception: pass

    try:
        def _on_fabric_type_change(*_a):
            try: ctx._update_supplier_fabric_color_options()
            except Exception: pass
        ctx.sup_fabric_type_var.trace_add('write', _on_fabric_type_change)
    except Exception:
        pass

    # Disable combos until product chosen
    for combo in (ctx.sup_size_combo, ctx.sup_fabric_type_combo, ctx.sup_fabric_color_combo, ctx.sup_print_name_combo):
        try: combo.state(['disabled'])
        except Exception: pass

    # Buttons
    # Shift action buttons after adding a new field
    tk.Button(entry_bar, text="➕ הוסף", command=ctx._add_supplier_line, bg='#27ae60', fg='white').grid(row=1,column=16,padx=6)
    tk.Button(entry_bar, text="🗑️ מחק נבחר", command=ctx._delete_supplier_selected, bg='#e67e22', fg='white').grid(row=1,column=17,padx=4)
    tk.Button(entry_bar, text="❌ נקה הכל", command=ctx._clear_supplier_lines, bg='#e74c3c', fg='white').grid(row=1,column=18,padx=4)

    # Lines tree
    cols = ('product','size','fabric_type','fabric_color','fabric_category','print_name','quantity','note')
    ctx.supplier_tree = ttk.Treeview(lines_frame, columns=cols, show='headings', height=10)
    headers = {'product':'מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','fabric_category':'קטגורית בד','print_name':'שם פרינט','quantity':'כמות','note':'הערה'}
    widths = {'product':160,'size':80,'fabric_type':110,'fabric_color':90,'fabric_category':120,'print_name':110,'quantity':70,'note':220}
    for c in cols:
        ctx.supplier_tree.heading(c, text=headers[c])
        ctx.supplier_tree.column(c, width=widths[c], anchor='center')
    vs = ttk.Scrollbar(lines_frame, orient='vertical', command=ctx.supplier_tree.yview)
    ctx.supplier_tree.configure(yscroll=vs.set)
    ctx.supplier_tree.pack(side='left', fill='both', expand=True, padx=(4,0), pady=4)
    vs.pack(side='right', fill='y', pady=4)

    # Transport section
    pkg_frame = ttk.LabelFrame(container, text="הובלה", padding=8)
    pkg_frame.pack(fill='x', padx=10, pady=(4,4))
    ctx.sup_pkg_type_var = tk.StringVar(value='שקית קטנה')
    ctx.sup_pkg_qty_var = tk.StringVar()
    ctx.sup_pkg_driver_var = tk.StringVar()
    tk.Label(pkg_frame, text="פריט הובלה:").grid(row=0,column=0,sticky='w',padx=4,pady=2)
    ctx.sup_pkg_type_combo = ttk.Combobox(pkg_frame, textvariable=ctx.sup_pkg_type_var, state='readonly', width=14, values=['שקית קטנה','שק','בד'])
    ctx.sup_pkg_type_combo.grid(row=0,column=1,sticky='w',padx=4,pady=2)
    tk.Label(pkg_frame, text="כמות:").grid(row=0,column=2,sticky='w',padx=4,pady=2)
    tk.Entry(pkg_frame, textvariable=ctx.sup_pkg_qty_var, width=8).grid(row=0,column=3,sticky='w',padx=4,pady=2)
    tk.Label(pkg_frame, text="שם המוביל:").grid(row=0,column=4,sticky='w',padx=4,pady=2)
    ctx.sup_pkg_driver_combo = ttk.Combobox(pkg_frame, textvariable=ctx.sup_pkg_driver_var, width=16, state='readonly')
    ctx.sup_pkg_driver_combo.grid(row=0,column=5,sticky='w',padx=4,pady=2)
    try: ctx._refresh_driver_names_for_intake()
    except Exception: pass
    tk.Button(pkg_frame, text="➕ הוסף", command=ctx._add_supplier_package_line, bg='#27ae60', fg='white').grid(row=0,column=6,padx=8)
    tk.Button(pkg_frame, text="🗑️ מחק נבחר", command=ctx._delete_selected_supplier_package, bg='#e67e22', fg='white').grid(row=0,column=7,padx=4)
    tk.Button(pkg_frame, text="❌ נקה", command=ctx._clear_supplier_packages, bg='#e74c3c', fg='white').grid(row=0,column=8,padx=4)
    ctx.sup_packages_tree = ttk.Treeview(pkg_frame, columns=('type','quantity','driver'), show='headings', height=4)
    ctx.sup_packages_tree.heading('type', text='פריט הובלה')
    ctx.sup_packages_tree.heading('quantity', text='כמות')
    ctx.sup_packages_tree.heading('driver', text='שם המוביל')
    ctx.sup_packages_tree.column('type', width=120, anchor='center')
    ctx.sup_packages_tree.column('quantity', width=70, anchor='center')
    ctx.sup_packages_tree.column('driver', width=110, anchor='center')
    ctx.sup_packages_tree.grid(row=1,column=0,columnspan=9, sticky='ew', padx=2, pady=(6,2))

    # Save + summary
    bottom_actions = tk.Frame(container, bg='#f7f9fa')
    bottom_actions.pack(fill='x', padx=10, pady=6)
    tk.Button(bottom_actions, text="💾 שמור קליטה", command=ctx._save_supplier_receipt, bg='#2c3e50', fg='white', font=('Arial',11,'bold')).pack(side='right', padx=4)
    ctx.supplier_summary_var = tk.StringVar(value="0 שורות | 0 כמות")
    tk.Label(container, textvariable=ctx.supplier_summary_var, bg='#34495e', fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')
