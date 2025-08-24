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

    # New fields: Arrival date and supplier document number
    tk.Label(form, text="תאריך הגעה:", font=('Arial',10,'bold')).grid(row=1,column=0,sticky='w',padx=4,pady=4)
    ctx.supplier_arrival_date_var = tk.StringVar()
    try:
        from datetime import datetime
        ctx.supplier_arrival_date_var.set(datetime.now().strftime('%Y-%m-%d'))
    except Exception:
        ctx.supplier_arrival_date_var.set('')
    tk.Entry(form, textvariable=ctx.supplier_arrival_date_var, width=15).grid(row=1,column=1,sticky='w',padx=4,pady=4)

    tk.Label(form, text="מס' מסמך ספק:", font=('Arial',10,'bold')).grid(row=1,column=2,sticky='w',padx=4,pady=4)
    ctx.supplier_doc_number_var = tk.StringVar()
    tk.Entry(form, textvariable=ctx.supplier_doc_number_var, width=18).grid(row=1,column=3,sticky='w',padx=4,pady=4)

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
    # Optional Barcode for categories like Fabrics
    ctx.sup_barcode_var = tk.StringVar()
    # Returned from drawing + Drawing ID
    ctx.sup_returned_from_drawing_var = tk.StringVar(value='לא')
    ctx.sup_drawing_id_var = tk.StringVar()

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

    # Variant controls
    ctx.sup_size_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_size_var, width=10, state='readonly')
    ctx.sup_fabric_type_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_fabric_type_var, width=12, state='readonly')
    ctx.sup_fabric_color_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_fabric_color_var, width=10, state='readonly')
    # Fabric category is auto-filled from products_catalog; show as read-only entry
    ctx.sup_fabric_category_entry = ttk.Entry(entry_bar, textvariable=ctx.sup_fabric_category_var, width=14, state='readonly')
    ctx.sup_print_name_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_print_name_var, width=12, state='readonly')
    # Sub Category (from categories.json) – default to 'בגדים תפורים'
    ctx.sup_category_var = tk.StringVar(value='בגדים תפורים')
    ctx.sup_category_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_category_var, width=14, state='readonly')
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
        cats = sorted({c for c in cats} | {'בגדים תפורים'})
        # Normalize default if exists in list
        try:
            if 'בגדים תפורים' in cats:
                ctx.sup_category_var.set('בגדים תפורים')
        except Exception:
            pass
        ctx.sup_category_combo['values'] = cats
    except Exception:
        pass
    # Returned from drawing selector and Drawing ID selection (from drawings_data.json)
    ctx.sup_returned_from_drawing_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_returned_from_drawing_var, width=10, state='readonly', values=['לא','כן'])
    ctx.sup_drawing_id_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_drawing_id_var, width=10, state='disabled')
    # Populate drawing IDs from data_processor
    try:
        drawings = getattr(ctx.data_processor, 'drawings_data', []) or []
        drawing_ids = [str(d.get('id')) for d in drawings if d.get('id') is not None]
        ctx.sup_drawing_id_combo['values'] = drawing_ids
    except Exception:
        pass

    # Main Category (values from data_processor.main_categories) – default 'בגדים'
    ctx.sup_main_category_var = tk.StringVar(value='בגדים')
    ctx.sup_main_category_combo = ttk.Combobox(entry_bar, textvariable=ctx.sup_main_category_var, width=14, state='readonly')
    try:
        names = [c.get('name','') for c in getattr(ctx.data_processor, 'main_categories', [])]
        ctx.sup_main_category_combo['values'] = names
        if 'בגדים' in names:
            try:
                ctx.sup_main_category_var.set('בגדים')
            except Exception:
                pass
    except Exception:
        try:
            import json, os
            path = os.path.join(os.getcwd(), 'main_categories.json')
            names = []
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    names = [ (d.get('name') or '').strip() for d in data if (d.get('name') or '').strip() ]
            ctx.sup_main_category_combo['values'] = names
            if 'בגדים' in names:
                try:
                    ctx.sup_main_category_var.set('בגדים')
                except Exception:
                    pass
        except Exception:
            pass

    # Labels and dynamic layout like Delivery Note
    label_texts = {
        'main_category': 'קטגוריה ראשית',
        'model_name': 'מוצר',
        'sizes': 'מידה',
        'fabric_type': 'סוג בד',
        'fabric_color': 'צבע בד',
        'fabric_category': 'קטגורית בד',
        'print_name': 'שם פרינט',
        'barcode': 'בר קוד',
        'sub_category': 'תת קטגוריה',
        'returned_from_drawing': 'חזר מציור',
        'drawing_id': "מס' ציור",
        'quantity': 'כמות',
        'note': 'הערה',
    }
    label_widgets = {k: tk.Label(entry_bar, text=v, bg='#f7f9fa') for k,v in label_texts.items()}

    qty_entry = tk.Entry(entry_bar, textvariable=ctx.sup_qty_var, width=7)
    note_entry = tk.Entry(entry_bar, textvariable=ctx.sup_note_var, width=18)

    # Barcode input
    sup_barcode_entry = tk.Entry(entry_bar, textvariable=ctx.sup_barcode_var, width=14)

    field_pairs = {
        'main_category': (label_widgets['main_category'], ctx.sup_main_category_combo),
        'model_name': (label_widgets['model_name'], ctx.sup_product_combo),
        'sizes': (label_widgets['sizes'], ctx.sup_size_combo),
        'fabric_type': (label_widgets['fabric_type'], ctx.sup_fabric_type_combo),
        'fabric_color': (label_widgets['fabric_color'], ctx.sup_fabric_color_combo),
        'fabric_category': (label_widgets['fabric_category'], ctx.sup_fabric_category_entry),
        'print_name': (label_widgets['print_name'], ctx.sup_print_name_combo),
        'barcode': (label_widgets['barcode'], sup_barcode_entry),
        'sub_category': (label_widgets['sub_category'], ctx.sup_category_combo),
        'returned_from_drawing': (label_widgets['returned_from_drawing'], ctx.sup_returned_from_drawing_combo),
        'drawing_id': (label_widgets['drawing_id'], ctx.sup_drawing_id_combo),
        'quantity': (label_widgets['quantity'], qty_entry),
        'note': (label_widgets['note'], note_entry),
    }

    # Action buttons placed after fields dynamically
    btn_add = tk.Button(entry_bar, text="➕ הוסף", command=ctx._add_supplier_line, bg='#27ae60', fg='white')
    btn_del = tk.Button(entry_bar, text="🗑️ מחק נבחר", command=ctx._delete_supplier_selected, bg='#e67e22', fg='white')
    btn_clr = tk.Button(entry_bar, text="❌ נקה הכל", command=ctx._clear_supplier_lines, bg='#e74c3c', fg='white')

    def _find_main_category_by_name(name: str):
        try:
            for c in getattr(ctx.data_processor, 'main_categories', []) or []:
                if (c.get('name') or '').strip() == (name or '').strip():
                    return c
        except Exception:
            pass
        return None

    def _apply_layout_for_main_category():
        selected = (ctx.sup_main_category_var.get() or '').strip()
        order = ['main_category','model_name','sizes','fabric_type','fabric_color','fabric_category','print_name','barcode']
        mc_has_barcode = False
        if selected:
            rec = _find_main_category_by_name(selected)
            fields = []
            if rec:
                fields = (rec.get('fields') or [])
                if not fields and hasattr(ctx.data_processor, 'get_main_category_fields'):
                    try:
                        fields = ctx.data_processor.get_main_category_fields(rec.get('id')) or []
                    except Exception:
                        fields = []
            mc_has_barcode = ('barcode' in (fields or []))
            visible_keys = ['main_category','model_name'] + [k for k in order if k in fields]
        else:
            visible_keys = ['main_category','model_name']

        # Always keep sub_category, returned-from-drawing, drawing_id, quantity, note
        tail = ['sub_category','returned_from_drawing','drawing_id','quantity','note']

        # Hide all first
        for key,(lbl,inp) in field_pairs.items():
            try:
                lbl.grid_remove(); inp.grid_remove()
            except Exception:
                pass

        col = 0
        for key in visible_keys + tail:
            lbl, inp = field_pairs[key]
            lbl.grid(row=0, column=col, sticky='w', padx=2)
            inp.grid(row=1, column=col, sticky='w', padx=2)
            col += 2

        btn_add.grid(row=1, column=col, padx=6)
        btn_del.grid(row=1, column=col+1, padx=4)
        btn_clr.grid(row=1, column=col+2, padx=4)

        # If barcode isn't enabled for this main category, clear it
        try:
            if not mc_has_barcode:
                ctx.sup_barcode_var.set('')
        except Exception:
            pass
        # Update qty label text (unit) after layout
        try:
            _update_qty_label()
        except Exception:
            pass

    # Initial layout
    _apply_layout_for_main_category()

    # React to main category changes: update layout and filter products + reset dependent fields
    def _on_main_category_change(*_):
        try:
            _apply_layout_for_main_category()
        except Exception:
            pass
        try:
            ctx._refresh_supplier_products_allowed()
        except Exception:
            pass
        try:
            ctx.sup_product_var.set('')
            ctx.sup_size_var.set('')
            ctx.sup_fabric_type_var.set('')
            ctx.sup_fabric_color_var.set('')
            ctx.sup_print_name_var.set('')
            ctx.sup_fabric_category_var.set('')
        except Exception:
            pass
        for combo in (ctx.sup_size_combo, ctx.sup_fabric_type_combo, ctx.sup_fabric_color_combo, ctx.sup_print_name_combo):
            try:
                combo.set('')
                combo.state(['disabled'])
            except Exception:
                pass
        # Reset qty label (product cleared)
        try:
            _update_qty_label()
        except Exception:
            pass

    try:
        ctx.sup_main_category_var.trace_add('write', _on_main_category_change)
    except Exception:
        pass

    # Enable/disable drawing ID entry by toggle
    def _toggle_drawing_id(*_a):
        try:
            if ctx.sup_returned_from_drawing_var.get() == 'כן':
                # refresh values in case drawings list changed
                try:
                    drawings = getattr(ctx.data_processor, 'drawings_data', []) or []
                    drawing_ids = [str(d.get('id')) for d in drawings if d.get('id') is not None]
                    ctx.sup_drawing_id_combo['values'] = drawing_ids
                except Exception:
                    pass
                ctx.sup_drawing_id_combo.config(state='readonly')
            else:
                ctx.sup_drawing_id_var.set('')
                ctx.sup_drawing_id_combo.config(state='disabled')
        except Exception:
            pass
    try:
        ctx.sup_returned_from_drawing_var.trace_add('write', _toggle_drawing_id)
    except Exception:
        pass

    def _on_product_change(*_a):
        try:
            ctx._update_supplier_size_options()
            ctx._update_supplier_fabric_type_options()
            ctx._update_supplier_fabric_color_options()
            ctx._update_supplier_print_name_options()
        except Exception:
            pass
        # Update quantity unit label based on product
        try:
            _update_qty_label()
        except Exception:
            pass
    try:
        ctx.sup_product_var.trace_add('write', _on_product_change)
        # also recompute fabric category when product changes
        ctx.sup_product_var.trace_add('write', lambda *_: getattr(ctx, '_update_supplier_fabric_category_auto', lambda: None)())
    except Exception: pass

    try:
        def _on_fabric_type_change(*_a):
            try: ctx._update_supplier_fabric_color_options()
            except Exception: pass
        ctx.sup_fabric_type_var.trace_add('write', _on_fabric_type_change)
        ctx.sup_fabric_type_var.trace_add('write', lambda *_: getattr(ctx, '_update_supplier_fabric_category_auto', lambda: None)())
    except Exception:
        pass

    # Also update fabric category on size/color/print changes
    try:
        ctx.sup_size_var.trace_add('write', lambda *_: getattr(ctx, '_update_supplier_fabric_category_auto', lambda: None)())
    except Exception: pass
    try:
        ctx.sup_fabric_color_var.trace_add('write', lambda *_: getattr(ctx, '_update_supplier_fabric_category_auto', lambda: None)())
    except Exception: pass
    try:
        ctx.sup_print_name_var.trace_add('write', lambda *_: getattr(ctx, '_update_supplier_fabric_category_auto', lambda: None)())
    except Exception: pass

    # Disable combos until product chosen
    for combo in (ctx.sup_size_combo, ctx.sup_fabric_type_combo, ctx.sup_fabric_color_combo, ctx.sup_print_name_combo):
        try: combo.state(['disabled'])
        except Exception: pass

    # Buttons placed dynamically above with the layout

    # Helper: compute and show unit type next to Quantity label
    def _compute_unit_for_product(name: str) -> str:
        try:
            catalog = getattr(ctx.data_processor, 'products_catalog', []) or []
        except Exception:
            catalog = []
        name = (name or '').strip()
        if not name:
            return ''
        units = []
        try:
            for rec in catalog:
                if (rec.get('name') or '').strip() == name:
                    ut = (rec.get('unit_type') or '').strip()
                    if ut:
                        units.append(ut)
        except Exception:
            pass
        if not units:
            return ''
        try:
            from collections import Counter
            return Counter(units).most_common(1)[0][0]
        except Exception:
            return units[0]

    def _update_qty_label():
        try:
            unit = _compute_unit_for_product(ctx.sup_product_var.get())
            lbl = label_widgets.get('quantity')
            if not lbl:
                return
            if unit:
                lbl.config(text=f"כמות ({unit})")
            else:
                lbl.config(text='כמות')
        except Exception:
            pass

    # Lines tree
    cols = ('product','size','fabric_type','fabric_color','fabric_category','print_name','barcode','category','returned_from_drawing','drawing_id','quantity','note')
    ctx.supplier_tree = ttk.Treeview(lines_frame, columns=cols, show='headings', height=10)
    headers = {
        'product':'מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד',
        'fabric_category':'קטגורית בד','print_name':'שם פרינט','barcode':'בר קוד','category':'קטגוריה',
        'returned_from_drawing':'חזר מציור','drawing_id':"מס' ציור",
        'quantity':'כמות','note':'הערה'
    }
    widths = {'product':160,'size':80,'fabric_type':110,'fabric_color':90,'fabric_category':120,'print_name':110,'barcode':110,'category':110,'returned_from_drawing':90,'drawing_id':80,'quantity':70,'note':220}
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
