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

    # New fields in entry details: arrival date and supplier doc number
    tk.Label(form, text="תאריך הגעה:", font=('Arial',10,'bold')).grid(row=1,column=0,sticky='w',padx=4,pady=4)
    ctx.dn_arrival_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
    # DateEntry if available; else Entry + popup calendar
    try:
        DateEntry = None
        try:
            from tkcalendar import DateEntry  # type: ignore
        except Exception:
            DateEntry = None
        if DateEntry is not None:
            dn_arrival_entry = DateEntry(form, textvariable=ctx.dn_arrival_date_var, width=12, date_pattern='yyyy-mm-dd', locale='he_IL')
            try:
                dn_arrival_entry.set_date(datetime.now())
            except Exception:
                pass
        else:
            dn_arrival_entry = tk.Entry(form, textvariable=ctx.dn_arrival_date_var, width=12)
        dn_arrival_entry.grid(row=1,column=1,sticky='w',padx=(4,0),pady=4)
        try:
            tk.Button(form, text='📅', width=2, command=lambda e=dn_arrival_entry,v=ctx.dn_arrival_date_var: ctx._open_date_picker(e, v)).grid(row=1,column=1,sticky='w',padx=(120,0),pady=4)
        except Exception:
            pass
    except Exception:
        tk.Entry(form, textvariable=ctx.dn_arrival_date_var, width=15).grid(row=1,column=1,sticky='w',padx=4,pady=4)

    tk.Label(form, text="מס' מסמך ספק:", font=('Arial',10,'bold')).grid(row=1,column=2,sticky='w',padx=4,pady=4)
    ctx.dn_supplier_doc_number_var = tk.StringVar()
    tk.Entry(form, textvariable=ctx.dn_supplier_doc_number_var, width=18).grid(row=1,column=3,sticky='w',padx=4,pady=4)

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
    # Default print name empty; will be populated only when applicable
    ctx.dn_print_name_var = tk.StringVar(value='')
    ctx.dn_fabric_category_var = tk.StringVar()
    # Optional Barcode (visible per Main Category fields)
    ctx.dn_barcode_var = tk.StringVar()

    # Product list
    ctx._delivery_products_allowed = []
    ctx._refresh_delivery_products_allowed(initial=True)
    ctx._delivery_products_allowed_full = list(ctx._delivery_products_allowed)
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

    # Define all input widgets and drive their visibility by Main Category selection
    # Create inputs
    ctx.dn_size_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_size_var, width=10, state='readonly')
    ctx.dn_fabric_type_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_fabric_type_var, width=12, state='readonly')
    ctx.dn_fabric_color_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_fabric_color_var, width=10, state='readonly')
    # Fabric category is auto-filled from products_catalog match; show as read-only entry (not user-selectable)
    ctx.dn_fabric_category_entry = ttk.Entry(entry_bar, textvariable=ctx.dn_fabric_category_var, width=14, state='readonly')
    # Print name: selection from catalog (filtered by product or main category)
    ctx.dn_print_name_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_print_name_var, width=12, state='readonly')
    try:
        # initial population of print-name options
        if hasattr(ctx, '_refresh_delivery_print_name_options'):
            ctx._refresh_delivery_print_name_options()
    except Exception:
        pass

    # Sub Category (from categories.json)
    # Standardize default to 'גזרות שלא נתפרו' (instead of the older 'גזרות לא תפורות')
    ctx.dn_category_var = tk.StringVar(value='גזרות שלא נתפרו')
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
        # Ensure standardized default subcategory exists in values
        cats = sorted({c for c in cats} | {'גזרות שלא נתפרו'})
        ctx.dn_category_combo['values'] = cats
        # If default not present, it's already injected above; keep default selection
    except Exception:
        pass

    # Main Category (from main_categories.json via data_processor)
    ctx.dn_main_category_var = tk.StringVar(value='בגדים')
    ctx.dn_main_category_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_main_category_var, width=14, state='readonly')
    try:
        names = [c.get('name','') for c in getattr(ctx.data_processor, 'main_categories', [])]
        ctx.dn_main_category_combo['values'] = names
        # Normalize default if not in list
        if 'בגדים' in names:
            try:
                ctx.dn_main_category_var.set('בגדים')
            except Exception:
                pass
        else:
            # leave blank if not available
            pass
    except Exception:
        # fallback: direct file read
        try:
            import json, os
            path = os.path.join(os.getcwd(), 'main_categories.json')
            names = []
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    names = [ (d.get('name') or '').strip() for d in data if (d.get('name') or '').strip() ]
            ctx.dn_main_category_combo['values'] = names
            if 'בגדים' in names:
                try:
                    ctx.dn_main_category_var.set('בגדים')
                except Exception:
                    pass
        except Exception:
            pass

    # Build labels for each field and a mapping for dynamic layout
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
        'quantity': 'כמות',
        'note': 'הערה',
    }
    label_widgets = {k: tk.Label(entry_bar, text=v, bg='#f7f9fa') for k,v in label_texts.items()}

    # Create dedicated qty/note entries so we can layout them dynamically
    dn_qty_entry = tk.Entry(entry_bar, textvariable=ctx.dn_qty_var, width=7)
    dn_note_entry = tk.Entry(entry_bar, textvariable=ctx.dn_note_var, width=18)

    # Pair inputs with their labels by logical keys
    # Barcode input widget
    dn_barcode_entry = tk.Entry(entry_bar, textvariable=ctx.dn_barcode_var, width=14)

    field_pairs = {
        'main_category': (label_widgets['main_category'], ctx.dn_main_category_combo),
        'model_name': (label_widgets['model_name'], ctx.dn_product_combo),
        'sizes': (label_widgets['sizes'], ctx.dn_size_combo),
        'fabric_type': (label_widgets['fabric_type'], ctx.dn_fabric_type_combo),
        'fabric_color': (label_widgets['fabric_color'], ctx.dn_fabric_color_combo),
    'fabric_category': (label_widgets['fabric_category'], ctx.dn_fabric_category_entry),
    'print_name': (label_widgets['print_name'], ctx.dn_print_name_combo),
        'barcode': (label_widgets['barcode'], dn_barcode_entry),
        'sub_category': (label_widgets['sub_category'], ctx.dn_category_combo),
        'quantity': (label_widgets['quantity'], dn_qty_entry),
        'note': (label_widgets['note'], dn_note_entry),
    }

    # Create action buttons but position them dynamically after laying out fields
    btn_add = tk.Button(entry_bar, text="➕ הוסף", command=ctx._add_delivery_line, bg='#27ae60', fg='white')
    btn_del = tk.Button(entry_bar, text="🗑️ מחק נבחר", command=ctx._delete_delivery_selected, bg='#e67e22', fg='white')
    btn_clr = tk.Button(entry_bar, text="❌ נקה הכל", command=ctx._clear_delivery_lines, bg='#e74c3c', fg='white')

    def _find_main_category_by_name(name: str):
        try:
            for c in getattr(ctx.data_processor, 'main_categories', []) or []:
                if (c.get('name') or '').strip() == (name or '').strip():
                    return c
        except Exception:
            pass
        return None

    def _apply_layout_for_main_category():
        # Determine which logical keys to show based on selected main category fields
        selected = (ctx.dn_main_category_var.get() or '').strip()
        visible_keys = []
        # Base order for fields in the row
        order = ['main_category','model_name','sizes','fabric_type','fabric_color','fabric_category','print_name','barcode','sub_category']
        mc_has_print = False
        mc_has_barcode = False
        mc_has_sub = False
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
            mc_has_print = ('print_name' in (fields or []))
            mc_has_barcode = ('barcode' in (fields or []))
            mc_has_sub = ('sub_category' in (fields or []))
            # Always show main_category and model_name; add the rest according to configured fields
            visible_keys = ['main_category','model_name'] + [k for k in order if k in fields]
        else:
            # When no main category selected: show minimal inputs
            visible_keys = ['main_category','model_name']

        # Always include quantity and note at the end
        tail = ['quantity','note']

        # First, hide everything
        for key, (lbl, inp) in field_pairs.items():
            try:
                lbl.grid_remove(); inp.grid_remove()
            except Exception:
                pass

        # Re-grid visible fields in compact order
        col = 0
        for key in visible_keys + tail:
            lbl, inp = field_pairs[key]
            lbl.grid(row=0, column=col, sticky='w', padx=2)
            inp.grid(row=1, column=col, sticky='w', padx=2)
            col += 2

        # Position action buttons after the last field
        btn_add.grid(row=1, column=col, padx=6)
        btn_del.grid(row=1, column=col+1, padx=4)
        btn_clr.grid(row=1, column=col+2, padx=4)

        # If print_name is not part of this main category, clear its value
        try:
            if not mc_has_print:
                ctx.dn_print_name_var.set('')
        except Exception:
            pass
        # If sub_category is not part of this main category, clear its value
        try:
            if not mc_has_sub:
                ctx.dn_category_var.set('')
        except Exception:
            pass
        # If barcode is not part of this main category, clear it
        try:
            if not mc_has_barcode:
                ctx.dn_barcode_var.set('')
        except Exception:
            pass
        # Update quantity label unit after layout (in case product cleared)
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
        # Refresh allowed products by main category
        try:
            ctx._refresh_delivery_products_allowed()
        except Exception:
            pass
        # Refresh print-name options by selected main category (fallback list)
        try:
            if hasattr(ctx, '_refresh_delivery_print_name_options'):
                ctx._refresh_delivery_print_name_options()
                # If print_name isn't a field in this main category, clear it and clear list
                try:
                    mc = (ctx.dn_main_category_var.get() or '').strip()
                    has_print = False
                    if mc:
                        for c in getattr(ctx.data_processor, 'main_categories', []) or []:
                            if (c.get('name') or '').strip() == mc:
                                fields = (c.get('fields') or [])
                                if (not fields) and hasattr(ctx.data_processor, 'get_main_category_fields'):
                                    try:
                                        fields = ctx.data_processor.get_main_category_fields(c.get('id')) or []
                                    except Exception:
                                        fields = []
                                has_print = 'print_name' in (fields or [])
                                break
                    if not has_print:
                        ctx.dn_print_name_var.set('')
                        try:
                            ctx.dn_print_name_combo['values'] = []
                        except Exception:
                            pass
                    else:
                        # Clear selection if not valid under the refreshed list
                        try:
                            vals = list(ctx.dn_print_name_combo['values'] or [])
                        except Exception:
                            vals = []
                        cur = (ctx.dn_print_name_var.get() or '').strip()
                        if cur and vals and cur not in vals:
                            ctx.dn_print_name_var.set('')
                except Exception:
                    pass
        except Exception:
            pass
        # Clear product and dependent selections as their domains changed
        try:
            ctx.dn_product_var.set('')
            ctx.dn_size_var.set('')
            ctx.dn_fabric_type_var.set('')
            ctx.dn_fabric_color_var.set('')
            # keep print name as free text; keep fabric category auto blank until next compute
            ctx.dn_fabric_category_var.set('')
        except Exception:
            pass
        # Disable combos until product chosen
        for combo in (ctx.dn_size_combo, ctx.dn_fabric_type_combo, ctx.dn_fabric_color_combo):
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
        ctx.dn_main_category_var.trace_add('write', _on_main_category_change)
    except Exception:
        pass

    def _on_product_change(*_a):
        try:
            ctx._update_delivery_size_options()
            ctx._update_delivery_fabric_type_options()
            ctx._update_delivery_fabric_color_options()
            # Update print-name options to those available for this product (or fallback by main category)
            if hasattr(ctx, '_refresh_delivery_print_name_options'):
                ctx._refresh_delivery_print_name_options()
                # Respect main category: if print_name is not allowed, clear value/options
                try:
                    mc = (ctx.dn_main_category_var.get() or '').strip()
                    has_print = False
                    if mc:
                        for c in getattr(ctx.data_processor, 'main_categories', []) or []:
                            if (c.get('name') or '').strip() == mc:
                                fields = (c.get('fields') or [])
                                if (not fields) and hasattr(ctx.data_processor, 'get_main_category_fields'):
                                    try:
                                        fields = ctx.data_processor.get_main_category_fields(c.get('id')) or []
                                    except Exception:
                                        fields = []
                                has_print = 'print_name' in (fields or [])
                                break
                    if not has_print:
                        ctx.dn_print_name_var.set('')
                        try:
                            ctx.dn_print_name_combo['values'] = []
                        except Exception:
                            pass
                except Exception:
                    pass
            # also recompute fabric category when product changes
            if hasattr(ctx, '_update_delivery_fabric_category_auto'):
                ctx._update_delivery_fabric_category_auto()
        except Exception:
            pass
        # Update quantity unit label based on selected product
        try:
            _update_qty_label()
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

    # Buttons are positioned dynamically by _apply_layout_for_main_category

    # Helper to update the Quantity label with unit type from products catalog
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
            # fallback: first
            return units[0]

    def _update_qty_label():
        try:
            unit = _compute_unit_for_product(ctx.dn_product_var.get())
            lbl = label_widgets.get('quantity')
            if not lbl:
                return
            if unit:
                lbl.config(text=f"כמות ({unit})")
            else:
                lbl.config(text='כמות')
        except Exception:
            pass

    # Include a Barcode column to support categories like "בדים"; for other categories it will stay empty
    cols = ('product','size','fabric_type','fabric_color','fabric_category','print_name','barcode','category','quantity','note')
    ctx.delivery_tree = ttk.Treeview(lines_frame, columns=cols, show='headings', height=10)
    headers = {'product':'מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','fabric_category':'קטגורית בד','print_name':'שם פרינט','barcode':'בר קוד','category':'קטגוריה','quantity':'כמות','note':'הערה'}
    widths = {'product':160,'size':80,'fabric_type':110,'fabric_color':90,'fabric_category':120,'print_name':110,'barcode':110,'category':110,'quantity':70,'note':220}
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
