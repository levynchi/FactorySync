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

    lbls = ["מוצר","מידה","סוג בד","צבע בד","שם פרינט","כמות","הערה"]
    for i,lbl in enumerate(lbls):
        tk.Label(entry_bar, text=lbl, bg='#f7f9fa').grid(row=0,column=i*2,sticky='w',padx=2)

    ctx.dn_size_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_size_var, width=10, state='readonly')
    ctx.dn_fabric_type_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_fabric_type_var, width=12, state='readonly')
    ctx.dn_fabric_color_combo = ttk.Combobox(entry_bar, textvariable=ctx.dn_fabric_color_var, width=10, state='readonly')
    dn_print_entry = tk.Entry(entry_bar, textvariable=ctx.dn_print_name_var, width=12)

    widgets = [
        ctx.dn_product_combo,
        ctx.dn_size_combo,
        ctx.dn_fabric_type_combo,
        ctx.dn_fabric_color_combo,
        dn_print_entry,
        tk.Entry(entry_bar, textvariable=ctx.dn_qty_var, width=7),
        tk.Entry(entry_bar, textvariable=ctx.dn_note_var, width=18)
    ]
    for i,w in enumerate(widgets):
        w.grid(row=1,column=i*2,sticky='w',padx=2)

    def _on_product_change(*_a):
        try:
            ctx._update_delivery_size_options()
            ctx._update_delivery_fabric_type_options()
            ctx._update_delivery_fabric_color_options()
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
            except Exception:
                pass
        ctx.dn_fabric_type_var.trace_add('write', _on_fabric_type_change)
    except Exception:
        pass

    for combo in (ctx.dn_size_combo, ctx.dn_fabric_type_combo, ctx.dn_fabric_color_combo):
        try: combo.state(['disabled'])
        except Exception: pass
