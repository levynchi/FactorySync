import tkinter as tk
from tkinter import ttk, messagebox

# UI builder for the "שליחת בדים" sub-tab under Delivery Note

def build_fabrics_send_tab(ctx, container: tk.Frame):
    header = tk.Frame(container, bg='#f7f9fa'); header.pack(fill='x', padx=10, pady=(8,4))
    tk.Label(header, text='שליחת בדים', font=('Arial',12,'bold'), bg='#f7f9fa').pack(side='right')

    # Supplier selection (required)
    sup_bar = tk.Frame(container, bg='#f7f9fa'); sup_bar.pack(fill='x', padx=10, pady=(0,4))
    tk.Label(sup_bar, text='ספק יעד:', bg='#f7f9fa').pack(side='right', padx=(8,4))
    ctx.fs_supplier_var = tk.StringVar()
    ctx.fs_supplier_combo = ttk.Combobox(sup_bar, textvariable=ctx.fs_supplier_var, state='readonly', width=30)
    try:
        if hasattr(ctx, '_get_supplier_names'):
            ctx.fs_supplier_combo['values'] = ctx._get_supplier_names()
    except Exception:
        pass
    ctx.fs_supplier_combo.pack(side='right')

    # Barcode scan bar
    bar = tk.Frame(container, bg='#f7f9fa'); bar.pack(fill='x', padx=10, pady=(0,6))
    tk.Label(bar, text='בר קוד:', bg='#f7f9fa').pack(side='right', padx=(8,4))
    ctx.fs_barcode_var = tk.StringVar()
    entry = tk.Entry(bar, textvariable=ctx.fs_barcode_var, width=24)
    entry.pack(side='right')
    try:
        entry.bind('<Return>', lambda e: ctx._fs_add_fabric_by_barcode())
    except Exception:
        pass
    tk.Button(bar, text='➕ הוסף', command=ctx._fs_add_fabric_by_barcode, bg='#27ae60', fg='white').pack(side='right', padx=6)
    tk.Button(bar, text='🗑️ הסר נבחר', command=ctx._fs_remove_selected).pack(side='left', padx=6)
    tk.Button(bar, text='🧹 נקה הכל', command=ctx._fs_clear_all).pack(side='left')

    # Table of selected fabrics to ship
    table_wrap = tk.Frame(container, bg='#ffffff', relief='groove', bd=1)
    table_wrap.pack(fill='both', expand=True, padx=10, pady=6)
    cols = ('barcode','fabric_type','color_name','color_no','design_code','width','net_kg','meters','price','location','status')
    headers = {'barcode':'ברקוד','fabric_type':'סוג בד','color_name':'צבע','color_no':'מס׳ צבע','design_code':'Desen Kodu','width':'רוחב','net_kg':'ק"ג נטו','meters':'מטרים','price':'מחיר','location':'מיקום','status':'סטטוס'}
    widths = {'barcode':140,'fabric_type':140,'color_name':110,'color_no':80,'design_code':110,'width':60,'net_kg':80,'meters':80,'price':80,'location':90,'status':110}
    ctx.fs_tree = ttk.Treeview(table_wrap, columns=cols, show='headings', height=12)
    for c in cols:
        ctx.fs_tree.heading(c, text=headers[c])
        ctx.fs_tree.column(c, width=widths[c], anchor='center')
    vs = ttk.Scrollbar(table_wrap, orient='vertical', command=ctx.fs_tree.yview)
    ctx.fs_tree.configure(yscroll=vs.set)
    ctx.fs_tree.grid(row=0, column=0, sticky='nsew')
    vs.grid(row=0, column=1, sticky='ns')
    table_wrap.grid_rowconfigure(0, weight=1)
    table_wrap.grid_columnconfigure(0, weight=1)

    # Transport section (bottom like entry tab)
    pkg_frame = ttk.LabelFrame(container, text="הובלה", padding=8)
    pkg_frame.pack(fill='x', padx=10, pady=(0,6))
    if not hasattr(ctx, '_fs_packages'):
        ctx._fs_packages = []
    ctx.fs_pkg_type_var = tk.StringVar(value='בד')
    ctx.fs_pkg_qty_var = tk.StringVar()
    ctx.fs_pkg_driver_var = tk.StringVar()
    tk.Label(pkg_frame, text="פריט הובלה:").grid(row=0,column=0,sticky='w',padx=4,pady=2)
    ctx.fs_pkg_type_combo = ttk.Combobox(pkg_frame, textvariable=ctx.fs_pkg_type_var, state='readonly', width=14, values=['בד','שק','שקית קטנה'])
    ctx.fs_pkg_type_combo.grid(row=0,column=1,sticky='w',padx=4,pady=2)
    tk.Label(pkg_frame, text="כמות:").grid(row=0,column=2,sticky='w',padx=4,pady=2)
    tk.Entry(pkg_frame, textvariable=ctx.fs_pkg_qty_var, width=8).grid(row=0,column=3,sticky='w',padx=4,pady=2)
    tk.Label(pkg_frame, text="שם המוביל:").grid(row=0,column=4,sticky='w',padx=4,pady=2)
    ctx.fs_pkg_driver_combo = ttk.Combobox(pkg_frame, textvariable=ctx.fs_pkg_driver_var, width=16, state='readonly')
    ctx.fs_pkg_driver_combo.grid(row=0,column=5,sticky='w',padx=4,pady=2)
    try:
        # reuse same drivers as delivery
        ctx._refresh_driver_names_for_delivery()
    except Exception:
        pass
    tk.Button(pkg_frame, text="➕ הוסף", command=ctx._fs_add_package_line, bg='#27ae60', fg='white').grid(row=0,column=6,padx=8)
    tk.Button(pkg_frame, text="🗑️ מחק נבחר", command=ctx._fs_delete_selected_package, bg='#e67e22', fg='white').grid(row=0,column=7,padx=4)
    tk.Button(pkg_frame, text="❌ נקה", command=ctx._fs_clear_packages, bg='#e74c3c', fg='white').grid(row=0,column=8,padx=4)
    ctx.fs_packages_tree = ttk.Treeview(pkg_frame, columns=('type','quantity','driver'), show='headings', height=4)
    ctx.fs_packages_tree.heading('type', text='פריט הובלה')
    ctx.fs_packages_tree.heading('quantity', text='כמות')
    ctx.fs_packages_tree.heading('driver', text='שם המוביל')
    ctx.fs_packages_tree.column('type', width=120, anchor='center')
    ctx.fs_packages_tree.column('quantity', width=70, anchor='center')
    ctx.fs_packages_tree.column('driver', width=110, anchor='center')
    ctx.fs_packages_tree.grid(row=1,column=0,columnspan=9, sticky='ew', padx=2, pady=(6,2))

    actions = tk.Frame(container, bg='#f7f9fa'); actions.pack(fill='x', padx=10, pady=(0,8))
    tk.Button(actions, text='💾 שמור שליחת בדים', command=ctx._fs_save_shipment, bg='#2c3e50', fg='white', font=('Arial',11,'bold')).pack(side='right')
    ctx.fs_summary_var = tk.StringVar(value='0 בדים')
    ctx.fs_supplier_summary_var = tk.StringVar(value='ספק: -')
    status_bar = tk.Frame(container)
    status_bar.pack(fill='x', side='bottom')
    tk.Label(status_bar, textvariable=ctx.fs_summary_var, bg='#34495e', fg='white', anchor='w', padx=10).pack(fill='x', side='left', expand=True)
    tk.Label(status_bar, textvariable=ctx.fs_supplier_summary_var, bg='#2c3e50', fg='white', anchor='e', padx=10).pack(fill='x', side='right')
    # Update supplier summary on change
    try:
        def _on_sup_change(event=None):
            try:
                name = (ctx.fs_supplier_var.get() or '').strip() or '-'
                ctx.fs_supplier_summary_var.set(f"ספק: {name}")
            except Exception:
                pass
        ctx.fs_supplier_combo.bind('<<ComboboxSelected>>', _on_sup_change)
    except Exception:
        pass
