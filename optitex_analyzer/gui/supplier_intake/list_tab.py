import tkinter as tk
from tkinter import ttk

def build_list_tab(ctx, container: tk.Frame):
    ctx.supplier_receipts_tree = ttk.Treeview(container, columns=('id','date','supplier','total','packages'), show='headings')
    for col, txt, w in (
        ('id','ID',60), ('date','תאריך',110), ('supplier','ספק',180), ('total','סה"כ כמות',90), ('packages','הובלה',140)
    ):
        ctx.supplier_receipts_tree.heading(col, text=txt)
        ctx.supplier_receipts_tree.column(col, width=w, anchor='center')
    vs = ttk.Scrollbar(container, orient='vertical', command=ctx.supplier_receipts_tree.yview)
    ctx.supplier_receipts_tree.configure(yscroll=vs.set)
    ctx.supplier_receipts_tree.grid(row=0,column=0,sticky='nsew', padx=6, pady=6)
    vs.grid(row=0,column=1,sticky='ns', pady=6)
    container.grid_columnconfigure(0, weight=1)
    container.grid_rowconfigure(0, weight=1)
    try:
        ctx.supplier_receipts_tree.bind('<Double-1>', lambda e: ctx._open_supplier_receipt_details())
        ctx.supplier_receipts_tree.bind('<Return>', lambda e: ctx._open_supplier_receipt_details())
    except Exception:
        pass
    tk.Button(container, text="🔄 רענן", command=ctx._refresh_supplier_intake_list, bg='#3498db', fg='white').grid(row=1,column=0,sticky='e', padx=6, pady=(0,6))
    ctx._refresh_supplier_intake_list()
