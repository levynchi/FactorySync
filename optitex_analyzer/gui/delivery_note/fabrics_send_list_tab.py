import tkinter as tk
from tkinter import ttk

# UI builder for the saved fabrics shipments list

def build_fabrics_send_list_tab(ctx, container: tk.Frame):
    header = tk.Frame(container, bg='#f7f9fa')
    header.pack(fill='x', padx=10, pady=(8,4))
    tk.Label(header, text='שליחות בדים שמורות', font=('Arial',12,'bold'), bg='#f7f9fa').pack(side='right')

    cols = ('id','date','count','packages','delete')
    table_wrap = tk.Frame(container, bg='#ffffff')
    table_wrap.pack(fill='both', expand=True, padx=10, pady=6)

    ctx.fabrics_shipments_tree = ttk.Treeview(table_wrap, columns=cols, show='headings')
    for col, txt, w in (
        ('id','ID',60),
        ('date','תאריך',110),
        ('count','מס׳ ברקודים',100),
        ('packages','הובלה',160),
        ('delete','מחיקה',70),
    ):
        ctx.fabrics_shipments_tree.heading(col, text=txt)
        ctx.fabrics_shipments_tree.column(col, width=w, anchor='center')

    vs = ttk.Scrollbar(table_wrap, orient='vertical', command=ctx.fabrics_shipments_tree.yview)
    ctx.fabrics_shipments_tree.configure(yscroll=vs.set)
    ctx.fabrics_shipments_tree.grid(row=0, column=0, sticky='nsew')
    vs.grid(row=0, column=1, sticky='ns')
    table_wrap.grid_columnconfigure(0, weight=1)
    table_wrap.grid_rowconfigure(0, weight=1)

    try:
        ctx.fabrics_shipments_tree.bind('<Button-1>', lambda e: ctx._fs_on_click_shipments(e))
    except Exception:
        pass

    btns = tk.Frame(container, bg='#f7f9fa')
    btns.pack(fill='x', padx=10, pady=(0,6))
    tk.Button(btns, text='🔄 רענן', command=ctx._fs_refresh_shipments_list, bg='#3498db', fg='white').pack(side='right')
    ctx._fs_refresh_shipments_list()
