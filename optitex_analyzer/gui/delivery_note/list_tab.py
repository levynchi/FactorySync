import tkinter as tk
from tkinter import ttk, messagebox

# This module defines a function that builds the Saved Delivery Notes list sub-tab.

def build_list_tab(ctx, container: tk.Frame):
    # Saved delivery notes list tab
    ctx.delivery_notes_tree = ttk.Treeview(container, columns=('id','date','supplier','total','packages'), show='headings')
    for col, txt, w in (
        ('id','ID',60),('date','תאריך',110),('supplier','ספק',180),('total','סה"כ כמות',90),('packages','הובלה',140)
    ):
        ctx.delivery_notes_tree.heading(col, text=txt)
        ctx.delivery_notes_tree.column(col, width=w, anchor='center')
    vs2 = ttk.Scrollbar(container, orient='vertical', command=ctx.delivery_notes_tree.yview)
    ctx.delivery_notes_tree.configure(yscroll=vs2.set)
    ctx.delivery_notes_tree.grid(row=0,column=0,sticky='nsew', padx=6, pady=6)
    vs2.grid(row=0,column=1,sticky='ns', pady=6)
    container.grid_columnconfigure(0, weight=1)
    container.grid_rowconfigure(0, weight=1)
    refresh_btn = tk.Button(container, text="🔄 רענן", command=ctx._refresh_delivery_notes_list, bg='#3498db', fg='white')
    refresh_btn.grid(row=1,column=0,sticky='e', padx=6, pady=(0,6))
    # כפתור צפייה בתעודה
    view_btn = tk.Button(container, text="👁 צפה", command=ctx._open_selected_delivery_note_view, bg='#2c3e50', fg='white')
    view_btn.grid(row=1,column=0,sticky='e', padx=60, pady=(0,6))
    # פתיחת פירוט בדאבל-קליק על שורה
    ctx.delivery_notes_tree.bind('<Double-1>', ctx._open_selected_delivery_note_view)
    ctx._refresh_delivery_notes_list()
