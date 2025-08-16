import tkinter as tk
from tkinter import ttk


def build_list_tab(ctx, container: tk.Frame):
    # Add a delete column at the end for per-row deletion
    cols = ('id', 'date', 'supplier', 'total', 'packages', 'delete')
    ctx.supplier_receipts_tree = ttk.Treeview(container, columns=cols, show='headings')
    for col, txt, w in (
        ('id', 'ID', 60),
        ('date', 'תאריך', 110),
        ('supplier', 'ספק', 180),
        ('total', 'סה"כ כמות', 90),
        ('packages', 'הובלה', 160),
        ('delete', 'מחיקה', 70),
    ):
        ctx.supplier_receipts_tree.heading(col, text=txt)
        ctx.supplier_receipts_tree.column(col, width=w, anchor='center')

    vs = ttk.Scrollbar(container, orient='vertical', command=ctx.supplier_receipts_tree.yview)
    ctx.supplier_receipts_tree.configure(yscroll=vs.set)
    ctx.supplier_receipts_tree.grid(row=0, column=0, sticky='nsew', padx=6, pady=6)
    vs.grid(row=0, column=1, sticky='ns', pady=6)
    container.grid_columnconfigure(0, weight=1)
    container.grid_rowconfigure(0, weight=1)

    try:
        ctx.supplier_receipts_tree.bind('<Double-1>', lambda e: ctx._open_supplier_receipt_details())
        ctx.supplier_receipts_tree.bind('<Return>', lambda e: ctx._open_supplier_receipt_details())
        ctx.supplier_receipts_tree.bind('<Button-1>', lambda e: _on_click_list(ctx, e))
    except Exception:
        pass

    tk.Button(
        container,
        text="🔄 רענן",
        command=ctx._refresh_supplier_intake_list,
        bg='#3498db',
        fg='white'
    ).grid(row=1, column=0, sticky='e', padx=6, pady=(0, 6))

    ctx._refresh_supplier_intake_list()


def _on_click_list(ctx, event):
    # Detect clicks on delete column and trigger deletion
    tree = ctx.supplier_receipts_tree
    region = tree.identify('region', event.x, event.y)
    if region != 'cell':
        return
    col_id = tree.identify_column(event.x)
    if col_id != '#6':  # delete column index
        return
    row_id = tree.identify_row(event.y)
    if not row_id:
        return
    values = tree.item(row_id, 'values')
    if not values:
        return
    try:
        rec_id = int(values[0])
    except Exception:
        return
    try:
        import tkinter.messagebox as mbox
        if not mbox.askyesno("אישור", f"למחוק קליטה שמורה ID {rec_id}?"):
            return
    except Exception:
        pass
    # Only delete supplier_intake kind
    try:
        # Refresh to get latest kind
        ctx.data_processor.refresh_supplier_receipts()
        target = None
        for rec in ctx.data_processor.supplier_receipts:
            if int(rec.get('id', -1)) == rec_id and rec.get('receipt_kind') == 'supplier_intake':
                target = rec
                break
        if not target:
            return
        if ctx.data_processor.delete_supplier_intake(rec_id):
            # Remove from UI
            tree.delete(row_id)
            # Refresh shipments tab to remove packages from this intake
            try:
                if hasattr(ctx, '_notify_new_receipt_saved'):
                    ctx._notify_new_receipt_saved()
                elif hasattr(ctx, '_refresh_shipments_table'):
                    ctx._refresh_shipments_table()
            except Exception:
                pass
    except Exception:
        pass
