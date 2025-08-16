import tkinter as tk
from tkinter import ttk, messagebox

# This module defines a function that builds the Saved Delivery Notes list sub-tab.

def build_list_tab(ctx, container: tk.Frame):
    # Saved delivery notes list tab with per-row delete column
    columns = ('id','date','supplier','total','packages','delete')
    ctx.delivery_notes_tree = ttk.Treeview(container, columns=columns, show='headings')
    for col, txt, w in (
        ('id','ID',60),('date','תאריך',110),('supplier','ספק',180),('total','סה"כ כמות',90),('packages','הובלה',140),('delete','מחיקה',70)
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
    # מחיקה בלחיצה על עמודת האייקון
    ctx.delivery_notes_tree.bind('<Button-1>', lambda e: _on_click_delete(ctx, e))
    ctx._refresh_delivery_notes_list()


def _on_click_delete(ctx, event):
    tree = ctx.delivery_notes_tree
    region = tree.identify('region', event.x, event.y)
    if region != 'cell':
        return
    col_id = tree.identify_column(event.x)
    # delete column is the 6th (#6)
    if col_id != '#6':
        return
    row_id = tree.identify_row(event.y)
    if not row_id:
        return
    values = tree.item(row_id, 'values')
    if not values:
        return
    try:
        note_id = int(values[0])
    except Exception:
        return
    try:
        if not messagebox.askyesno("אישור", f"למחוק תעודת משלוח שמורה ID {note_id}?"):
            return
    except Exception:
        pass
    try:
        # וידוא שזה באמת delivery_note
        ctx.data_processor.refresh_supplier_receipts()
        target = None
        for rec in ctx.data_processor.delivery_notes:
            if int(rec.get('id', -1)) == note_id:
                target = rec; break
        if not target:
            return
        if hasattr(ctx.data_processor, 'delete_delivery_note'):
            ok = ctx.data_processor.delete_delivery_note(note_id)
        else:
            # תאימות לאחור לא צפויה להידרש, אך נשאיר נתיב שמירה
            ok = False
        if ok:
            tree.delete(row_id)
            # רענון טאב הובלות כדי להסיר חבילות מאותה תעודה שנמחקה
            try:
                if hasattr(ctx, '_notify_new_receipt_saved'):
                    ctx._notify_new_receipt_saved()
                elif hasattr(ctx, '_refresh_shipments_table'):
                    ctx._refresh_shipments_table()
            except Exception:
                pass
            # הודעה למשתמש על מחיקת ההובלות המשויכות
            try:
                messagebox.showinfo(
                    "נמחק",
                    f"תעודת משלוח {note_id} נמחקה.\nפריטי ההובלה (חבילות) מתעודה זו הוסרו מסיכום הובלות."
                )
            except Exception:
                pass
    except Exception:
        pass
