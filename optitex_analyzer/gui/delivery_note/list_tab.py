import tkinter as tk
from tkinter import ttk, messagebox
from .. import theme

# This module defines a function that builds the Saved Delivery Notes list sub-tab.

def build_list_tab(ctx, container: tk.Frame):
    # Create a main frame with scrollbar
    main_frame = tk.Frame(container, bg=theme.PAGE_BG)
    main_frame.pack(fill='both', expand=True)
    
    # Create canvas and scrollbar
    canvas = tk.Canvas(main_frame, bg=theme.PAGE_BG)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg=theme.PAGE_BG)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Pack canvas and scrollbar
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # Configure canvas to expand properly
    def configure_scroll_region(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
        # Make sure the scrollable frame takes full width
        canvas_width = event.width
        canvas.itemconfig(canvas.find_all()[0], width=canvas_width)
    
    canvas.bind('<Configure>', configure_scroll_region)
    
    # Bind mousewheel to canvas
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    # Use scrollable_frame as the new container
    container = scrollable_frame
    
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
    # הכפתורים נמצאים במכל נפרד כדי שלא יחפפו זה את זה.
    actions = tk.Frame(container, bg=theme.PAGE_BG)
    actions.grid(row=1, column=0, sticky='e', padx=6, pady=(0, 6))
    refresh_btn = tk.Button(actions, text="🔄 רענן", command=ctx._refresh_delivery_notes_list,
                            bg=theme.PRIMARY, fg='white')
    refresh_btn.pack(side='right', padx=(0, 4))
    # כפתור צפייה בתעודה, שממנו אפשר גם לפתוח את הקובץ באקסל.
    view_btn = tk.Button(actions, text="👁 צפה", command=ctx._open_selected_delivery_note_view,
                         bg=theme.DARK, fg='white')
    view_btn.pack(side='right')
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
            try:
                if hasattr(ctx, '_refresh_label_inventory_ui'):
                    ctx._refresh_label_inventory_ui()
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
