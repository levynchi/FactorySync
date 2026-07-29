import tkinter as tk
from tkinter import ttk
from .. import theme

# UI builder for the saved fabrics shipments list

def build_fabrics_send_list_tab(ctx, container: tk.Frame):
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
    
    header = tk.Frame(container, bg=theme.PAGE_BG)
    header.pack(fill='x', padx=10, pady=(8,4))
    tk.Label(header, text='שליחות בדים שמורות', font=(theme.FONT_FAMILY,12,'bold'), bg=theme.PAGE_BG).pack(side='right')

    # Minimal columns per request: ID (internal), Date, Count of rolls, Delete
    cols = (
        'id','date','count','delete'
    )
    table_wrap = tk.Frame(container, bg=theme.CARD_BG)
    table_wrap.pack(fill='both', expand=True, padx=10, pady=6)

    ctx.fabrics_shipments_tree = ttk.Treeview(table_wrap, columns=cols, show='headings')
    for col, txt, w in (
        ('id','ID',60),
        ('date','תאריך',140),
        ('count','מס׳ גלילים',120),
        ('delete','מחיקה',80),
    ):
        ctx.fabrics_shipments_tree.heading(col, text=txt)
        anchor = 'center'
        ctx.fabrics_shipments_tree.column(col, width=w, anchor=anchor)

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

    btns = tk.Frame(container, bg=theme.PAGE_BG)
    btns.pack(fill='x', padx=10, pady=(0,6))
    # View button similar to delivery notes list
    try:
        tk.Button(btns, text='👁 צפה', command=ctx._open_selected_fabrics_shipment_view, bg=theme.DARK, fg='white').pack(side='right', padx=(0,8))
        ctx.fabrics_shipments_tree.bind('<Double-1>', ctx._open_selected_fabrics_shipment_view)
    except Exception:
        pass
    tk.Button(btns, text='🔄 רענן', command=ctx._fs_refresh_shipments_list, bg=theme.PRIMARY, fg='white').pack(side='right')
    ctx._fs_refresh_shipments_list()
