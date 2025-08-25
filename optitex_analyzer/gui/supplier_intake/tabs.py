import tkinter as tk
from tkinter import ttk
from .methods import SupplierIntakeMethodsMixin

class SupplierIntakeTabMixin(SupplierIntakeMethodsMixin):
    """Compose the Supplier Intake tab (entry + saved list)."""
    def _create_supplier_intake_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="תעודת קליטה")
        tk.Label(tab, text="תעודת קליטה (הזנה ידנית)", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)

        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=4, pady=4)
        entry_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        list_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        fabrics_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(entry_wrapper, text="קליטה")
        inner_nb.add(list_wrapper, text="קליטות שמורות")
        inner_nb.add(fabrics_wrapper, text="קליטת בדים")

        self._build_supplier_entry_tab(entry_wrapper)
        self._build_supplier_list_tab(list_wrapper)
        self._build_fabrics_intake_tab(fabrics_wrapper)

    def _build_supplier_entry_tab(self, container: tk.Frame):
        from .entry_tab import build_entry_tab
        build_entry_tab(self, container)

    def _build_supplier_list_tab(self, container: tk.Frame):
        from .list_tab import build_list_tab
        build_list_tab(self, container)

    def _build_fabrics_intake_tab(self, container: tk.Frame):
        """UI עבור 'קליטת בדים': סריקת ברקוד, תצוגת נתונים ושמירה לשינוי סטטוס במלאי הבדים."""
        header = tk.Frame(container, bg='#f7f9fa'); header.pack(fill='x', padx=10, pady=(8,4))
        tk.Label(header, text='קליטת בדים לפי ברקוד (מנתוני מלאי)', font=('Arial',12,'bold'), bg='#f7f9fa').pack(side='right')

        bar = tk.Frame(container, bg='#f7f9fa'); bar.pack(fill='x', padx=10, pady=(0,6))
        tk.Label(bar, text='בר קוד:', bg='#f7f9fa').pack(side='right', padx=(8,4))
        self.fi_barcode_var = tk.StringVar()
        entry = tk.Entry(bar, textvariable=self.fi_barcode_var, width=24)
        entry.pack(side='right')
        try:
            entry.bind('<Return>', lambda e: self._fi_add_fabric_by_barcode())
        except Exception:
            pass
        tk.Button(bar, text='➕ הוסף', command=self._fi_add_fabric_by_barcode, bg='#27ae60', fg='white').pack(side='right', padx=6)
        tk.Button(bar, text='🗑️ הסר נבחר', command=self._fi_remove_selected).pack(side='left', padx=6)
        tk.Button(bar, text='🧹 נקה הכל', command=self._fi_clear_all).pack(side='left')

        # טבלת פריטי בד שנבחרו לקליטה
        table_wrap = tk.Frame(container, bg='#ffffff', relief='groove', bd=1)
        table_wrap.pack(fill='both', expand=True, padx=10, pady=6)
        cols = ('barcode','fabric_type','color_name','color_no','design_code','width','net_kg','meters','price','location','status')
        headers = {'barcode':'ברקוד','fabric_type':'סוג בד','color_name':'צבע','color_no':'מס׳ צבע','design_code':'Desen Kodu','width':'רוחב','net_kg':'ק"ג נטו','meters':'מטרים','price':'מחיר','location':'מיקום','status':'סטטוס'}
        widths = {'barcode':140,'fabric_type':140,'color_name':110,'color_no':80,'design_code':110,'width':60,'net_kg':80,'meters':80,'price':80,'location':90,'status':110}
        self.fi_tree = ttk.Treeview(table_wrap, columns=cols, show='headings', height=12)
        for c in cols:
            self.fi_tree.heading(c, text=headers[c])
            self.fi_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(table_wrap, orient='vertical', command=self.fi_tree.yview)
        self.fi_tree.configure(yscroll=vs.set)
        self.fi_tree.grid(row=0, column=0, sticky='nsew')
        vs.grid(row=0, column=1, sticky='ns')
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        actions = tk.Frame(container, bg='#f7f9fa'); actions.pack(fill='x', padx=10, pady=(0,8))
        tk.Button(actions, text='💾 שמור קליטת בדים', command=self._fi_save_receipt, bg='#2c3e50', fg='white', font=('Arial',11,'bold')).pack(side='right')
        self.fi_summary_var = tk.StringVar(value='0 בדים')
        tk.Label(container, textvariable=self.fi_summary_var, bg='#34495e', fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')
