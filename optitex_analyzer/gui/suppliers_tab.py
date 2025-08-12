import tkinter as tk
from tkinter import ttk, messagebox

class SuppliersTabMixin:
    """Mixin לטאב ספקים: הוספה, מחיקה והצגה."""

    def _create_suppliers_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="ספקים")
        tk.Label(tab, text="ניהול ספקים", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)

        form = ttk.LabelFrame(tab, text="הוספת ספק", padding=10)
        form.pack(fill='x', padx=10, pady=6)

        self.sup_business_name_var = tk.StringVar()
        self.sup_first_name_var = tk.StringVar()
        self.sup_phone_var = tk.StringVar()
        self.sup_address_var = tk.StringVar()
        self.sup_business_var = tk.StringVar()
        self.sup_notes_var = tk.StringVar()

        labels = [
            ("שם עסק", self.sup_business_name_var, 18),
            ("שם פרטי", self.sup_first_name_var, 14),
            ("טלפון", self.sup_phone_var, 14),
            ("כתובת", self.sup_address_var, 25),
            ("מס' עסק", self.sup_business_var, 14),
            ("הערות", self.sup_notes_var, 25),
        ]
        for i,(lbl,var,w) in enumerate(labels):
            tk.Label(form, text=f"{lbl}:", font=('Arial',10,'bold')).grid(row=0,column=i*2,sticky='w',padx=4,pady=4)
            tk.Entry(form, textvariable=var, width=w).grid(row=0,column=i*2+1,sticky='w',padx=2,pady=4)

        tk.Button(form, text="➕ הוסף", command=self._add_supplier_record, bg='#27ae60', fg='white').grid(row=0, column=len(labels)*2, padx=8)
        tk.Button(form, text="🗑️ מחק נבחר", command=self._delete_selected_supplier, bg='#e67e22', fg='white').grid(row=0, column=len(labels)*2+1, padx=4)

        tree_frame = ttk.LabelFrame(tab, text="ספקים", padding=6)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=6)
        cols = ('id','business_name','first_name','phone','address','business_number','notes','created_at')
        self.suppliers_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=14)
        headers = {
            'id':'ID','business_name':'שם עסק','first_name':'שם פרטי','phone':'טלפון','address':'כתובת',
            'business_number':'מספר עסק','notes':'הערות','created_at':'נוצר'
        }
        widths = {'id':45,'business_name':150,'first_name':110,'phone':110,'address':200,'business_number':110,'notes':160,'created_at':140}
        for c in cols:
            self.suppliers_tree.heading(c, text=headers[c])
            self.suppliers_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(tree_frame, orient='vertical', command=self.suppliers_tree.yview)
        self.suppliers_tree.configure(yscroll=vs.set)
        self.suppliers_tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self._load_suppliers_into_tree()

    def _load_suppliers_into_tree(self):
        if not hasattr(self, 'suppliers_tree'): return
        for item in self.suppliers_tree.get_children():
            self.suppliers_tree.delete(item)
        try:
            for rec in getattr(self.data_processor, 'suppliers', []):
                self.suppliers_tree.insert('', 'end', values=(
                    rec.get('id'),
                    rec.get('business_name') or rec.get('name'),
                    rec.get('first_name',''),
                    rec.get('phone'),
                    rec.get('address'),
                    rec.get('business_number'),
                    rec.get('notes'),
                    rec.get('created_at')
                ))
        except Exception:
            pass

    def _add_supplier_record(self):
        business_name = self.sup_business_name_var.get().strip()
        if not business_name:
            messagebox.showerror("שגיאה", "חובה להזין שם עסק")
            return
        try:
            new_id = self.data_processor.add_supplier(
                business_name,
                self.sup_phone_var.get(),
                self.sup_address_var.get(),
                self.sup_business_var.get(),
                self.sup_notes_var.get(),
                self.sup_first_name_var.get(),
            )
            self.suppliers_tree.insert('', 'end', values=(new_id, business_name, self.sup_first_name_var.get(), self.sup_phone_var.get(), self.sup_address_var.get(), self.sup_business_var.get(), self.sup_notes_var.get(), ''))
            self.sup_business_name_var.set(''); self.sup_first_name_var.set(''); self.sup_phone_var.set(''); self.sup_address_var.set(''); self.sup_business_var.set(''); self.sup_notes_var.set('')
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _delete_selected_supplier(self):
        sel = self.suppliers_tree.selection()
        if not sel:
            return
        ids = []
        for item in sel:
            vals = self.suppliers_tree.item(item, 'values')
            if vals:
                ids.append(int(vals[0]))
        deleted_any = False
        for _id in ids:
            if self.data_processor.delete_supplier(_id):
                deleted_any = True
        if deleted_any:
            self._load_suppliers_into_tree()
