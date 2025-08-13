import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import json
import os

class ShipmentsTabMixin:
    """Mixin לטאב 'הובלות' המציג שורות אריזה מכל הקליטות והתעודות.

    כל רשומת קליטה / תעודת משלוח יכולה להכיל רשימת packages (package_type, quantity).
    הטאב מציג פירוק שטוח של כל החבילות עם מקורן.
    """

    def _create_shipments_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="הובלות")

        # פנימי: Notebook לתת-טאבים (סיכום / מובילים)
        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=5, pady=5)

        # --- עמוד סיכום הובלות ---
        shipments_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(shipments_page, text='סיכום הובלות')
        tk.Label(shipments_page, text="הובלות - סיכום הובלה מקליטות ותעודות", font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=6)

        toolbar = tk.Frame(shipments_page, bg='#f7f9fa')
        toolbar.pack(fill='x', padx=8, pady=(0,4))
        tk.Button(toolbar, text="🔄 רענן", command=self._refresh_shipments_table, bg='#3498db', fg='white').pack(side='right', padx=4)

        columns = ('id','kind','date','package_type','quantity','driver')
        self.shipments_tree = ttk.Treeview(shipments_page, columns=columns, show='headings', height=17)
        headers = {
            'id': 'מספר תעודה',
            'kind': 'סוג',
            'date': 'תאריך',
            'package_type': 'פריט הובלה',
            'quantity': 'כמות',
            'driver': 'שם המוביל'
        }
        widths = {'id':110,'kind':90,'date':110,'package_type':140,'quantity':80,'driver':110}
        for c in columns:
            self.shipments_tree.heading(c, text=headers[c])
            self.shipments_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(shipments_page, orient='vertical', command=self.shipments_tree.yview)
        self.shipments_tree.configure(yscroll=vs.set)
        self.shipments_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=6)
        vs.pack(side='left', fill='y', pady=6)
        self._refresh_shipments_table()

        # --- עמוד מובילים ---
        drivers_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(drivers_page, text='מובילים')

        form = tk.LabelFrame(drivers_page, text='הוספת / עדכון מוביל', bg='#f7f9fa')
        form.pack(fill='x', padx=10, pady=8)
        tk.Label(form, text='שם:', bg='#f7f9fa').grid(row=0, column=0, padx=4, pady=4, sticky='e')
        self.driver_name_var = tk.StringVar()
        tk.Entry(form, textvariable=self.driver_name_var, width=30).grid(row=0, column=1, padx=4, pady=4)
        tk.Label(form, text='טלפון:', bg='#f7f9fa').grid(row=0, column=2, padx=4, pady=4, sticky='e')
        self.driver_phone_var = tk.StringVar()
        tk.Entry(form, textvariable=self.driver_phone_var, width=20).grid(row=0, column=3, padx=4, pady=4)
        tk.Button(form, text='➕ שמור/עדכן', command=self._add_or_update_driver, bg='#27ae60', fg='white').grid(row=0, column=4, padx=6, pady=4)
        tk.Button(form, text='🗑 מחק נבחר', command=self._delete_driver, bg='#c0392b', fg='white').grid(row=0, column=5, padx=6, pady=4)

        drivers_columns = ('name','phone')
        self.drivers_tree = ttk.Treeview(drivers_page, columns=drivers_columns, show='headings', height=15)
        self.drivers_tree.heading('name', text='שם מוביל')
        self.drivers_tree.heading('phone', text='טלפון')
        self.drivers_tree.column('name', width=180, anchor='center')
        self.drivers_tree.column('phone', width=140, anchor='center')
        dscroll = ttk.Scrollbar(drivers_page, orient='vertical', command=self.drivers_tree.yview)
        self.drivers_tree.configure(yscroll=dscroll.set)
        self.drivers_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=6)
        dscroll.pack(side='left', fill='y', pady=6)

        self._load_drivers()
        self._refresh_drivers_table()

    # ---- Drivers management ----
    def _drivers_file_path(self):
        return os.path.join(os.getcwd(), 'drivers.json')

    def _load_drivers(self):
        self._drivers = []
        try:
            path = self._drivers_file_path()
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        self._drivers = [d for d in data if isinstance(d, dict) and 'name' in d]
        except Exception:
            self._drivers = []

    def _save_drivers(self):
        try:
            path = self._drivers_file_path()
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self._drivers, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror('שגיאה', f'שגיאה בשמירת מובילים: {e}')

    def _refresh_drivers_table(self):
        if not hasattr(self, 'drivers_tree'):
            return
        for iid in self.drivers_tree.get_children():
            self.drivers_tree.delete(iid)
        for d in self._drivers:
            self.drivers_tree.insert('', 'end', values=(d.get('name',''), d.get('phone','')))

    def _add_or_update_driver(self):
        name = (self.driver_name_var.get() or '').strip()
        phone = (self.driver_phone_var.get() or '').strip()
        if not name:
            messagebox.showwarning('חסר שם', 'נא להזין שם מוביל')
            return
        # עדכון אם קיים
        for d in self._drivers:
            if d.get('name') == name:
                d['phone'] = phone
                break
        else:
            self._drivers.append({'name': name, 'phone': phone})
        self._save_drivers()
        self._refresh_drivers_table()
        self.driver_name_var.set('')
        self.driver_phone_var.set('')

    def _delete_driver(self):
        sel = self.drivers_tree.selection()
        if not sel:
            return
        values = self.drivers_tree.item(sel[0], 'values')
        if not values:
            return
        name = values[0]
        self._drivers = [d for d in self._drivers if d.get('name') != name]
        self._save_drivers()
        self._refresh_drivers_table()

    # ---- Data build ----
    def _refresh_shipments_table(self):
        """רענון טבלת ההובלות ע"י איסוף כל ה-packages מכל הרשומות."""
        try:
            # לוודא טעינה עדכנית מהדיסק
            if hasattr(self.data_processor, 'refresh_supplier_receipts'):
                self.data_processor.refresh_supplier_receipts()
            supplier_intakes = getattr(self.data_processor, 'supplier_intakes', [])
            delivery_notes = getattr(self.data_processor, 'delivery_notes', [])
            rows = []
            def collect(source_list):
                for rec in source_list:
                    rec_id = rec.get('id')
                    kind = rec.get('receipt_kind') or ''
                    date_str = rec.get('date') or ''
                    # validate / normalize date for sorting
                    try:
                        sort_dt = datetime.strptime(date_str, '%Y-%m-%d')
                    except Exception:
                        try:
                            sort_dt = datetime.strptime(date_str[:10], '%Y-%m-%d')
                        except Exception:
                            sort_dt = datetime.min
                    for pkg in rec.get('packages', []) or []:
                        rows.append({
                            'rec_id': rec_id,
                            'kind': 'קליטה' if kind == 'supplier_intake' else 'הובלה' if kind == 'delivery_note' else kind,
                            'date': date_str,
                            'sort_dt': sort_dt,
                            'package_type': pkg.get('package_type',''),
                            'quantity': pkg.get('quantity',''),
                            'driver': pkg.get('driver','')
                        })
            collect(supplier_intakes)
            collect(delivery_notes)
            # מיון תאריך יורד ואז מספר תעודה יורד
            rows.sort(key=lambda r: (r['sort_dt'], r['rec_id']), reverse=True)
        except Exception:
            rows = []
        if hasattr(self, 'shipments_tree'):
            for iid in self.shipments_tree.get_children():
                self.shipments_tree.delete(iid)
            for r in rows:
                self.shipments_tree.insert('', 'end', values=(r['rec_id'], r['kind'], r['date'], r['package_type'], r['quantity'], r.get('driver','')))

    # ---- Hook from save actions ----
    def _notify_new_receipt_saved(self):
        """קריאה מהטאבים של קליטה / תעודת משלוח לאחר שמירה כדי לרענן הובלות."""
        try:
            if hasattr(self, '_refresh_shipments_table'):
                self._refresh_shipments_table()
        except Exception:
            pass
