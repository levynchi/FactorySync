import tkinter as tk
from tkinter import ttk
from datetime import datetime

class ShipmentsTabMixin:
    """Mixin לטאב 'משלוחים' המציג שורות אריזה מכל הקליטות והתעודות.

    כל רשומת קליטה / תעודת משלוח יכולה להכיל רשימת packages (package_type, quantity).
    הטאב מציג פירוק שטוח של כל החבילות עם מקורן.
    """

    def _create_shipments_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="משלוחים")
        tk.Label(tab, text="משלוחים - סיכום צורות אריזה מקליטות ותעודות", font=('Arial',14,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=6)

        # Toolbar
        toolbar = tk.Frame(tab, bg='#f7f9fa')
        toolbar.pack(fill='x', padx=8, pady=(0,4))
        tk.Button(toolbar, text="🔄 רענן", command=self._refresh_shipments_table, bg='#3498db', fg='white').pack(side='right', padx=4)

        # Treeview
        columns = ('id','kind','date','package_type','quantity')
        self.shipments_tree = ttk.Treeview(tab, columns=columns, show='headings', height=18)
        headers = {
            'id': 'מספר תעודה',
            'kind': 'סוג',
            'date': 'תאריך',
            'package_type': 'צורת אריזה',
            'quantity': 'כמות'
        }
        widths = {'id':110,'kind':90,'date':110,'package_type':140,'quantity':80}
        for c in columns:
            self.shipments_tree.heading(c, text=headers[c])
            self.shipments_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(tab, orient='vertical', command=self.shipments_tree.yview)
        self.shipments_tree.configure(yscroll=vs.set)
        self.shipments_tree.pack(side='left', fill='both', expand=True, padx=(10,0), pady=6)
        vs.pack(side='left', fill='y', pady=6)

        self._refresh_shipments_table()

    # ---- Data build ----
    def _refresh_shipments_table(self):
        """רענון טבלת המשלוחים ע"י איסוף כל ה-packages מכל הרשומות."""
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
                            'kind': 'קליטה' if kind == 'supplier_intake' else 'משלוח' if kind == 'delivery_note' else kind,
                            'date': date_str,
                            'sort_dt': sort_dt,
                            'package_type': pkg.get('package_type',''),
                            'quantity': pkg.get('quantity','')
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
                self.shipments_tree.insert('', 'end', values=(r['rec_id'], r['kind'], r['date'], r['package_type'], r['quantity']))

    # ---- Hook from save actions ----
    def _notify_new_receipt_saved(self):
        """קריאה מהטאבים של קליטה / תעודת משלוח לאחר שמירה כדי לרענן משלוחים."""
        try:
            if hasattr(self, '_refresh_shipments_table'):
                self._refresh_shipments_table()
        except Exception:
            pass
