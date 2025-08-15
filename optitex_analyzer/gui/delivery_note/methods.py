import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class DeliveryNoteMethodsMixin:
    def _refresh_driver_names_for_delivery(self):
        """טעינת שמות המובילים לקומבובוקס תעודת משלוח."""
        try:
            import json, os
            path = os.path.join(os.getcwd(), 'drivers.json')
            names = []
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        for d in data:
                            name = (d.get('name') or '').strip()
                            if name:
                                names.append(name)
            names = sorted({n for n in names})
            if hasattr(self, 'pkg_driver_combo'):
                self.pkg_driver_combo['values'] = names
        except Exception:
            pass

        # Internal state lists
        self._delivery_lines = []
        self._packages = []

    # ---- Helpers (products / variants) ----
    def _refresh_delivery_products_allowed(self, initial: bool = False):
        try:
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
            names = sorted({ (rec.get('name') or '').strip() for rec in catalog if rec.get('name') })
            self._delivery_products_allowed = names
            self._delivery_products_allowed_full = list(names)
            if hasattr(self, 'dn_product_combo'):
                try:
                    cur = self.dn_product_var.get()
                    self.dn_product_combo['values'] = self._delivery_products_allowed_full
                    if cur and cur not in self._delivery_products_allowed_full:
                        self.dn_product_var.set('')
                except Exception: pass
        except Exception:
            self._delivery_products_allowed = []
            self._delivery_products_allowed_full = []

    def _update_delivery_size_options(self):
        product = (self.dn_product_var.get() or '').strip(); sizes = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        sz = (rec.get('size') or '').strip()
                        if sz: sizes.append(sz)
                import re
                def _size_key(s):
                    m = re.match(r"(\d+)", s); base = int(m.group(1)) if m else 0
                    return (base, s)
                sizes = sorted({s for s in sizes}, key=_size_key)
            except Exception: sizes = []
        if hasattr(self, 'dn_size_combo'):
            try:
                self.dn_size_combo['values'] = sizes
                if sizes:
                    self.dn_size_combo.state(['!disabled','readonly'])
                    if self.dn_size_var.get() not in sizes: self.dn_size_var.set('')
                else:
                    self.dn_size_var.set(''); self.dn_size_combo.set('')
                    try: self.dn_size_combo.state(['disabled'])
                    except Exception: pass
            except Exception: pass

    def _update_delivery_fabric_type_options(self):
        product = (self.dn_product_var.get() or '').strip(); fabric_types = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        ft = (rec.get('fabric_type') or '').strip()
                        if ft: fabric_types.append(ft)
                fabric_types = sorted({f for f in fabric_types})
            except Exception: fabric_types = []
        if hasattr(self, 'dn_fabric_type_combo'):
            try:
                self.dn_fabric_type_combo['values'] = fabric_types
                if fabric_types:
                    self.dn_fabric_type_combo.state(['!disabled','readonly'])
                    if self.dn_fabric_type_var.get() not in fabric_types: self.dn_fabric_type_var.set('')
                else:
                    self.dn_fabric_type_var.set(''); self.dn_fabric_type_combo.set('')
                    try: self.dn_fabric_type_combo.state(['disabled'])
                    except Exception: pass
            except Exception: pass

    def _update_delivery_fabric_color_options(self):
        product = (self.dn_product_var.get() or '').strip(); chosen_ft = (self.dn_fabric_type_var.get() or '').strip(); colors = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        if chosen_ft and (rec.get('fabric_type') or '').strip() != chosen_ft:
                            continue
                        col = (rec.get('fabric_color') or '').strip()
                        if col: colors.append(col)
                colors = sorted({c for c in colors})
            except Exception: colors = []
        if hasattr(self, 'dn_fabric_color_combo'):
            try:
                self.dn_fabric_color_combo['values'] = colors
                if colors:
                    self.dn_fabric_color_combo.state(['!disabled','readonly'])
                    if self.dn_fabric_color_var.get() not in colors:
                        self.dn_fabric_color_var.set('לבן' if 'לבן' in colors else '')
                else:
                    self.dn_fabric_color_var.set(''); self.dn_fabric_color_combo.set('')
                    try: self.dn_fabric_color_combo.state(['disabled'])
                    except Exception: pass
            except Exception: pass

    # ---- Lines ops ----
    def _add_delivery_line(self):
        product = self.dn_product_var.get().strip(); size = self.dn_size_var.get().strip(); qty_raw = self.dn_qty_var.get().strip(); note = self.dn_note_var.get().strip()
        fabric_type = self.dn_fabric_type_var.get().strip(); fabric_color = self.dn_fabric_color_var.get().strip(); print_name = self.dn_print_name_var.get().strip() or 'חלק'
        if not product or not qty_raw:
            messagebox.showerror("שגיאה", "חובה לבחור מוצר ולהזין כמות"); return
        if self._delivery_products_allowed and product not in self._delivery_products_allowed:
            messagebox.showerror("שגיאה", "יש לבחור מוצר מהרשימה בלבד"); return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי"); return
        line = {'product': product, 'size': size, 'fabric_type': fabric_type, 'fabric_color': fabric_color, 'print_name': print_name, 'quantity': qty, 'note': note}
        self._delivery_lines.append(line)
        self.delivery_tree.insert('', 'end', values=(product,size,fabric_type,fabric_color,print_name,qty,note))
        self.dn_size_var.set(''); self.dn_qty_var.set(''); self.dn_note_var.set('')
        try: self.dn_product_combo['values'] = self._delivery_products_allowed_full
        except Exception: pass
        self._update_delivery_summary()

    def _delete_delivery_selected(self):
        sel = self.delivery_tree.selection()
        if not sel: return
        all_items = self.delivery_tree.get_children(); indices = [all_items.index(i) for i in sel]
        for item in sel: self.delivery_tree.delete(item)
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._delivery_lines): del self._delivery_lines[idx]
        self._update_delivery_summary()

    def _clear_delivery_lines(self):
        self._delivery_lines = []
        for item in self.delivery_tree.get_children(): self.delivery_tree.delete(item)
        self._update_delivery_summary()

    def _update_delivery_summary(self):
        total_rows = len(self._delivery_lines); total_qty = sum(l.get('quantity',0) for l in self._delivery_lines)
        self.delivery_summary_var.set(f"{total_rows} שורות | {total_qty} כמות")

    def _save_delivery_note(self):
        supplier = self.dn_supplier_name_var.get().strip(); date_str = self.dn_date_var.get().strip()
        valid_names = set(self._get_supplier_names()) if hasattr(self,'_get_supplier_names') else set()
        if not supplier or (valid_names and supplier not in valid_names):
            messagebox.showerror("שגיאה", "יש לבחור שם ספק מהרשימה"); return
        if not self._delivery_lines:
            messagebox.showerror("שגיאה", "אין שורות לשמירה"); return
        # Confirm saving without any packages defined (parallel to supplier intake tab)
        try:
            if not self._packages:
                proceed = messagebox.askyesno(
                    "אישור",
                    "לא הוזנו צורות אריזה (שקיות / שקים / בדים).\nהאם לשמור את התעודה ללא כמות שקים?"
                )
                if not proceed:
                    return
        except Exception:
            pass
        try:
            # שימוש בשיטה החדשה לאחר פיצול הקבצים (עם נפילה אחורה)
            if hasattr(self.data_processor, 'add_delivery_note'):
                new_id = self.data_processor.add_delivery_note(supplier, date_str, self._delivery_lines, packages=self._packages)
            else:
                new_id = self.data_processor.add_supplier_receipt(supplier, date_str, self._delivery_lines, packages=self._packages, receipt_kind='delivery_note')
            messagebox.showinfo("הצלחה", f"תעודה נשמרה (ID: {new_id})")
            self._clear_delivery_lines()
            self._clear_packages()
            try:
                self._refresh_delivery_notes_list()
            except Exception:
                pass
            # עדכון טאב הובלות אם קיים
            try:
                self._notify_new_receipt_saved()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    # ---- Packages ops ----
    def _add_package_line(self):
        pkg_type = (self.pkg_type_var.get() or '').strip()
        qty_raw = (self.pkg_qty_var.get() or '').strip()
        if not pkg_type or not qty_raw:
            messagebox.showerror("שגיאה", "חובה לבחור פריט הובלה ולהזין כמות")
            return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי")
            return
        driver = (getattr(self, 'pkg_driver_var', tk.StringVar()).get() or '').strip()
        record = {'package_type': pkg_type, 'quantity': qty, 'driver': driver}
        self._packages.append(record)
        self.packages_tree.insert('', 'end', values=(pkg_type, qty, driver))
        self.pkg_qty_var.set('')

    def _delete_selected_package(self):
        sel = self.packages_tree.selection()
        if not sel: return
        all_items = self.packages_tree.get_children(); indices = [all_items.index(i) for i in sel]
        for item in sel: self.packages_tree.delete(item)
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._packages): del self._packages[idx]

    def _clear_packages(self):
        self._packages = []
        for item in self.packages_tree.get_children():
            self.packages_tree.delete(item)

    # ---- Delivery notes list ----
    def _refresh_delivery_notes_list(self):
        try:
            # Always refresh from disk to include external additions
            self.data_processor.refresh_supplier_receipts()
            notes = self.data_processor.get_delivery_notes()
        except Exception:
            notes = []
        # Clear
        if hasattr(self, 'delivery_notes_tree'):
            for iid in self.delivery_notes_tree.get_children():
                self.delivery_notes_tree.delete(iid)
            for rec in notes:
                pkg_summary = ', '.join(f"{p.get('package_type')}:{p.get('quantity')}" for p in rec.get('packages', [])[:4])
                if len(rec.get('packages', [])) > 4:
                    pkg_summary += ' ...'
                self.delivery_notes_tree.insert('', 'end', values=(rec.get('id'), rec.get('date'), rec.get('supplier'), rec.get('total_quantity'), pkg_summary))

    # ---- View single delivery note ----
    def _open_selected_delivery_note_view(self, event=None):
        try:
            if not hasattr(self, 'delivery_notes_tree'): return
            sel = self.delivery_notes_tree.selection()
            if not sel: return
            vals = self.delivery_notes_tree.item(sel[0], 'values')
            if not vals: return
            note_id = vals[0]
            # איתור הרשומה המלאה
            try:
                self.data_processor.refresh_supplier_receipts()
            except Exception:
                pass
            full_list = getattr(self.data_processor, 'delivery_notes', [])
            rec = None
            for r in full_list:
                if str(r.get('id')) == str(note_id):
                    rec = r; break
            if not rec:
                messagebox.showerror("שגיאה", "תעודה לא נמצאה")
                return
            win = tk.Toplevel(self.notebook)
            win.title(f"תעודת משלוח #{rec.get('id')}")
            win.geometry('760x520')
            win.transient(self.notebook.winfo_toplevel())
            header = tk.Frame(win, pady=6)
            header.pack(fill='x')
            def _lbl(text):
                return tk.Label(header, text=text, font=('Arial',10,'bold'))
            _lbl(f"ID: {rec.get('id')}").grid(row=0,column=0,padx=6,sticky='w')
            _lbl(f"תאריך: {rec.get('date')}").grid(row=0,column=1,padx=6,sticky='w')
            _lbl(f"ספק: {rec.get('supplier')}").grid(row=0,column=2,padx=6,sticky='w')
            # תיקון מחרוזת f לא נכונה שגרמה להצגת הטקסט {rec.get('total_quantity')} במקום הערך
            _lbl(f"סה\"כ כמות: {rec.get('total_quantity')}").grid(row=0,column=3,padx=6,sticky='w')
            # קונטיינר לפאנלים
            body = tk.PanedWindow(win, orient='vertical')
            body.pack(fill='both', expand=True, padx=8, pady=4)
            # שורות מוצרים
            lines_frame = tk.LabelFrame(body, text='שורות מוצרים')
            body.add(lines_frame, stretch='always')
            lines_cols = ('product','size','fabric_type','fabric_color','print_name','quantity','note')
            lines_tree = ttk.Treeview(lines_frame, columns=lines_cols, show='headings', height=8)
            headers_map = {'product':'מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','print_name':'פרינט','quantity':'כמות','note':'הערה'}
            widths_map = {'product':140,'size':70,'fabric_type':110,'fabric_color':110,'print_name':110,'quantity':60,'note':160}
            for c in lines_cols:
                lines_tree.heading(c, text=headers_map[c])
                lines_tree.column(c, width=widths_map[c], anchor='center')
            for line in rec.get('lines', []) or []:
                lines_tree.insert('', 'end', values=(line.get('product'), line.get('size'), line.get('fabric_type'), line.get('fabric_color'), line.get('print_name'), line.get('quantity'), line.get('note')))
            lines_tree.pack(fill='both', expand=True, padx=4, pady=4)
            # חבילות הובלה
            pkg_frame = tk.LabelFrame(body, text='פרטי הובלה / חבילות')
            body.add(pkg_frame, stretch='always')
            pkg_cols = ('package_type','quantity','driver')
            pkg_tree = ttk.Treeview(pkg_frame, columns=pkg_cols, show='headings', height=6)
            pkg_headers = {'package_type':'פריט הובלה','quantity':'כמות','driver':'מוביל'}
            pkg_widths = {'package_type':140,'quantity':70,'driver':120}
            for c in pkg_cols:
                pkg_tree.heading(c, text=pkg_headers[c])
                pkg_tree.column(c, width=pkg_widths[c], anchor='center')
            for p in rec.get('packages', []) or []:
                pkg_tree.insert('', 'end', values=(p.get('package_type'), p.get('quantity'), p.get('driver')))
            pkg_tree.pack(fill='both', expand=True, padx=4, pady=4)
            tk.Button(win, text='סגור', command=win.destroy).pack(pady=4)
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", str(e))
            except Exception:
                pass
