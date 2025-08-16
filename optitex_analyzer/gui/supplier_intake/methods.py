import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class SupplierIntakeMethodsMixin:
    def _refresh_driver_names_for_intake(self):
        """טעינת שמות המובילים מקובץ drivers.json והחלתם בקומבובוקס."""
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
            if hasattr(self, 'sup_pkg_driver_combo'):
                self.sup_pkg_driver_combo['values'] = names
        except Exception:
            pass

    def _refresh_supplier_products_allowed(self, initial: bool = False):
        """רענון רשימת המוצרים האפשריים מתוך קטלוג המוצרים."""
        try:
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
            names = sorted({ (rec.get('name') or '').strip() for rec in catalog if rec.get('name') })
            self._supplier_products_allowed = names
            self._supplier_products_allowed_full = list(names)
            if hasattr(self, 'sup_product_combo'):
                try:
                    current = self.sup_product_var.get()
                    self.sup_product_combo['values'] = self._supplier_products_allowed_full
                    if current and current not in self._supplier_products_allowed_full:
                        self.sup_product_var.set('')
                except Exception:
                    pass
        except Exception:
            self._supplier_products_allowed = []
            self._supplier_products_allowed_full = []

    def _update_supplier_size_options(self):
        """עדכון רשימת המידות (וריאנטים) עבור המוצר שנבחר מקטלוג המוצרים."""
        product = (self.sup_product_var.get() or '').strip()
        sizes = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        sz = (rec.get('size') or '').strip()
                        if sz:
                            sizes.append(sz)
                import re
                def _size_key(s):
                    m = re.match(r"(\d+)", s)
                    base = int(m.group(1)) if m else 0
                    return (base, s)
                sizes = sorted({s for s in sizes}, key=_size_key)
            except Exception:
                sizes = []
        if hasattr(self, 'sup_size_combo'):
            try:
                self.sup_size_combo['values'] = sizes
                if sizes:
                    self.sup_size_combo.state(['!disabled','readonly'])
                    if self.sup_size_var.get() not in sizes:
                        self.sup_size_var.set('')
                else:
                    self.sup_size_var.set('')
                    self.sup_size_combo.set('')
                    try:
                        self.sup_size_combo.state(['disabled'])
                    except Exception:
                        pass
            except Exception:
                pass

    def _update_supplier_fabric_type_options(self):
        """עדכון רשימת סוגי הבד הזמינים עבור המוצר הנבחר (מהווריאנטים בקטלוג)."""
        product = (self.sup_product_var.get() or '').strip()
        fabric_types = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        ft = (rec.get('fabric_type') or '').strip()
                        if ft:
                            fabric_types.append(ft)
                fabric_types = sorted({f for f in fabric_types})
            except Exception:
                fabric_types = []
        if hasattr(self, 'sup_fabric_type_combo'):
            try:
                self.sup_fabric_type_combo['values'] = fabric_types
                if fabric_types:
                    self.sup_fabric_type_combo.state(['!disabled','readonly'])
                    if self.sup_fabric_type_var.get() not in fabric_types:
                        self.sup_fabric_type_var.set('')
                else:
                    self.sup_fabric_type_var.set('')
                    self.sup_fabric_type_combo.set('')
                    try:
                        self.sup_fabric_type_combo.state(['disabled'])
                    except Exception:
                        pass
            except Exception:
                pass

    def _update_supplier_fabric_color_options(self):
        """עדכון רשימת צבעי הבד הזמינים עבור המוצר / סוג בד שנבחרו."""
        product = (self.sup_product_var.get() or '').strip()
        chosen_ft = (self.sup_fabric_type_var.get() or '').strip()
        colors = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        if chosen_ft and (rec.get('fabric_type') or '').strip() != chosen_ft:
                            continue
                        col = (rec.get('fabric_color') or '').strip()
                        if col:
                            colors.append(col)
                colors = sorted({c for c in colors})
            except Exception:
                colors = []
        if hasattr(self, 'sup_fabric_color_combo'):
            try:
                self.sup_fabric_color_combo['values'] = colors
                if colors:
                    self.sup_fabric_color_combo.state(['!disabled','readonly'])
                    if self.sup_fabric_color_var.get() not in colors:
                        self.sup_fabric_color_var.set('לבן' if 'לבן' in colors else '')
                else:
                    self.sup_fabric_color_var.set('')
                    self.sup_fabric_color_combo.set('')
                    try:
                        self.sup_fabric_color_combo.state(['disabled'])
                    except Exception:
                        pass
            except Exception:
                pass

    def _update_supplier_print_name_options(self):
        """עדכון רשימת שמות פרינט בהתאם למוצר שנבחר."""
        product = (self.sup_product_var.get() or '').strip()
        names = []
        if product:
            try:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        pn = (rec.get('print_name') or '').strip()
                        if pn:
                            names.append(pn)
                names = sorted({n for n in names})
                if not names:
                    try:
                        names = [r.get('name') for r in getattr(self.data_processor, 'product_print_names', []) if r.get('name')]
                    except Exception:
                        names = []
            except Exception:
                names = []
        if hasattr(self, 'sup_print_name_combo'):
            try:
                if product:
                    self.sup_print_name_combo['values'] = names
                    if names:
                        try:
                            self.sup_print_name_combo.state(['!disabled','readonly'])
                        except Exception:
                            pass
                        current = (self.sup_print_name_var.get() or '').strip()
                        if 'חלק' in names:
                            if current not in names:
                                self.sup_print_name_var.set('חלק')
                        else:
                            if current not in names:
                                self.sup_print_name_var.set(names[0] if names else '')
                    else:
                        self.sup_print_name_var.set('')
                        self.sup_print_name_combo.set('')
                        try:
                            self.sup_print_name_combo.state(['disabled'])
                        except Exception:
                            pass
                else:
                    self.sup_print_name_var.set('')
                    self.sup_print_name_combo.set('')
                    self.sup_print_name_combo['values'] = []
                    try:
                        self.sup_print_name_combo.state(['disabled'])
                    except Exception:
                        pass
            except Exception:
                pass

    def _refresh_supplier_print_name_options(self):
        """טעינת רשימת שמות הפרינט מקובץ print_names.json דרך ה-DataProcessor."""
        try:
            names = [r.get('name') for r in getattr(self.data_processor, 'product_print_names', []) if r.get('name')]
        except Exception:
            names = []
        if hasattr(self, 'sup_print_name_combo'):
            try:
                self.sup_print_name_combo['values'] = names
                if names:
                    try:
                        self.sup_print_name_combo.state(['!disabled','readonly'])
                    except Exception:
                        pass
                    current = (self.sup_print_name_var.get() or '').strip()
                    if 'חלק' in names:
                        if current not in names:
                            self.sup_print_name_var.set('חלק')
                    else:
                        if current not in names:
                            self.sup_print_name_var.set(names[0] if names else '')
                else:
                    self.sup_print_name_var.set('')
                    self.sup_print_name_combo.set('')
                    try:
                        self.sup_print_name_combo.state(['disabled'])
                    except Exception:
                        pass
            except Exception:
                pass

    def _add_supplier_line(self):
        product = self.sup_product_var.get().strip(); size = self.sup_size_var.get().strip(); qty_raw = self.sup_qty_var.get().strip(); note = self.sup_note_var.get().strip()
        fabric_type = self.sup_fabric_type_var.get().strip(); fabric_color = self.sup_fabric_color_var.get().strip(); print_name = self.sup_print_name_var.get().strip() or 'חלק'
        fabric_category = getattr(self, 'sup_fabric_category_var', None)
        fabric_category = fabric_category.get().strip() if fabric_category else ''
        if not product or not qty_raw:
            messagebox.showerror("שגיאה", "חובה לבחור מוצר ולהזין כמות")
            return
        if hasattr(self, '_supplier_products_allowed') and self._supplier_products_allowed and product not in self._supplier_products_allowed:
            messagebox.showerror("שגיאה", "יש לבחור מוצר מהרשימה בלבד"); return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי"); return
        returned_from_drawing = (getattr(self, 'sup_returned_from_drawing_var', tk.StringVar(value='לא')).get() or 'לא').strip()
        drawing_id = (getattr(self, 'sup_drawing_id_var', tk.StringVar()).get() or '').strip()
        # if not returned, ignore drawing_id
        if returned_from_drawing != 'כן':
            drawing_id = ''
        line = {
            'product': product,
            'size': size,
            'fabric_type': fabric_type,
            'fabric_color': fabric_color,
            'fabric_category': fabric_category,
            'print_name': print_name,
            'returned_from_drawing': returned_from_drawing,
            'drawing_id': drawing_id,
            'quantity': qty,
            'note': note
        }
        self._supplier_lines.append(line)
        self.supplier_tree.insert('', 'end', values=(product,size,fabric_type,fabric_color,fabric_category,print_name,returned_from_drawing,drawing_id,qty,note))
        self.sup_size_var.set(''); self.sup_qty_var.set(''); self.sup_note_var.set('')
        try:
            self.sup_returned_from_drawing_var.set('לא')
            self.sup_drawing_id_var.set('')
        except Exception:
            pass
        if hasattr(self, 'sup_product_combo'):
            try:
                self.sup_product_combo['values'] = self._supplier_products_allowed_full
            except Exception:
                pass
        self._update_supplier_summary()

    def _delete_supplier_selected(self):
        sel = self.supplier_tree.selection()
        if not sel: return
        all_items = self.supplier_tree.get_children(); indices = [all_items.index(i) for i in sel]
        for item in sel: self.supplier_tree.delete(item)
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._supplier_lines): del self._supplier_lines[idx]
        self._update_supplier_summary()

    def _clear_supplier_lines(self):
        self._supplier_lines = []
        for item in self.supplier_tree.get_children(): self.supplier_tree.delete(item)
        self._update_supplier_summary()

    def _update_supplier_summary(self):
        total_rows = len(self._supplier_lines)
        total_qty = sum(l.get('quantity',0) for l in self._supplier_lines)
        self.supplier_summary_var.set(f"{total_rows} שורות | {total_qty} כמות")

    def _save_supplier_receipt(self):
        supplier = self.supplier_name_var.get().strip(); date_str = self.supplier_date_var.get().strip()
        valid_names = set(self._get_supplier_names()) if hasattr(self,'_get_supplier_names') else set()
        if not supplier or (valid_names and supplier not in valid_names):
            messagebox.showerror("שגיאה", "יש לבחור שם ספק מהרשימה"); return
        if not self._supplier_lines: messagebox.showerror("שגיאה", "אין שורות לשמירה"); return
        try:
            if not self._supplier_packages:
                proceed = messagebox.askyesno(
                    "אישור",
                    "לא הוזנו פריטי הובלה (שקיות / שקים / בדים).\nהאם לשמור את הקליטה ללא פריטי הובלה?"
                )
                if not proceed:
                    return
        except Exception:
            pass
        try:
            if hasattr(self.data_processor, 'add_supplier_intake'):
                new_id = self.data_processor.add_supplier_intake(supplier, date_str, self._supplier_lines, packages=self._supplier_packages)
            else:
                new_id = self.data_processor.add_supplier_receipt(supplier, date_str, self._supplier_lines, packages=self._supplier_packages, receipt_kind='supplier_intake')
            messagebox.showinfo("הצלחה", f"קליטה נשמרה (ID: {new_id})")
            self._clear_supplier_lines()
            self._clear_supplier_packages()
            try:
                self._refresh_supplier_intake_list()
            except Exception:
                pass
            try:
                self._notify_new_receipt_saved()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _update_supplier_fabric_category_auto(self):
        """מילוי אוטומטי של קטגורית הבד לפי התאמת וריאנט בקטלוג המוצרים."""
        try:
            product = (self.sup_product_var.get() or '').strip()
            if not product:
                self.sup_fabric_category_var.set(''); return
            size = (self.sup_size_var.get() or '').strip()
            ft = (self.sup_fabric_type_var.get() or '').strip()
            col = (self.sup_fabric_color_var.get() or '').strip()
            pn = (self.sup_print_name_var.get() or '').strip()
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
            best = None; best_score = -1
            for rec in catalog:
                if (rec.get('name') or '').strip() != product:
                    continue
                score = 0
                if size and (rec.get('size') or '').strip() == size: score += 1
                if ft and (rec.get('fabric_type') or '').strip() == ft: score += 1
                if col and (rec.get('fabric_color') or '').strip() == col: score += 1
                if pn and (rec.get('print_name') or '').strip() == pn: score += 1
                if score > best_score:
                    best_score = score; best = rec
            if best and best.get('fabric_category'):
                self.sup_fabric_category_var.set((best.get('fabric_category') or '').strip())
            else:
                self.sup_fabric_category_var.set('')
        except Exception:
            try: self.sup_fabric_category_var.set('')
            except Exception: pass

    def _add_supplier_package_line(self):
        pkg_type = (self.sup_pkg_type_var.get() or '').strip()
        qty_raw = (self.sup_pkg_qty_var.get() or '').strip()
        if not pkg_type or not qty_raw:
            messagebox.showerror("שגיאה", "חובה לבחור פריט הובלה ולהזין כמות")
            return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי")
            return
        driver = (getattr(self, 'sup_pkg_driver_var', tk.StringVar()).get() or '').strip()
        record = {'package_type': pkg_type, 'quantity': qty, 'driver': driver}
        self._supplier_packages.append(record)
        self.sup_packages_tree.insert('', 'end', values=(pkg_type, qty, driver))
        self.sup_pkg_qty_var.set('')

    def _delete_selected_supplier_package(self):
        sel = self.sup_packages_tree.selection()
        if not sel: return
        all_items = self.sup_packages_tree.get_children(); indices = [all_items.index(i) for i in sel]
        for item in sel: self.sup_packages_tree.delete(item)
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._supplier_packages): del self._supplier_packages[idx]

    def _clear_supplier_packages(self):
        self._supplier_packages = []
        for item in self.sup_packages_tree.get_children():
            self.sup_packages_tree.delete(item)

    def _refresh_supplier_intake_list(self):
        try:
            self.data_processor.refresh_supplier_receipts()
            receipts = [r for r in self.data_processor.supplier_receipts if r.get('receipt_kind') == 'supplier_intake']
        except Exception:
            receipts = []
        if hasattr(self, 'supplier_receipts_tree'):
            for iid in self.supplier_receipts_tree.get_children():
                self.supplier_receipts_tree.delete(iid)
            for rec in receipts:
                pkg_summary = ', '.join(f"{p.get('package_type')}:{p.get('quantity')}" for p in rec.get('packages', [])[:4])
                if len(rec.get('packages', [])) > 4:
                    pkg_summary += ' ...'
                # Add delete icon cell at the end
                self.supplier_receipts_tree.insert(
                    '', 'end',
                    values=(rec.get('id'), rec.get('date'), rec.get('supplier'), rec.get('total_quantity'), pkg_summary, '🗑')
                )

    def _open_supplier_receipt_details(self):
        """פתיחת חלון פרטים עבור תעודת קליטה נבחרת מתוך 'קליטות שמורות'."""
        try:
            sel = self.supplier_receipts_tree.selection()
            if not sel:
                return
            values = self.supplier_receipts_tree.item(sel[0], 'values') or []
            if not values:
                return
            rec_id = int(values[0])
        except Exception:
            return
        try:
            self.data_processor.refresh_supplier_receipts()
        except Exception:
            pass
        rec = None
        try:
            for r in getattr(self.data_processor, 'supplier_receipts', []) or []:
                if r.get('receipt_kind') == 'supplier_intake' and int(r.get('id', -1)) == rec_id:
                    rec = r; break
        except Exception:
            rec = None
        if not rec:
            try:
                messagebox.showwarning("שגיאה", "הרשומה לא נמצאה")
            except Exception:
                pass
            return
        win = tk.Toplevel(self.root)
        win.title(f"תעודת קליטה #{rec_id}")
        win.transient(self.root)
        try:
            win.grab_set()
        except Exception:
            pass
        header = tk.Frame(win, bg='#f7f9fa')
        header.pack(fill='x', padx=10, pady=8)
        tk.Label(header, text=f"ספק: {rec.get('supplier','')}", bg='#f7f9fa', font=('Arial',11,'bold')).pack(side='right', padx=8)
        tk.Label(header, text=f"תאריך: {rec.get('date','')}", bg='#f7f9fa').pack(side='right', padx=8)
        tk.Label(header, text=f"ID: {rec.get('id','')}", bg='#f7f9fa').pack(side='right', padx=8)
        tk.Label(header, text=f"סה\"כ כמות: {rec.get('total_quantity',0)}", bg='#f7f9fa').pack(side='right', padx=8)

        body = tk.Frame(win, bg='#f7f9fa')
        body.pack(fill='both', expand=True, padx=10, pady=(0,10))

        lines_frame = tk.LabelFrame(body, text='שורות תעודה', bg='#f7f9fa')
        lines_frame.pack(fill='both', expand=True, pady=6)
        cols = ('product','size','fabric_type','fabric_color','fabric_category','print_name','returned_from_drawing','drawing_id','quantity','note')
        tree = ttk.Treeview(lines_frame, columns=cols, show='headings', height=8)
        headers = {
            'product':'מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד',
            'fabric_category':'קטגורית בד','print_name':'שם פרינט',
            'returned_from_drawing':'חזר מציור','drawing_id':'"מס\' ציור"','quantity':'כמות','note':'הערה'
        }
        widths = {
            'product':180,'size':80,'fabric_type':100,'fabric_color':90,
            'fabric_category':120,'print_name':110,
            'returned_from_drawing':90,'drawing_id':80,'quantity':70,'note':220
        }
        for c in cols:
            tree.heading(c, text=headers[c])
            tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(lines_frame, orient='vertical', command=tree.yview)
        tree.configure(yscroll=vs.set)
        tree.pack(side='left', fill='both', expand=True, padx=(6,0), pady=6)
        vs.pack(side='left', fill='y', pady=6)
        for ln in rec.get('lines', []) or []:
            tree.insert('', 'end', values=(
                ln.get('product',''), ln.get('size',''), ln.get('fabric_type',''), ln.get('fabric_color',''), ln.get('fabric_category',''),
                ln.get('print_name',''), ln.get('returned_from_drawing','לא'), ln.get('drawing_id',''), ln.get('quantity',''), ln.get('note','')
            ))

        pk_frame = tk.LabelFrame(body, text='פריטי הובלה', bg='#f7f9fa')
        pk_frame.pack(fill='x', pady=6)
        pk_cols = ('type','quantity','driver')
        pk_tree = ttk.Treeview(pk_frame, columns=pk_cols, show='headings', height=4)
        pk_tree.heading('type', text='פריט הובלה')
        pk_tree.heading('quantity', text='כמות')
        pk_tree.heading('driver', text='שם המוביל')
        pk_tree.column('type', width=120, anchor='center')
        pk_tree.column('quantity', width=70, anchor='center')
        pk_tree.column('driver', width=110, anchor='center')
        pk_tree.pack(fill='x', padx=6, pady=6)
        for p in rec.get('packages', []) or []:
            pk_tree.insert('', 'end', values=(p.get('package_type',''), p.get('quantity',''), p.get('driver','')))

        btns = tk.Frame(win, bg='#f7f9fa')
        btns.pack(fill='x', padx=10, pady=(0,10))
        tk.Button(btns, text='סגור', command=win.destroy).pack(side='left')
