import tkinter as tk
from tkinter import ttk, messagebox
import os
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
            # Optional filtering by selected main category using product_model_names list
            selected_mc = ''
            try:
                if hasattr(self, 'dn_main_category_var'):
                    selected_mc = (self.dn_main_category_var.get() or '').strip()
            except Exception:
                selected_mc = ''
            if selected_mc:
                model_names = getattr(self.data_processor, 'product_model_names', []) or []
                allowed_models = { (m.get('name') or '').strip() for m in model_names if (m.get('main_category') or 'בגדים') == selected_mc and (m.get('name') or '').strip() }
                names = sorted({ (rec.get('name') or '').strip() for rec in catalog if (rec.get('name') or '').strip() in allowed_models })
            else:
                names = sorted({ (rec.get('name') or '').strip() for rec in catalog if rec.get('name') })
            self._delivery_products_allowed = names
            self._delivery_products_allowed_full = list(names)
            if hasattr(self, 'dn_product_combo'):
                try:
                    cur = self.dn_product_var.get()
                    self.dn_product_combo['values'] = self._delivery_products_allowed_full
                    if cur and cur not in self._delivery_products_allowed_full:
                        # Clear selection if it no longer valid under current main category
                        self.dn_product_var.set('')
                except Exception:
                    pass
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

    def _update_delivery_fabric_category_auto(self):
        """Auto-fill fabric category based on a matching product variant from products_catalog.
        Match keys: name, size, fabric_type, fabric_color, print_name (when available)."""
        try:
            product = (self.dn_product_var.get() or '').strip()
            if not product:
                self.dn_fabric_category_var.set(''); return
            size = (self.dn_size_var.get() or '').strip()
            ft = (self.dn_fabric_type_var.get() or '').strip()
            col = (self.dn_fabric_color_var.get() or '').strip()
            pn = (self.dn_print_name_var.get() or '').strip()
            catalog = getattr(self.data_processor, 'products_catalog', []) or []
            best = None
            best_score = -1
            for rec in catalog:
                if (rec.get('name') or '').strip() != product:
                    continue
                score = 0
                # optional exact matches add score
                if size and (rec.get('size') or '').strip() == size: score += 1
                if ft and (rec.get('fabric_type') or '').strip() == ft: score += 1
                if col and (rec.get('fabric_color') or '').strip() == col: score += 1
                if pn and (rec.get('print_name') or '').strip() == pn: score += 1
                if score > best_score:
                    best_score = score; best = rec
            if best and best.get('fabric_category'):
                self.dn_fabric_category_var.set((best.get('fabric_category') or '').strip())
            else:
                # fallback: blank
                self.dn_fabric_category_var.set('')

            # Also auto-fill print name from best-matching variant if user hasn't picked one yet
            try:
                if hasattr(self, 'dn_print_name_var'):
                    # Only if 'print_name' is enabled for the selected main category
                    mc = ''
                    try:
                        mc = (self.dn_main_category_var.get() or '').strip() if hasattr(self, 'dn_main_category_var') else ''
                    except Exception:
                        mc = ''
                    allow_print = False
                    try:
                        if mc:
                            mcs = getattr(self.data_processor, 'main_categories', []) or []
                            for c in mcs:
                                if (c.get('name') or '').strip() == mc:
                                    fields = (c.get('fields') or [])
                                    # Optional resolver API
                                    if (not fields) and hasattr(self.data_processor, 'get_main_category_fields'):
                                        try:
                                            fields = self.data_processor.get_main_category_fields(c.get('id')) or []
                                        except Exception:
                                            fields = []
                                    allow_print = 'print_name' in (fields or [])
                                    break
                    except Exception:
                        allow_print = False
                    if allow_print:
                        cur_pn = (self.dn_print_name_var.get() or '').strip()
                        cand_pn = ((best or {}).get('print_name') or '').strip()
                        if cand_pn and not cur_pn:
                            # Ensure combobox (if exists) contains the candidate
                            if hasattr(self, 'dn_print_name_combo'):
                                try:
                                    vals = list(self.dn_print_name_combo['values'] or [])
                                    if cand_pn not in vals:
                                        vals.append(cand_pn)
                                        seen = set(); vals = [x for x in vals if not (x in seen or seen.add(x))]
                                        self.dn_print_name_combo['values'] = vals
                                except Exception:
                                    pass
                            self.dn_print_name_var.set(cand_pn)
                    else:
                        # Ensure it stays empty when print_name not relevant
                        try:
                            self.dn_print_name_var.set('')
                        except Exception:
                            pass
            except Exception:
                pass
        except Exception:
            try:
                self.dn_fabric_category_var.set('')
            except Exception:
                pass

    def _refresh_delivery_print_name_options(self):
        """Refresh available print names optionally filtered by selected main category when product doesn’t constrain them."""
        try:
            product = (self.dn_product_var.get() or '').strip()
            names = []
            if product:
                catalog = getattr(self.data_processor, 'products_catalog', []) or []
                for rec in catalog:
                    if (rec.get('name') or '').strip() == product:
                        pn = (rec.get('print_name') or '').strip()
                        if pn:
                            names.append(pn)
                names = sorted({n for n in names})
            if not names:
                selected_mc = ''
                try:
                    if hasattr(self, 'dn_main_category_var'):
                        selected_mc = (self.dn_main_category_var.get() or '').strip()
                except Exception:
                    selected_mc = ''
                plist = getattr(self.data_processor, 'product_print_names', []) or []
                if selected_mc:
                    names = [ (r.get('name') or '') for r in plist if r.get('name') and (r.get('main_category') or 'בגדים') == selected_mc ]
                else:
                    names = [ (r.get('name') or '') for r in plist if r.get('name') ]
        except Exception:
            names = []
        # If a print-name combobox exists (future-proof), set it
        try:
            if hasattr(self, 'dn_print_name_combo'):
                self.dn_print_name_combo['values'] = names
        except Exception:
            pass

    # ---- Lines ops ----
    def _add_delivery_line(self):
        product = self.dn_product_var.get().strip(); size = self.dn_size_var.get().strip(); qty_raw = self.dn_qty_var.get().strip(); note = self.dn_note_var.get().strip()
        fabric_type = self.dn_fabric_type_var.get().strip(); fabric_color = self.dn_fabric_color_var.get().strip()
        # Respect main category: only take print_name if the field exists for this category
        try:
            mc = (self.dn_main_category_var.get() or '').strip() if hasattr(self, 'dn_main_category_var') else ''
        except Exception:
            mc = ''
        allow_print = False
        allow_barcode = False
        try:
            if mc:
                mcs = getattr(self.data_processor, 'main_categories', []) or []
                for c in mcs:
                    if (c.get('name') or '').strip() == mc:
                        fields = (c.get('fields') or [])
                        if (not fields) and hasattr(self.data_processor, 'get_main_category_fields'):
                            try:
                                fields = self.data_processor.get_main_category_fields(c.get('id')) or []
                            except Exception:
                                fields = []
                        allow_print = 'print_name' in (fields or [])
                        allow_barcode = 'barcode' in (fields or [])
                        break
        except Exception:
            allow_print = False
            allow_barcode = False
        raw_print = (self.dn_print_name_var.get() or '').strip()
        print_name = raw_print if allow_print else ''
        raw_barcode = ''
        try:
            raw_barcode = (self.dn_barcode_var.get() or '').strip()
        except Exception:
            raw_barcode = ''
        barcode = raw_barcode if allow_barcode else ''
        fabric_category = getattr(self, 'dn_fabric_category_var', None)
        fabric_category = fabric_category.get().strip() if fabric_category else ''
        if not product or not qty_raw:
            messagebox.showerror("שגיאה", "חובה לבחור מוצר ולהזין כמות"); return
        if self._delivery_products_allowed and product not in self._delivery_products_allowed:
            messagebox.showerror("שגיאה", "יש לבחור מוצר מהרשימה בלבד"); return
        try:
            qty = int(qty_raw); assert qty > 0
        except Exception:
            messagebox.showerror("שגיאה", "כמות חייבת להיות מספר חיובי"); return
        # Respect main category for sub_category as well
        category = getattr(self, 'dn_category_var', None)
        category = category.get().strip() if category else ''
        try:
            mc_fields = []
            mc = (self.dn_main_category_var.get() or '').strip() if hasattr(self, 'dn_main_category_var') else ''
            if mc:
                for c in getattr(self.data_processor, 'main_categories', []) or []:
                    if (c.get('name') or '').strip() == mc:
                        mc_fields = (c.get('fields') or [])
                        if (not mc_fields) and hasattr(self.data_processor, 'get_main_category_fields'):
                            try:
                                mc_fields = self.data_processor.get_main_category_fields(c.get('id')) or []
                            except Exception:
                                mc_fields = []
                        break
            if 'sub_category' not in (mc_fields or []):
                category = ''
        except Exception:
            pass
        line = {'product': product, 'size': size, 'fabric_type': fabric_type, 'fabric_color': fabric_color, 'fabric_category': fabric_category, 'print_name': print_name, 'barcode': barcode, 'category': category, 'quantity': qty, 'note': note}
        self._delivery_lines.append(line)
        # columns: product,size,fabric_type,fabric_color,fabric_category,print_name,barcode,category,quantity,note
        self.delivery_tree.insert('', 'end', values=(product,size,fabric_type,fabric_color,fabric_category,print_name,barcode,category,qty,note))
        self.dn_size_var.set(''); self.dn_qty_var.set(''); self.dn_note_var.set('')
        try:
            self.dn_product_combo['values'] = self._delivery_products_allowed_full
        except Exception:
            pass
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
        # אם רשימת ה-packages בזיכרון ריקה אך יש שורות בטבלה, נשחזר מהרשימה המוצגת
        try:
            if (not self._packages) and hasattr(self, 'packages_tree'):
                items = self.packages_tree.get_children()
                rebuilt = []
                for it in items:
                    vals = self.packages_tree.item(it, 'values') or ()
                    if vals:
                        try:
                            rebuilt.append({'package_type': vals[0], 'quantity': int(vals[1]), 'driver': vals[2] if len(vals) > 2 else ''})
                        except Exception:
                            # אם הכמות אינה מספר – נשמור כמחרוזת כדי לא לאבד נתון
                            rebuilt.append({'package_type': vals[0], 'quantity': vals[1], 'driver': vals[2] if len(vals) > 2 else ''})
                if rebuilt:
                    self._packages = rebuilt
        except Exception:
            pass
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
                # add delete icon in the last column to match ('id','date','supplier','total','packages','delete')
                self.delivery_notes_tree.insert('', 'end', values=(rec.get('id'), rec.get('date'), rec.get('supplier'), rec.get('total_quantity'), pkg_summary, '🗑'))

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
            # Include barcode column for categories like Fabrics; remains empty otherwise
            lines_cols = ('product','size','fabric_type','fabric_color','fabric_category','print_name','barcode','quantity','note')
            lines_tree = ttk.Treeview(lines_frame, columns=lines_cols, show='headings', height=8)
            headers_map = {'product':'מוצר','size':'מידה','fabric_type':'סוג בד','fabric_color':'צבע בד','fabric_category':'קטגורית בד','print_name':'פרינט','barcode':'בר קוד','quantity':'כמות','note':'הערה'}
            widths_map = {'product':140,'size':70,'fabric_type':110,'fabric_color':110,'fabric_category':120,'print_name':110,'barcode':110,'quantity':60,'note':160}
            for c in lines_cols:
                lines_tree.heading(c, text=headers_map[c])
                lines_tree.column(c, width=widths_map[c], anchor='center')
            for line in rec.get('lines', []) or []:
                lines_tree.insert('', 'end', values=(
                    line.get('product'), line.get('size'), line.get('fabric_type'), line.get('fabric_color'),
                    line.get('fabric_category',''), line.get('print_name'), line.get('barcode',''), line.get('quantity'), line.get('note')
                ))
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
            # Actions: Open in Excel + Close
            btns = tk.Frame(win)
            btns.pack(fill='x', pady=6, padx=4)

            def _export_dn_to_excel_and_open():
                try:
                    # Lazy import to avoid hard dependency during normal UI flow
                    from openpyxl import Workbook
                    from openpyxl.styles import Border, Side, Alignment, Font
                    try:
                        from openpyxl.worksheet.page import PageMargins  # type: ignore
                    except Exception:
                        PageMargins = None  # type: ignore
                    try:
                        from openpyxl.drawing.image import Image as XLImage  # type: ignore
                    except Exception:
                        XLImage = None  # type: ignore
                except Exception as e:
                    try:
                        messagebox.showerror("שגיאה", f"נדרש openpyxl לצורך יצוא לאקסל:\n{e}")
                    except Exception:
                        pass
                    return
                try:
                    # Prepare workbook
                    wb = Workbook(); ws = wb.active; ws.title = "תעודת משלוח"
                    # Right-to-Left sheet for Hebrew
                    try:
                        ws.sheet_view.rightToLeft = True
                    except Exception:
                        pass
                    # Page setup: A4 portrait, fit to width
                    try:
                        ws.page_setup.orientation = 'portrait'
                    except Exception:
                        pass
                    try:
                        # A4 paper size code is 9 in Excel
                        ws.page_setup.paperSize = 9
                    except Exception:
                        pass
                    try:
                        ws.page_setup.fitToWidth = 1
                        ws.page_setup.fitToHeight = 0
                    except Exception:
                        pass
                    try:
                        if PageMargins:
                            # Excel Narrow preset: Top/Bottom 1.91 cm (~0.75"), Left/Right 0.64 cm (~0.25"), Header/Footer 0.76 cm (~0.3")
                            ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.75, bottom=0.75, header=0.3, footer=0.3)
                    except Exception:
                        pass
                    # Business header (from settings): Business Name + VAT/Type line + optional logo
                    try:
                        s = getattr(self, 'settings', None)
                        def _sget(k, default=""):
                            try:
                                return (s.get(k, default) if s else default) or ""
                            except Exception:
                                return ""
                        biz_name = _sget('business.name')
                        biz_type = _sget('business.type')
                        biz_vat  = _sget('business.vat_id')
                        biz_logo = _sget('business.logo_path')
                    except Exception:
                        biz_name = biz_type = biz_vat = biz_logo = ""
                    # Defer logo insertion until after columns are autosized; remember path and desired height
                    pending_logo_path = biz_logo if (biz_logo and os.path.exists(biz_logo)) else ""
                    desired_header_height_in = 48.75 / 25.4  # ≈1.92 inches

                    # Document info header (keys and values)
                    doc_info_start = ws.max_row + 1
                    ws.append(["מסמך", f"תעודת משלוח #{rec.get('id')}"])
                    # תאריך בפורמט D/M/Y
                    try:
                        raw_date = (rec.get('date') or '').strip()
                        def _parse_dt(s: str):
                            for fmt in ("%Y-%m-%d","%d/%m/%Y","%d-%m-%Y","%Y/%m/%d","%m/%d/%Y","%d.%m.%Y"):
                                try:
                                    # Return date object to avoid time component in Excel
                                    return datetime.strptime(s, fmt).date()
                                except Exception:
                                    pass
                            return None
                        d_obj = _parse_dt(raw_date)
                        if d_obj is not None:
                            ws.append(["תאריך", d_obj])
                            try:
                                # Apply Excel date number format on the just-added cell
                                r = ws.max_row
                                dc = ws.cell(row=r, column=2)
                                dc.number_format = 'd/m/yyyy'
                                dc.alignment = Alignment(horizontal='right')
                            except Exception:
                                pass
                        else:
                            # fallback quick reformat (YYYY-MM-DD -> D/M/Y)
                            parts = raw_date.replace('.', '-').replace('/', '-').split('-')
                            if len(parts) == 3 and len(parts[0]) == 4:
                                try:
                                    y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                                    ws.append(["תאריך", f"{d}/{m}/{y}"])
                                except Exception:
                                    ws.append(["תאריך", raw_date])
                            else:
                                ws.append(["תאריך", raw_date])
                    except Exception:
                        ws.append(["תאריך", rec.get('date')])
                    ws.append(["ספק", rec.get('supplier')])
                    doc_info_end = ws.max_row
                    try:
                        # Style: right-align, bold keys, slightly larger font
                        for row in ws.iter_rows(min_row=doc_info_start, max_row=doc_info_end, min_col=1, max_col=2):
                            try:
                                row[0].font = Font(bold=True, size=12)
                                row[0].alignment = Alignment(horizontal='right')
                            except Exception:
                                pass
                            try:
                                row[1].font = Font(size=12)
                                row[1].alignment = Alignment(horizontal='right')
                            except Exception:
                                pass
                        # אם הוזנה ישירות כתאריך - מספר פורמט ד/מ/שנה (אימות נוסף)
                        try:
                            dcell = ws.cell(row=doc_info_start+1, column=2)
                            dcell.number_format = 'd/m/yyyy'
                        except Exception:
                            pass
                    except Exception:
                        pass
                    ws.append([])
                    # Lines table (A4-ready): Model | Size | Description (fabric type + print) | Quantity
                    ws.append(["שורות מוצרים"])
                    try:
                        # Merge and style the section title
                        r = ws.max_row
                        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
                        tcell = ws.cell(row=r, column=1)
                        tcell.font = Font(bold=True, size=12)
                        tcell.alignment = Alignment(horizontal='right')
                    except Exception:
                        pass
                    header_row = ["שם הדגם", "מידה", "תיאור", "כמות"]
                    ws.append(header_row)
                    start_data_row = ws.max_row + 1
                    for line in (rec.get('lines', []) or []):
                        model = line.get('product')
                        size = (line.get('size') or '')
                        fabric_type = (line.get('fabric_type') or '')
                        color = (line.get('fabric_color') or '')
                        print_name = (line.get('print_name') or '')
                        # Compose description: fabric_type + color + print_name (skip empties)
                        parts = [p for p in [fabric_type, color, print_name] if p]
                        desc = " | ".join(parts)
                        qty = line.get('quantity')
                        ws.append([model, size, desc, qty])
                    # שורת סה"כ כמות מתחת לטבלה (תגית בעמודה C והמספר בעמודה D)
                    try:
                        try:
                            total_qty = sum(int((l or {}).get('quantity', 0) or 0) for l in (rec.get('lines', []) or []))
                        except Exception:
                            total_qty = (rec.get('total_quantity') or 0)
                        # A, B empty; C label; D numeric total
                        ws.append([None, None, "סה\"כ כמות", total_qty])
                        r_tot = ws.max_row
                        try:
                            ws.cell(row=r_tot, column=3).font = Font(bold=True, size=12)
                            ws.cell(row=r_tot, column=3).alignment = Alignment(horizontal='right')
                            ws.cell(row=r_tot, column=4).alignment = Alignment(horizontal='center')
                        except Exception:
                            pass
                    except Exception:
                        pass
                    end_data_row = ws.max_row
                    # Style: black grid on lines area + fonts and alignments
                    try:
                        thin = Side(style='thin', color='000000')
                        border = Border(left=thin, right=thin, top=thin, bottom=thin)
                        # Bold header row
                        for cell in ws[ws.cell(row=start_data_row-1, column=1).row]:
                            if cell.row == start_data_row-1:
                                try:
                                    cell.font = Font(bold=True, size=12)
                                except Exception:
                                    pass
                        for row in ws.iter_rows(min_row=start_data_row-1, max_row=end_data_row, min_col=1, max_col=4):
                            for cell in row:
                                try:
                                    cell.border = border
                                except Exception:
                                    pass
                                try:
                                    # Data font size a bit larger for readability
                                    if cell.row >= start_data_row:
                                        cell.font = Font(size=11)
                                except Exception:
                                    pass
                            # Align quantity center/right
                            try:
                                row[3].alignment = Alignment(horizontal='center')
                            except Exception:
                                pass
                            # Align Hebrew text to the right for Model/Size/Description
                            try:
                                row[0].alignment = Alignment(horizontal='right')
                                row[1].alignment = Alignment(horizontal='right')
                                row[2].alignment = Alignment(horizontal='right')
                            except Exception:
                                pass
                    except Exception:
                        pass
                    # Autosize simple columns
                    try:
                        for col in ws.columns:
                            max_len = 0; col_letter = col[0].column_letter
                            for cell in col:
                                try:
                                    val = str(cell.value) if cell.value is not None else ""
                                    if len(val) > max_len: max_len = len(val)
                                except Exception:
                                    pass
                            ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)
                    except Exception:
                        pass
                    # Insert the logo banner at the very top, sized to the full printable A4 width (independent of column widths)
                    try:
                        if XLImage and pending_logo_path:
                            # Compute physical printable width of A4 and use it directly
                            try:
                                A4_WIDTH_INCH = 8.27
                                lm = float(getattr(ws.page_margins, 'left', 0.25) or 0.25)
                                rm = float(getattr(ws.page_margins, 'right', 0.25) or 0.25)
                                printable_inch = max(4.0, A4_WIDTH_INCH - lm - rm)
                                # Safety pad so the banner never touches printable edge and won't be clipped
                                safety_in = 0.35
                                if printable_inch > safety_in:
                                    printable_inch -= safety_in
                                total_px = int(printable_inch * 96.0)
                            except Exception:
                                total_px = int(8.0 * 96.0)
                            # Prepare image
                            img_obj = XLImage(pending_logo_path)
                            img_obj.width = max(200, int(total_px))
                            img_obj.height = int(desired_header_height_in * 96.0)
                            # Insert two rows at top so banner sits above all content
                            try:
                                ws.insert_rows(1, amount=2)
                            except Exception:
                                pass
                            try:
                                img_obj.anchor = 'A1'
                            except Exception:
                                pass
                            ws.add_image(img_obj)
                            # Set row 1 height to banner height in points (1in = 72pt)
                            try:
                                ws.row_dimensions[1].height = desired_header_height_in * 72.0
                            except Exception:
                                pass
                            # Keep a small spacer row
                            try:
                                ws.row_dimensions[2].height = 6
                            except Exception:
                                pass
                    except Exception:
                        pass
                    # Save to exports folder
                    base_dir = os.path.join(os.getcwd(), 'exports', 'delivery_notes')
                    try:
                        os.makedirs(base_dir, exist_ok=True)
                    except Exception:
                        pass
                    safe_id = rec.get('id')
                    safe_date = (rec.get('date') or '').replace('/', '-').replace(':', '-')
                    fname = f"delivery_note_{safe_id}_{safe_date}.xlsx" if safe_id is not None else f"delivery_note_{safe_date}.xlsx"
                    out_path = os.path.join(base_dir, fname)
                    try:
                        wb.save(out_path)
                    except PermissionError as e:
                        # Fallback: add timestamp suffix if file is locked/open
                        try:
                            ts = datetime.now().strftime('%H%M%S')
                            alt_path = out_path.replace('.xlsx', f'_{ts}.xlsx')
                            wb.save(alt_path)
                            out_path = alt_path
                        except Exception:
                            try:
                                messagebox.showerror("שגיאה", f"שמירת קובץ נכשלה (קובץ פתוח?):\n{e}")
                            except Exception:
                                pass
                            return
                    except Exception as e:
                        try:
                            messagebox.showerror("שגיאה", f"שמירת קובץ נכשלה:\n{e}")
                        except Exception:
                            pass
                        return
                    # Open in Excel (Windows association)
                    try:
                        os.startfile(out_path)  # type: ignore[attr-defined]
                    except Exception as e:
                        try:
                            messagebox.showinfo("נשמר", f"הקובץ נשמר ב:\n{out_path}\n(לא הצלחתי לפתוח אוטומטית)")
                        except Exception:
                            pass

                except Exception as e:
                    try:
                        messagebox.showerror("שגיאה", str(e))
                    except Exception:
                        pass

            tk.Button(btns, text='🖨 פתח באקסל', command=_export_dn_to_excel_and_open, bg='#27ae60', fg='white').pack(side='right', padx=4)
            tk.Button(btns, text='סגור', command=win.destroy).pack(side='right')
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", str(e))
            except Exception:
                pass
