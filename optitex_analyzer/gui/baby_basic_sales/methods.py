import os
from datetime import datetime
from tkinter import messagebox, filedialog

from optitex_analyzer.core.data_processor import parse_numeric_value
from .. import theme


class BabyBasicSalesMethodsMixin:
    """Logic for Baby Basic wholesale notes, price list and account."""

    def _bb_money(self, value) -> str:
        try:
            return f"{parse_numeric_value(value):,.2f}"
        except Exception:
            return "0.00"

    def _bb_refresh_customer_combos(self):
        partner = self._bb_partner_name()
        names = []
        try:
            names = self.data_processor.get_baby_basic_customers()
        except Exception:
            names = []
        if partner not in names:
            names = [partner] + [n for n in names if n != partner]
        if getattr(self, 'bb_customer_combo', None) is not None:
            try:
                self.bb_customer_combo['values'] = names
                current = (self.bb_customer_var.get() if hasattr(self, 'bb_customer_var') else '').strip()
                if not current:
                    self.bb_customer_var.set(partner)
            except Exception:
                pass

    def _bb_products(self):
        try:
            return self.data_processor.get_baby_basic_products()
        except Exception:
            return []

    def _bb_product_by_barcode(self, barcode: str):
        barcode = str(barcode or '').strip()
        for p in self._bb_products():
            if p.get('barcode') == barcode:
                return p
        return None

    def _bb_filter_products(self, query: str = ''):
        q = (query or '').strip().lower()
        terms = q.split()
        result = []
        for p in self._bb_products():
            blob = ' '.join([
                str(p.get('print_name') or ''),
                str(p.get('item_name') or ''),
                str(p.get('size') or ''),
                str(p.get('fabric') or ''),
                str(p.get('barcode') or ''),
                str(p.get('color') or ''),
            ]).lower()
            if terms and not all(t in blob for t in terms):
                continue
            result.append(p)
        return result

    def _bb_clear_tree(self, tree):
        if tree is None:
            return
        for iid in tree.get_children():
            tree.delete(iid)

    def _bb_inline_edit(self, tree, event, column_id, on_commit):
        col = tree.identify_column(event.x)
        row = tree.identify_row(event.y)
        if not row or col != column_id:
            return
        bbox = tree.bbox(row, col)
        if not bbox:
            return
        x, y, w, h = bbox
        vals = list(tree.item(row, 'values'))
        col_index = int(column_id.replace('#', '')) - 1
        current = vals[col_index] if 0 <= col_index < len(vals) else ''
        import tkinter as tk
        entry = tk.Entry(tree, justify='center')
        entry.insert(0, '' if str(current) in ('0', '0.00') else str(current))
        entry.select_range(0, 'end')
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()

        def commit(_e=None):
            raw = entry.get().strip()
            try:
                entry.destroy()
            except Exception:
                pass
            on_commit(row, raw)

        entry.bind('<Return>', commit)
        entry.bind('<KP_Enter>', commit)
        entry.bind('<FocusOut>', commit)
        entry.bind('<Escape>', lambda _e: entry.destroy())

    # ---------- מחירון ----------
    def _bb_refresh_price_table(self):
        tree = getattr(self, 'bb_price_tree', None)
        if tree is None:
            return
        self._bb_clear_tree(tree)
        query = self.bb_price_search_var.get() if hasattr(self, 'bb_price_search_var') else ''
        for p in self._bb_filter_products(query):
            tree.insert('', 'end', iid=p['barcode'], values=(
                p.get('print_name') or p.get('item_name') or '',
                p.get('size', ''),
                p.get('fabric', ''),
                p.get('color', ''),
                p.get('pack_qty', 5),
                p.get('barcode', ''),
                self._bb_money(p.get('price')),
            ))
        theme.stripe_tree(tree)

    def _bb_edit_price_cell(self, event):
        def on_commit(barcode, raw):
            try:
                price = parse_numeric_value(raw)
            except Exception:
                price = 0
            self.data_processor.set_baby_basic_price(barcode, price)
            self._bb_refresh_price_table()
            self._bb_refresh_note_products()
        self._bb_inline_edit(self.bb_price_tree, event, '#7', on_commit)

    def _bb_apply_price_to_family(self):
        tree = getattr(self, 'bb_price_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("מחירון", "בחר מוצר בטבלה ואז החל את המחיר לכל המידות של אותו דגם.")
            return
        src = self._bb_product_by_barcode(sel[0])
        if not src:
            return
        try:
            price = parse_numeric_value(tree.item(sel[0], 'values')[-1])
        except Exception:
            price = src.get('price') or 0
        if price <= 0:
            messagebox.showwarning("מחירון", "קודם הזן מחיר לשורה שנבחרה (דאבל-קליק על עמודת המחיר).")
            return
        family_key = (src.get('print_name') or src.get('item_name') or '').strip()
        fabric = (src.get('fabric') or '').strip()
        updates = {}
        for p in self._bb_products():
            same_name = (p.get('print_name') or p.get('item_name') or '').strip() == family_key
            same_fabric = (p.get('fabric') or '').strip() == fabric
            if same_name and same_fabric:
                updates[p['barcode']] = price
        if not updates:
            return
        self.data_processor.set_baby_basic_prices_bulk(updates)
        self._bb_refresh_price_table()
        self._bb_refresh_note_products()
        messagebox.showinfo("מחירון", f"עודכן מחיר {self._bb_money(price)} ₪ ל-{len(updates)} פריטים של «{family_key}».")

    # ---------- תעודת סחורה ----------
    def _bb_refresh_note_products(self):
        tree = getattr(self, 'bb_note_products_tree', None)
        if tree is None:
            return
        self._bb_clear_tree(tree)
        query = self.bb_note_search_var.get() if hasattr(self, 'bb_note_search_var') else ''
        qty_map = getattr(self, '_bb_note_qty_by_barcode', {})
        for p in self._bb_filter_products(query):
            bc = p['barcode']
            tree.insert('', 'end', iid=bc, values=(
                p.get('print_name') or p.get('item_name') or '',
                p.get('size', ''),
                p.get('fabric', ''),
                bc,
                self._bb_money(p.get('price')),
                int(qty_map.get(bc, 0) or 0),
            ))
        theme.stripe_tree(tree)

    def _bb_edit_note_qty_cell(self, event):
        def on_commit(barcode, raw):
            try:
                qty = int(float(raw)) if raw else 0
            except Exception:
                qty = 0
            if qty < 0:
                qty = 0
            if not hasattr(self, '_bb_note_qty_by_barcode'):
                self._bb_note_qty_by_barcode = {}
            if qty > 0:
                self._bb_note_qty_by_barcode[barcode] = qty
            else:
                self._bb_note_qty_by_barcode.pop(barcode, None)
            if self.bb_note_products_tree.exists(barcode):
                vals = list(self.bb_note_products_tree.item(barcode, 'values'))
                vals[-1] = qty
                self.bb_note_products_tree.item(barcode, values=vals)

        self._bb_inline_edit(self.bb_note_products_tree, event, '#6', on_commit)

    def _bb_refresh_note_lines(self):
        tree = getattr(self, 'bb_note_lines_tree', None)
        if tree is None:
            return
        self._bb_clear_tree(tree)
        lines = getattr(self, '_bb_note_lines', [])
        total = 0.0
        qty_total = 0
        for i, line in enumerate(lines):
            total += parse_numeric_value(line.get('line_total'))
            qty_total += int(line.get('quantity') or 0)
            tree.insert('', 'end', iid=str(i), values=(
                line.get('print_name') or line.get('item_name') or '',
                line.get('size', ''),
                line.get('fabric', ''),
                line.get('barcode', ''),
                int(line.get('quantity') or 0),
                self._bb_money(line.get('unit_price')),
                self._bb_money(line.get('line_total')),
            ))
        theme.stripe_tree(tree)
        if hasattr(self, 'bb_note_total_var'):
            self.bb_note_total_var.set(f"סה״כ יחידות: {qty_total}    סה״כ לתשלום: {self._bb_money(total)} ₪")

    def _bb_add_selected_to_note(self):
        qty_map = {bc: q for bc, q in getattr(self, '_bb_note_qty_by_barcode', {}).items() if int(q or 0) > 0}
        if not qty_map:
            messagebox.showwarning("תעודה", "לא הוזנה כמות. לחץ פעמיים על עמודת «כמות» ואז הוסף לתעודה.")
            return
        if not hasattr(self, '_bb_note_lines'):
            self._bb_note_lines = []
        existing = {str(l.get('barcode')): l for l in self._bb_note_lines}
        missing_price = []
        for bc, qty in qty_map.items():
            prod = self._bb_product_by_barcode(bc)
            if not prod:
                continue
            price = parse_numeric_value(prod.get('price'))
            if price <= 0:
                missing_price.append(prod.get('print_name') or prod.get('item_name') or bc)
            qty = int(qty)
            if bc in existing:
                existing[bc]['quantity'] = int(existing[bc].get('quantity') or 0) + qty
                existing[bc]['unit_price'] = price
                existing[bc]['line_total'] = round(existing[bc]['quantity'] * price, 2)
            else:
                self._bb_note_lines.append({
                    'barcode': bc,
                    'item_name': prod.get('item_name', ''),
                    'print_name': prod.get('print_name') or prod.get('item_name') or '',
                    'size': prod.get('size', ''),
                    'fabric': prod.get('fabric', ''),
                    'pack_qty': prod.get('pack_qty', 5),
                    'quantity': qty,
                    'unit_price': price,
                    'line_total': round(qty * price, 2),
                })
        self._bb_note_qty_by_barcode = {}
        self._bb_refresh_note_products()
        self._bb_refresh_note_lines()
        if missing_price:
            messagebox.showwarning(
                "מחירון חסר",
                "הפריטים הבאים נוספו בלי מחיר. עדכן במחירון או דאבל-קליק על מחיר בשורת התעודה:\n• "
                + "\n• ".join(missing_price[:12])
            )

    def _bb_edit_line_price_cell(self, event):
        def on_commit(iid, raw):
            try:
                idx = int(iid)
            except Exception:
                return
            lines = getattr(self, '_bb_note_lines', [])
            if idx < 0 or idx >= len(lines):
                return
            price = parse_numeric_value(raw)
            qty = int(lines[idx].get('quantity') or 0)
            lines[idx]['unit_price'] = round(price, 2)
            lines[idx]['line_total'] = round(qty * price, 2)
            self._bb_refresh_note_lines()
        self._bb_inline_edit(self.bb_note_lines_tree, event, '#6', on_commit)

    def _bb_remove_selected_note_line(self):
        tree = getattr(self, 'bb_note_lines_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        try:
            idx = int(sel[0])
            if 0 <= idx < len(self._bb_note_lines):
                self._bb_note_lines.pop(idx)
        except Exception:
            return
        self._bb_refresh_note_lines()

    def _bb_clear_current_note(self):
        self._bb_note_lines = []
        self._bb_note_qty_by_barcode = {}
        if hasattr(self, 'bb_note_comment_var'):
            self.bb_note_comment_var.set('')
        if hasattr(self, 'bb_note_date_var'):
            self.bb_note_date_var.set(datetime.now().strftime('%Y-%m-%d'))
        self._bb_refresh_note_products()
        self._bb_refresh_note_lines()

    def _bb_save_note(self):
        customer = (self.bb_customer_var.get() if hasattr(self, 'bb_customer_var') else '').strip()
        if not customer:
            customer = self._bb_partner_name()
            if hasattr(self, 'bb_customer_var'):
                self.bb_customer_var.set(customer)
        date_str = (self.bb_note_date_var.get() if hasattr(self, 'bb_note_date_var') else '').strip()
        comment = (self.bb_note_comment_var.get() if hasattr(self, 'bb_note_comment_var') else '').strip()
        lines = list(getattr(self, '_bb_note_lines', []) or [])
        try:
            new_id = self.data_processor.add_baby_basic_note(customer, date_str, lines, note=comment)
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
            return
        self._bb_refresh_customer_combos()
        self._bb_refresh_saved_notes()
        self._bb_refresh_account()
        self._bb_clear_current_note()
        messagebox.showinfo("נשמר", f"תעודת סחורה #{new_id} נשמרה.")
        rec = self.data_processor.get_baby_basic_note(new_id)
        if rec:
            self._bb_open_note_view(rec)

    def _bb_refresh_saved_notes(self):
        tree = getattr(self, 'bb_saved_notes_tree', None)
        if tree is None:
            return
        self._bb_clear_tree(tree)
        notes = sorted(self.data_processor.get_baby_basic_notes(), key=lambda n: int(n.get('id') or 0), reverse=True)
        for rec in notes:
            tree.insert('', 'end', iid=str(rec.get('id')), values=(
                rec.get('id', ''),
                rec.get('date', ''),
                rec.get('customer', ''),
                rec.get('total_quantity', 0),
                self._bb_money(rec.get('total_amount')),
                rec.get('note', ''),
            ))
        theme.stripe_tree(tree)

    def _bb_open_selected_note(self, event=None):
        tree = getattr(self, 'bb_saved_notes_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        rec = self.data_processor.get_baby_basic_note(sel[0])
        if rec:
            self._bb_open_note_view(rec)

    def _bb_delete_selected_note(self):
        tree = getattr(self, 'bb_saved_notes_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        rec = self.data_processor.get_baby_basic_note(sel[0])
        if not rec:
            return
        if not messagebox.askyesno("מחיקה", f"למחוק תעודה #{rec.get('id')} לשותף {rec.get('customer')}?"):
            return
        self.data_processor.delete_baby_basic_note(rec.get('id'))
        self._bb_refresh_saved_notes()
        self._bb_refresh_account()
        self._bb_refresh_customer_combos()

    def _bb_open_note_view(self, rec: dict):
        import tkinter as tk
        from tkinter import ttk
        win = tk.Toplevel(self.root)
        win.title(f"תעודת סחורה בייבי בייסיק #{rec.get('id')}")
        win.configure(bg=theme.PAGE_BG)
        win.geometry("920x560")
        tk.Label(
            win,
            text=f"תעודת סחורה #{rec.get('id')}  |  {rec.get('customer')}  |  {rec.get('date')}",
            font=(theme.FONT_FAMILY, 13, 'bold'),
            bg=theme.PAGE_BG,
            fg=theme.DARK,
        ).pack(pady=(10, 4))
        if rec.get('note'):
            tk.Label(win, text=rec.get('note'), bg=theme.PAGE_BG, fg=theme.SUBTEXT).pack()

        cols = ('print_name', 'size', 'fabric', 'color', 'barcode', 'qty', 'price', 'total')
        headers = {
            'print_name': 'מוצר', 'size': 'מידה', 'fabric': 'בד', 'color': 'צבע',
            'barcode': 'ברקוד', 'qty': 'יחידות', 'price': 'מחיר ליחידה', 'total': 'סה״כ',
        }
        tree = theme.make_treeview(win, columns=cols, show='headings', height=14)
        for c in cols:
            tree.heading(c, text=headers[c])
            tree.column(c, width=90 if c in ('size', 'qty', 'price', 'total', 'color') else 160, anchor='center')
        vs = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.pack(side='left', fill='both', expand=True, padx=(12, 0), pady=8)
        vs.pack(side='left', fill='y', pady=8)
        for line in rec.get('lines') or []:
            tree.insert('', 'end', values=(
                line.get('print_name') or line.get('item_name') or '',
                line.get('size', ''),
                line.get('fabric', ''),
                line.get('color', '') or '',
                line.get('barcode', ''),
                line.get('quantity', 0),
                self._bb_money(line.get('unit_price')),
                self._bb_money(line.get('line_total')),
            ))
        theme.stripe_tree(tree)

        footer = tk.Frame(win, bg=theme.PAGE_BG)
        footer.pack(fill='x', padx=12, pady=(0, 10))
        tk.Label(
            footer,
            text=f"סה״כ יחידות: {rec.get('total_quantity', 0)}    סה״כ: {self._bb_money(rec.get('total_amount'))} ₪",
            font=(theme.FONT_FAMILY, 11, 'bold'),
            bg=theme.PAGE_BG,
        ).pack(side='right')
        theme.make_button(footer, "ייצוא לאקסל", kind="success", command=lambda: self._bb_export_note_excel(rec)).pack(side='left')
        theme.make_button(footer, "PDF להדפסה (A4)", kind="primary", command=lambda: self._bb_export_note_pdf(rec)).pack(side='left', padx=6)

    def _bb_export_note_excel(self, rec: dict, path: str | None = None):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
        except Exception as e:
            messagebox.showerror("שגיאה", f"נדרש openpyxl לייצוא לאקסל:\n{e}")
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "תעודת סחורה"
        try:
            ws.sheet_view.rightToLeft = True
        except Exception:
            pass
        ws.page_setup.orientation = 'portrait'
        ws.page_setup.paperSize = 9
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        thin = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1'),
        )
        header_fill = PatternFill('solid', fgColor='1E293B')
        header_font = Font(name='Arial', bold=True, color='FFFFFF', size=11)
        title_font = Font(name='Arial', bold=True, size=16, color='1E293B')
        center = Alignment(horizontal='center', vertical='center')

        biz_name = ''
        try:
            biz_name = (self.settings.get('business.name', '') or '') if getattr(self, 'settings', None) else ''
        except Exception:
            biz_name = ''
        ws.merge_cells('A1:H1')
        ws['A1'] = biz_name or 'בייבי בייסיק'
        ws['A1'].font = title_font
        ws['A1'].alignment = center
        ws.merge_cells('A2:H2')
        ws['A2'] = f"תעודת סחורה #{rec.get('id')}"
        ws['A2'].font = Font(name='Arial', bold=True, size=13)
        ws['A2'].alignment = center
        ws['A4'] = 'לקוח'
        ws['B4'] = rec.get('customer', '')
        ws['C4'] = 'תאריך'
        ws['D4'] = rec.get('date', '')
        if rec.get('note'):
            ws['A5'] = 'הערה'
            ws.merge_cells('B5:H5')
            ws['B5'] = rec.get('note')

        headers = ['מוצר', 'מידה', 'בד', 'צבע', 'ברקוד', 'יחידות', 'מחיר ליחידה', 'סה״כ']
        start_row = 7
        for col, h in enumerate(headers, 1):
            cell = ws.cell(start_row, col, h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = thin
        for i, line in enumerate(rec.get('lines') or [], 1):
            row = start_row + i
            values = [
                line.get('print_name') or line.get('item_name') or '',
                line.get('size', ''),
                line.get('fabric', ''),
                line.get('color', '') or 'לבן',
                line.get('barcode', ''),
                int(line.get('quantity') or 0),
                round(parse_numeric_value(line.get('unit_price')), 2),
                round(parse_numeric_value(line.get('line_total')), 2),
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row, col, val)
                cell.alignment = center
                cell.border = thin
                if col >= 7:
                    cell.number_format = '#,##0.00'
        total_row = start_row + 1 + len(rec.get('lines') or [])
        ws.cell(total_row, 5, 'סה״כ').font = Font(name='Arial', bold=True)
        ws.cell(total_row, 6, rec.get('total_quantity', 0)).font = Font(name='Arial', bold=True)
        total_cell = ws.cell(total_row, 8, round(parse_numeric_value(rec.get('total_amount')), 2))
        total_cell.font = Font(name='Arial', bold=True)
        total_cell.number_format = '#,##0.00'
        widths = [28, 12, 12, 10, 18, 12, 14, 14]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w

        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_notes')
        os.makedirs(export_dir, exist_ok=True)
        safe_id = rec.get('id')
        safe_date = str(rec.get('date') or '').replace(':', '-')
        default_name = f"baby_basic_note_{safe_id}_{safe_date}.xlsx"
        if not path:
            path = filedialog.asksaveasfilename(
                title="שמירת תעודת סחורה",
                defaultextension=".xlsx",
                initialdir=export_dir,
                initialfile=default_name,
                filetypes=[("Excel", "*.xlsx")],
            )
        if not path:
            return
        wb.save(path)
        try:
            os.startfile(path)
        except Exception:
            pass

    def _bb_print_selected_note_pdf(self):
        tree = getattr(self, 'bb_saved_notes_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("הדפסה", "בחר תעודה מהרשימה.")
            return
        rec = self.data_processor.get_baby_basic_note(sel[0])
        if rec:
            self._bb_export_note_pdf(rec)

    def _bb_export_note_pdf(self, rec: dict, path: str | None = None):
        try:
            from optitex_analyzer.core.baby_basic_note_pdf import generate_baby_basic_note_pdf
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן ליצור PDF:\n{e}")
            return
        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_notes')
        os.makedirs(export_dir, exist_ok=True)
        safe_id = rec.get('id')
        safe_date = str(rec.get('date') or '').replace(':', '-')
        default_name = f"baby_basic_note_{safe_id}_{safe_date}.pdf"
        if not path:
            path = os.path.join(export_dir, default_name)
        biz_name = ''
        try:
            biz_name = (self.settings.get('business.name', '') or '') if getattr(self, 'settings', None) else ''
        except Exception:
            biz_name = ''
        try:
            generate_baby_basic_note_pdf(rec, path, biz_name=biz_name or 'בייבי בייסיק')
        except Exception as e:
            messagebox.showerror("שגיאה", f"יצירת PDF נכשלה:\n{e}")
            return
        try:
            os.startfile(path)
        except Exception:
            messagebox.showinfo("PDF", f"הקובץ נשמר:\n{path}")

    # ---------- חשבון שותף: בייבי בייסיק ----------
    def _bb_partner_name(self) -> str:
        return getattr(self.data_processor, 'BABY_BASIC_PARTNER_NAME', 'בייבי בייסיק')

    def _bb_refresh_account(self):
        acc = self.data_processor.get_baby_basic_account()
        if hasattr(self, 'bb_acc_supplied_var'):
            self.bb_acc_supplied_var.set(f"{self._bb_money(acc['supplied'])} ₪")
        if hasattr(self, 'bb_acc_paid_var'):
            self.bb_acc_paid_var.set(f"{self._bb_money(acc['paid'])} ₪")
        if hasattr(self, 'bb_acc_balance_var'):
            self.bb_acc_balance_var.set(f"{self._bb_money(acc['balance'])} ₪")
        if hasattr(self, 'bb_acc_balance_label'):
            color = theme.DANGER if acc['balance'] > 0.009 else theme.SUCCESS
            self.bb_acc_balance_label.configure(fg=color)

        notes_tree = getattr(self, 'bb_acc_notes_tree', None)
        if notes_tree is not None:
            self._bb_clear_tree(notes_tree)
            notes = sorted(acc['notes'], key=lambda n: (str(n.get('date') or ''), int(n.get('id') or 0)), reverse=True)
            for rec in notes:
                notes_tree.insert('', 'end', iid=f"n{rec.get('id')}", values=(
                    rec.get('date', ''),
                    rec.get('id', ''),
                    rec.get('customer', ''),
                    rec.get('total_quantity', 0),
                    self._bb_money(rec.get('total_amount')),
                    rec.get('note', ''),
                ))
            theme.stripe_tree(notes_tree)

        pay_tree = getattr(self, 'bb_acc_pay_tree', None)
        if pay_tree is not None:
            self._bb_clear_tree(pay_tree)
            pays = sorted(acc['payments'], key=lambda p: (str(p.get('date') or ''), int(p.get('id') or 0)), reverse=True)
            for rec in pays:
                pay_tree.insert('', 'end', iid=str(rec.get('id')), values=(
                    rec.get('date', ''),
                    rec.get('id', ''),
                    rec.get('customer') or self._bb_partner_name(),
                    self._bb_money(rec.get('amount')),
                    rec.get('note', ''),
                ))
            theme.stripe_tree(pay_tree)

        hist_tree = getattr(self, 'bb_acc_hist_tree', None)
        if hist_tree is not None:
            self._bb_clear_tree(hist_tree)
            rows = []
            running = 0.0
            events = []
            for n in acc['notes']:
                events.append(('note', n.get('date'), int(n.get('id') or 0), n))
            for p in acc['payments']:
                events.append(('pay', p.get('date'), int(p.get('id') or 0), p))
            events.sort(key=lambda x: (str(x[1] or ''), 0 if x[0] == 'note' else 1, x[2]))
            for kind, date, _id, rec in events:
                if kind == 'note':
                    amount = parse_numeric_value(rec.get('total_amount'))
                    running += amount
                    rows.append((date, 'תעודת סחורה', rec.get('id'), rec.get('customer') or self._bb_partner_name(), amount, 0, running, rec.get('note', '')))
                else:
                    amount = parse_numeric_value(rec.get('amount'))
                    running -= amount
                    rows.append((date, 'תשלום', rec.get('id'), rec.get('customer') or self._bb_partner_name(), 0, amount, running, rec.get('note', '')))
            for i, row in enumerate(reversed(rows)):
                hist_tree.insert('', 'end', iid=str(i), values=(
                    row[0], row[1], row[2], row[3],
                    self._bb_money(row[4]) if row[4] else '',
                    self._bb_money(row[5]) if row[5] else '',
                    self._bb_money(row[6]),
                    row[7],
                ))
            theme.stripe_tree(hist_tree)

        self._bb_refresh_customer_combos()

    def _bb_add_payment(self):
        date_str = (self.bb_pay_date_var.get() if hasattr(self, 'bb_pay_date_var') else '').strip()
        amount = parse_numeric_value(self.bb_pay_amount_var.get() if hasattr(self, 'bb_pay_amount_var') else 0)
        note = (self.bb_pay_note_var.get() if hasattr(self, 'bb_pay_note_var') else '').strip()
        try:
            self.data_processor.add_baby_basic_payment(date_str=date_str, amount=amount, note=note)
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
            return
        if hasattr(self, 'bb_pay_amount_var'):
            self.bb_pay_amount_var.set('')
        if hasattr(self, 'bb_pay_note_var'):
            self.bb_pay_note_var.set('')
        self._bb_refresh_account()

    def _bb_delete_selected_payment(self):
        tree = getattr(self, 'bb_acc_pay_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        if not messagebox.askyesno("מחיקה", f"למחוק תשלום #{sel[0]}?"):
            return
        self.data_processor.delete_baby_basic_payment(sel[0])
        self._bb_refresh_account()

    def _bb_open_account_note(self, event=None):
        tree = getattr(self, 'bb_acc_notes_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        note_id = str(sel[0]).lstrip('n')
        rec = self.data_processor.get_baby_basic_note(note_id)
        if rec:
            self._bb_open_note_view(rec)
