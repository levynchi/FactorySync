import os
from datetime import datetime
from tkinter import messagebox, filedialog

from optitex_analyzer.core.data_processor import parse_numeric_value
from .. import theme


class BabyBasicSalesMethodsMixin:
    """Logic for Baby Basic wholesale notes, price list and account."""

    BB_ROOM_COUNT_EXTRA_MODELS = ('אוברול', 'חולצות טורקיה', 'גופייה טורקיה')
    BB_ROOM_COUNT_COLORS = (
        'חום', 'כחול', 'שחור', 'מנטה', 'אפור בהיר', 'לבן',
        'ורוד בהיר', 'כחול נייבי', 'כחול מלאנז', 'ורוד בייבי מלאנז', 'פוקסיה',
    )
    BB_ROOM_COUNT_PRINTS = (
        'משולשים', 'מנומר', 'נקודות', 'חלל', 'באר שבע בנות',
        'פרפרים', 'באר שבע בנים', 'נוצות', 'אוהלים', 'עלים',
    )
    BB_ROOM_COUNT_FABRICS = ('פלנל', 'טריקו', 'ופל')

    def _bb_money(self, value) -> str:
        try:
            return f"{parse_numeric_value(value):,.2f}"
        except Exception:
            return "0.00"

    # ---------- מקור סחורה (ייצור מקומי / סחורה טורקית OMER) ----------
    BB_SOURCE_FILTER_ALL = 'הכל'

    def _bb_sources(self) -> dict:
        return dict(getattr(self.data_processor, 'BABY_BASIC_SOURCES', None) or {
            'local': 'ייצור מקומי',
            'turkey': 'סחורה טורקית (OMER)',
        })

    def _bb_source_label(self, key) -> str:
        norm = getattr(self.data_processor, 'normalize_baby_basic_source', None)
        k = norm(key) if norm else (key or 'local')
        return self._bb_sources().get(k, '')

    def _bb_source_labels(self) -> list:
        # סדר קבוע: ייצור מקומי ואז סחורה טורקית
        srcs = self._bb_sources()
        return [srcs[k] for k in ('local', 'turkey') if k in srcs] + [v for k, v in srcs.items() if k not in ('local', 'turkey')]

    def _bb_source_filter_labels(self) -> list:
        return [self.BB_SOURCE_FILTER_ALL] + self._bb_source_labels()

    def _bb_source_key_from_label(self, label: str) -> str:
        """תווית בעברית -> מפתח ('local' / 'turkey'); 'הכל' או ריק -> ''."""
        label = str(label or '').strip()
        if not label or label == self.BB_SOURCE_FILTER_ALL:
            return ''
        for k, v in self._bb_sources().items():
            if v == label:
                return k
        norm = getattr(self.data_processor, 'normalize_baby_basic_source', None)
        return norm(label) if norm else 'local'

    # ---------- סוג תעודה (אספקה לבייבי בייסיק / משיכה לאריה) ----------
    BB_KIND_FORM_LABELS = {
        'supply': 'אספקה לבייבי בייסיק',
        'withdrawal': 'משיכה לאריה (מהחדר של בייבי בייסיק)',
    }

    def _bb_kinds(self) -> dict:
        return dict(getattr(self.data_processor, 'BABY_BASIC_NOTE_KINDS', None) or {
            'supply': 'תעודת סחורה',
            'withdrawal': 'משיכה לאריה',
        })

    def _bb_normalize_kind(self, kind) -> str:
        norm = getattr(self.data_processor, 'normalize_baby_basic_note_kind', None)
        if norm:
            return norm(kind)
        return 'withdrawal' if str(kind or '').strip().lower() == 'withdrawal' else 'supply'

    def _bb_kind_label(self, kind) -> str:
        """תווית קצרה לטבלאות: «תעודת סחורה» / «משיכה לאריה»."""
        return self._bb_kinds().get(self._bb_normalize_kind(kind), '')

    def _bb_is_withdrawal(self, rec: dict) -> bool:
        return self._bb_normalize_kind((rec or {}).get('kind')) == 'withdrawal'

    def _bb_kind_form_labels(self) -> list:
        return [self.BB_KIND_FORM_LABELS['supply'], self.BB_KIND_FORM_LABELS['withdrawal']]

    def _bb_current_note_kind(self) -> str:
        var = getattr(self, 'bb_note_kind_var', None)
        label = str(var.get() if var is not None else '').strip()
        for key, lbl in self.BB_KIND_FORM_LABELS.items():
            if lbl == label:
                return key
        return 'supply'

    def _bb_note_title(self, rec: dict) -> str:
        """כותרת תעודה לפי סוג: «תעודת סחורה #N» / «משיכת סחורה לאריה #N»."""
        if self._bb_is_withdrawal(rec):
            return f"משיכת סחורה לאריה #{rec.get('id')}"
        return f"תעודת סחורה #{rec.get('id')}"

    def _bb_on_note_kind_change(self):
        """משיכה: השותף נעול על בייבי בייסיק, הטקסטים הופכים לזיכוי."""
        is_wd = self._bb_current_note_kind() == 'withdrawal'
        combo = getattr(self, 'bb_customer_combo', None)
        if combo is not None:
            try:
                if is_wd:
                    if hasattr(self, 'bb_customer_var'):
                        self.bb_customer_var.set(self._bb_partner_name())
                    combo.configure(state='disabled')
                else:
                    combo.configure(state='normal')
            except Exception:
                pass
        btn = getattr(self, 'bb_note_save_btn', None)
        if btn is not None:
            try:
                btn.configure(text='שמור משיכה' if is_wd else 'שמור תעודה')
            except Exception:
                pass
        if hasattr(self, 'bb_note_kind_hint_var'):
            self.bb_note_kind_hint_var.set(
                'אריה לוקח מהחדר של בייבי בייסיק — הסכום מזכה את בייבי בייסיק ויורד מהיתרה.' if is_wd else ''
            )
        self._bb_refresh_note_lines()

    def _bb_current_account_source(self) -> str:
        var = getattr(self, 'bb_acc_filter_var', None)
        return self._bb_source_key_from_label(var.get() if var is not None else '')

    def _bb_on_account_filter_change(self):
        # ברירת מחדל לתשלום חדש = הפילוח שנבחר (אם לא "הכל")
        src = self._bb_current_account_source()
        if src and hasattr(self, 'bb_pay_source_var'):
            self.bb_pay_source_var.set(self._bb_source_label(src))
        self._bb_refresh_account()

    def _bb_set_account_note_source(self, source: str):
        tree = getattr(self, 'bb_acc_notes_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("מקור סחורה", "בחר תעודה אחת או יותר מהרשימה.")
            return
        for iid in sel:
            note_id = str(iid).lstrip('n')
            self.data_processor.set_baby_basic_note_source(note_id, source)
        self._bb_refresh_account()
        self._bb_refresh_saved_notes()

    def _bb_set_saved_note_source(self, source: str):
        tree = getattr(self, 'bb_saved_notes_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("מקור סחורה", "בחר תעודה אחת או יותר מהרשימה.")
            return
        for iid in sel:
            self.data_processor.set_baby_basic_note_source(iid, source)
        self._bb_refresh_saved_notes()
        self._bb_refresh_account()

    def _bb_set_account_payment_source(self, source: str):
        tree = getattr(self, 'bb_acc_pay_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("מקור סחורה", "בחר תשלום אחד או יותר מהרשימה.")
            return
        for iid in sel:
            self.data_processor.set_baby_basic_payment_source(iid, source)
        self._bb_refresh_account()

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
        # פריטים ללא ברקוד (אוברול, טריקו, סדינים וכו') — מפתחות nb| במחירון
        nb_items = self._bb_filter_unbarcoded_prices(query)
        if nb_items:
            tree.insert('', 'end', iid='__nb_header__', values=('— פריטים ללא ברקוד —', '', '', '', '', '', ''), tags=('section',))
            for it in nb_items:
                tree.insert('', 'end', iid=it['key'], values=(
                    it['model'],
                    it['size'],
                    it['fabric'],
                    ' / '.join(x for x in (it['color'], it['print']) if x),
                    '',
                    'ללא ברקוד',
                    self._bb_money(it['price']),
                ))
        theme.stripe_tree(tree)
        try:
            tree.tag_configure('section', background=theme.PANEL_BG, foreground=theme.PRIMARY_DARK, font=theme.FONT_BODY_BOLD)
        except Exception:
            pass

    def _bb_unbarcoded_prices(self):
        try:
            items = self.data_processor.get_baby_basic_unbarcoded_price_items()
        except Exception:
            items = []
        items.sort(key=lambda it: (it['model'], it['fabric'], self._bb_room_size_sort_key(it['size']), it['color'], it['print']))
        return items

    def _bb_filter_unbarcoded_prices(self, query: str = ''):
        terms = (query or '').strip().lower().split()
        result = []
        for it in self._bb_unbarcoded_prices():
            blob = ' '.join([it['model'], it['fabric'], it['color'], it['print'], it['size'], 'ללא ברקוד']).lower()
            if terms and not all(t in blob for t in terms):
                continue
            result.append(it)
        return result

    def _bb_edit_price_cell(self, event):
        def on_commit(key, raw):
            if key == '__nb_header__':
                return
            try:
                price = parse_numeric_value(raw)
            except Exception:
                price = 0
            self.data_processor.set_baby_basic_price(key, price)
            self._bb_refresh_price_table()
            self._bb_refresh_note_products()
        self._bb_inline_edit(self.bb_price_tree, event, '#7', on_commit)

    def _bb_add_unbarcoded_price_item(self):
        """חלון קטן להוספת פריט ללא ברקוד למחירון: דגם / בד / צבע / הדפס / מידה / מחיר."""
        import tkinter as tk
        from tkinter import ttk
        opts = self._bb_note_list_catalog() if hasattr(self, '_bb_note_list_catalog') else {}
        win = tk.Toplevel(self.root)
        win.title("הוספת פריט ללא ברקוד למחירון")
        win.configure(bg=theme.PAGE_BG)
        win.transient(self.root)
        win.grab_set()
        fields = (
            ('דגם', 'model', opts.get('models') or []),
            ('בד', 'fabric', opts.get('fabrics') or []),
            ('צבע', 'color', opts.get('colors') or []),
            ('הדפס', 'print', opts.get('prints') or []),
            ('מידה', 'size', opts.get('sizes') or []),
        )
        vars_ = {}
        for row, (label, name, values) in enumerate(fields):
            tk.Label(win, text=label, bg=theme.PAGE_BG).grid(row=row, column=1, sticky='e', padx=6, pady=4)
            var = tk.StringVar()
            ttk.Combobox(win, textvariable=var, values=list(values), width=28, justify='right').grid(row=row, column=0, padx=6, pady=4)
            vars_[name] = var
        tk.Label(win, text='מחיר ליחידה ₪', bg=theme.PAGE_BG).grid(row=len(fields), column=1, sticky='e', padx=6, pady=4)
        price_var = tk.StringVar()
        price_entry = ttk.Entry(win, textvariable=price_var, width=12, justify='center')
        price_entry.grid(row=len(fields), column=0, sticky='w', padx=6, pady=4)
        if vars_['color'].get() == '' and 'לבן' in (opts.get('colors') or []):
            vars_['color'].set('לבן')

        def save(_e=None):
            model = vars_['model'].get().strip()
            if not model:
                messagebox.showwarning("מחירון", "הזן דגם.", parent=win)
                return
            price = parse_numeric_value(price_var.get())
            if price <= 0:
                messagebox.showwarning("מחירון", "הזן מחיר גדול מאפס.", parent=win)
                return
            key = self.data_processor.baby_basic_unbarcoded_price_key(
                model, vars_['fabric'].get(), vars_['color'].get(), vars_['print'].get(), vars_['size'].get())
            self.data_processor.set_baby_basic_price(key, price)
            fabric = vars_['fabric'].get().strip()
            if fabric and hasattr(self.data_processor, 'remember_baby_basic_model_fabric'):
                self.data_processor.remember_baby_basic_model_fabric(model, fabric)
            if hasattr(self.data_processor, 'add_baby_basic_note_list_item'):
                self.data_processor.add_baby_basic_note_list_item('models', model)
                self.data_processor.add_baby_basic_note_list_item('sizes', vars_['size'].get())
            win.destroy()
            self._bb_refresh_price_table()
            if hasattr(self, '_bb_refresh_note_list_combos'):
                self._bb_refresh_note_list_combos()
            try:
                self.bb_price_tree.see(key)
                self.bb_price_tree.selection_set(key)
            except Exception:
                pass

        btns = tk.Frame(win, bg=theme.PAGE_BG)
        btns.grid(row=len(fields) + 1, column=0, columnspan=2, pady=8)
        theme.make_button(btns, "שמור", kind="primary", command=save).pack(side='right', padx=4)
        theme.make_button(btns, "ביטול", kind="secondary", command=win.destroy).pack(side='right', padx=4)
        price_entry.bind('<Return>', save)
        win.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - win.winfo_height()) // 3
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _bb_delete_unbarcoded_price_item(self):
        tree = getattr(self, 'bb_price_tree', None)
        if tree is None:
            return
        sel = [k for k in tree.selection() if str(k).startswith('nb|')]
        if not sel:
            messagebox.showwarning("מחירון", "בחר שורה של פריט ללא ברקוד למחיקה.")
            return
        if not messagebox.askyesno("מחירון", f"למחוק {len(sel)} פריטים ללא ברקוד מהמחירון?"):
            return
        for key in sel:
            self.data_processor.delete_baby_basic_price(key)
        self._bb_refresh_price_table()

    def _bb_apply_price_to_family(self):
        tree = getattr(self, 'bb_price_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("מחירון", "בחר מוצר בטבלה ואז החל את המחיר לכל המידות של אותו דגם.")
            return
        if str(sel[0]).startswith('nb|'):
            self._bb_apply_unbarcoded_price_to_family(sel[0])
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

    def _bb_apply_unbarcoded_price_to_family(self, key: str):
        items = {it['key']: it for it in self._bb_unbarcoded_prices()}
        src = items.get(key)
        if not src:
            return
        tree = self.bb_price_tree
        try:
            price = parse_numeric_value(tree.item(key, 'values')[-1])
        except Exception:
            price = src.get('price') or 0
        if price <= 0:
            messagebox.showwarning("מחירון", "קודם הזן מחיר לשורה שנבחרה (דאבל-קליק על עמודת המחיר).")
            return
        updates = {
            k: price for k, it in items.items()
            if it['model'] == src['model'] and it['fabric'] == src['fabric']
        }
        self.data_processor.set_baby_basic_prices_bulk(updates)
        self._bb_refresh_price_table()
        messagebox.showinfo("מחירון", f"עודכן מחיר {self._bb_money(price)} ₪ ל-{len(updates)} מידות של «{src['model']}» ({src['fabric']}).")

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
                self._bb_line_color(line),
                line.get('print', '') or '',
                self._bb_display_barcode(line.get('barcode')),
                int(line.get('quantity') or 0),
                self._bb_money(line.get('unit_price')),
                self._bb_money(line.get('line_total')),
            ))
        theme.stripe_tree(tree)
        if hasattr(self, 'bb_note_total_var'):
            label = 'סה״כ לזיכוי' if self._bb_current_note_kind() == 'withdrawal' else 'סה״כ לתשלום'
            self.bb_note_total_var.set(f"סה״כ יחידות: {qty_total}    {label}: {self._bb_money(total)} ₪")

    def _bb_add_selected_to_note(self):
        qty_map = {bc: q for bc, q in getattr(self, '_bb_note_qty_by_barcode', {}).items() if int(q or 0) > 0}
        if not qty_map:
            messagebox.showwarning("תעודה", "לא הוזנה כמות. לחץ פעמיים על עמודת «כמות» ואז הוסף לתעודה.")
            return
        if not hasattr(self, '_bb_note_lines'):
            self._bb_note_lines = []
        existing = {str(l.get('barcode')): l for l in self._bb_note_lines if str(l.get('barcode') or '').strip()}
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
                    'color': prod.get('color', '') or '',
                    'print': '',
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
        self._bb_inline_edit(self.bb_note_lines_tree, event, '#8', on_commit)

    def _bb_display_barcode(self, barcode) -> str:
        bc = str(barcode or '').strip()
        return bc if bc else '—'

    def _bb_line_color(self, line: dict) -> str:
        color = str((line or {}).get('color') or '').strip()
        if str((line or {}).get('barcode') or '').strip():
            return color or 'לבן'
        return color

    def _bb_note_list_catalog(self):
        stored = {}
        try:
            stored = self.data_processor.get_baby_basic_note_lists() or {}
        except Exception:
            stored = {}
        models, fabrics, sizes = [], [], []
        for p in self._bb_products():
            name = (p.get('print_name') or p.get('item_name') or '').strip()
            if name:
                models.append(name)
            fabric = (p.get('fabric') or '').strip()
            if fabric:
                fabrics.append(fabric)
            size = (p.get('size') or '').strip()
            if size:
                sizes.append(size)
        defaults = getattr(self.data_processor, 'BABY_BASIC_NOTE_LIST_DEFAULTS', {}) or {}
        sizes = self._bb_room_merge_names(defaults.get('sizes'), stored.get('sizes'), sizes)
        return {
            'models': self._bb_room_merge_names(defaults.get('models'), stored.get('models'), models),
            'fabrics': self._bb_room_merge_names(defaults.get('fabrics'), stored.get('fabrics'), fabrics),
            'colors': self._bb_room_merge_names(defaults.get('colors'), stored.get('colors')),
            'prints': self._bb_room_merge_names(defaults.get('prints'), stored.get('prints')),
            'sizes': sorted(sizes, key=self._bb_room_size_sort_key),
        }

    def _bb_refresh_note_list_combos(self, event=None):
        opts = self._bb_note_list_catalog()
        mapping = (
            ('bb_note_nb_model_combo', 'models'),
            ('bb_note_nb_fabric_combo', 'fabrics'),
            ('bb_note_nb_color_combo', 'colors'),
            ('bb_note_nb_print_combo', 'prints'),
            ('bb_note_nb_size_combo', 'sizes'),
        )
        for attr, key in mapping:
            combo = getattr(self, attr, None)
            if combo is None:
                continue
            values = opts.get(key) or []
            try:
                if hasattr(combo, 'set_completion_list'):
                    combo.set_completion_list(values)
                else:
                    combo['values'] = values
            except Exception:
                pass

    def _bb_nb_focus(self, widget):
        if widget is None:
            return
        try:
            widget.focus_set()
            if hasattr(widget, 'select_range'):
                widget.select_range(0, 'end')
                widget.icursor('end')
        except Exception:
            pass

    def _bb_nb_get(self, field: str) -> str:
        """ערך שדה בטופס ללא ברקוד — כולל השלמה מוצעת שעדיין לא אושרה."""
        combo = getattr(self, f'bb_note_nb_{field}_combo', None)
        if combo is not None and hasattr(combo, 'get_value'):
            try:
                return str(combo.get_value() or '').strip()
            except Exception:
                pass
        var = getattr(self, f'bb_note_nb_{field}_var', None)
        try:
            return str(var.get() or '').strip() if var is not None else ''
        except Exception:
            return ''

    def _bb_on_nb_model_chosen(self, event=None, advance=False):
        model = self._bb_nb_get('model')
        fabrics = []
        try:
            fabrics = list(self.data_processor.get_baby_basic_model_fabrics(model) or [])
        except Exception:
            fabrics = []
        combo = getattr(self, 'bb_note_nb_fabric_combo', None)
        if combo is not None and fabrics:
            existing = []
            try:
                existing = list(getattr(combo, '_completion_list', None) or combo['values'] or [])
            except Exception:
                existing = []
            ordered = self._bb_room_merge_names(fabrics, existing)
            try:
                if hasattr(combo, 'set_completion_list'):
                    combo.set_completion_list(ordered)
                else:
                    combo['values'] = ordered
            except Exception:
                pass
        if len(fabrics) == 1 and hasattr(self, 'bb_note_nb_fabric_var'):
            self.bb_note_nb_fabric_var.set(fabrics[0])
        self._bb_fill_unbarcoded_price()
        if not advance:
            return
        if len(fabrics) == 1:
            self._bb_nb_focus(getattr(self, 'bb_note_nb_color_combo', None))
        else:
            self._bb_nb_focus(getattr(self, 'bb_note_nb_fabric_combo', None))

    def _bb_nb_advance(self, from_widget=None):
        model_combo = getattr(self, 'bb_note_nb_model_combo', None)
        qty_entry = getattr(self, 'bb_note_nb_qty_entry', None)
        price_entry = getattr(self, 'bb_note_nb_price_entry', None)
        if from_widget is model_combo:
            self._bb_on_nb_model_chosen(advance=True)
            return
        if from_widget is qty_entry:
            price = parse_numeric_value(self.bb_note_nb_price_var.get() if hasattr(self, 'bb_note_nb_price_var') else 0)
            if price > 0:
                self._bb_add_unbarcoded_line()
            else:
                self._bb_nb_focus(price_entry)
            return
        if from_widget is price_entry:
            self._bb_add_unbarcoded_line()
            return
        order = [
            getattr(self, 'bb_note_nb_fabric_combo', None),
            getattr(self, 'bb_note_nb_color_combo', None),
            getattr(self, 'bb_note_nb_print_combo', None),
            getattr(self, 'bb_note_nb_size_combo', None),
            qty_entry,
            price_entry,
        ]
        nxt = None
        seen = False
        for widget in order:
            if widget is None:
                continue
            if seen:
                nxt = widget
                break
            if widget is from_widget:
                seen = True
        self._bb_fill_unbarcoded_price()
        if nxt is not None:
            self._bb_nb_focus(nxt)

    def _bb_unbarcoded_price_key(self):
        return self.data_processor.baby_basic_unbarcoded_price_key(
            self._bb_nb_get('model'),
            self._bb_nb_get('fabric'),
            self._bb_nb_get('color'),
            self._bb_nb_get('print'),
            self._bb_nb_get('size'),
        )

    def _bb_fill_unbarcoded_price(self, event=None):
        if not hasattr(self, 'bb_note_nb_price_var'):
            return
        try:
            price = self.data_processor.get_baby_basic_unbarcoded_price(
                self._bb_nb_get('model'),
                self._bb_nb_get('fabric'),
                self._bb_nb_get('color'),
                self._bb_nb_get('print'),
                self._bb_nb_get('size'),
            )
        except Exception:
            price = 0
        if price > 0:
            self.bb_note_nb_price_var.set(f"{price:.2f}")

    def _bb_add_unbarcoded_line(self):
        for field in ('model', 'fabric', 'color', 'print', 'size'):
            combo = getattr(self, f'bb_note_nb_{field}_combo', None)
            if combo is not None and hasattr(combo, 'commit'):
                try:
                    combo.commit()
                except Exception:
                    pass
        model = self._bb_nb_get('model')
        fabric = self._bb_nb_get('fabric')
        color = self._bb_nb_get('color')
        print_val = self._bb_nb_get('print')
        size = self._bb_nb_get('size')
        raw_qty = (self.bb_note_nb_qty_var.get() if hasattr(self, 'bb_note_nb_qty_var') else '').strip()
        raw_price = (self.bb_note_nb_price_var.get() if hasattr(self, 'bb_note_nb_price_var') else '').strip()
        if not model:
            messagebox.showwarning("תעודה", "בחר או הקלד דגם.")
            return
        try:
            qty = int(float(raw_qty)) if raw_qty else 0
        except Exception:
            qty = 0
        if qty <= 0:
            messagebox.showwarning("תעודה", "הזן כמות גדולה מאפס.")
            return
        price = parse_numeric_value(raw_price)
        if not hasattr(self, '_bb_note_lines'):
            self._bb_note_lines = []
        key = (model, fabric, color, print_val, size)
        merged = False
        for line in self._bb_note_lines:
            if str(line.get('barcode') or '').strip():
                continue
            existing = (
                str(line.get('print_name') or '').strip(),
                str(line.get('fabric') or '').strip(),
                str(line.get('color') or '').strip(),
                str(line.get('print') or '').strip(),
                str(line.get('size') or '').strip(),
            )
            if existing == key:
                line['quantity'] = int(line.get('quantity') or 0) + qty
                line['unit_price'] = round(price, 2)
                line['line_total'] = round(line['quantity'] * price, 2)
                merged = True
                break
        if not merged:
            self._bb_note_lines.append({
                'barcode': '',
                'item_name': model,
                'print_name': model,
                'size': size,
                'fabric': fabric,
                'color': color,
                'print': print_val,
                'pack_qty': 1,
                'quantity': qty,
                'unit_price': round(price, 2),
                'line_total': round(qty * price, 2),
            })
        self._bb_persist_note_list_values(model, fabric, color, print_val, size)
        if model and fabric:
            try:
                self.data_processor.remember_baby_basic_model_fabric(model, fabric)
            except Exception:
                pass
        if price > 0:
            try:
                self.data_processor.set_baby_basic_price(self._bb_unbarcoded_price_key(), price)
            except Exception:
                pass
        if hasattr(self, 'bb_note_nb_color_var'):
            self.bb_note_nb_color_var.set('')
        if hasattr(self, 'bb_note_nb_print_var'):
            self.bb_note_nb_print_var.set('')
        if hasattr(self, 'bb_note_nb_qty_var'):
            self.bb_note_nb_qty_var.set('')
        if hasattr(self, 'bb_note_nb_price_var'):
            self.bb_note_nb_price_var.set('')
        self._bb_refresh_note_list_combos()
        self._bb_refresh_note_lines()
        self._bb_nb_focus(getattr(self, 'bb_note_nb_color_combo', None))
        if price <= 0:
            messagebox.showwarning(
                "מחירון חסר",
                "השורה נוספה בלי מחיר. הזן מחיר בשדה או דאבל-קליק על עמודת המחיר בשורת התעודה.",
            )

    def _bb_persist_note_list_values(self, model, fabric, color, print_val, size):
        mapping = (
            ('models', model),
            ('fabrics', fabric),
            ('colors', color),
            ('prints', print_val),
            ('sizes', size),
        )
        for key, value in mapping:
            name = str(value or '').strip()
            if not name:
                continue
            try:
                self.data_processor.add_baby_basic_note_list_item(key, name)
            except Exception:
                pass

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
        source = self._bb_source_key_from_label(self.bb_note_source_var.get() if hasattr(self, 'bb_note_source_var') else '') or 'local'
        kind = self._bb_current_note_kind()
        lines = list(getattr(self, '_bb_note_lines', []) or [])
        try:
            new_id = self.data_processor.add_baby_basic_note(customer, date_str, lines, note=comment, source=source, kind=kind)
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
            return
        self._bb_refresh_customer_combos()
        self._bb_refresh_saved_notes()
        self._bb_refresh_account()
        self._bb_clear_current_note()
        if kind == 'withdrawal':
            messagebox.showinfo("נשמר", f"משיכה לאריה #{new_id} נשמרה — הסכום זוכה לבייבי בייסיק.")
        else:
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
                self._bb_kind_label(rec.get('kind')),
                rec.get('customer', ''),
                self._bb_source_label(rec.get('source')),
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
        if not messagebox.askyesno("מחיקה", f"למחוק {self._bb_note_title(rec)} לשותף {rec.get('customer')}?"):
            return
        self.data_processor.delete_baby_basic_note(rec.get('id'))
        self._bb_refresh_saved_notes()
        self._bb_refresh_account()
        self._bb_refresh_customer_combos()

    def _bb_open_note_view(self, rec: dict):
        import tkinter as tk
        from tkinter import ttk
        win = tk.Toplevel(self.root)
        is_wd = self._bb_is_withdrawal(rec)
        win.title(f"{self._bb_note_title(rec)} — בייבי בייסיק")
        win.configure(bg=theme.PAGE_BG)
        win.geometry("920x560")
        tk.Label(
            win,
            text=f"{self._bb_note_title(rec)}  |  {rec.get('customer')}  |  {rec.get('date')}  |  {self._bb_source_label(rec.get('source'))}",
            font=(theme.FONT_FAMILY, 13, 'bold'),
            bg=theme.PAGE_BG,
            fg=theme.DARK,
        ).pack(pady=(10, 4))
        if is_wd:
            tk.Label(
                win,
                text="סחורה שאריה לקח מהחדר של בייבי בייסיק — הסכום מזכה את בייבי בייסיק.",
                bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=theme.FONT_SMALL,
            ).pack()
        if rec.get('note'):
            tk.Label(win, text=rec.get('note'), bg=theme.PAGE_BG, fg=theme.SUBTEXT).pack()

        cols = ('print_name', 'size', 'fabric', 'color', 'print', 'barcode', 'qty', 'price', 'total')
        headers = {
            'print_name': 'מוצר', 'size': 'מידה', 'fabric': 'בד', 'color': 'צבע רקע',
            'print': 'הדפס', 'barcode': 'ברקוד', 'qty': 'יחידות', 'price': 'מחיר ליחידה', 'total': 'סה״כ',
        }
        tree = theme.make_treeview(win, columns=cols, show='headings', height=14)
        for c in cols:
            tree.heading(c, text=headers[c])
            tree.column(c, width=90 if c in ('size', 'qty', 'price', 'total', 'color', 'print') else 150, anchor='center')
        vs = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.pack(side='left', fill='both', expand=True, padx=(12, 0), pady=8)
        vs.pack(side='left', fill='y', pady=8)
        for line in rec.get('lines') or []:
            tree.insert('', 'end', values=(
                line.get('print_name') or line.get('item_name') or '',
                line.get('size', ''),
                line.get('fabric', ''),
                self._bb_line_color(line),
                line.get('print', '') or '',
                self._bb_display_barcode(line.get('barcode')),
                line.get('quantity', 0),
                self._bb_money(line.get('unit_price')),
                self._bb_money(line.get('line_total')),
            ))
        theme.stripe_tree(tree)

        footer = tk.Frame(win, bg=theme.PAGE_BG)
        footer.pack(fill='x', padx=12, pady=(0, 10))
        tk.Label(
            footer,
            text=f"סה״כ יחידות: {rec.get('total_quantity', 0)}    {'סה״כ לזיכוי' if is_wd else 'סה״כ'}: {self._bb_money(rec.get('total_amount'))} ₪",
            font=(theme.FONT_FAMILY, 11, 'bold'),
            bg=theme.PAGE_BG,
        ).pack(side='right')
        theme.make_button(footer, "ייצוא לאקסל", kind="success", command=lambda: self._bb_export_note_excel(rec)).pack(side='left')
        theme.make_button(footer, "PDF להדפסה (A4)", kind="primary", command=lambda: self._bb_export_note_pdf(rec)).pack(side='left', padx=6)
        theme.make_button(footer, "PDF ללא מחירים", kind="secondary", command=lambda: self._bb_export_note_pdf(rec, show_prices=False)).pack(side='left')

    def _bb_export_note_excel(self, rec: dict, path: str | None = None):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
        except Exception as e:
            messagebox.showerror("שגיאה", f"נדרש openpyxl לייצוא לאקסל:\n{e}")
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "משיכה לאריה" if self._bb_is_withdrawal(rec) else "תעודת סחורה"
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
        ws.merge_cells('A1:I1')
        ws['A1'] = biz_name or 'בייבי בייסיק'
        ws['A1'].font = title_font
        ws['A1'].alignment = center
        ws.merge_cells('A2:I2')
        ws['A2'] = self._bb_note_title(rec)
        ws['A2'].font = Font(name='Arial', bold=True, size=13)
        ws['A2'].alignment = center
        ws['A4'] = 'לקוח'
        ws['B4'] = rec.get('customer', '')
        ws['C4'] = 'תאריך'
        ws['D4'] = rec.get('date', '')
        if rec.get('note'):
            ws['A5'] = 'הערה'
            ws.merge_cells('B5:I5')
            ws['B5'] = rec.get('note')

        headers = ['מוצר', 'מידה', 'בד', 'צבע רקע', 'הדפס', 'ברקוד', 'יחידות', 'מחיר ליחידה', 'סה״כ']
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
                self._bb_line_color(line),
                line.get('print', '') or '',
                self._bb_display_barcode(line.get('barcode')),
                int(line.get('quantity') or 0),
                round(parse_numeric_value(line.get('unit_price')), 2),
                round(parse_numeric_value(line.get('line_total')), 2),
            ]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row, col, val)
                cell.alignment = center
                cell.border = thin
                if col >= 8:
                    cell.number_format = '#,##0.00'
        total_row = start_row + 1 + len(rec.get('lines') or [])
        ws.cell(total_row, 6, 'סה״כ').font = Font(name='Arial', bold=True)
        ws.cell(total_row, 7, rec.get('total_quantity', 0)).font = Font(name='Arial', bold=True)
        total_cell = ws.cell(total_row, 9, round(parse_numeric_value(rec.get('total_amount')), 2))
        total_cell.font = Font(name='Arial', bold=True)
        total_cell.number_format = '#,##0.00'
        widths = [28, 12, 12, 12, 14, 16, 12, 14, 14]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w

        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_notes')
        os.makedirs(export_dir, exist_ok=True)
        safe_id = rec.get('id')
        safe_date = str(rec.get('date') or '').replace(':', '-')
        prefix = 'baby_basic_withdrawal' if self._bb_is_withdrawal(rec) else 'baby_basic_note'
        default_name = f"{prefix}_{safe_id}_{safe_date}.xlsx"
        if not path:
            path = filedialog.asksaveasfilename(
                title=f"שמירת {self._bb_note_title(rec)}",
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

    def _bb_print_selected_note_pdf(self, show_prices: bool = True):
        tree = getattr(self, 'bb_saved_notes_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("הדפסה", "בחר תעודה מהרשימה.")
            return
        rec = self.data_processor.get_baby_basic_note(sel[0])
        if rec:
            self._bb_export_note_pdf(rec, show_prices=show_prices)

    def _bb_export_note_pdf(self, rec: dict, path: str | None = None, show_prices: bool = True):
        try:
            from optitex_analyzer.core.baby_basic_note_pdf import generate_baby_basic_note_pdf
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן ליצור PDF:\n{e}")
            return
        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_notes')
        os.makedirs(export_dir, exist_ok=True)
        safe_id = rec.get('id')
        safe_date = str(rec.get('date') or '').replace(':', '-')
        suffix = '' if show_prices else '_no_prices'
        is_wd = self._bb_is_withdrawal(rec)
        prefix = 'baby_basic_withdrawal' if is_wd else 'baby_basic_note'
        default_name = f"{prefix}_{safe_id}_{safe_date}{suffix}.pdf"
        if not path:
            path = os.path.join(export_dir, default_name)
        biz_name = ''
        try:
            biz_name = (self.settings.get('business.name', '') or '') if getattr(self, 'settings', None) else ''
        except Exception:
            biz_name = ''
        pdf_kwargs = {}
        if is_wd:
            pdf_kwargs = {
                'title': f"משיכת סחורה לאריה #{rec.get('id')} — מהחדר של בייבי בייסיק",
                'subtitle': 'סחורה שאריה לקח מהחדר של בייבי בייסיק — הסכום מזכה את בייבי בייסיק ויורד מהיתרה',
                'total_label': 'סה״כ לזיכוי',
            }
        try:
            generate_baby_basic_note_pdf(rec, path, biz_name=biz_name or 'בייבי בייסיק', show_prices=show_prices, **pdf_kwargs)
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
        acc = self.data_processor.get_baby_basic_account(source=self._bb_current_account_source())
        if hasattr(self, 'bb_acc_supplied_var'):
            self.bb_acc_supplied_var.set(f"{self._bb_money(acc['supplied'])} ₪")
        if hasattr(self, 'bb_acc_withdrawn_var'):
            self.bb_acc_withdrawn_var.set(f"{self._bb_money(acc.get('withdrawn', 0))} ₪")
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
            all_notes = list(acc['notes']) + list(acc.get('withdrawals') or [])
            notes = sorted(all_notes, key=lambda n: (str(n.get('date') or ''), int(n.get('id') or 0)), reverse=True)
            for rec in notes:
                notes_tree.insert('', 'end', iid=f"n{rec.get('id')}", values=(
                    rec.get('date', ''),
                    rec.get('id', ''),
                    self._bb_kind_label(rec.get('kind')),
                    rec.get('customer', ''),
                    self._bb_source_label(rec.get('source')),
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
                    self._bb_source_label(rec.get('source')),
                    self._bb_money(rec.get('amount')),
                    rec.get('note', ''),
                ))
            theme.stripe_tree(pay_tree)

        hist_tree = getattr(self, 'bb_acc_hist_tree', None)
        if hist_tree is not None:
            self._bb_clear_tree(hist_tree)
            rows = self._bb_build_ledger_rows(acc)
            for i, row in enumerate(reversed(rows)):
                hist_tree.insert('', 'end', iid=str(i), values=(
                    row['date'], row['type'], row['id'], row['partner'], row['source'],
                    self._bb_money(row['debit']) if row['debit'] else '',
                    self._bb_money(row['credit']) if row['credit'] else '',
                    self._bb_money(row['balance']),
                    row['note'],
                ))
            theme.stripe_tree(hist_tree)

        self._bb_refresh_customer_combos()

    def _bb_build_ledger_rows(self, acc: dict) -> list:
        """שורות כרטסת כרונולוגיות (תעודות + תשלומים) עם יתרה רצה — משותף למסך ול-PDF."""
        rows = []
        running = 0.0
        events = []
        for n in acc.get('notes') or []:
            events.append(('note', n.get('date'), int(n.get('id') or 0), n))
        for n in acc.get('withdrawals') or []:
            events.append(('withdrawal', n.get('date'), int(n.get('id') or 0), n))
        for p in acc.get('payments') or []:
            events.append(('pay', p.get('date'), int(p.get('id') or 0), p))
        order = {'note': 0, 'withdrawal': 1, 'pay': 2}
        events.sort(key=lambda x: (str(x[1] or ''), order.get(x[0], 9), x[2]))
        for kind, date, _id, rec in events:
            src_label = self._bb_source_label(rec.get('source'))
            partner = rec.get('customer') or self._bb_partner_name()
            if kind == 'note':
                amount = parse_numeric_value(rec.get('total_amount'))
                running += amount
                debit, credit, kind_label = amount, 0, 'תעודת סחורה'
            elif kind == 'withdrawal':
                # אריה לקח מהחדר של בייבי בייסיק — זיכוי, יורד מהיתרה
                amount = parse_numeric_value(rec.get('total_amount'))
                running -= amount
                debit, credit, kind_label = 0, amount, 'משיכה לאריה'
            else:
                amount = parse_numeric_value(rec.get('amount'))
                running -= amount
                debit, credit, kind_label = 0, amount, 'תשלום'
            rows.append({
                'date': date or '',
                'type': kind_label,
                'id': rec.get('id'),
                'partner': partner,
                'source': src_label,
                'debit': debit,
                'credit': credit,
                'balance': round(running, 2),
                'note': rec.get('note', '') or '',
            })
        return rows

    def _bb_print_ledger_pdf(self):
        try:
            from optitex_analyzer.core.baby_basic_note_pdf import generate_baby_basic_ledger_pdf
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן ליצור PDF:\n{e}")
            return
        source = self._bb_current_account_source()
        acc = self.data_processor.get_baby_basic_account(source=source) or {}
        rows = self._bb_build_ledger_rows(acc)
        if not rows:
            messagebox.showwarning("הדפסה", "אין תנועות להדפסה בפילוח הנוכחי.")
            return
        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_account')
        os.makedirs(export_dir, exist_ok=True)
        suffix = f"_{source}" if source else ''
        path = os.path.join(export_dir, f"baby_basic_ledger{suffix}_{datetime.now().strftime('%Y%m%d')}.pdf")
        try:
            generate_baby_basic_ledger_pdf(
                rows,
                path,
                biz_name=self._bb_biz_name() or 'בייבי בייסיק',
                partner=self._bb_partner_name(),
                source_label=acc.get('source_label') or 'הכל',
                supplied=acc.get('supplied'),
                paid=acc.get('paid'),
                balance=acc.get('balance'),
                withdrawn=acc.get('withdrawn'),
                headers={'credit': 'שולם / זיכוי'},
            )
        except Exception as e:
            messagebox.showerror("שגיאה", f"יצירת PDF נכשלה:\n{e}")
            return
        try:
            os.startfile(path)
        except Exception:
            messagebox.showinfo("PDF", f"הקובץ נשמר:\n{path}")

    def _bb_add_payment(self):
        date_str = (self.bb_pay_date_var.get() if hasattr(self, 'bb_pay_date_var') else '').strip()
        amount = parse_numeric_value(self.bb_pay_amount_var.get() if hasattr(self, 'bb_pay_amount_var') else 0)
        note = (self.bb_pay_note_var.get() if hasattr(self, 'bb_pay_note_var') else '').strip()
        source = self._bb_source_key_from_label(self.bb_pay_source_var.get() if hasattr(self, 'bb_pay_source_var') else '') or 'local'
        try:
            self.data_processor.add_baby_basic_payment(date_str=date_str, amount=amount, note=note, source=source)
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

    def _bb_biz_name(self) -> str:
        try:
            return (self.settings.get('business.name', '') or '') if getattr(self, 'settings', None) else ''
        except Exception:
            return ''

    def _bb_print_selected_payment_pdf(self, event=None):
        tree = getattr(self, 'bb_acc_pay_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("הדפסה", "בחר תשלום מהרשימה.")
            return
        rec = self.data_processor.get_baby_basic_payment(sel[0])
        if rec:
            self._bb_export_payment_pdf(rec)

    def _bb_export_payment_pdf(self, rec: dict, path: str | None = None):
        try:
            from optitex_analyzer.core.baby_basic_note_pdf import generate_baby_basic_payment_pdf
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן ליצור PDF:\n{e}")
            return
        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_payments')
        os.makedirs(export_dir, exist_ok=True)
        safe_id = rec.get('id')
        safe_date = str(rec.get('date') or '').replace(':', '-')
        default_name = f"baby_basic_payment_{safe_id}_{safe_date}.pdf"
        if not path:
            path = os.path.join(export_dir, default_name)
        acc = {}
        try:
            # הסיכום ב-PDF לפי מקור התשלום עצמו (טורקי / מקומי)
            acc = self.data_processor.get_baby_basic_account(source=rec.get('source') or '') or {}
        except Exception:
            acc = {}
        try:
            generate_baby_basic_payment_pdf(
                rec,
                path,
                biz_name=self._bb_biz_name() or 'בייבי בייסיק',
                partner=self._bb_partner_name(),
                supplied=acc.get('supplied'),
                paid=acc.get('paid'),
                balance=acc.get('balance'),
                withdrawn=acc.get('withdrawn'),
            )
        except Exception as e:
            messagebox.showerror("שגיאה", f"יצירת PDF נכשלה:\n{e}")
            return
        try:
            os.startfile(path)
        except Exception:
            messagebox.showinfo("PDF", f"הקובץ נשמר:\n{path}")

    def _bb_print_payments_report_pdf(self):
        try:
            from optitex_analyzer.core.baby_basic_note_pdf import generate_baby_basic_payments_report_pdf
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן ליצור PDF:\n{e}")
            return
        # הדו"ח לפי הפילוח שנבחר במסך (הכל / טורקי / מקומי)
        source = self._bb_current_account_source()
        acc = self.data_processor.get_baby_basic_account(source=source) or {}
        payments = acc.get('payments') or ([] if source else self.data_processor.get_baby_basic_payments())
        if not payments:
            messagebox.showwarning("הדפסה", "אין תשלומים להדפסה בפילוח הנוכחי.")
            return
        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_payments')
        os.makedirs(export_dir, exist_ok=True)
        suffix = f"_{source}" if source else ''
        path = os.path.join(export_dir, f"baby_basic_payments{suffix}_{datetime.now().strftime('%Y%m%d')}.pdf")
        try:
            generate_baby_basic_payments_report_pdf(
                payments,
                path,
                biz_name=self._bb_biz_name() or 'בייבי בייסיק',
                partner=self._bb_partner_name(),
                supplied=acc.get('supplied'),
                paid=acc.get('paid'),
                balance=acc.get('balance'),
            )
        except Exception as e:
            messagebox.showerror("שגיאה", f"יצירת PDF נכשלה:\n{e}")
            return
        try:
            os.startfile(path)
        except Exception:
            messagebox.showinfo("PDF", f"הקובץ נשמר:\n{path}")

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

    # ---------- ספירת מלאי חדר (בלי ברקודים) ----------
    def _bb_room_size_sort_key(self, size: str):
        order = ['0-3', '3-6', '6-12', '12-18', '18-24', '24-30']
        s = str(size or '').strip()
        try:
            return (0, order.index(s), '')
        except ValueError:
            pass
        try:
            return (1, float(s), '')
        except ValueError:
            return (2, 0.0, s)

    def _bb_room_merge_names(self, *groups):
        names = []
        seen = set()
        for group in groups:
            for name in group or []:
                name = str(name or '').strip()
                if name and name not in seen:
                    seen.add(name)
                    names.append(name)
        return names

    def _bb_room_extra_models(self):
        try:
            extra = self.data_processor.get_baby_basic_room_count_extra_models()
        except Exception:
            extra = []
        return self._bb_room_merge_names(self.BB_ROOM_COUNT_EXTRA_MODELS, extra)

    def _bb_room_extra_colors(self):
        try:
            extra = self.data_processor.get_baby_basic_room_count_extra_colors()
        except Exception:
            extra = []
        return self._bb_room_merge_names(self.BB_ROOM_COUNT_COLORS, extra)

    def _bb_room_extra_prints(self):
        try:
            extra = self.data_processor.get_baby_basic_room_count_extra_prints()
        except Exception:
            extra = []
        return self._bb_room_merge_names(self.BB_ROOM_COUNT_PRINTS, extra)

    def _bb_room_extra_fabrics(self):
        try:
            extra = self.data_processor.get_baby_basic_room_count_extra_fabrics()
        except Exception:
            extra = []
        catalog = []
        for p in self._bb_products():
            name = (p.get('fabric') or '').strip()
            if name:
                catalog.append(name)
        return self._bb_room_merge_names(self.BB_ROOM_COUNT_FABRICS, catalog, extra)

    def _bb_room_catalog_for_model(self, model: str = ''):
        model = (model or '').strip()
        sizes = set()
        catalog_models = set()
        extras = self._bb_room_extra_models()
        extra_set = set(extras)
        for p in self._bb_products():
            name = (p.get('print_name') or p.get('item_name') or '').strip()
            if name:
                catalog_models.add(name)
            if model and name != model:
                continue
            size = (p.get('size') or '').strip()
            if size:
                sizes.add(size)
        rest = sorted(m for m in catalog_models if m not in extra_set)
        return {
            'models': extras + rest,
            'colors': self._bb_room_extra_colors(),
            'prints': self._bb_room_extra_prints(),
            'fabrics': self._bb_room_extra_fabrics(),
            'sizes': sorted(sizes, key=self._bb_room_size_sort_key),
        }

    def _bb_refresh_room_combos(self, event=None):
        model = (self.bb_room_model_var.get() if hasattr(self, 'bb_room_model_var') else '').strip()
        opts = self._bb_room_catalog_for_model(model)
        all_opts = self._bb_room_catalog_for_model('')
        if getattr(self, 'bb_room_model_combo', None) is not None:
            try:
                self.bb_room_model_combo['values'] = all_opts['models']
            except Exception:
                pass
        if getattr(self, 'bb_room_color_combo', None) is not None:
            try:
                self.bb_room_color_combo['values'] = all_opts['colors']
            except Exception:
                pass
        if getattr(self, 'bb_room_print_combo', None) is not None:
            try:
                self.bb_room_print_combo['values'] = all_opts['prints']
            except Exception:
                pass
        if getattr(self, 'bb_room_fabric_combo', None) is not None:
            try:
                self.bb_room_fabric_combo['values'] = all_opts['fabrics']
            except Exception:
                pass
        if getattr(self, 'bb_room_size_combo', None) is not None:
            try:
                sizes = opts['sizes'] or all_opts['sizes']
                self.bb_room_size_combo['values'] = sizes
            except Exception:
                pass

    def _bb_on_room_model_change(self, event=None):
        self._bb_refresh_room_combos()

    def _bb_refresh_room_lines(self):
        tree = getattr(self, 'bb_room_lines_tree', None)
        if tree is None:
            return
        self._bb_clear_tree(tree)
        lines = getattr(self, '_bb_room_lines', [])
        total = 0
        for i, line in enumerate(lines):
            qty = int(line.get('quantity') or 0)
            total += qty
            tree.insert('', 'end', iid=str(i), values=(
                line.get('print_name') or '',
                line.get('fabric') or '',
                line.get('color') or '',
                line.get('print') or '',
                line.get('size') or '',
                qty,
            ))
        theme.stripe_tree(tree)
        if hasattr(self, 'bb_room_total_var'):
            self.bb_room_total_var.set(f"סה״כ יחידות: {total}")

    def _bb_add_room_line(self):
        model = (self.bb_room_model_var.get() if hasattr(self, 'bb_room_model_var') else '').strip()
        fabric = (self.bb_room_fabric_var.get() if hasattr(self, 'bb_room_fabric_var') else '').strip()
        color = (self.bb_room_color_var.get() if hasattr(self, 'bb_room_color_var') else '').strip()
        print_val = (self.bb_room_print_var.get() if hasattr(self, 'bb_room_print_var') else '').strip()
        size = (self.bb_room_size_var.get() if hasattr(self, 'bb_room_size_var') else '').strip()
        raw = (self.bb_room_qty_var.get() if hasattr(self, 'bb_room_qty_var') else '').strip()
        if not model:
            messagebox.showwarning("ספירת מלאי", "בחר או הקלד דגם.")
            return
        try:
            qty = int(float(raw)) if raw else 0
        except Exception:
            qty = 0
        if qty <= 0:
            messagebox.showwarning("ספירת מלאי", "הזן כמות גדולה מאפס.")
            return
        if not hasattr(self, '_bb_room_lines'):
            self._bb_room_lines = []
        key = (model, fabric, color, print_val, size)
        merged = False
        for line in self._bb_room_lines:
            existing = (
                str(line.get('print_name') or '').strip(),
                str(line.get('fabric') or '').strip(),
                str(line.get('color') or '').strip(),
                str(line.get('print') or '').strip(),
                str(line.get('size') or '').strip(),
            )
            if existing == key:
                line['quantity'] = int(line.get('quantity') or 0) + qty
                merged = True
                break
        if not merged:
            self._bb_room_lines.append({
                'print_name': model,
                'fabric': fabric,
                'color': color,
                'print': print_val,
                'size': size,
                'quantity': qty,
            })
        self._bb_persist_room_line_extras([{
            'print_name': model,
            'fabric': fabric,
            'color': color,
            'print': print_val,
        }])
        if hasattr(self, 'bb_room_qty_var'):
            self.bb_room_qty_var.set('')
        self._bb_refresh_room_combos()
        self._bb_refresh_room_lines()

    def _bb_edit_room_qty_cell(self, event):
        def on_commit(iid, raw):
            try:
                idx = int(iid)
            except Exception:
                return
            lines = getattr(self, '_bb_room_lines', [])
            if idx < 0 or idx >= len(lines):
                return
            try:
                qty = int(float(raw)) if raw else 0
            except Exception:
                qty = 0
            if qty <= 0:
                lines.pop(idx)
            else:
                lines[idx]['quantity'] = qty
            self._bb_refresh_room_lines()
        self._bb_inline_edit(self.bb_room_lines_tree, event, '#6', on_commit)

    def _bb_remove_selected_room_line(self):
        tree = getattr(self, 'bb_room_lines_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        try:
            idx = int(sel[0])
            if 0 <= idx < len(self._bb_room_lines):
                self._bb_room_lines.pop(idx)
        except Exception:
            return
        self._bb_refresh_room_lines()

    def _bb_clear_current_room_count(self):
        self._bb_room_lines = []
        if hasattr(self, 'bb_room_note_var'):
            self.bb_room_note_var.set('')
        if hasattr(self, 'bb_room_date_var'):
            self.bb_room_date_var.set(datetime.now().strftime('%Y-%m-%d'))
        if hasattr(self, 'bb_room_qty_var'):
            self.bb_room_qty_var.set('')
        self._bb_refresh_room_lines()

    def _bb_save_room_count(self):
        date_str = (self.bb_room_date_var.get() if hasattr(self, 'bb_room_date_var') else '').strip()
        note = (self.bb_room_note_var.get() if hasattr(self, 'bb_room_note_var') else '').strip()
        lines = list(getattr(self, '_bb_room_lines', []) or [])
        try:
            new_id = self.data_processor.add_baby_basic_room_count(date_str, lines, note=note)
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
            return
        self._bb_clear_current_room_count()
        self._bb_refresh_saved_room_counts()
        messagebox.showinfo("נשמר", f"ספירת מלאי חדר #{new_id} נשמרה.")
        rec = self.data_processor.get_baby_basic_room_count(new_id)
        if rec:
            self._bb_open_room_count_view(rec)

    def _bb_persist_room_line_extras(self, lines):
        catalog_names = set()
        catalog_fabrics = set(self.BB_ROOM_COUNT_FABRICS)
        for p in self._bb_products():
            name = (p.get('print_name') or p.get('item_name') or '').strip()
            if name:
                catalog_names.add(name)
            fname = (p.get('fabric') or '').strip()
            if fname:
                catalog_fabrics.add(fname)
        for line in lines or []:
            model = str(line.get('print_name') or '').strip()
            fabric = str(line.get('fabric') or '').strip()
            color = str(line.get('color') or '').strip()
            print_val = str(line.get('print') or '').strip()
            if model and model not in catalog_names and model not in self.BB_ROOM_COUNT_EXTRA_MODELS:
                try:
                    self.data_processor.add_baby_basic_room_count_extra_model(model)
                    catalog_names.add(model)
                except Exception:
                    pass
            if fabric and fabric not in catalog_fabrics:
                try:
                    self.data_processor.add_baby_basic_room_count_extra_fabric(fabric)
                    catalog_fabrics.add(fabric)
                except Exception:
                    pass
            if color and color not in self.BB_ROOM_COUNT_COLORS:
                try:
                    self.data_processor.add_baby_basic_room_count_extra_color(color)
                except Exception:
                    pass
            if print_val and print_val not in self.BB_ROOM_COUNT_PRINTS:
                try:
                    self.data_processor.add_baby_basic_room_count_extra_print(print_val)
                except Exception:
                    pass

    def _bb_import_room_count_excel(self):
        initial = os.path.join(os.getcwd(), 'exports', 'baby_basic_room_counts')
        if not os.path.isdir(initial):
            initial = os.getcwd()
        path = filedialog.askopenfilename(
            title='ייבוא ספירת מלאי מאקסל',
            initialdir=initial,
            filetypes=[('Excel', '*.xlsx'), ('All files', '*.*')],
        )
        if not path:
            return
        try:
            parsed = self.data_processor.parse_baby_basic_room_count_excel(path)
        except Exception as e:
            messagebox.showerror("שגיאה", f"קריאת הקובץ נכשלה:\n{e}")
            return
        lines = list(parsed.get('lines') or [])
        self._bb_persist_room_line_extras(lines)
        try:
            new_id = self.data_processor.add_baby_basic_room_count(
                parsed.get('date') or '',
                lines,
                note=parsed.get('note') or '',
            )
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
            return
        self._bb_refresh_room_combos()
        self._bb_refresh_saved_room_counts()
        messagebox.showinfo("יובא", f"ספירת מלאי חדר #{new_id} נשמרה מתוך הקובץ.\n{len(lines)} שורות.")
        rec = self.data_processor.get_baby_basic_room_count(new_id)
        if rec:
            self._bb_open_room_count_view(rec)

    def _bb_refresh_saved_room_counts(self):
        tree = getattr(self, 'bb_saved_room_tree', None)
        if tree is None:
            return
        self._bb_clear_tree(tree)
        counts = sorted(
            self.data_processor.get_baby_basic_room_counts(),
            key=lambda n: int(n.get('id') or 0),
            reverse=True,
        )
        for rec in counts:
            tree.insert('', 'end', iid=str(rec.get('id')), values=(
                rec.get('id', ''),
                rec.get('date', ''),
                rec.get('total_quantity', 0),
                rec.get('note', ''),
            ))
        theme.stripe_tree(tree)

    def _bb_open_selected_room_count(self, event=None):
        tree = getattr(self, 'bb_saved_room_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        rec = self.data_processor.get_baby_basic_room_count(sel[0])
        if rec:
            self._bb_open_room_count_view(rec)

    def _bb_delete_selected_room_count(self):
        tree = getattr(self, 'bb_saved_room_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        rec = self.data_processor.get_baby_basic_room_count(sel[0])
        if not rec:
            return
        if not messagebox.askyesno("מחיקה", f"למחוק ספירת מלאי #{rec.get('id')} מתאריך {rec.get('date')}?"):
            return
        self.data_processor.delete_baby_basic_room_count(rec.get('id'))
        self._bb_refresh_saved_room_counts()

    def _bb_previous_room_count(self, rec: dict):
        counts = sorted(
            self.data_processor.get_baby_basic_room_counts(),
            key=lambda n: (str(n.get('date') or ''), int(n.get('id') or 0)),
        )
        prev = None
        rec_id = int(rec.get('id') or 0)
        for item in counts:
            if int(item.get('id') or 0) == rec_id:
                break
            prev = item
        return prev

    def _bb_open_room_count_view(self, rec: dict):
        import tkinter as tk
        from tkinter import ttk
        win = tk.Toplevel(self.root)
        win.title(f"ספירת מלאי חדר #{rec.get('id')}")
        win.configure(bg=theme.PAGE_BG)
        win.geometry("1140x520")
        tk.Label(
            win,
            text=f"ספירת מלאי חדר #{rec.get('id')}  |  {rec.get('date')}",
            font=(theme.FONT_FAMILY, 13, 'bold'),
            bg=theme.PAGE_BG,
            fg=theme.DARK,
        ).pack(pady=(10, 4))
        if rec.get('note'):
            tk.Label(win, text=rec.get('note'), bg=theme.PAGE_BG, fg=theme.SUBTEXT).pack()

        prev = self._bb_previous_room_count(rec)
        prev_map = {}
        if prev:
            for line in prev.get('lines') or []:
                key = (
                    str(line.get('print_name') or '').strip(),
                    str(line.get('fabric') or '').strip(),
                    str(line.get('color') or '').strip(),
                    str(line.get('print') or '').strip(),
                    str(line.get('size') or '').strip(),
                    str(line.get('area') or '').strip(),
                    str(line.get('box') or '').strip(),
                )
                prev_map[key] = int(line.get('quantity') or 0)
            tk.Label(
                win,
                text=f"השוואה לספירה #{prev.get('id')} מתאריך {prev.get('date')}",
                bg=theme.PAGE_BG,
                fg=theme.SUBTEXT,
                font=theme.FONT_SMALL,
            ).pack()

        cols = ('print_name', 'fabric', 'color', 'print', 'size', 'area', 'box', 'qty', 'delta')
        headers = {
            'print_name': 'דגם',
            'fabric': 'סוג בד',
            'color': 'צבע רקע',
            'print': 'הדפס',
            'size': 'מידה',
            'area': 'אזור',
            'box': 'קופסא',
            'qty': 'כמות',
            'delta': 'שינוי',
        }
        tree = theme.make_treeview(win, columns=cols, show='headings', height=14)
        for c in cols:
            tree.heading(c, text=headers[c])
            w = 80 if c in ('size', 'qty', 'delta', 'area', 'box') else (110 if c in ('fabric', 'color', 'print') else 160)
            tree.column(c, width=w, anchor='center')
        vs = ttk.Scrollbar(win, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.pack(side='left', fill='both', expand=True, padx=(12, 0), pady=8)
        vs.pack(side='left', fill='y', pady=8)
        for line in rec.get('lines') or []:
            qty = int(line.get('quantity') or 0)
            key = (
                str(line.get('print_name') or '').strip(),
                str(line.get('fabric') or '').strip(),
                str(line.get('color') or '').strip(),
                str(line.get('print') or '').strip(),
                str(line.get('size') or '').strip(),
                str(line.get('area') or '').strip(),
                str(line.get('box') or '').strip(),
            )
            delta = ''
            if prev_map:
                old = prev_map.get(key, 0)
                diff = qty - old
                if diff > 0:
                    delta = f"+{diff}"
                elif diff < 0:
                    delta = str(diff)
                else:
                    delta = '0'
            tree.insert('', 'end', values=(
                line.get('print_name') or '',
                line.get('fabric') or '',
                line.get('color') or '',
                line.get('print') or '',
                line.get('size') or '',
                line.get('area') or '',
                line.get('box') or '',
                qty,
                delta,
            ))
        theme.stripe_tree(tree)

        footer = tk.Frame(win, bg=theme.PAGE_BG)
        footer.pack(fill='x', padx=12, pady=(0, 10))
        tk.Label(
            footer,
            text=f"סה״כ יחידות: {rec.get('total_quantity', 0)}",
            font=(theme.FONT_FAMILY, 11, 'bold'),
            bg=theme.PAGE_BG,
        ).pack(side='right')
        theme.make_button(footer, "PDF להדפסה (A4)", kind="primary", command=lambda: self._bb_export_room_count_pdf(rec)).pack(side='left')

    def _bb_print_selected_room_count_pdf(self):
        tree = getattr(self, 'bb_saved_room_tree', None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("הדפסה", "בחר ספירה מהרשימה.")
            return
        rec = self.data_processor.get_baby_basic_room_count(sel[0])
        if rec:
            self._bb_export_room_count_pdf(rec)

    def _bb_export_room_count_pdf(self, rec: dict, path: str | None = None):
        try:
            from optitex_analyzer.core.baby_basic_note_pdf import generate_baby_basic_room_count_pdf
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן ליצור PDF:\n{e}")
            return
        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_room_counts')
        os.makedirs(export_dir, exist_ok=True)
        safe_id = rec.get('id')
        safe_date = str(rec.get('date') or '').replace(':', '-')
        default_name = f"baby_basic_room_count_{safe_id}_{safe_date}.pdf"
        if not path:
            path = os.path.join(export_dir, default_name)
        try:
            generate_baby_basic_room_count_pdf(rec, path, biz_name=self._bb_biz_name() or 'בייבי בייסיק')
        except Exception as e:
            messagebox.showerror("שגיאה", f"יצירת PDF נכשלה:\n{e}")
            return
        try:
            os.startfile(path)
        except Exception:
            messagebox.showinfo("PDF", f"הקובץ נשמר:\n{path}")

    def _bb_print_count_sheet_pdf(self):
        try:
            from optitex_analyzer.core.baby_basic_note_pdf import generate_baby_basic_count_sheet_pdf
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן ליצור PDF:\n{e}")
            return
        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_room_counts')
        os.makedirs(export_dir, exist_ok=True)
        path = os.path.join(
            export_dir,
            f"baby_basic_count_sheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
        )
        try:
            generate_baby_basic_count_sheet_pdf(
                file_path=path,
                biz_name=self._bb_biz_name() or 'בייבי בייסיק',
                location='חדר בייבי בייסיק',
            )
        except Exception as e:
            messagebox.showerror("שגיאה", f"יצירת PDF נכשלה:\n{e}")
            return
        try:
            os.startfile(path)
        except Exception:
            messagebox.showinfo("PDF", f"הקובץ נשמר:\n{path}")

    def _bb_export_room_counts_summary_excel(self, path: str | None = None):
        counts = list(self.data_processor.get_baby_basic_room_counts() or [])
        if not counts:
            messagebox.showinfo("ייצוא", "אין ספירות שמורות לייצוא.")
            return
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Border, Font, Side, PatternFill
            from openpyxl.utils import get_column_letter
        except Exception as e:
            messagebox.showerror("שגיאה", f"נדרש openpyxl לייצוא לאקסל:\n{e}")
            return

        def is_colorful(color, fabric, print_val) -> bool:
            if print_val:
                return True
            return bool(color) and color != 'לבן' and fabric in ('פלנל', 'טריקו')

        summary = {}
        detail = {}
        for rec in counts:
            for line in rec.get('lines') or []:
                qty = int(line.get('quantity') or 0)
                if qty <= 0:
                    continue
                model = str(line.get('print_name') or '').strip()
                if not model:
                    continue
                size = str(line.get('size') or '').strip()
                fabric = str(line.get('fabric') or '').strip()
                color = str(line.get('color') or '').strip()
                print_val = str(line.get('print') or '').strip()
                area = str(line.get('area') or '').strip()
                box = str(line.get('box') or '').strip()
                colorful = 'כן' if is_colorful(color, fabric, print_val) else 'לא'
                skey = (model, size, fabric, colorful)
                summary[skey] = summary.get(skey, 0) + qty
                dkey = (model, fabric, color, print_val, size, area, box)
                detail[dkey] = detail.get(dkey, 0) + qty

        if not summary:
            messagebox.showinfo("ייצוא", "אין שורות עם כמות בספירות השמורות.")
            return

        thin = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1'),
        )
        header_fill = PatternFill('solid', fgColor='1E293B')
        header_font = Font(name='Arial', bold=True, color='FFFFFF', size=11)
        total_fill = PatternFill('solid', fgColor='FFF2CC')
        center = Alignment(horizontal='center', vertical='center')
        wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)

        def style_header(ws, ncols):
            for col in range(1, ncols + 1):
                cell = ws.cell(1, col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = wrap
                cell.border = thin

        def style_data(ws, row, ncols):
            for col in range(1, ncols + 1):
                cell = ws.cell(row, col)
                cell.alignment = center
                cell.border = thin

        def write_total(ws, row, label_col, qty_col, last_data_row, ncols):
            ws.cell(row, label_col, 'סה״כ יחידות').font = Font(name='Arial', bold=True)
            cell = ws.cell(row, qty_col, f'=SUM({get_column_letter(qty_col)}2:{get_column_letter(qty_col)}{last_data_row})')
            cell.font = Font(name='Arial', bold=True)
            cell.number_format = '#,##0'
            for col in range(1, ncols + 1):
                ws.cell(row, col).fill = total_fill
                ws.cell(row, col).border = thin
                ws.cell(row, col).alignment = center

        wb = Workbook()
        ws_sum = wb.active
        ws_sum.title = 'סיכום דגם-מידה'
        try:
            ws_sum.sheet_view.rightToLeft = True
        except Exception:
            pass
        sum_headers = ['דגם', 'מידה', 'סוג בד', 'צבעוני', 'סה״כ כמות']
        ws_sum.append(sum_headers)
        style_header(ws_sum, len(sum_headers))
        sum_rows = sorted(
            summary.items(),
            key=lambda item: (
                item[0][0],
                self._bb_room_size_sort_key(item[0][1]),
                item[0][1],
                item[0][2],
                0 if item[0][3] == 'לא' else 1,
            ),
        )
        for (model, size, fabric, colorful), qty in sum_rows:
            ws_sum.append([model, size, fabric, colorful, qty])
            style_data(ws_sum, ws_sum.max_row, 5)
            ws_sum.cell(ws_sum.max_row, 5).number_format = '#,##0'
        last_sum = ws_sum.max_row
        write_total(ws_sum, last_sum + 1, 1, 5, last_sum, 5)
        for i, w in enumerate([32, 14, 12, 12, 14], 1):
            ws_sum.column_dimensions[get_column_letter(i)].width = w
        ws_sum.freeze_panes = 'A2'
        ws_sum.auto_filter.ref = f'A1:E{last_sum}'

        ws_det = wb.create_sheet('פירוט')
        try:
            ws_det.sheet_view.rightToLeft = True
        except Exception:
            pass
        det_headers = ['דגם', 'סוג בד', 'צבע רקע', 'הדפס', 'צבעוני', 'מידה', 'אזור', 'מספר קופסא', 'כמות']
        ws_det.append(det_headers)
        style_header(ws_det, len(det_headers))
        det_rows = sorted(
            detail.items(),
            key=lambda item: (
                item[0][0],
                self._bb_room_size_sort_key(item[0][4]),
                item[0][4],
                item[0][2],
                item[0][3],
                item[0][5],
                item[0][1],
                item[0][6],
            ),
        )
        for (model, fabric, color, print_val, size, area, box), qty in det_rows:
            colorful = 'כן' if is_colorful(color, fabric, print_val) else 'לא'
            ws_det.append([model, fabric, color, print_val, colorful, size, area, box, qty])
            style_data(ws_det, ws_det.max_row, 9)
            ws_det.cell(ws_det.max_row, 9).number_format = '#,##0'
        last_det = ws_det.max_row
        write_total(ws_det, last_det + 1, 1, 9, last_det, 9)
        for i, w in enumerate([28, 12, 16, 22, 12, 12, 22, 14, 12], 1):
            ws_det.column_dimensions[get_column_letter(i)].width = w
        ws_det.freeze_panes = 'A2'
        ws_det.auto_filter.ref = f'A1:I{last_det}'

        ws_src = wb.create_sheet('מקור')
        try:
            ws_src.sheet_view.rightToLeft = True
        except Exception:
            pass
        src_headers = ['מס׳ ספירה', 'תאריך', 'סה״כ יחידות', 'הערה']
        ws_src.append(src_headers)
        style_header(ws_src, len(src_headers))
        for rec in sorted(counts, key=lambda r: int(r.get('id') or 0)):
            ws_src.append([
                rec.get('id', ''),
                rec.get('date', ''),
                int(rec.get('total_quantity') or 0),
                rec.get('note', ''),
            ])
            style_data(ws_src, ws_src.max_row, 4)
            ws_src.cell(ws_src.max_row, 3).number_format = '#,##0'
        last_src = ws_src.max_row
        write_total(ws_src, last_src + 1, 1, 3, last_src, 4)
        for i, w in enumerate([12, 14, 14, 70], 1):
            ws_src.column_dimensions[get_column_letter(i)].width = w
        ws_src.freeze_panes = 'A2'
        ws_src.auto_filter.ref = f'A1:D{last_src}'

        export_dir = os.path.join(os.getcwd(), 'exports', 'baby_basic_room_counts')
        os.makedirs(export_dir, exist_ok=True)
        default_name = f"baby_basic_room_counts_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        if not path:
            path = os.path.join(export_dir, default_name)
        try:
            wb.save(path)
        except Exception as e:
            messagebox.showerror("שגיאה", f"שמירת הקובץ נכשלה:\n{e}")
            return
        try:
            os.startfile(path)
        except Exception:
            messagebox.showinfo("Excel", f"הקובץ נשמר:\n{path}")
