"""Reusable label generator UI: Rivhit products -> queue -> PDF/print."""
import os
import re
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from . import theme


class LabelGeneratorFrame(ttk.Frame):
    """סאב-טאב ליצירת מדבקות מתבנית, ממוצרי ריווחית."""

    def __init__(
        self,
        parent,
        data_processor,
        *,
        logo_path=None,
        default_category='הכל',
        title='יצירת דף מדבקות מתבנית',
        sumatra_print_fn=None,
        draw_border=True,
        default_pack_size=3,
    ):
        super().__init__(parent)
        self.data_processor = data_processor
        self.logo_path = logo_path
        self.default_category = default_category
        self._sumatra_print_fn = sumatra_print_fn
        self.draw_border = draw_border
        self.default_pack_size = 5 if int(default_pack_size or 3) == 5 else 3

        self._queue_data = []
        self._all_products = []
        self._qty_by_barcode = {}
        self._last_pdf = None

        self._build_ui(title)
        self.refresh_products()

    def _build_ui(self, title):
        # אזור גלילה — כדי להגיע לכפתורי PDF/הדפסה מתחת לתור
        main_frame = tk.Frame(self)
        main_frame.pack(fill="both", expand=True)
        canvas = tk.Canvas(main_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        body = tk.Frame(canvas)
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win_id = canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(win_id, width=event.width)
        canvas.bind("<Configure>", _configure_scroll_region)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        tk.Label(body, text=title, font=(theme.FONT_FAMILY, 14, 'bold')).pack(pady=(10, 5))

        sel = ttk.LabelFrame(body, text="בחירת מוצרים מריווחית", padding=10)
        sel.pack(fill="x", padx=15, pady=10)

        row1 = tk.Frame(sel)
        row1.pack(fill="x", pady=5)
        ttk.Label(row1, text="קטגוריה:").pack(side="right", padx=5)
        self._category_var = tk.StringVar(value=self.default_category)
        self._category_cb = ttk.Combobox(row1, textvariable=self._category_var, width=22, state="readonly")
        self._category_cb.pack(side="right", padx=5)
        self._category_cb.bind('<<ComboboxSelected>>', lambda e: self._refresh_products_table())

        ttk.Label(row1, text="חיפוש:").pack(side="right", padx=(20, 5))
        self._search_var = tk.StringVar()
        search_entry = ttk.Entry(row1, textvariable=self._search_var, width=28)
        search_entry.pack(side="right", padx=5)
        search_entry.bind('<KeyRelease>', lambda e: self._refresh_products_table())

        tk.Button(row1, text="🧹 נקה סינון", bg=theme.MUTED, fg="white",
                  command=self._clear_filters).pack(side="right", padx=10)
        tk.Button(row1, text="➕ הוסף הכל לתור", bg=theme.SUCCESS, fg="white",
                  font=(theme.FONT_FAMILY, 10, 'bold'), command=self._add_all_to_queue).pack(side="left", padx=10)

        row_pack = tk.Frame(sel)
        row_pack.pack(fill="x", pady=(0, 5))
        ttk.Label(row_pack, text="גודל מארז:").pack(side="right", padx=5)
        self._pack_size_var = tk.IntVar(value=self.default_pack_size)
        pack_frame = tk.Frame(row_pack)
        pack_frame.pack(side="right", padx=5)
        for val in (3, 5):
            tk.Radiobutton(
                pack_frame, text=f"{val} יחידות", variable=self._pack_size_var, value=val,
                command=self._on_pack_size_changed,
            ).pack(side="right", padx=6)

        tk.Label(
            sel,
            text="טיפ: לחץ פעמיים על תא 'כמות' כדי להזין כמות לכל מוצר, ואז 'הוסף הכל לתור'. "
                 "גודל המארז (3/5) חל על כל המדבקות — ניתן לשנות לפני יצירת PDF או הדפסה. "
                 "לעריכת שדות המדבקה - בטאב ריווחית, דאבל-קליק על מוצר.",
            fg=theme.MUTED, font=(theme.FONT_FAMILY, 8), justify="right", anchor="e",
        ).pack(fill="x", pady=(2, 6))

        prod_frame = tk.Frame(sel)
        prod_frame.pack(fill="x")
        pcols = ("name", "size", "fabric", "barcode", "qty")
        pheaders = {"name": "שם המוצר", "size": "מידה", "fabric": "סוג בד",
                    "barcode": "ברקוד", "qty": "כמות"}
        self._products_tree = ttk.Treeview(prod_frame, columns=pcols, show="headings", height=8)
        for c in pcols:
            self._products_tree.heading(c, text=pheaders[c])
            w = 70 if c in ("size", "qty") else (150 if c == "fabric" else 240 if c == "name" else 140)
            self._products_tree.column(c, width=w, anchor="center")
        pvs = ttk.Scrollbar(prod_frame, orient="vertical", command=self._products_tree.yview)
        self._products_tree.configure(yscrollcommand=pvs.set)
        self._products_tree.grid(row=0, column=0, sticky="nsew")
        pvs.grid(row=0, column=1, sticky="ns")
        prod_frame.grid_columnconfigure(0, weight=1)
        self._products_tree.bind('<Double-1>', self._edit_qty_cell)

        queue_frame = ttk.LabelFrame(body, text="תור מדבקות", padding=10)
        queue_frame.pack(fill="x", padx=15, pady=10)

        cols = ("print_name", "size", "fabric", "pack", "barcode", "qty")
        headers = {"print_name": "שם להדפסה", "size": "מידה", "fabric": "סוג בד",
                   "pack": "מארז", "barcode": "ברקוד", "qty": "כמות"}
        self._queue_tree = ttk.Treeview(queue_frame, columns=cols, show="headings", height=8)
        for c in cols:
            self._queue_tree.heading(c, text=headers[c])
            w = 70 if c in ("size", "pack", "qty") else 160
            self._queue_tree.column(c, width=w, anchor="center")
        vs = ttk.Scrollbar(queue_frame, orient="vertical", command=self._queue_tree.yview)
        self._queue_tree.configure(yscrollcommand=vs.set)
        self._queue_tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        queue_frame.grid_columnconfigure(0, weight=1)

        self._summary_var = tk.StringVar(value='סה"כ: 0 מדבקות | 0 דפים (15 לדף)')
        tk.Label(queue_frame, textvariable=self._summary_var, anchor="e", justify="right",
                 font=(theme.FONT_FAMILY, 11, 'bold'), fg=theme.DARK).grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        actions = tk.Frame(body)
        actions.pack(fill="x", padx=15, pady=(0, 16))
        tk.Button(actions, text="🗑️ מחק נבחר", bg=theme.WARNING, fg="white",
                  command=self._delete_selected).pack(side="left", padx=5)
        tk.Button(actions, text="🧹 נקה הכל", bg=theme.MUTED, fg="white",
                  command=self._clear_queue).pack(side="left", padx=5)
        tk.Button(actions, text="🖨️ הדפס", bg=theme.PRIMARY_DARK, fg="white",
                  font=(theme.FONT_FAMILY, 11, 'bold'), command=self._print).pack(side="right", padx=5)
        tk.Button(actions, text="🧾 צור דף מדבקות (PDF)", bg=theme.SUCCESS, fg="white",
                  font=(theme.FONT_FAMILY, 11, 'bold'), command=self._generate_pdf).pack(side="right", padx=5)
        tk.Button(actions, text="📁 ייצא קבצים בודדים (PDF)", bg=theme.PURPLE, fg="white",
                  font=(theme.FONT_FAMILY, 11, 'bold'), command=self._export_single_pdfs).pack(side="right", padx=5)

    def _rivhit_products(self):
        dp = self.data_processor
        return list(getattr(dp, 'rivhit_products', []) or []) if dp else []

    def refresh_products(self):
        """רענון רשימת המוצרים, בורר הקטגוריות וטבלת המוצרים מתוך מוצרי ריווחית."""
        items = self._rivhit_products()
        pairs = []
        for p in items:
            name = str(p.get('item_name', '')).strip()
            bc = str(p.get('item_part_num', '')).strip()
            if not name and not bc:
                continue
            label = f"{name} | {bc}" if bc else name
            pairs.append((label, p))
        self._all_products = pairs
        self._update_categories()
        self._refresh_products_table()

    def _update_categories(self):
        dp = self.data_processor
        cats = {str(p.get('compute_0036', '')).strip()
                for _, p in self._all_products
                if str(p.get('compute_0036', '')).strip()}
        groups = set((getattr(dp, 'rivhit_groups', {}) or {}).keys()) if dp else set()
        all_cats = sorted(cats | groups)
        values = ['הכל'] + all_cats
        self._category_cb['values'] = values
        if self._category_var.get() not in values:
            fallback = self.default_category if self.default_category in values else 'הכל'
            self._category_var.set(fallback)

    def _clear_filters(self):
        self._search_var.set('')
        self._category_var.set(self.default_category)
        self._refresh_products_table()

    def _current_pack_size(self) -> int:
        """גודל מארז גלובלי להדפסה — 3 או 5 בלבד."""
        try:
            val = int(self._pack_size_var.get())
        except Exception:
            val = self.default_pack_size
        return 5 if val == 5 else 3

    def _on_pack_size_changed(self):
        """רענון עמודת מארז בתור כשמחליפים בורר גלובלי."""
        if self._queue_data:
            self._refresh_queue_tree()

    def _filtered_products(self):
        category = self._category_var.get()
        q = (self._search_var.get() or '').strip().lower()
        terms = q.split()
        result = []
        for lbl, p in self._all_products:
            if category and category != 'הכל':
                if str(p.get('compute_0036', '')).strip() != category:
                    continue
            if terms and not all(t in lbl.lower() for t in terms):
                continue
            result.append((lbl, p))
        return result

    def _refresh_products_table(self):
        for item in self._products_tree.get_children():
            self._products_tree.delete(item)
        seen = set()
        for lbl, p in self._filtered_products():
            bc = str(p.get('item_part_num', '')).strip()
            if not bc or bc in seen:
                continue
            seen.add(bc)
            f = self.data_processor.get_rivhit_label_fields(bc, product=p)
            name = f.get('print_name', '') or str(p.get('item_name', '')).strip()
            qty = int(self._qty_by_barcode.get(bc, 0) or 0)
            self._products_tree.insert("", "end", iid=bc, values=(
                name, f.get('size', ''), f.get('fabric', ''), bc, qty,
            ))

    def _edit_qty_cell(self, event):
        tree = self._products_tree
        col = tree.identify_column(event.x)
        row = tree.identify_row(event.y)
        if not row or col != f"#{len(tree['columns'])}":
            return
        x, y, w, h = tree.bbox(row, col)
        cur = self._qty_by_barcode.get(row, 0)
        entry = tk.Entry(tree, justify='center')
        entry.insert(0, str(cur or ''))
        entry.select_range(0, 'end')
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()

        def commit(_e=None):
            val = entry.get().strip()
            try:
                qty = int(val) if val else 0
                if qty < 0:
                    qty = 0
            except ValueError:
                qty = 0
            if qty > 0:
                self._qty_by_barcode[row] = qty
            else:
                self._qty_by_barcode.pop(row, None)
            entry.destroy()
            if tree.exists(row):
                vals = list(tree.item(row, 'values'))
                vals[-1] = qty
                tree.item(row, values=vals)

        entry.bind('<Return>', commit)
        entry.bind('<KP_Enter>', commit)
        entry.bind('<FocusOut>', commit)
        entry.bind('<Escape>', lambda _e: entry.destroy())

    def _add_all_to_queue(self):
        to_add = {bc: q for bc, q in self._qty_by_barcode.items() if int(q or 0) > 0}
        if not to_add:
            messagebox.showwarning("אזהרה", "לא הוזנה כמות לאף מוצר. לחץ פעמיים על תא 'כמות' כדי להזין.")
            return
        prod_by_bc = {str(p.get('item_part_num', '')).strip(): p for _, p in self._all_products}
        existing = {it.get('barcode'): it for it in self._queue_data}
        added = 0
        for bc, qty in to_add.items():
            p = prod_by_bc.get(bc)
            if not p:
                continue
            qty = int(qty)
            if bc in existing:
                existing[bc]['qty'] = int(existing[bc].get('qty', 0) or 0) + qty
                added += 1
                continue
            f = self.data_processor.get_rivhit_label_fields(bc, product=p)
            self._queue_data.append({
                "print_name": f.get('print_name', ''),
                "size": f.get('size', ''),
                "size_unit": f.get('size_unit', ''),
                "fabric": f.get('fabric', ''),
                "image": f.get('image', ''),
                "model_code": f.get('model_code', ''),
                "brand": self.data_processor.brand_key_from_category(p.get('compute_0036', '')),
                "barcode": bc,
                "qty": qty,
            })
            added += 1
        self._qty_by_barcode.clear()
        self._refresh_products_table()
        self._refresh_queue_tree()
        messagebox.showinfo("נוסף לתור", f"נוספו/עודכנו {added} מוצרים בתור")

    def _refresh_queue_tree(self):
        pack = self._current_pack_size()
        for item in self._queue_tree.get_children():
            self._queue_tree.delete(item)
        for idx, it in enumerate(self._queue_data):
            self._queue_tree.insert("", "end", iid=str(idx), values=(
                it.get("print_name", ""), it.get("size", ""), it.get("fabric", ""),
                pack, it.get("barcode", ""), it.get("qty", 1),
            ))
        self._update_summary()

    def _update_summary(self):
        try:
            from optitex_analyzer.core.label_sheet import PER_PAGE
            per_page = PER_PAGE or 15
        except Exception:
            per_page = 15
        total = sum(int(it.get('qty', 0) or 0) for it in self._queue_data)
        pages = (total + per_page - 1) // per_page if total > 0 else 0
        self._summary_var.set(f'סה"כ: {total} מדבקות | {pages} דפים ({per_page} לדף)')

    def _delete_selected(self):
        sel = self._queue_tree.selection()
        if not sel:
            messagebox.showwarning("אזהרה", "אנא בחר שורה למחיקה")
            return
        idx = int(sel[0])
        del self._queue_data[idx]
        self._refresh_queue_tree()

    def _clear_queue(self):
        if not self._queue_data:
            return
        if messagebox.askyesno("אישור", "לנקות את כל התור?"):
            self._queue_data.clear()
            self._refresh_queue_tree()

    def _expand_items(self):
        pack_qty = self._current_pack_size()
        expanded = []
        for it in self._queue_data:
            for _ in range(int(it.get("qty", 1) or 1)):
                expanded.append({
                    "print_name": it.get("print_name", ""),
                    "size": it.get("size", ""),
                    "size_unit": it.get("size_unit", ""),
                    "fabric": it.get("fabric", ""),
                    "pack_qty": pack_qty,
                    "image": it.get("image", ""),
                    "barcode": it.get("barcode", ""),
                })
        return expanded

    @staticmethod
    def _safe_filename_part(text):
        """מנקה טקסט לשם קובץ באנגלית: רק A-Za-z0-9, מקף וקו תחתון."""
        text = str(text or '').strip()
        text = re.sub(r'\s+', '-', text)
        text = re.sub(r'[^A-Za-z0-9_\-]', '', text)
        text = re.sub(r'-{2,}', '-', text)
        return text.strip('-_')

    def _export_single_pdfs(self):
        """מייצא כל מוצר בתור לקובץ PDF נפרד בגודל מדבקה בודדת, בשם אנגלי."""
        if not self._queue_data:
            messagebox.showwarning("אזהרה", "התור ריק")
            return
        from optitex_analyzer.core.label_sheet import build_single_label_pdf
        out_dir = os.path.join(os.getcwd(), "exports", "labels", "singles")
        os.makedirs(out_dir, exist_ok=True)
        pack_qty = self._current_pack_size()
        exported = 0
        missing_code = []
        errors = []
        brand_prefix = {'baby_basic': 'baby basic', 'arie': 'ARYE'}
        for it in self._queue_data:
            code = self._safe_filename_part(it.get('model_code', ''))
            if not code:
                missing_code.append(it.get('print_name', '') or it.get('barcode', ''))
                code = self._safe_filename_part(it.get('barcode', '')) or 'label'
            size_part = self._safe_filename_part(it.get('size', ''))
            prefix = brand_prefix.get(str(it.get('brand', '') or 'arie'), 'ARYE')
            base = f"{code}_{size_part}" if size_part else code
            fname = f"{prefix} {base}.pdf"
            file_path = os.path.join(out_dir, fname)
            item = {
                "print_name": it.get("print_name", ""),
                "size": it.get("size", ""),
                "size_unit": it.get("size_unit", ""),
                "fabric": it.get("fabric", ""),
                "pack_qty": pack_qty,
                "image": it.get("image", ""),
                "barcode": it.get("barcode", ""),
            }
            try:
                build_single_label_pdf(item, file_path, logo_path=self.logo_path,
                                       draw_border=self.draw_border)
                exported += 1
            except Exception as e:
                errors.append(f"{fname}: {e}")
        if missing_code:
            messagebox.showwarning(
                "חסר קוד דגם",
                "למוצרים הבאים אין קוד דגם (נעשה שימוש בברקוד כשם הקובץ):\n"
                + "\n".join(missing_code)
                + "\n\nניתן להגדיר קוד דגם בטאב ריווחית - דאבל-קליק על מוצר.")
        if errors:
            messagebox.showerror("שגיאות בייצוא", "\n".join(errors))
        if exported:
            messagebox.showinfo("הצלחה", f"יוצאו {exported} קבצים לתיקייה:\n{out_dir}")
            try:
                os.startfile(out_dir)
            except Exception:
                pass

    def _generate_pdf(self):
        if not self._queue_data:
            messagebox.showwarning("אזהרה", "התור ריק")
            return
        from datetime import datetime
        from optitex_analyzer.core.label_sheet import build_label_sheet_pdf
        out_dir = os.path.join(os.getcwd(), "exports", "labels")
        file_path = os.path.join(out_dir, f"labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        try:
            items = self._expand_items()
            count = build_label_sheet_pdf(items, file_path, logo_path=self.logo_path,
                                          draw_border=self.draw_border)
            self._last_pdf = file_path
            messagebox.showinfo("הצלחה", f"נוצרו {count} מדבקות בקובץ:\n{file_path}")
            try:
                os.startfile(file_path)
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה ביצירת הקובץ:\n{e}")

    def _print(self):
        if not self._queue_data:
            messagebox.showwarning("אזהרה", "התור ריק")
            return
        from datetime import datetime
        from optitex_analyzer.core.label_sheet import build_label_sheet_pdf
        out_dir = os.path.join(os.getcwd(), "exports", "labels")
        file_path = os.path.join(out_dir, f"labels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        try:
            items = self._expand_items()
            build_label_sheet_pdf(items, file_path, logo_path=self.logo_path,
                                  draw_border=self.draw_border)
            self._last_pdf = file_path
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה ביצירת הקובץ:\n{e}")
            return
        try:
            if os.name == 'nt':
                printed = False
                if self._sumatra_print_fn:
                    printed = self._sumatra_print_fn(file_path)
                if not printed:
                    os.startfile(file_path, "print")
            else:
                subprocess.run(['lpr', file_path], check=True)
            messagebox.showinfo("הצלחה", "דף המדבקות נשלח להדפסה")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בהדפסה:\n{e}")
