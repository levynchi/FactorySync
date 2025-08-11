import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# סימן כיווניות RTL לטקסט עברי (Right-To-Left Mark)
RLM = '\u200f'

class DrawingsManagerTabMixin:
    """Mixin עבור טאב מנהל ציורים."""
    def _create_drawings_manager_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa'); self.notebook.add(tab, text="מנהל ציורים")
        self._drawings_tab = tab
        tk.Label(tab, text="מנהל ציורים - טבלה מקומית", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=10)
        actions = tk.Frame(tab, bg='#f7f9fa'); actions.pack(fill='x', padx=12, pady=(0,8))
        left = tk.Frame(actions, bg='#f7f9fa'); left.pack(side='left')
        tk.Button(left, text="🔄 רענן", command=self._refresh_drawings_tree, bg='#3498db', fg='white', font=('Arial',10,'bold'), width=10).pack(side='left', padx=4)
        tk.Button(left, text="📊 ייצא לאקסל", command=self._export_drawings_to_excel_tab, bg='#27ae60', fg='white', font=('Arial',10,'bold'), width=12).pack(side='left', padx=4)
        right = tk.Frame(actions, bg='#f7f9fa'); right.pack(side='right')
        tk.Button(right, text="🗑️ מחק הכל", command=self._clear_all_drawings_tab, bg='#e74c3c', fg='white', font=('Arial',10,'bold'), width=10).pack(side='right', padx=4)
        tk.Button(right, text="❌ מחק נבחר", command=self._delete_selected_drawing_tab, bg='#e67e22', fg='white', font=('Arial',10,'bold'), width=10).pack(side='right', padx=4)
        table_frame = tk.Frame(tab, bg='#ffffff'); table_frame.pack(fill='both', expand=True, padx=12, pady=8)
        # Added 'print' column for per-row print action
        cols = ("id","file_name","created_at","products","total_quantity","status","print")
        self.drawings_tree = ttk.Treeview(table_frame, columns=cols, show='headings')
        headers = {"id":"ID","file_name":"שם הקובץ","created_at":"תאריך יצירה","products":"מוצרים","total_quantity":"סך כמויות","status":"סטטוס","print":"הדפס"}
        widths = {"id":70,"file_name":260,"created_at":140,"products":80,"total_quantity":90,"status":90,"print":60}
        for c in cols:
            self.drawings_tree.heading(c, text=headers[c])
            self.drawings_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(table_frame, orient='vertical', command=self.drawings_tree.yview)
        self.drawings_tree.configure(yscroll=vs.set)
        self.drawings_tree.grid(row=0,column=0,sticky='nsew'); vs.grid(row=0,column=1,sticky='ns')
        table_frame.grid_columnconfigure(0,weight=1); table_frame.grid_rowconfigure(0,weight=1)
        self.drawings_tree.bind('<Double-1>', self._on_drawings_double_click)
        self.drawings_tree.bind('<Button-3>', self._on_drawings_right_click)
        self.drawings_tree.bind('<Button-1>', self._on_drawings_click)
        self._drawing_status_menu = tk.Menu(self.drawings_tree, tearoff=0)
        for st in ("טרם נשלח","נשלח","הוחזר","נחתך"):
            self._drawing_status_menu.add_command(label=st, command=lambda s=st: self._change_selected_drawing_status(s))
        self.drawings_stats_var = tk.StringVar(value="אין נתונים")
        tk.Label(tab, textvariable=self.drawings_stats_var, bg='#34495e', fg='white', anchor='w', padx=10, font=('Arial',10)).pack(fill='x', side='bottom')
        self._populate_drawings_tree(); self._update_drawings_stats()

    def _show_drawings_manager_tab(self):
        for i in range(len(self.notebook.tabs())):
            if self.notebook.tab(i, 'text') == "מנהל ציורים": self.notebook.select(i); break

    def _populate_drawings_tree(self):
        for item in self.drawings_tree.get_children(): self.drawings_tree.delete(item)
        for record in self.data_processor.drawings_data:
            products_count = len(record.get('מוצרים', [])); total_quantity = record.get('סך כמויות', 0)
            self.drawings_tree.insert('', 'end', values=(
                record.get('id',''),
                record.get('שם הקובץ',''),
                record.get('תאריך יצירה',''),
                products_count,
                f"{total_quantity:.1f}" if isinstance(total_quantity,(int,float)) else total_quantity,
                record.get('status','נשלח'),
                "🖨️"  # print icon
            ))

    def _update_drawings_stats(self):
        total_drawings = len(self.data_processor.drawings_data); total_quantity = sum(r.get('סך כמויות', 0) for r in self.data_processor.drawings_data)
        self.drawings_stats_var.set(f"סך הכל: {total_drawings} ציורים | סך כמויות: {total_quantity:.1f}")

    def _refresh_drawings_tree(self):
        if hasattr(self.data_processor, 'refresh_drawings_data'):
            try: self.data_processor.refresh_drawings_data()
            except Exception: pass
        self._populate_drawings_tree(); self._update_drawings_stats()

    def _on_drawings_double_click(self, event):
        item_id = self.drawings_tree.focus();
        if not item_id: return
        vals = self.drawings_tree.item(item_id, 'values');
        if not vals: return
        try: rec_id = int(vals[0])
        except Exception: return
        record = self.data_processor.get_drawing_by_id(rec_id) if hasattr(self.data_processor, 'get_drawing_by_id') else None
        if not record: return
        self._show_drawing_details(record)

    def _on_drawings_right_click(self, event):
        row_id = self.drawings_tree.identify_row(event.y)
        if row_id: self.drawings_tree.selection_set(row_id)
        sel = self.drawings_tree.selection()
        if not sel: return
        menu = tk.Menu(self.drawings_tree, tearoff=0)
        menu.add_command(label="📋 הצג פרטים", command=lambda: self._on_drawings_double_click(None))
        menu.add_cascade(label="סטטוס", menu=self._drawing_status_menu)
        menu.add_separator(); menu.add_command(label="🗑️ מחק", command=self._delete_selected_drawing_tab)
        try: menu.tk_popup(event.x_root, event.y_root)
        finally: menu.grab_release()

    def _change_selected_drawing_status(self, new_status):
        sel = self.drawings_tree.selection();
        if not sel: return
        vals = self.drawings_tree.item(sel[0], 'values');
        if not vals: return
        try: rec_id = int(vals[0])
        except Exception: return
        if hasattr(self.data_processor, 'update_drawing_status') and self.data_processor.update_drawing_status(rec_id, new_status):
            # status column is index 5 in the defined columns order
            new_vals = list(vals)
            if len(new_vals) >= 6:
                new_vals[5] = new_status
            self.drawings_tree.item(sel[0], values=new_vals)

    # === New: Click handling for print column ===
    def _on_drawings_click(self, event):
        try:
            col_id = self.drawings_tree.identify_column(event.x)  # e.g. '#1'
            row_id = self.drawings_tree.identify_row(event.y)
            if not row_id:
                return
            # Determine column name
            columns = self.drawings_tree['columns']
            idx = int(col_id.replace('#','')) - 1
            if 0 <= idx < len(columns) and columns[idx] == 'print':
                vals = self.drawings_tree.item(row_id, 'values')
                if not vals: return
                try:
                    rec_id = int(vals[0])
                except Exception:
                    return
                record = self.data_processor.get_drawing_by_id(rec_id) if hasattr(self.data_processor, 'get_drawing_by_id') else None
                if record:
                    self._print_drawing_record(record)
                return 'break'  # prevent selection change flicker
        except Exception:
            pass

    def _print_drawing_record(self, record):
        """Open a printable-style window and offer system print if possible."""
        # Header info first (RTL lines)
        header_lines = []
        header_lines.append(f"ציור: {record.get('שם הקובץ','')}")
        header_lines.append(f"ID: {record.get('id','')}")
        header_lines.append(f"תאריך יצירה: {record.get('תאריך יצירה','')}")
        if 'סוג בד' in record:
            header_lines.append(f"סוג בד: {record.get('סוג בד')}")
        header_lines.append(f"סטטוס: {record.get('status','')}")

        # Flatten products into single table rows: product(model), size, quantity
        rows = []
        import re
        def _size_key(size_str: str):
            """מפתח מיון לגדלים (תמיכה בפורמטים שונים כמו 0-3, 12m-18m, 3, 24-30)."""
            if not isinstance(size_str, str):
                return 0, size_str
            # החלפת אות m (months) כדי להשוות מספרים בלבד
            cleaned = size_str.lower().replace('m','')
            # קח את הספרות הראשונות
            m = re.match(r"(\d+)", cleaned.strip())
            base = int(m.group(1)) if m else 0
            # עדיפות: טווח לפני מספר בודד? נשתמש באורך לקביעת סדר יציב
            return base, cleaned
        for product in record.get('מוצרים', []):
            prod_name = product.get('שם המוצר','')
            # מיון לוגי של המידות לפני יצירת השורות
            sorted_sizes = sorted(product.get('מידות', []), key=lambda si: _size_key(si.get('מידה','')))
            for size_info in sorted_sizes:
                size = size_info.get('מידה','')
                qty = size_info.get('כמות',0)
                rows.append((prod_name, size, qty))

        # Determine column widths (in characters)
        prod_w = max([len(str(r[0])) for r in rows] + [4])
        size_w = max([len(str(r[1])) for r in rows] + [4])
        qty_w  = max([len(str(r[2])) for r in rows] + [4])

        # Build table header (remember RTL: we want Product on right, then Size, then Quantity on left)
        # Using monospaced font; we compose as Product | Size | Quantity in logical order, RLM will render RTL
        table_lines = []
        header_row = f"{ 'דגם'.ljust(prod_w) }  { 'מידה'.ljust(size_w) }  { 'כמות'.rjust(qty_w) }"
        sep_row = '-' * len(header_row)
        table_lines.append(header_row)
        table_lines.append(sep_row)
        for prod, size, qty in rows:
            line = f"{ str(prod).ljust(prod_w) }  { str(size).ljust(size_w) }  { str(qty).rjust(qty_w) }"
            table_lines.append(line)

        content = "\n".join([RLM + l for l in header_lines]) + "\n\n" + "\n".join([RLM + l for l in table_lines])

        top = tk.Toplevel(self.root); top.title("תצוגת הדפסה"); top.geometry('600x700'); top.configure(bg='#f0f0f0')
        txt = scrolledtext.ScrolledText(top, font=('Courier New',10), wrap='word')
        txt.pack(fill='both', expand=True, padx=8, pady=8)
        txt.tag_configure('rtl', justify='right')
        txt.insert(tk.END, content, 'rtl'); txt.config(state='disabled')
        btns = tk.Frame(top, bg='#f0f0f0'); btns.pack(fill='x', pady=6)
        tk.Button(btns, text="הדפס", command=lambda: self._attempt_system_print(content), bg='#2c3e50', fg='white', width=12).pack(side='left', padx=10)
        tk.Button(btns, text="סגור", command=top.destroy, bg='#95a5a6', fg='white', width=12).pack(side='right', padx=10)

    def _attempt_system_print(self, text_content: str):
        """Try to print using Windows default printer. Fallback: copy to clipboard."""
        try:
            import tempfile, os, sys
            # ודא שהקובץ מתחיל ב-RLM כדי לשמר יישור RTL גם בהדפסה טקסטואלית
            if not text_content.startswith(RLM):
                text_content = RLM + text_content
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.txt', mode='w', encoding='utf-8')
            tmp.write(text_content)
            tmp.close()
            # Windows specific printing via shell
            if sys.platform.startswith('win'):
                try:
                    os.startfile(tmp.name, 'print')  # type: ignore[attr-defined]
                    messagebox.showinfo("הודעה", "נשלח להדפסה (ייתכן עיכוב שניות ספורות)")
                    return
                except Exception:
                    pass
            # Fallback copy to clipboard
            self.root.clipboard_clear(); self.root.clipboard_append(text_content)
            messagebox.showinfo("הודעה", "לא ניתן להדפיס ישירות. התוכן הועתק ללוח – הדבק לקובץ ונסה להדפיס.")
        except Exception as e:
            messagebox.showerror("שגיאה", f"כשל בהדפסה: {e}")

    def _show_drawing_details(self, record):
        top = tk.Toplevel(self.root); top.title(f"פרטי ציור - {record.get('שם הקובץ','')}"); top.geometry('900x700'); top.configure(bg='#f0f0f0')
        tk.Label(top, text=f"פרטי ציור: {record.get('שם הקובץ','')}", font=('Arial',14,'bold'), bg='#f0f0f0', anchor='e', justify='right').pack(pady=10, fill='x')
        info = tk.LabelFrame(top, text="מידע כללי", bg='#f0f0f0'); info.pack(fill='x', padx=12, pady=6)
        base_txt = (
            f"ID: {record.get('id','')}\n"
            f"תאריך יצירה: {record.get('תאריך יצירה','')}\n"
            f"מספר מוצרים: {len(record.get('מוצרים', []))}\n"
            f"סך הכמויות: {record.get('סך כמויות',0)}"
        )
        if 'סוג בד' in record:
            base_txt += f"\nסוג בד: {record.get('סוג בד')}"
        status_val = record.get('status','')
        base_txt += f"\nסטטוס: {status_val}"
        tk.Label(info, text=base_txt, bg='#f0f0f0', justify='right', anchor='e').pack(fill='x', padx=8, pady=6)
        tk.Label(top, text="פירוט מוצרים ומידות:", font=('Arial',12,'bold'), bg='#f0f0f0', anchor='e', justify='right').pack(anchor='e', padx=12, pady=(6,2), fill='x')
        st = scrolledtext.ScrolledText(top, height=20, font=('Courier New',10), wrap='word')
        st.pack(fill='both', expand=True, padx=12, pady=4)
        st.tag_configure('rtl', justify='right')
        layers_used = None
        if status_val == 'נחתך':
            layers_used = self._get_layers_for_drawing(record.get('id'))
        overall_expected = 0
        for product in record.get('מוצרים', []):
            st.insert(tk.END, RLM + f"\n📦 {product.get('שם המוצר','')}\n", 'rtl')
            st.insert(tk.END, RLM + "="*60 + "\n", 'rtl')
            total_prod_q = 0; total_expected_product = 0
            # מיון לוגי של המידות להצגה עקבית
            import re
            def _size_key(size_str: str):
                if not isinstance(size_str, str):
                    return 0, size_str
                cleaned = size_str.lower().replace('m','')
                m = re.match(r"(\d+)", cleaned.strip())
                base = int(m.group(1)) if m else 0
                return base, cleaned
            product_sizes_sorted = sorted(product.get('מידות', []), key=lambda si: _size_key(si.get('מידה','')))
            for size_info in product_sizes_sorted:
                size = size_info.get('מידה',''); quantity = size_info.get('כמות',0); note = size_info.get('הערה',''); total_prod_q += quantity
                line = f"   מידה {size:>8}: {quantity:>8}"
                if layers_used and isinstance(layers_used, int) and layers_used > 0:
                    expected_qty = quantity * layers_used; total_expected_product += expected_qty; overall_expected += expected_qty
                    line += f"  | לאחר גזירה (שכבות {layers_used}): {expected_qty}"
                if note: line += f"  - {note}"
                st.insert(tk.END, RLM + line + "\n", 'rtl')
            st.insert(tk.END, RLM + f"\nסך עבור מוצר זה: {total_prod_q}", 'rtl')
            if total_expected_product:
                st.insert(tk.END, RLM + f" | סך צפוי לאחר גזירה: {total_expected_product}", 'rtl')
            st.insert(tk.END, RLM + "\n" + "-"*60 + "\n", 'rtl')
        if layers_used and overall_expected:
            st.insert(tk.END, RLM + f"\n➡ סך כמות צפויה לאחר גזירה לכל הציור: {overall_expected}\n", 'rtl')
        st.config(state='disabled')
        tk.Button(top, text="סגור", command=top.destroy, bg='#95a5a6', fg='white', font=('Arial',11,'bold'), width=12).pack(pady=10)

    def _get_layers_for_drawing(self, drawing_id):
        try:
            did_str = str(drawing_id)
            candidates = [r for r in getattr(self.data_processor, 'returned_drawings_data', []) if str(r.get('drawing_id')) == did_str and r.get('layers')]
            if not candidates: return None
            from datetime import datetime as _dt
            def _dtparse(rec):
                ts = rec.get('created_at') or rec.get('date')
                try: return _dt.strptime(ts, "%Y-%m-%d %H:%M:%S")
                except Exception: return _dt.min
            candidates.sort(key=_dtparse, reverse=True)
            layers_val = candidates[0].get('layers')
            try: return int(layers_val)
            except Exception: return None
        except Exception: return None

    def _delete_selected_drawing_tab(self):
        sel = self.drawings_tree.selection();
        if not sel: return
        vals = self.drawings_tree.item(sel[0], 'values');
        if not vals: return
        rec_id = vals[0]; file_name = vals[1]
        if not messagebox.askyesno("אישור מחיקה", f"למחוק את הציור:\n{file_name}? פעולה זו אינה הפיכה"): return
        deleted = False
        if hasattr(self.data_processor, 'delete_drawing'):
            try: deleted = self.data_processor.delete_drawing(rec_id)
            except Exception: deleted = False
        if not deleted:
            before = len(self.data_processor.drawings_data)
            self.data_processor.drawings_data = [r for r in self.data_processor.drawings_data if str(r.get('id')) != str(rec_id)]
            if len(self.data_processor.drawings_data) < before and hasattr(self.data_processor, 'save_drawings_data'):
                try: self.data_processor.save_drawings_data(); deleted = True
                except Exception: pass
        if deleted:
            self._refresh_drawings_tree(); messagebox.showinfo("הצלחה", "נמחק בהצלחה")
        else:
            messagebox.showerror("שגיאה", "המחיקה נכשלה")

    def _clear_all_drawings_tab(self):
        if not self.data_processor.drawings_data:
            messagebox.showinfo("מידע", "אין ציורים למחיקה"); return
        if not messagebox.askyesno("אישור מחיקה", f"למחוק את כל {len(self.data_processor.drawings_data)} הציורים? הפעולה לא ניתנת לשחזור"): return
        cleared = False
        if hasattr(self.data_processor, 'clear_all_drawings'):
            try: cleared = self.data_processor.clear_all_drawings()
            except Exception: cleared = False
        if cleared:
            self._refresh_drawings_tree(); messagebox.showinfo("הצלחה", "כל הציורים נמחקו")
        else:
            messagebox.showerror("שגיאה", "מחיקה נכשלה")

    def _export_drawings_to_excel_tab(self):
        if not self.data_processor.drawings_data:
            messagebox.showwarning("אזהרה", "אין ציורים לייצוא"); return
        file_path = filedialog.asksaveasfilename(title="ייצא ציורים לאקסל", defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx"),("All files","*.*")])
        if not file_path: return
        try:
            self.data_processor.export_drawings_to_excel(file_path)
            messagebox.showinfo("הצלחה", f"הציורים יוצאו אל:\n{file_path}")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
