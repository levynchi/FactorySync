import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class ReturnedDrawingTabMixin:
    """Mixin עבור טאב 'ציורים שנחתכו' (לשעבר קליטת ציור חוזר)."""
    def _create_returned_drawing_tab(self):
        """Create standalone top-level 'cut drawings' tab (legacy)."""
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="ציורים שנחתכו")
        self._build_returned_drawings_content(tab)

    # ===== Embedded builder =====
    def _build_returned_drawings_content(self, container: tk.Widget):
        """Build the returned / cut drawings UI inside an arbitrary container (for embedding)."""
        tk.Label(container, text="קליטת ציור שנחתך / חזר מגזירה", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)

        # Notebook inside container (scan + list)
        inner_nb = ttk.Notebook(container)
        inner_nb.pack(fill='both', expand=True, padx=8, pady=5)

        # --- Scan sub-tab ---
        scan_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(scan_tab, text="סריקת ציור")
        form = ttk.LabelFrame(scan_tab, text="פרטי ציור שנחתך", padding=12)
        form.pack(fill='x', padx=8, pady=6)

        # Row 0
        tk.Label(form, text="ציור ID:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=0, column=0, pady=4, sticky='w')
        self.return_drawing_id_var = tk.StringVar()
        # קומבובוקס לבחירת ID מתוך טבלת הציורים (ID – שם קובץ)
        from tkinter import ttk as _ttk_internal  # שמירה אם ערך צבוע ע"י כלים
        self.return_drawing_id_combo = ttk.Combobox(form, textvariable=self.return_drawing_id_var, width=32, state='readonly')
        self.return_drawing_id_combo.grid(row=0, column=1, pady=4, sticky='w')
        # רענון נתונים בזמן פתיחה / דרישה
        def _on_combo_drop(*_a):
            try: self._refresh_return_drawing_id_options()
            except Exception: pass
        try:
            self.return_drawing_id_combo.bind('<Button-1>', lambda e: _on_combo_drop())
            self.return_drawing_id_combo.bind('<FocusIn>', lambda e: _on_combo_drop())
        except Exception: pass
        # כפתור רענון קטן ליד
        tk.Button(form, text="↺", width=3, command=lambda: self._refresh_return_drawing_id_options(), bg='#3498db', fg='white').grid(row=0, column=1, sticky='e', padx=(0,4))
        tk.Label(form, text="שם הספק:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=0, column=2, pady=4, sticky='w')
        self.return_source_var = tk.StringVar()
        tk.Entry(form, textvariable=self.return_source_var, width=25).grid(row=0, column=3, pady=4, sticky='w')

        # Row 1
        tk.Label(form, text="תאריך:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=1, column=0, pady=4, sticky='w')
        self.return_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(form, textvariable=self.return_date_var, width=20).grid(row=1, column=1, pady=4, sticky='w')
        tk.Label(form, text="שכבות:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=1, column=2, pady=4, sticky='w')
        self.return_layers_var = tk.StringVar()
        tk.Entry(form, textvariable=self.return_layers_var, width=10).grid(row=1, column=3, pady=4, sticky='w')

        tk.Label(form, text="סרוק ברקודים (Enter מוסיף)").grid(row=2, column=0, columnspan=4, pady=(6,2), sticky='w')

        scan_frame = ttk.LabelFrame(scan_tab, text="ברקודים שנסרקו (בד שנחתך)", padding=8)
        scan_frame.pack(fill='both', expand=True, padx=8, pady=4)

        self.barcode_var = tk.StringVar()
        be = tk.Entry(scan_frame, textvariable=self.barcode_var, font=('Consolas',12), width=32)
        be.pack(pady=4, anchor='w')
        be.bind('<Return>', self._handle_barcode_enter)

        cols = ('barcode','fabric_type','color_name','color_no','design_code','width','net_kg','meters','price','location')
        headers = {'barcode':'ברקוד','fabric_type':'סוג בד','color_name':'צבע','color_no':'מס׳ צבע','design_code':'Desen','width':'רוחב','net_kg':'נטו','meters':'מטרים','price':'מחיר','location':'מיקום'}
        widths = {'barcode':110,'fabric_type':150,'color_name':90,'color_no':70,'design_code':90,'width':55,'net_kg':60,'meters':65,'price':55,'location':70}
        self.scanned_fabrics_tree = ttk.Treeview(scan_frame, columns=cols, show='headings', height=11)
        for c in cols:
            self.scanned_fabrics_tree.heading(c, text=headers[c])
            self.scanned_fabrics_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(scan_frame, orient='vertical', command=self.scanned_fabrics_tree.yview)
        self.scanned_fabrics_tree.configure(yscroll=vs.set)
        self.scanned_fabrics_tree.pack(side='left', fill='both', expand=True, padx=(4,0), pady=4)
        vs.pack(side='right', fill='y', pady=4)

        btns = tk.Frame(scan_frame, bg='#f7f9fa')
        btns.pack(fill='x', pady=4)
        tk.Button(btns, text="🗑️ מחק נבחר", command=self._delete_selected_barcode, bg='#e67e22', fg='white').pack(side='left', padx=4)
        tk.Button(btns, text="❌ נקה הכל", command=self._clear_all_barcodes, bg='#e74c3c', fg='white').pack(side='left', padx=4)
        tk.Button(btns, text="💾 שמור ציור שנחתך", command=self._save_returned_drawing, bg='#27ae60', fg='white').pack(side='right', padx=4)

        self.return_summary_var = tk.StringVar(value="0 ברקודים נסרקו")
        tk.Label(scan_tab, textvariable=self.return_summary_var, bg='#2c3e50', fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')

        # --- List sub-tab ---
        list_tab = tk.Frame(inner_nb, bg='#ffffff')
        inner_nb.add(list_tab, text="רשימת ציורים שנחתכו")
        lf = tk.Frame(list_tab, bg='#ffffff')
        lf.pack(fill='both', expand=True, padx=6, pady=6)
        rcols = ('id','drawing_id','date','barcodes_count','delete')
        self.returned_drawings_tree = ttk.Treeview(lf, columns=rcols, show='headings')
        h = {'id':'ID','drawing_id':'ציור','date':'תאריך','barcodes_count':'# ברקודים','delete':'מחיקה'}
        w = {'id':60,'drawing_id':140,'date':110,'barcodes_count':90,'delete':70}
        for c in rcols:
            self.returned_drawings_tree.heading(c, text=h[c])
            self.returned_drawings_tree.column(c, width=w[c], anchor='center')
        lsv = ttk.Scrollbar(lf, orient='vertical', command=self.returned_drawings_tree.yview)
        self.returned_drawings_tree.configure(yscroll=lsv.set)
        self.returned_drawings_tree.grid(row=0, column=0, sticky='nsew')
        lsv.grid(row=0, column=1, sticky='ns')
        lf.grid_columnconfigure(0, weight=1)
        lf.grid_rowconfigure(0, weight=1)
        self.returned_drawings_tree.bind('<Double-1>', self._on_returned_drawing_double_click)
        self.returned_drawings_tree.bind('<Button-1>', self._on_returned_drawings_click)

        self._scanned_barcodes = []
        self._populate_returned_drawings_table()
        # לאתחל אפשרויות ID לאחר יצירת הקומפוננטה
        try: self._refresh_return_drawing_id_options()
        except Exception: pass

    # רענון רשימת ה-ID-ים הזמינים לבחירה מהציורים
    def _refresh_return_drawing_id_options(self):
        try:
            data = getattr(self.data_processor, 'drawings_data', []) or []
            # הצגת פורמט: ID – שם קובץ (חתוך ל-40 תווים)
            options = []
            for rec in data:
                rid = rec.get('id')
                if rid is None: continue
                name = (rec.get('שם הקובץ','') or '')
                if len(name) > 40:
                    name = name[:37] + '...'
                options.append(f"{rid} - {name}")
            # מיון לפי ID מספרי אם אפשר
            def _id_key(txt):
                try: return int(str(txt).split('-',1)[0].strip())
                except Exception: return 0
            options.sort(key=_id_key)
            if hasattr(self, 'return_drawing_id_combo'):
                cur = self.return_drawing_id_var.get()
                self.return_drawing_id_combo['values'] = options
                if cur not in options:
                    # השארת ערך ריק עד בחירת המשתמש
                    pass
                # בעת בחירה – נעדכן שה-StringVar מכיל רק את ה-ID עצמו (לוגיקה פנימית)
                def _on_selected(event=None):
                    full = self.return_drawing_id_var.get()
                    if ' - ' in full:
                        self.return_drawing_id_var.set(full.split(' - ',1)[0].strip())
                        # לשמור את הטקסט המלא בתצוגה:
                        try:
                            self.return_drawing_id_combo.set(full)
                        except Exception: pass
                try: self.return_drawing_id_combo.unbind('<<ComboboxSelected>>')
                except Exception: pass
                self.return_drawing_id_combo.bind('<<ComboboxSelected>>', _on_selected)
        except Exception:
            pass

    # Handlers & logic (copied from main file methods)
    def _handle_barcode_enter(self, event=None):
        code = self.barcode_var.get().strip()
        if not code: return
        if self._scanned_barcodes and self._scanned_barcodes[-1] == code:
            self.barcode_var.set(""); return
        fabric = next((rec for rec in reversed(self.data_processor.fabrics_inventory) if str(rec.get('barcode')) == code), None)
        if not fabric:
            messagebox.showwarning("ברקוד לא נמצא", f"הברקוד {code} לא קיים במלאי הבדים"); self.barcode_var.set(""); return
        status = fabric.get('status', 'במלאי')
        if status == 'נגזר':
            messagebox.showwarning("ברקוד כבר נגזר", f"הברקוד {code} כבר מסומן כ'נגזר'"); self.barcode_var.set(""); return
        if code in self._scanned_barcodes:
            messagebox.showinfo("כפילות", f"הברקוד {code} כבר סרוק ברשימה"); self.barcode_var.set(""); return
        self._scanned_barcodes.append(code)
        values = (
            fabric.get('barcode',''), fabric.get('fabric_type',''), fabric.get('color_name',''), fabric.get('color_no',''),
            fabric.get('design_code',''), fabric.get('width',''), f"{fabric.get('net_kg',0):.2f}", f"{fabric.get('meters',0):.2f}", f"{fabric.get('price',0):.2f}", fabric.get('location','')
        )
        self.scanned_fabrics_tree.insert('', 'end', values=values)
        self.barcode_var.set(""); self._update_return_summary()

    def _delete_selected_barcode(self):
        sel = self.scanned_fabrics_tree.selection()
        if not sel: return
        all_items = self.scanned_fabrics_tree.get_children(); indices = [all_items.index(i) for i in sel]
        for item in sel: self.scanned_fabrics_tree.delete(item)
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self._scanned_barcodes): del self._scanned_barcodes[idx]
        self._update_return_summary()

    def _clear_all_barcodes(self):
        self._scanned_barcodes = []
        for item in self.scanned_fabrics_tree.get_children(): self.scanned_fabrics_tree.delete(item)
        self._update_return_summary()

    def _update_return_summary(self):
        self.return_summary_var.set(f"{len(self._scanned_barcodes)} ברקודים נסרקו")

    def _save_returned_drawing(self):
        drawing_id = self.return_drawing_id_var.get().strip(); date_str = self.return_date_var.get().strip()
        source = getattr(self, 'return_source_var', tk.StringVar()).get().strip() if hasattr(self, 'return_source_var') else ''
        layers_raw = getattr(self, 'return_layers_var', tk.StringVar()).get().strip() if hasattr(self, 'return_layers_var') else ''
        try: layers_val = int(layers_raw) if layers_raw else None
        except ValueError: layers_val = None
        if not drawing_id: messagebox.showerror("שגיאה", "אנא הכנס ציור ID"); return
        if not source: messagebox.showerror("שגיאה", "חובה למלא 'שם ספק'"); return
        if layers_val is None: messagebox.showerror("שגיאה", "חובה למלא 'שכבות' (מספר שלם)"); return
        if layers_val <= 0: messagebox.showerror("שגיאה", "ערך 'שכבות' חייב להיות גדול מ-0"); return
        if not self._scanned_barcodes: messagebox.showerror("שגיאה", "אין ברקודים לשמירה"); return
        try:
            new_id = self.data_processor.add_returned_drawing(drawing_id, date_str, self._scanned_barcodes, source=source or None, layers=layers_val)
            try:
                did = int(drawing_id)
            except ValueError:
                did = None
            if did is not None and hasattr(self.data_processor, 'update_drawing_status'):
                if self.data_processor.update_drawing_status(did, "נחתך"):
                    try: self._refresh_drawings_tree()
                    except Exception: pass
            updated = 0; unique_codes = set(self._scanned_barcodes)
            for code in unique_codes:
                if self.data_processor.update_fabric_status(code, "נגזר"): updated += 1
            if hasattr(self, 'fabrics_tree'):
                try: self._refresh_fabrics_table()
                except Exception: pass
            messagebox.showinfo("הצלחה", f"הקליטה נשמרה! ID: {new_id}\nעודכנו {updated} גלילים ל'נגזר'")
            self._clear_all_barcodes(); self._populate_returned_drawings_table()
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _populate_returned_drawings_table(self):
        for item in self.returned_drawings_tree.get_children(): self.returned_drawings_tree.delete(item)
        for rec in self.data_processor.returned_drawings_data:
            self.returned_drawings_tree.insert('', 'end', values=(rec.get('id',''), rec.get('drawing_id',''), rec.get('date',''), len(rec.get('barcodes', [])), '🗑'))

    def _on_returned_drawing_double_click(self, event):
        item_id = self.returned_drawings_tree.focus();
        if not item_id: return
        vals = self.returned_drawings_tree.item(item_id, 'values');
        if not vals: return
        rec_id = vals[0]; record = None
        for r in self.data_processor.returned_drawings_data:
            if str(r.get('id')) == str(rec_id): record = r; break
        if not record: return
        barcodes = record.get('barcodes', [])
        if not barcodes:
            messagebox.showinfo("ברקודים", "אין ברקודים לרשומה זו"); return
        top = tk.Toplevel(self.root); top.title(f"ברקודים - ציור {record.get('drawing_id','')}"); top.geometry('400x400')
        lb = tk.Listbox(top, font=('Consolas', 11)); lb.pack(fill='both', expand=True, padx=8, pady=8)
        for c in barcodes: lb.insert(tk.END, c)
        tk.Label(top, text=f"סה""כ {len(barcodes)} ברקודים", anchor='w').pack(fill='x')

    def _on_returned_drawings_click(self, event):
        region = self.returned_drawings_tree.identify('region', event.x, event.y)
        if region != 'cell': return
        col = self.returned_drawings_tree.identify_column(event.x)
        if col != '#5': return
        item_id = self.returned_drawings_tree.identify_row(event.y)
        if not item_id: return
        vals = self.returned_drawings_tree.item(item_id, 'values')
        if not vals: return
        rec_id = vals[0]
        if not messagebox.askyesno("אישור", "למחוק קליטה זו? הפעולה לא ניתנת לשחזור"): return
        if self.data_processor.delete_returned_drawing(rec_id): self._populate_returned_drawings_table()
