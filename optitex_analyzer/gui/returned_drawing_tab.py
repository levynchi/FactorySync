import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

class ReturnedDrawingTabMixin:
    """Mixin עבור טאב 'קליטת ציור שנחתך' (לשעבר קליטת ציור חוזר)."""
    def _create_returned_drawing_tab(self):
        """Create standalone top-level 'cut drawings' tab (legacy)."""
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="קליטת ציור שנחתך")
        self._build_returned_drawings_content(tab)

    # ===== Embedded builder =====
    def _build_returned_drawings_content(self, container: tk.Widget):
        """Build the returned / cut drawings UI directly in the container (no inner tabs)."""
        tk.Label(container, text="קליטת ציור שנחתך / חזר מגזירה", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=(8,4))

        # --- Scan content directly ---
        form = ttk.LabelFrame(container, text="פרטי ציור שנחתך", padding=12)
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
        tk.Label(form, text="ספק (מוצג אוטומטית):", font=('Arial',10,'bold'), width=18, anchor='w').grid(row=0, column=2, pady=4, sticky='w')
        # תצוגה בלבד של שם הספק לפי הציור הנבחר (אין שדה הזנה)
        self.return_supplier_display_var = tk.StringVar(value="")
        tk.Label(form, textvariable=self.return_supplier_display_var, width=25, anchor='w').grid(row=0, column=3, pady=4, sticky='w')

        # Row 1
        tk.Label(form, text="תאריך:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=1, column=0, pady=4, sticky='w')
        self.return_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(form, textvariable=self.return_date_var, width=20).grid(row=1, column=1, pady=4, sticky='w')
        tk.Label(form, text="שכבות:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=1, column=2, pady=4, sticky='w')
        self.return_layers_var = tk.StringVar()
        tk.Entry(form, textvariable=self.return_layers_var, width=10).grid(row=1, column=3, pady=4, sticky='w')

        # Row 2 - פירוט מוצרים מהציור
        tk.Label(form, text="פירוט מוצרים:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=2, column=0, pady=4, sticky='w')
        
        # טבלת פירוט מוצרים
        details_frame = ttk.LabelFrame(container, text="פירוט מוצרים מהציור", padding=8)
        details_frame.pack(fill='x', padx=8, pady=4)
        
        # עמודות לטבלת פירוט
        detail_cols = ('product_name', 'size', 'quantity', 'note')
        detail_headers = {'product_name': 'דגם', 'size': 'מידה', 'quantity': 'כמות', 'note': 'הערה'}
        detail_widths = {'product_name': 200, 'size': 80, 'quantity': 60, 'note': 100}
        
        self.drawing_details_tree = ttk.Treeview(details_frame, columns=detail_cols, show='headings', height=6)
        for c in detail_cols:
            self.drawing_details_tree.heading(c, text=detail_headers[c])
            self.drawing_details_tree.column(c, width=detail_widths[c], anchor='center')
        
        detail_vs = ttk.Scrollbar(details_frame, orient='vertical', command=self.drawing_details_tree.yview)
        self.drawing_details_tree.configure(yscroll=detail_vs.set)
        self.drawing_details_tree.pack(side='left', fill='both', expand=True, padx=(4,0), pady=4)
        detail_vs.pack(side='right', fill='y', pady=4)

        tk.Label(form, text="סרוק ברקודים (Enter מוסיף)").grid(row=3, column=0, columnspan=4, pady=(6,2), sticky='w')

        scan_frame = ttk.LabelFrame(container, text="ברקודים שנסרקו (בד שנחתך)", padding=8)
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
        tk.Label(container, textvariable=self.return_summary_var, bg='#2c3e50', fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')

        self._scanned_barcodes = []
        # לאתחל אפשרויות ID לאחר יצירת הקומפוננטה
        try:
            self._refresh_return_drawing_id_options()
        except Exception:
            pass

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
                    # עדכון הצגת הספק האוטומטית
                    try:
                        self._update_return_supplier_display()
                    except Exception:
                        pass
                    # עדכון פירוט המוצרים
                    try:
                        self._update_drawing_details()
                    except Exception:
                        pass
                try: self.return_drawing_id_combo.unbind('<<ComboboxSelected>>')
                except Exception: pass
                self.return_drawing_id_combo.bind('<<ComboboxSelected>>', _on_selected)
            # עדכון ראשוני של הספק אם כבר נבחר ID
            try:
                self._update_return_supplier_display()
            except Exception:
                pass
            # עדכון ראשוני של פירוט המוצרים
            try:
                self._update_drawing_details()
            except Exception:
                pass
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
        drawing_id = self.return_drawing_id_var.get().strip()
        # שכבות עבור חישובים עתידיים (לא נשמר כרשומה)
        layers_raw = getattr(self, 'return_layers_var', tk.StringVar()).get().strip() if hasattr(self, 'return_layers_var') else ''
        try:
            layers_val = int(layers_raw) if layers_raw else None
        except ValueError:
            layers_val = None
        if not drawing_id:
            messagebox.showerror("שגיאה", "אנא הכנס ציור ID"); return
        if layers_val is None:
            messagebox.showerror("שגיאה", "חובה למלא 'שכבות' (מספר שלם)"); return
        if layers_val <= 0:
            messagebox.showerror("שגיאה", "ערך 'שכבות' חייב להיות גדול מ-0"); return
        if not self._scanned_barcodes:
            messagebox.showerror("שגיאה", "אין ברקודים לשמירה"); return
        try:
            # עדכון סטטוס הציור בלשונית 'מנהל ציורים' ל"נחתך" (פירוק בטוח של ה-ID גם אם מוצג כ"ID - שם")
            raw = drawing_id
            if isinstance(raw, str) and ' - ' in raw:
                raw = raw.split(' - ', 1)[0].strip()
            did = int(str(raw).strip())
            if hasattr(self.data_processor, 'update_drawing_status'):
                self.data_processor.update_drawing_status(did, "נחתך")
            # שמירת שכבות לציור כדי לאפשר חישוב "מה נגזר אצל הספק" במאזן
            try:
                if hasattr(self.data_processor, 'update_drawing_layers'):
                    self.data_processor.update_drawing_layers(did, layers_val)
            except Exception:
                pass
            try:
                self._refresh_drawings_tree()
            except Exception:
                pass
            updated = 0; unique_codes = set(self._scanned_barcodes)
            for code in unique_codes:
                if self.data_processor.update_fabric_status(code, "נגזר"): updated += 1
            if hasattr(self, 'fabrics_tree'):
                try: self._refresh_fabrics_table()
                except Exception: pass
            messagebox.showinfo("הצלחה", f"עודכנו {updated} גלילים ל'נגזר' והסטטוס עודכן ל'נחתך'")
            self._clear_all_barcodes()
            # רענון טאב מאזן מוצרים (אם קיים) כדי להציג את החישוב החדש
            try:
                if hasattr(self, '_refresh_balance_views'):
                    self._refresh_balance_views()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _get_supplier_for_drawing_id(self, drawing_id: str) -> str:
        """מחזיר את שם הספק ('נמען') עבור ציור לפי ID אם קיים, אחרת מחרוזת ריקה."""
        try:
            did = str(drawing_id).strip()
            for rec in getattr(self.data_processor, 'drawings_data', []) or []:
                if str(rec.get('id')) == did:
                    name = (rec.get('נמען') or '').strip()
                    return name
        except Exception:
            pass
        return ""

    def _update_return_supplier_display(self):
        """עדכון תצוגת שם הספק בשדה הקריאה בלבד לפי ה-ID הנבחר."""
        try:
            sel = self.return_drawing_id_var.get()
            # אם המשתנה מכיל כבר את ה-ID בלבד – נשתמש בו; אחרת נחלץ מהטקסט המוצג
            if ' - ' in sel:
                did = sel.split(' - ', 1)[0].strip()
            else:
                did = sel.strip()
            supplier = self._get_supplier_for_drawing_id(did) if did else ""
            if hasattr(self, 'return_supplier_display_var'):
                self.return_supplier_display_var.set(supplier)
        except Exception:
            pass

    def _update_drawing_details(self):
        """עדכון טבלת פירוט המוצרים לפי הציור הנבחר."""
        try:
            # ניקוי הטבלה הקיימת
            if hasattr(self, 'drawing_details_tree'):
                for item in self.drawing_details_tree.get_children():
                    self.drawing_details_tree.delete(item)
            
            # קבלת ID הציור
            sel = self.return_drawing_id_var.get()
            if ' - ' in sel:
                drawing_id = sel.split(' - ', 1)[0].strip()
            else:
                drawing_id = sel.strip()
            
            if not drawing_id:
                return
            
            # חיפוש הציור בנתונים
            drawing_data = getattr(self.data_processor, 'drawings_data', []) or []
            drawing_record = None
            for record in drawing_data:
                if str(record.get('id')) == str(drawing_id):
                    drawing_record = record
                    break
            
            if not drawing_record:
                return
            
            # הוספת המוצרים לטבלה
            products = drawing_record.get('מוצרים', [])
            for product in products:
                product_name = product.get('שם המוצר', '')
                sizes = product.get('מידות', [])
                
                for size_info in sizes:
                    size = size_info.get('מידה', '')
                    quantity = size_info.get('כמות', 0)
                    note = size_info.get('הערה', '')
                    
                    self.drawing_details_tree.insert('', 'end', values=(
                        product_name, size, quantity, note
                    ))
        except Exception:
            pass

    # הוסר: טאב 'רשימת ציורים שנחתכו' והפונקציות הקשורות אליו
