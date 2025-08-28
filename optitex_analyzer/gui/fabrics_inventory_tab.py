import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class FabricsInventoryTabMixin:
    """Mixin לטאב מלאי בדים."""
    def _create_fabrics_inventory_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa'); self.notebook.add(tab, text="מלאי בדים")
        tk.Label(tab, text="מלאי בדים", font=('Arial', 16, 'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)
        # Action bar
        actions = tk.Frame(tab, bg='#f7f9fa'); actions.pack(fill='x', padx=15, pady=5)
        tk.Button(actions, text="⬇️ הורד תבנית אקסל למשלוח", command=self._export_fabrics_template_excel, bg='#27ae60', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5)
        tk.Button(actions, text="📥 הכנס משלוח בדים (CSV)", command=self._import_fabrics_csv, bg='#2980b9', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5)
        tk.Button(actions, text="🔄 רענן", command=self._refresh_fabrics_table, bg='#3498db', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5)
        inner_notebook = ttk.Notebook(tab); inner_notebook.pack(fill='both', expand=True, padx=10, pady=(0,5))
        inventory_tab = tk.Frame(inner_notebook, bg='#ffffff'); inner_notebook.add(inventory_tab, text="נתוני מלאי")
        unbarcoded_tab = tk.Frame(inner_notebook, bg='#ffffff'); inner_notebook.add(unbarcoded_tab, text="בדים בלי ברקוד")

        # Inventory table
        table_frame = tk.Frame(inventory_tab, bg='#ffffff'); table_frame.pack(fill='both', expand=True, padx=5, pady=5)
        cols = ('barcode','fabric_type','color_name','color_no','design_code','width','net_kg','meters','price','location','status')
        self.fabrics_tree = ttk.Treeview(table_frame, columns=cols, show='headings')
        headers = {'barcode':'ברקוד','fabric_type':'סוג בד','color_name':'צבע','color_no':'מס׳ צבע','design_code':'Desen Kodu','width':'רוחב','net_kg':'ק"ג נטו','meters':'מטרים','price':'מחיר','location':'מיקום','status':'סטטוס'}
        widths = {'barcode':120,'fabric_type':140,'color_name':110,'color_no':80,'design_code':110,'width':60,'net_kg':80,'meters':80,'price':80,'location':90,'status':80}
        for c in cols:
            self.fabrics_tree.heading(c, text=headers[c]); self.fabrics_tree.column(c, width=widths[c], anchor='center')
        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.fabrics_tree.yview); self.fabrics_tree.configure(yscroll=vsb.set)
        self.fabrics_tree.grid(row=0,column=0,sticky='nsew'); vsb.grid(row=0,column=1,sticky='ns')
        table_frame.grid_columnconfigure(0,weight=1); table_frame.grid_rowconfigure(0,weight=1)
        self._fabric_status_menu = tk.Menu(self.fabrics_tree, tearoff=0)
        for status in ("במלאי","נשלח","נגזר"):
            self._fabric_status_menu.add_command(label=status, command=lambda s=status: self._change_selected_fabric_status(s))
        self.fabrics_tree.bind('<Button-3>', self._on_fabrics_right_click)

        # Logs tab
        logs_tab = tk.Frame(inner_notebook, bg='#ffffff'); inner_notebook.add(logs_tab, text="קבצים שעלו")
        logs_frame = tk.Frame(logs_tab, bg='#ffffff'); logs_frame.pack(fill='both', expand=True, padx=5, pady=5)
        log_cols = ('id','file_name','imported_at','records_added','delete')
        self.fabrics_logs_tree = ttk.Treeview(logs_frame, columns=log_cols, show='headings')
        log_headers = {'id':'ID','file_name':'שם קובץ','imported_at':'תאריך העלאה','records_added':'רשומות','delete':'מחיקה'}
        log_widths = {'id':50,'file_name':220,'imported_at':140,'records_added':70,'delete':60}
        for c in log_cols:
            self.fabrics_logs_tree.heading(c, text=log_headers[c]); self.fabrics_logs_tree.column(c, width=log_widths[c], anchor='center')
        lsvb = ttk.Scrollbar(logs_frame, orient='vertical', command=self.fabrics_logs_tree.yview); self.fabrics_logs_tree.configure(yscroll=lsvb.set)
        self.fabrics_logs_tree.grid(row=0,column=0,sticky='nsew'); lsvb.grid(row=0,column=1,sticky='ns')
        logs_frame.grid_columnconfigure(0,weight=1); logs_frame.grid_rowconfigure(0,weight=1)
        self.fabrics_logs_tree.bind('<Button-1>', self._handle_logs_click)

        # Unbarcoded fabrics UI
        ub_actions = tk.Frame(unbarcoded_tab, bg='#ffffff'); ub_actions.pack(fill='x', padx=6, pady=6)
        tk.Button(ub_actions, text="➕ הוסף", command=self._ub_add_dialog, bg='#27ae60', fg='white').pack(side='right', padx=4)
        tk.Button(ub_actions, text="🗑️ מחק נבחר", command=self._ub_delete_selected, bg='#e67e22', fg='white').pack(side='right')
        ub_frame = tk.Frame(unbarcoded_tab, bg='#ffffff'); ub_frame.pack(fill='both', expand=True, padx=6, pady=(0,6))
        ub_cols = ('id','created_at','fabric_type','manufacturer','color','shade','notes')
        self.ub_tree = ttk.Treeview(ub_frame, columns=ub_cols, show='headings')
        ub_headers = {'id':'', 'created_at':'תאריך','fabric_type':'סוג בד','manufacturer':'יצרן הבד','color':'צבע','shade':'גוון','notes':'הערות'}
        ub_widths = {'id':60,'created_at':140,'fabric_type':160,'manufacturer':160,'color':100,'shade':80,'notes':240}
        for c in ub_cols:
            self.ub_tree.heading(c, text=ub_headers[c])
            if c == 'id':
                self.ub_tree.column(c, width=0, minwidth=0, stretch=False)
            else:
                self.ub_tree.column(c, width=ub_widths[c], anchor='center')
        ub_vsb = ttk.Scrollbar(ub_frame, orient='vertical', command=self.ub_tree.yview); self.ub_tree.configure(yscroll=ub_vsb.set)
        self.ub_tree.grid(row=0,column=0,sticky='nsew'); ub_vsb.grid(row=0,column=1,sticky='ns')
        ub_frame.grid_columnconfigure(0,weight=1); ub_frame.grid_rowconfigure(0,weight=1)
        self._populate_unbarcoded_table()

        # Footer summary
        self.fabrics_summary_var = tk.StringVar(value="אין נתונים")
        tk.Label(tab, textvariable=self.fabrics_summary_var, bg='#2c3e50', fg='white', anchor='w', padx=12, font=('Arial',10)).pack(fill='x', side='bottom')
        self._populate_fabrics_table(); self._populate_fabrics_logs(); self._update_fabrics_summary()

    def _export_fabrics_template_excel(self):
        """יוצר קובץ Excel ריק עם כותרות בסדר שהיבוא (CSV) מצפה לו."""
        # סדר ושמות העמודות כפי שהפונקציה import_fabrics_csv מצפה להם
        headers = [
            'BARCODE NO',
            'סוג בד',
            'COLOR NAME',
            'COLOR NO',
            'Desen Kodu',
            'WIDTH',
            'GR',
            'NET KG',
            'GROSS KG',
            'METER',
            'PRICE',
            'TOTAL',
            'location',
            'Last Modified',
            'מטרה',
        ]
        # בחירת נתיב שמירה
        from tkinter import filedialog, messagebox
        default_name = 'fabrics_shipment_template.xlsx'
        path = filedialog.asksaveasfilename(title='שמירת תבנית משלוח בדים', defaultextension='.xlsx', initialfile=default_name, filetypes=[('Excel','*.xlsx')])
        if not path:
            return
        try:
            # יצירת קובץ Excel עם הכותרות בלבד
            from openpyxl import Workbook  # type: ignore
            from openpyxl.styles import Font, Alignment  # type: ignore
            from openpyxl.utils import get_column_letter  # type: ignore
            wb = Workbook()
            ws = wb.active
            ws.title = 'Shipment'
            try:
                # תצוגת RTL כדי להקל על הזנה בעברית
                ws.sheet_view.rightToLeft = True
            except Exception:
                pass
            # כתיבת כותרות בשורה הראשונה
            for col_idx, name in enumerate(headers, start=1):
                c = ws.cell(row=1, column=col_idx, value=name)
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal='center')
                # רוחב עמודה אוטומטי בסיסי לפי אורך הטקסט
                try:
                    ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(28, len(name) + 4))
                except Exception:
                    pass
            # שורת עזרה אופציונלית (לא חובה)
            ws.cell(row=2, column=1, value='')
            wb.save(path)
            try:
                messagebox.showinfo('נוצר קובץ', f'הקובץ נשמר בהצלחה:\n{path}\n\nהערה: ליבוא בתוכנה יש לשמור/להמיר את הקובץ ל-CSV עם אותן כותרות.')
            except Exception:
                pass
        except Exception as e:
            try:
                messagebox.showerror('שגיאה', f'כשל ביצירת תבנית: {e}')
            except Exception:
                pass

    def _populate_fabrics_table(self):
        for item in self.fabrics_tree.get_children(): self.fabrics_tree.delete(item)
        for rec in self.data_processor.fabrics_inventory[-1000:]:
            self.fabrics_tree.insert('', 'end', values=(rec.get('barcode',''), rec.get('fabric_type',''), rec.get('color_name',''), rec.get('color_no',''), rec.get('design_code',''), rec.get('width',''), f"{rec.get('net_kg',0):.2f}", f"{rec.get('meters',0):.2f}", f"{rec.get('price',0):.2f}", rec.get('location',''), rec.get('status','במלאי')))

    def _on_fabrics_right_click(self, event):
        row_id = self.fabrics_tree.identify_row(event.y)
        if row_id:
            self.fabrics_tree.selection_set(row_id)
            try:
                self._fabric_status_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._fabric_status_menu.grab_release()

    def _change_selected_fabric_status(self, new_status):
        sel = self.fabrics_tree.selection()
        if not sel: return
        values = list(self.fabrics_tree.item(sel[0], 'values'))
        if not values: return
        barcode = values[0]
        if self.data_processor.update_fabric_status(barcode, new_status):
            values[-1] = new_status; self.fabrics_tree.item(sel[0], values=values)

    def _update_fabrics_summary(self):
        summary = self.data_processor.get_fabrics_summary()
        self.fabrics_summary_var.set(f"סה\"כ רשומות: {summary['total_records']} | מטרים: {summary['total_meters']:.2f} | ק\"ג נטו: {summary['total_net_kg']:.2f}")

    def _refresh_fabrics_table(self):
        self.data_processor.fabrics_inventory = self.data_processor.load_fabrics_inventory()
        self._populate_fabrics_table()
        if hasattr(self.data_processor, 'fabrics_import_logs'):
            self.data_processor.fabrics_import_logs = self.data_processor.load_fabrics_import_logs(); self._populate_fabrics_logs()
        # Refresh unbarcoded list
        try:
            self.data_processor.refresh_fabrics_unbarcoded(); self._populate_unbarcoded_table()
        except Exception:
            pass
        self._update_fabrics_summary()

    def _import_fabrics_csv(self):
        file_path = filedialog.askopenfilename(title="בחר קובץ CSV של משלוח בדים", filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if not file_path: return
        try:
            added = self.data_processor.import_fabrics_csv(file_path); self._refresh_fabrics_table(); messagebox.showinfo("הצלחה", f"נוספו {added} רשומות מהמשלוח")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))

    def _populate_fabrics_logs(self):
        for item in self.fabrics_logs_tree.get_children(): self.fabrics_logs_tree.delete(item)
        logs = getattr(self.data_processor, 'fabrics_import_logs', [])
        for log in sorted(logs, key=lambda x: x.get('id', 0)):
            self.fabrics_logs_tree.insert('', 'end', values=(log.get('id',''), log.get('file_name',''), log.get('imported_at',''), log.get('records_added',''), '🗑'))

    def _handle_logs_click(self, event):
        region = self.fabrics_logs_tree.identify('region', event.x, event.y)
        if region != 'cell': return
        col = self.fabrics_logs_tree.identify_column(event.x)
        if col != '#5': return
        item_id = self.fabrics_logs_tree.identify_row(event.y)
        if not item_id: return
        values = self.fabrics_logs_tree.item(item_id, 'values')
        if not values: return
        try: log_id = int(values[0])
        except Exception: return
        if not messagebox.askyesno("אישור", "למחוק רשומת לוג זו?"): return
        result = self.data_processor.delete_fabric_import_log_and_fabrics(log_id)
        if result.get('logs_deleted'):
            self._populate_fabrics_logs(); self._populate_fabrics_table()

    # ===== Unbarcoded fabrics helpers =====
    def _populate_unbarcoded_table(self):
        tree = getattr(self, 'ub_tree', None)
        if not tree: return
        for item in tree.get_children(): tree.delete(item)
        rows = getattr(self.data_processor, 'fabrics_unbarcoded', []) or []
        for r in rows:
            tree.insert('', 'end', values=(
                r.get('id',''),
                r.get('created_at',''),
                r.get('fabric_type',''),
                r.get('manufacturer',''),
                r.get('color',''),
                r.get('shade',''),
                r.get('notes','')
            ))

    def _ub_add_dialog(self):
        win = tk.Toplevel(self.root)
        win.title('הוספת בד ללא ברקוד')
        form = tk.Frame(win, padx=10, pady=10)
        form.pack(fill='both', expand=True)
        labels = ['סוג בד','יצרן הבד','צבע','גוון','הערות']
        vars_ = [tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar()]
        for i, lbl in enumerate(labels):
            tk.Label(form, text=lbl).grid(row=i, column=0, sticky='e', padx=4, pady=4)
            tk.Entry(form, textvariable=vars_[i], width=30).grid(row=i, column=1, sticky='w', padx=4, pady=4)
        btns = tk.Frame(form); btns.grid(row=len(labels), column=0, columnspan=2, sticky='e', pady=(8,0))
        def _do_add():
            try:
                new_id = self.data_processor.add_unbarcoded_fabric(vars_[0].get(), vars_[1].get(), vars_[2].get(), vars_[3].get(), vars_[4].get())
                self._populate_unbarcoded_table()
                try: messagebox.showinfo('נשמר', f'נוסף (ID: {new_id})')
                except Exception: pass
                win.destroy()
            except Exception as e:
                messagebox.showerror('שגיאה', str(e))
        tk.Button(btns, text='שמירה', command=_do_add, bg='#2c3e50', fg='white').pack(side='right', padx=4)
        tk.Button(btns, text='ביטול', command=win.destroy).pack(side='right')

    def _ub_delete_selected(self):
        sel = self.ub_tree.selection()
        if not sel: return
        item = sel[0]
        vals = self.ub_tree.item(item, 'values') or []
        if not vals: return
        try:
            rec_id = int(vals[0])
        except Exception:
            return
        try:
            if not messagebox.askyesno('אישור', f"למחוק רשומה {rec_id}?"):
                return
        except Exception:
            pass
        if self.data_processor.delete_unbarcoded_fabric(rec_id):
            self._populate_unbarcoded_table()
