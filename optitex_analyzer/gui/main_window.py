"""Main application window (clean single implementation)."""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from threading import Thread
from datetime import datetime


class MainWindow:
    def __init__(self, root, settings_manager, file_analyzer, data_processor):
        """Initialize main window and build all primary tabs."""
        self.root = root
        self.settings = settings_manager
        self.file_analyzer = file_analyzer
        self.data_processor = data_processor

        # Window basic setup
        self.root.title("FactorySync - ממיר אופטיטקס")
        # --- Safe geometry apply (prevents hidden / off-screen window) ---
        desired_geom = None
        try:
            desired_geom = self.settings.get("app.window_size", "1400x900")
            if not isinstance(desired_geom, str):
                desired_geom = "1400x900"
        except Exception:
            desired_geom = "1400x900"

        def _safe_apply_geometry(g: str):
            # Parse WxH+X+Y if exists
            import re
            scr_w = self.root.winfo_screenwidth()
            scr_h = self.root.winfo_screenheight()
            m = re.match(r"^(\d+)x(\d+)([+-]\d+)?([+-]\d+)?$", g.strip())
            if not m:
                return "1400x900+50+50"
            w = max(600, min(int(m.group(1)), scr_w))
            h = max(400, min(int(m.group(2)), scr_h))
            # Offsets
            x = 50
            y = 50
            if m.group(3) and m.group(4):
                try:
                    x = int(m.group(3))
                    y = int(m.group(4))
                except ValueError:
                    x, y = 50, 50
            # If off-screen adjust
            if x < 0 or x > scr_w - 100:
                x = 50
            if y < 0 or y > scr_h - 100:
                y = 50
            return f"{w}x{h}+{x}+{y}"

        safe_geom = _safe_apply_geometry(desired_geom or "1400x900")
        try:
            self.root.geometry(safe_geom)
        except Exception:
            self.root.geometry("1400x900+50+50")
        # Bring to front briefly (in case hidden behind other windows)
        self.root.update_idletasks()
        self.root.deiconify()
        self.root.lift()
        try:
            self.root.attributes('-topmost', True)
            self.root.after(500, lambda: self.root.attributes('-topmost', False))
        except Exception:
            pass

        # State vars
        self.rib_file = ""
        self.products_file = self.settings.get("app.products_file", "")
        if self.products_file and not os.path.exists(self.products_file):
            self.products_file = ""
        self.current_results = []
        self.drawings_manager_window = None  # legacy (window mode removed)

        # Notebook (main tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True)

        # Build tabs
        self._create_converter_tab()
        self._create_returned_drawing_tab()
        self._create_fabrics_inventory_tab()
        self._create_drawings_manager_tab()  # new main tab instead of popup window

        # Status bar + settings load
        self._create_status_bar()
        self._load_initial_settings()

    # ===== Converter Tab =====
    def _create_converter_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="ממיר קבצים")
        for builder in (self._create_files_section, self._create_options_section, self._create_action_buttons, self._create_results_section):
            orig = self.root; self.root = tab; builder(); self.root = orig

    # ===== Returned Drawing Tab =====
    def _create_returned_drawing_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="קליטת ציור חוזר")
        tk.Label(tab, text="קליטת ציור שחזר מייצור", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)
        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=8, pady=5)
        # --- Scan tab ---
        scan_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(scan_tab, text="סריקת ציור")
        form = ttk.LabelFrame(scan_tab, text="פרטי קליטה", padding=12)
        form.pack(fill='x', padx=8, pady=6)
        # Row 0
        tk.Label(form, text="ציור ID:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=0,column=0,pady=4,sticky='w')
        self.return_drawing_id_var = tk.StringVar()
        tk.Entry(form, textvariable=self.return_drawing_id_var, width=30).grid(row=0,column=1,pady=4,sticky='w')
        tk.Label(form, text="שם הספק:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=0,column=2,pady=4,sticky='w')
        self.return_source_var = tk.StringVar()
        tk.Entry(form, textvariable=self.return_source_var, width=25).grid(row=0,column=3,pady=4,sticky='w')
        # Row 1
        tk.Label(form, text="תאריך:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=1,column=0,pady=4,sticky='w')
        self.return_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(form, textvariable=self.return_date_var, width=20).grid(row=1,column=1,pady=4,sticky='w')
        tk.Label(form, text="שכבות:", font=('Arial',10,'bold'), width=12, anchor='w').grid(row=1,column=2,pady=4,sticky='w')
        self.return_layers_var = tk.StringVar()
        tk.Entry(form, textvariable=self.return_layers_var, width=10).grid(row=1,column=3,pady=4,sticky='w')
        # Instruction
        tk.Label(form, text="סרוק ברקודים (Enter מוסיף)").grid(row=2,column=0,columnspan=4,pady=(6,2),sticky='w')
        # Barcode list
        scan_frame = ttk.LabelFrame(scan_tab, text="ברקודים נסרקים", padding=8)
        scan_frame.pack(fill='both', expand=True, padx=8, pady=4)
        self.barcode_var = tk.StringVar()
        be = tk.Entry(scan_frame, textvariable=self.barcode_var, font=('Consolas',12), width=32)
        be.pack(pady=4, anchor='w')
        be.bind('<Return>', self._handle_barcode_enter)
        cols = ('barcode','fabric_type','color_name','color_no','design_code','width','net_kg','meters','price','location')
        self.scanned_fabrics_tree = ttk.Treeview(scan_frame, columns=cols, show='headings', height=11)
        headers = {'barcode':'ברקוד','fabric_type':'סוג בד','color_name':'צבע','color_no':'מס׳ צבע','design_code':'Desen','width':'רוחב','net_kg':'נטו','meters':'מטרים','price':'מחיר','location':'מיקום'}
        widths = {'barcode':110,'fabric_type':150,'color_name':90,'color_no':70,'design_code':90,'width':55,'net_kg':60,'meters':65,'price':55,'location':70}
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
        tk.Button(btns, text="💾 שמור קליטה", command=self._save_returned_drawing, bg='#27ae60', fg='white').pack(side='right', padx=4)
        self.return_summary_var = tk.StringVar(value="0 ברקודים נסרקו")
        tk.Label(scan_tab, textvariable=self.return_summary_var, bg='#2c3e50', fg='white', anchor='w', padx=10).pack(fill='x', side='bottom')
        # --- List tab ---
        list_tab = tk.Frame(inner_nb, bg='#ffffff')
        inner_nb.add(list_tab, text="רשימת ציורים שנקלטו")
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
        self.returned_drawings_tree.grid(row=0,column=0,sticky='nsew')
        lsv.grid(row=0,column=1,sticky='ns')
        lf.grid_columnconfigure(0,weight=1)
        lf.grid_rowconfigure(0,weight=1)
        self.returned_drawings_tree.bind('<Double-1>', self._on_returned_drawing_double_click)
        self.returned_drawings_tree.bind('<Button-1>', self._on_returned_drawings_click)
        self._scanned_barcodes = []
        self._populate_returned_drawings_table()

    def _handle_barcode_enter(self, event=None):
        code = self.barcode_var.get().strip()
        if not code:
            return
        # prevent immediate duplicate scan (במקרה שהסורק שולח פעמיים מהר)
        if self._scanned_barcodes and self._scanned_barcodes[-1] == code:
            self.barcode_var.set("")
            return
        # חיפוש הבד במלאי
        fabric = next((rec for rec in reversed(self.data_processor.fabrics_inventory) if str(rec.get('barcode')) == code), None)
        if not fabric:
            # לא קיים במלאי – לא מוסיפים לרשימת הסרוקים
            messagebox.showwarning("ברקוד לא נמצא", f"הברקוד {code} לא קיים במלאי הבדים")
            self.barcode_var.set("")
            return
        # בדיקה אם הבד כבר מסומן כנגזר
        status = fabric.get('status', 'במלאי')
        if status == 'נגזר':
            messagebox.showwarning("ברקוד כבר נגזר", f"הברקוד {code} כבר מסומן כ'נגזר' במלאי ולכן לא ניתן לקלוט אותו")
            self.barcode_var.set("")
            return
        # (אופציונלי) מניעת כפילות מוקדמת ברשימת הסרוקים
        if code in self._scanned_barcodes:
            messagebox.showinfo("כפילות", f"הברקוד {code} כבר סרוק ברשימה")
            self.barcode_var.set("")
            return
        # הכל תקין – מוסיפים
        self._scanned_barcodes.append(code)
        values = (
            fabric.get('barcode',''),
            fabric.get('fabric_type',''),
            fabric.get('color_name',''),
            fabric.get('color_no',''),
            fabric.get('design_code',''),
            fabric.get('width',''),
            f"{fabric.get('net_kg',0):.2f}",
            f"{fabric.get('meters',0):.2f}",
            f"{fabric.get('price',0):.2f}",
            fabric.get('location','')
        )
        self.scanned_fabrics_tree.insert('', 'end', values=values)
        self.barcode_var.set("")
        self._update_return_summary()

    def _delete_selected_barcode(self):
        if hasattr(self, 'scanned_fabrics_tree'):
            sel = self.scanned_fabrics_tree.selection()
            if not sel:
                return
            # חישוב אינדקסים לפי סדר הצגה
            all_items = self.scanned_fabrics_tree.get_children()
            indices = [all_items.index(i) for i in sel]
            for item in sel:
                self.scanned_fabrics_tree.delete(item)
            # מחיקת הברקודים מהרשימה בזיכרון (מהסוף להתחלה)
            for idx in sorted(indices, reverse=True):
                if 0 <= idx < len(self._scanned_barcodes):
                    del self._scanned_barcodes[idx]
        self._update_return_summary()

    def _clear_all_barcodes(self):
        self._scanned_barcodes = []
        if hasattr(self, 'scanned_fabrics_tree'):
            for item in self.scanned_fabrics_tree.get_children():
                self.scanned_fabrics_tree.delete(item)
        self._update_return_summary()

    def _update_return_summary(self):
        count = len(self._scanned_barcodes)
        self.return_summary_var.set(f"{count} ברקודים נסרקו")

    def _save_returned_drawing(self):
        drawing_id = self.return_drawing_id_var.get().strip()
        date_str = self.return_date_var.get().strip()
        source = getattr(self, 'return_source_var', tk.StringVar()).get().strip() if hasattr(self, 'return_source_var') else ''
        layers_raw = getattr(self, 'return_layers_var', tk.StringVar()).get().strip() if hasattr(self, 'return_layers_var') else ''
        try:
            layers_val = int(layers_raw) if layers_raw else None
        except ValueError:
            layers_val = None
        if not drawing_id:
            messagebox.showerror("שגיאה", "אנא הכנס ציור ID")
            return
        if not self._scanned_barcodes:
            messagebox.showerror("שגיאה", "אין ברקודים לשמירה")
            return
        try:
            new_id = self.data_processor.add_returned_drawing(drawing_id, date_str, self._scanned_barcodes, source=source or None, layers=layers_val)
            # עדכון סטטוס הבדים שנקלטו ל"נגזר"
            updated = 0
            unique_codes = set(self._scanned_barcodes)
            for code in unique_codes:
                if self.data_processor.update_fabric_status(code, "נגזר"):
                    updated += 1
            # רענון טבלת מלאי אם פתוחה
            if hasattr(self, 'fabrics_tree'):
                try:
                    self._refresh_fabrics_table()
                except Exception:
                    pass
            messagebox.showinfo("הצלחה", f"הקליטה נשמרה בהצלחה!\nID: {new_id}\nעודכנו סטטוסים ל-{updated} גלילים")
            self._clear_all_barcodes()
            self._populate_returned_drawings_table()
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
    
    def _populate_returned_drawings_table(self):
        if not hasattr(self, 'returned_drawings_tree'):
            return
        for item in self.returned_drawings_tree.get_children():
            self.returned_drawings_tree.delete(item)
        for rec in self.data_processor.returned_drawings_data:
            self.returned_drawings_tree.insert('', 'end', values=(
                rec.get('id',''),
                rec.get('drawing_id',''),
                rec.get('date',''),
                len(rec.get('barcodes', [])),
                '🗑'
            ))

    def _on_returned_drawing_double_click(self, event):
        item_id = self.returned_drawings_tree.focus()
        if not item_id:
            return
        vals = self.returned_drawings_tree.item(item_id, 'values')
        if not vals:
            return
        rec_id = vals[0]
        # מציאת הרשומה
        record = None
        for r in self.data_processor.returned_drawings_data:
            if str(r.get('id')) == str(rec_id):
                record = r
                break
        if not record:
            return
        barcodes = record.get('barcodes', [])
        if not barcodes:
            messagebox.showinfo("ברקודים", "אין ברקודים לרשומה זו")
            return
        # הצגת ברקודים בחלון נפרד
        top = tk.Toplevel(self.root)
        top.title(f"ברקודים - ציור {record.get('drawing_id','')}")
        top.geometry('400x400')
        lb = tk.Listbox(top, font=('Consolas', 11))
        lb.pack(fill='both', expand=True, padx=8, pady=8)
        for c in barcodes:
            lb.insert(tk.END, c)
        tk.Label(top, text=f"סה""כ {len(barcodes)} ברקודים", anchor='w').pack(fill='x')

    def _on_returned_drawings_click(self, event):
        region = self.returned_drawings_tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
        col = self.returned_drawings_tree.identify_column(event.x)
        if col != '#5':  # delete column
            return
        item_id = self.returned_drawings_tree.identify_row(event.y)
        if not item_id:
            return
        vals = self.returned_drawings_tree.item(item_id, 'values')
        if not vals:
            return
        rec_id = vals[0]
        if not messagebox.askyesno("אישור", "למחוק קליטה זו? הפעולה לא ניתנת לשחזור"):
            return
        if self.data_processor.delete_returned_drawing(rec_id):
            self._populate_returned_drawings_table()


    # ===== טאב מלאי בדים =====
    def _create_fabrics_inventory_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="מלאי בדים")

        tk.Label(tab, text="מלאי בדים", font=('Arial', 16, 'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)

        actions = tk.Frame(tab, bg='#f7f9fa')
        actions.pack(fill='x', padx=15, pady=5)
        tk.Button(actions, text="📥 הכנס משלוח בדים (CSV)", command=self._import_fabrics_csv, bg='#2980b9', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5)
        tk.Button(actions, text="🔄 רענן", command=self._refresh_fabrics_table, bg='#3498db', fg='white', font=('Arial', 10, 'bold')).pack(side='right', padx=5)

        inner_notebook = ttk.Notebook(tab)
        inner_notebook.pack(fill='both', expand=True, padx=10, pady=(0,5))

        # טאב מלאי
        inventory_tab = tk.Frame(inner_notebook, bg='#ffffff')
        inner_notebook.add(inventory_tab, text="נתוני מלאי")
        table_frame = tk.Frame(inventory_tab, bg='#ffffff'); table_frame.pack(fill='both', expand=True, padx=5, pady=5)
        cols = ('barcode','fabric_type','color_name','color_no','design_code','width','net_kg','meters','price','location','status')
        self.fabrics_tree = ttk.Treeview(table_frame, columns=cols, show='headings')
        headers = {'barcode':'ברקוד','fabric_type':'סוג בד','color_name':'צבע','color_no':'מס׳ צבע','design_code':'Desen Kodu','width':'רוחב','net_kg':'ק"ג נטו','meters':'מטרים','price':'מחיר','location':'מיקום','status':'סטטוס'}
        widths = {'barcode':120,'fabric_type':140,'color_name':110,'color_no':80,'design_code':110,'width':60,'net_kg':80,'meters':80,'price':80,'location':90,'status':80}
        for c in cols:
            self.fabrics_tree.heading(c, text=headers[c]); self.fabrics_tree.column(c, width=widths[c], anchor='center')
        vsb = ttk.Scrollbar(table_frame, orient='vertical', command=self.fabrics_tree.yview)
        self.fabrics_tree.configure(yscroll=vsb.set)
        self.fabrics_tree.grid(row=0,column=0,sticky='nsew'); vsb.grid(row=0,column=1,sticky='ns')
        table_frame.grid_columnconfigure(0,weight=1); table_frame.grid_rowconfigure(0,weight=1)
        # תפריט סטטוס
        self._fabric_status_menu = tk.Menu(self.fabrics_tree, tearoff=0)
        for status in ("במלאי","נשלח","נגזר"):
            self._fabric_status_menu.add_command(label=status, command=lambda s=status: self._change_selected_fabric_status(s))
        self.fabrics_tree.bind('<Button-3>', self._on_fabrics_right_click)

        # טאב לוגים
        logs_tab = tk.Frame(inner_notebook, bg='#ffffff'); inner_notebook.add(logs_tab, text="קבצים שעלו")
        logs_frame = tk.Frame(logs_tab, bg='#ffffff'); logs_frame.pack(fill='both', expand=True, padx=5, pady=5)
        log_cols = ('id','file_name','imported_at','records_added','delete')
        self.fabrics_logs_tree = ttk.Treeview(logs_frame, columns=log_cols, show='headings')
        log_headers = {'id':'ID','file_name':'שם קובץ','imported_at':'תאריך העלאה','records_added':'רשומות','delete':'מחיקה'}
        log_widths = {'id':50,'file_name':220,'imported_at':140,'records_added':70,'delete':60}
        for c in log_cols:
            self.fabrics_logs_tree.heading(c, text=log_headers[c]); self.fabrics_logs_tree.column(c, width=log_widths[c], anchor='center')
        lsvb = ttk.Scrollbar(logs_frame, orient='vertical', command=self.fabrics_logs_tree.yview)
        self.fabrics_logs_tree.configure(yscroll=lsvb.set)
        self.fabrics_logs_tree.grid(row=0,column=0,sticky='nsew'); lsvb.grid(row=0,column=1,sticky='ns')
        logs_frame.grid_columnconfigure(0,weight=1); logs_frame.grid_rowconfigure(0,weight=1)
        self.fabrics_logs_tree.bind('<Button-1>', self._handle_logs_click)

        # סיכום
        self.fabrics_summary_var = tk.StringVar(value="אין נתונים")
        tk.Label(tab, textvariable=self.fabrics_summary_var, bg='#2c3e50', fg='white', anchor='w', padx=12, font=('Arial',10)).pack(fill='x', side='bottom')

        self._populate_fabrics_table(); self._populate_fabrics_logs(); self._update_fabrics_summary()

    def _populate_fabrics_table(self):
        if not hasattr(self, 'fabrics_tree'):
            return
        for item in self.fabrics_tree.get_children():
            self.fabrics_tree.delete(item)
        for rec in self.data_processor.fabrics_inventory[-1000:]:  # מגביל להצגה אחרונה אם גדול
            self.fabrics_tree.insert('', 'end', values=(
                rec.get('barcode',''),
                rec.get('fabric_type',''),
                rec.get('color_name',''),
                rec.get('color_no',''),
                rec.get('design_code',''),
                rec.get('width',''),
                f"{rec.get('net_kg',0):.2f}",
                f"{rec.get('meters',0):.2f}",
                f"{rec.get('price',0):.2f}",
                rec.get('location',''),
                rec.get('status','במלאי')
            ))

    def _on_fabrics_right_click(self, event):
        # בחירה לפי מיקום סמן
        row_id = self.fabrics_tree.identify_row(event.y)
        if row_id:
            self.fabrics_tree.selection_set(row_id)
            try:
                self._fabric_status_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._fabric_status_menu.grab_release()

    def _change_selected_fabric_status(self, new_status):
        sel = self.fabrics_tree.selection()
        if not sel:
            return
        item = sel[0]
        values = list(self.fabrics_tree.item(item, 'values'))
        if not values:
            return
        barcode = values[0]
        # עדכון בזיכרון + דיסק
        if self.data_processor.update_fabric_status(barcode, new_status):
            values[-1] = new_status
            self.fabrics_tree.item(item, values=values)

    def _update_fabrics_summary(self):
        summary = self.data_processor.get_fabrics_summary()
        self.fabrics_summary_var.set(
            f"סה\"כ רשומות: {summary['total_records']} | מטרים: {summary['total_meters']:.2f} | ק""ג נטו: {summary['total_net_kg']:.2f}"
        )

    def _refresh_fabrics_table(self):
        self.data_processor.fabrics_inventory = self.data_processor.load_fabrics_inventory()
        self._populate_fabrics_table()
        # רענון לוגים
        if hasattr(self.data_processor, 'fabrics_import_logs'):
            self.data_processor.fabrics_import_logs = self.data_processor.load_fabrics_import_logs()
            self._populate_fabrics_logs()
        self._update_fabrics_summary()

    def _import_fabrics_csv(self):
        file_path = filedialog.askopenfilename(title="בחר קובץ CSV של משלוח בדים", filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if not file_path:
            return
        try:
            added = self.data_processor.import_fabrics_csv(file_path)
            self._refresh_fabrics_table()
            messagebox.showinfo("הצלחה", f"נוספו {added} רשומות מהמשלוח")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
    
    def _populate_fabrics_logs(self):
        if not hasattr(self, 'fabrics_logs_tree'):
            return
        for item in self.fabrics_logs_tree.get_children():
            self.fabrics_logs_tree.delete(item)
        logs = getattr(self.data_processor, 'fabrics_import_logs', [])
        for log in sorted(logs, key=lambda x: x.get('id', 0)):
            self.fabrics_logs_tree.insert('', 'end', values=(
                log.get('id',''),
                log.get('file_name',''),
                log.get('imported_at',''),
                log.get('records_added',''),
                '🗑'
            ))
    
    def _handle_logs_click(self, event):
        # זיהוי עמודה
        region = self.fabrics_logs_tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
        col = self.fabrics_logs_tree.identify_column(event.x)  # e.g. '#5'
        if col != '#5':
            return  # לא עמודת המחיקה
        item_id = self.fabrics_logs_tree.identify_row(event.y)
        if not item_id:
            return
        values = self.fabrics_logs_tree.item(item_id, 'values')
        if not values:
            return
        try:
            log_id = int(values[0])
        except Exception:
            return
        # אישור
        if not messagebox.askyesno("אישור", "למחוק רשומת לוג זו?"):
            return
        result = self.data_processor.delete_fabric_import_log_and_fabrics(log_id)
        if result.get('logs_deleted'):
            self._populate_fabrics_logs()
            self._populate_fabrics_table()
    
    def _create_files_section(self):
        """יצירת מקטע בחירת קבצים"""
        files_frame = ttk.LabelFrame(self.root, text="בחירת קבצים", padding=15)
        files_frame.pack(fill="x", padx=20, pady=10)
        
        # קובץ אופטיטקס
        rib_frame = tk.Frame(files_frame)
        rib_frame.pack(fill="x", pady=8)
        
        tk.Label(
            rib_frame,
            text="קובץ אופטיטקס:",
            font=('Arial', 10, 'bold'),
            width=15,
            anchor="w"
        ).pack(side="left")
        
        self.rib_label = tk.Label(
            rib_frame,
            text="לא נבחר קובץ",
            bg="white",
            relief="sunken",
            width=60,
            anchor="w",
            padx=5
        )
        self.rib_label.pack(side="left", padx=10)
        
        tk.Button(
            rib_frame,
            text="📁 בחר קובץ",
            command=self._select_rib_file,
            bg='#3498db',
            fg='white',
            font=('Arial', 9, 'bold'),
            width=12
        ).pack(side="right")
        
        # קובץ מוצרים
        products_frame = tk.Frame(files_frame)
        products_frame.pack(fill="x", pady=8)
        
        tk.Label(
            products_frame,
            text="קובץ מוצרים:",
            font=('Arial', 10, 'bold'),
            width=15,
            anchor="w"
        ).pack(side="left")
        
        self.products_label = tk.Label(
            products_frame,
            text="לא נבחר קובץ",
            bg="white",
            relief="sunken",
            width=60,
            anchor="w",
            padx=5
        )
        self.products_label.pack(side="left", padx=10)
        
        tk.Button(
            products_frame,
            text="📁 בחר קובץ",
            command=self._select_products_file,
            bg='#3498db',
            fg='white',
            font=('Arial', 9, 'bold'),
            width=12
        ).pack(side="right")
    
    def _create_options_section(self):
        """יצירת מקטע אפשרויות"""
        options_frame = ttk.LabelFrame(self.root, text="אפשרויות", padding=15)
        options_frame.pack(fill="x", padx=20, pady=10)
        
        # אפשרויות עיבוד
        processing_frame = tk.Frame(options_frame)
        processing_frame.pack(fill="x", pady=5)
        
        self.tubular_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            processing_frame,
            text="טיפול אוטומטי ב-Layout Tubular (חלוקה ב-2)",
            variable=self.tubular_var,
            font=('Arial', 10)
        ).pack(anchor="w")
        
        self.only_positive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            processing_frame,
            text="הצג רק מידות עם כמות גדולה מ-0",
            variable=self.only_positive_var,
            font=('Arial', 10)
        ).pack(anchor="w")
    
    def _create_action_buttons(self):
        """יצירת כפתורי הפעולה"""
        buttons_frame = tk.Frame(self.root, bg='#f0f0f0')
        buttons_frame.pack(fill="x", padx=20, pady=15)
        
        # שורה ראשונה
        row1 = tk.Frame(buttons_frame, bg='#f0f0f0')
        row1.pack(fill="x", pady=5)
        
        tk.Button(
            row1,
            text="🔍 נתח קבצים",
            command=self._analyze_files,
            bg='#27ae60',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=15
        ).pack(side="left", padx=5)
        
        tk.Button(
            row1,
            text="💾 שמור כ-Excel",
            command=self._save_excel,
            bg='#e67e22',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=15
        ).pack(side="left", padx=5)
        
        tk.Button(
            row1,
            text="️ נקה הכל",
            command=self._clear_all,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=15
        ).pack(side="right", padx=5)
        
        # שורה שנייה
        row2 = tk.Frame(buttons_frame, bg='#f0f0f0')
        row2.pack(fill="x", pady=5)
        
        tk.Button(
            row2,
            text=" הוסף לטבלה מקומית",
            command=self._add_to_local_table,
            bg='#16a085',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=18
        ).pack(side="left", padx=5)
        
        tk.Button(
            row2,
            text="📁 מנהל ציורים",
            command=self._show_drawings_manager_tab,
            bg='#2980b9',
            fg='white',
            font=('Arial', 11, 'bold'),
            height=2,
            width=18
        ).pack(side="left", padx=5)
    
    def _create_results_section(self):
        """יצירת אזור התוצאות"""
        results_frame = ttk.LabelFrame(self.root, text="תוצאות וסטטוס", padding=10)
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.results_text = scrolledtext.ScrolledText(
            results_frame,
            height=15,
            font=('Consolas', 10),
            wrap=tk.WORD,
            bg='#f8f9fa',
            fg='#2c3e50'
        )
        self.results_text.pack(fill="both", expand=True)
    
    def _create_status_bar(self):
        """יצירת שורת הסטטוס"""
        self.status_label = tk.Label(
            self.root,
            text="מוכן לעבודה",
            bg='#34495e',
            fg='white',
            anchor='w',
            padx=15,
            font=('Arial', 10)
        )
        self.status_label.pack(fill="x", side="bottom")
    
    def _load_initial_settings(self):
        """טעינת הגדרות ראשוניות"""
        # טעינה אוטומטית של קובץ מוצרים
        if self.settings.get("app.auto_load_products", True):
            products_file = self.settings.get("app.products_file", "קובץ מוצרים.xlsx")
            if os.path.exists(products_file):
                self.products_file = os.path.abspath(products_file)
                self.products_label.config(text=os.path.basename(products_file))
                self._update_status(f"נטען קובץ מוצרים: {os.path.basename(products_file)}")
    
    # File Selection Methods
    def _select_rib_file(self):
        """בחירת קובץ אופטיטקס"""
        file_path = filedialog.askopenfilename(
            title="בחר קובץ אופטיטקס אקסל אקספורט",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.rib_file = file_path
            self.rib_label.config(text=os.path.basename(file_path))
            self._update_status(f"נבחר קובץ אופטיטקס: {os.path.basename(file_path)}")
    
    def _select_products_file(self):
        """בחירת קובץ מוצרים"""
        file_path = filedialog.askopenfilename(
            title="בחר קובץ רשימת מוצרים",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.products_file = file_path
            self.products_label.config(text=os.path.basename(file_path))
            self._update_status(f"נבחר קובץ מוצרים: {os.path.basename(file_path)}")
    
    # Analysis Methods
    def _analyze_files(self):
        """ניתוח הקבצים"""
        if not self.rib_file or not self.products_file:
            messagebox.showerror("שגיאה", "אנא בחר את שני הקבצים")
            return
        
        self._clear_results()
        self._update_status("מנתח קבצים...")
        
        # הרצה בחוט נפרד
        Thread(target=self._analyze_files_thread, daemon=True).start()
    
    def _analyze_files_thread(self):
        """ביצוע הניתוח בחוט נפרד"""
        try:
            self._log_message("=== התחלת ניתוח ===")
            
            # טעינת מיפוי מוצרים
            self._log_message("טוען מיפוי מוצרים...")
            if not self.file_analyzer.load_products_mapping(self.products_file):
                raise Exception("שגיאה בטעינת קובץ מוצרים")
            
            products_count = len(self.file_analyzer.product_mapping)
            self._log_message(f"✅ נטען מיפוי עבור {products_count} מוצרים")
            
            # ניתוח הקובץ
            self._log_message("מנתח קובץ אופטיטקס...")
            results = self.file_analyzer.analyze_file(
                self.rib_file,
                self.tubular_var.get(),
                self.only_positive_var.get()
            )
            
            if not results:
                self._log_message("❌ לא נמצאו נתונים מתאימים")
                self._update_status("לא נמצאו נתונים")
                return
            
            # מיון התוצאות
            self.current_results = self.file_analyzer.sort_results()
            
            # הצגת סיכום
            summary = self.file_analyzer.get_analysis_summary()
            self._log_message(f"✅ נוצרה טבלה עם {summary['total_records']} רשומות")
            
            # הצגת מוצרים שנמצאו
            found_products = self.file_analyzer.get_products_found()
            if found_products:
                self._log_message("\n📦 מוצרים שנמצאו:")
                for file_name, product_name in found_products:
                    self._log_message(f"   {file_name} → {product_name}")
            
            # הצגת תוצאות מפורטות
            self._display_detailed_results()
            
            # הצגת סטטיסטיקות
            self._display_statistics(summary)
            
            self._update_status("הניתוח הושלם בהצלחה!")
            
        except Exception as e:
            error_msg = f"❌ שגיאה בניתוח: {str(e)}"
            self._log_message(error_msg)
            self._update_status("שגיאה בניתוח")
            messagebox.showerror("שגיאה", str(e))
    
    def _display_detailed_results(self):
        """הצגת תוצאות מפורטות"""
        self._log_message("\n=== תוצאות הניתוח ===")
        
        current_product = None
        for result in self.current_results:
            if current_product != result['שם המוצר']:
                current_product = result['שם המוצר']
                self._log_message(f"\n📦 {current_product}:")
                self._log_message("-" * 60)
            
            quantity_text = f"{result['כמות']}"
            if result['כמות מקורית'] != result['כמות']:
                quantity_text += f" (מקורי: {result['כמות מקורית']})"
            
            self._log_message(f"   מידה {result['מידה']:>8}: {quantity_text:>10} - {result['הערה']}")
    
    def _display_statistics(self, summary):
        """הצגת סטטיסטיקות"""
        self._log_message("\n" + "=" * 70)
        self._log_message(f"\n=== סיכום ===")
        self._log_message(f"מוצרים: {summary['unique_products']}")
        self._log_message(f"מידות שונות: {summary['unique_sizes']}")
        self._log_message(f"סך רשומות: {summary['total_records']}")
        self._log_message(f"סך כמויות: {summary['total_quantity']:.1f}")
        
        if summary['is_tubular']:
            self._log_message("🔄 הכמויות חולקו ב-2 בגלל Layout: Tubular")
    
    # Export Methods
    def _save_excel(self):
        """שמירה כ-Excel"""
        if not self.current_results:
            messagebox.showwarning("אזהרה", "אין נתונים לשמירה. אנא בצע ניתוח תחילה.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="שמור כ-Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.data_processor.export_to_excel(self.current_results, file_path)
                self._log_message(f"📁 הקובץ נשמר בהצלחה: {file_path}")
                self._update_status(f"נשמר: {os.path.basename(file_path)}")
                messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה!")
            except Exception as e:
                messagebox.showerror("שגיאה", str(e))
    
    def _add_to_local_table(self):
        """הוספה לטבלה המקומית"""
        if not self.current_results:
            messagebox.showwarning("אזהרה", "אין נתונים להוספה. אנא בצע ניתוח תחילה.")
            return
        
        try:
            record_id = self.data_processor.add_to_local_table(self.current_results, self.rib_file)
            
            self._log_message(f"\n✅ הציור נוסף לטבלה המקומית!")
            self._log_message(f"ID רשומה חדשה: {record_id}")
            
            file_name = os.path.splitext(os.path.basename(self.rib_file))[0] if self.rib_file else 'לא ידוע'
            total_quantity = sum(r['כמות'] for r in self.current_results)
            
            self._log_message(f"שם הקובץ: {file_name}")
            self._log_message(f"סך כמויות: {total_quantity}")
            
            self._update_status("נוסף לטבלה המקומית")
            messagebox.showinfo("הצלחה", f"הציור נוסף בהצלחה לטבלה המקומית!\nID: {record_id}")
            
        except Exception as e:
            error_msg = str(e)
            self._log_message(f"❌ שגיאה בהוספה: {error_msg}")
            messagebox.showerror("שגיאה", error_msg)
    
    # Window Management
    def _open_drawings_manager(self):
        """פתיחת מנהל הציורים"""
        if not self.drawings_manager_window:
            from .drawings_manager import DrawingsManagerWindow
            self.drawings_manager_window = DrawingsManagerWindow(
                self.root,
                self.data_processor
            )
        
        self.drawings_manager_window.show()

    # ===== טאב מנהל ציורים (חדש כטאב ראשי) =====
    def _create_drawings_manager_tab(self):
        """יוצר טאב חדש לניהול הציורים (במקום חלון נפרד)."""
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="מנהל ציורים")
        self._drawings_tab = tab
        tk.Label(tab, text="מנהל ציורים - טבלה מקומית", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=10)
        actions = tk.Frame(tab, bg='#f7f9fa'); actions.pack(fill='x', padx=12, pady=(0,8))
        # Left buttons group
        left = tk.Frame(actions, bg='#f7f9fa'); left.pack(side='left')
        tk.Button(left, text="🔄 רענן", command=self._refresh_drawings_tree, bg='#3498db', fg='white', font=('Arial',10,'bold'), width=10).pack(side='left', padx=4)
        tk.Button(left, text="📊 ייצא לאקסל", command=self._export_drawings_to_excel_tab, bg='#27ae60', fg='white', font=('Arial',10,'bold'), width=12).pack(side='left', padx=4)
        # Right buttons group
        right = tk.Frame(actions, bg='#f7f9fa'); right.pack(side='right')
        tk.Button(right, text="🗑️ מחק הכל", command=self._clear_all_drawings_tab, bg='#e74c3c', fg='white', font=('Arial',10,'bold'), width=10).pack(side='right', padx=4)
        tk.Button(right, text="❌ מחק נבחר", command=self._delete_selected_drawing_tab, bg='#e67e22', fg='white', font=('Arial',10,'bold'), width=10).pack(side='right', padx=4)
        # Table frame
        table_frame = tk.Frame(tab, bg='#ffffff'); table_frame.pack(fill='both', expand=True, padx=12, pady=8)
        cols = ("id","file_name","created_at","products","total_quantity","status")
        self.drawings_tree = ttk.Treeview(table_frame, columns=cols, show='headings')
        headers = {"id":"ID","file_name":"שם הקובץ","created_at":"תאריך יצירה","products":"מוצרים","total_quantity":"סך כמויות","status":"סטטוס"}
        widths = {"id":70,"file_name":280,"created_at":140,"products":80,"total_quantity":90,"status":90}
        for c in cols:
            self.drawings_tree.heading(c, text=headers[c]); self.drawings_tree.column(c, width=widths[c], anchor='center')
        vs = ttk.Scrollbar(table_frame, orient='vertical', command=self.drawings_tree.yview); self.drawings_tree.configure(yscroll=vs.set)
        self.drawings_tree.grid(row=0,column=0,sticky='nsew'); vs.grid(row=0,column=1,sticky='ns')
        table_frame.grid_columnconfigure(0,weight=1); table_frame.grid_rowconfigure(0,weight=1)
        self.drawings_tree.bind('<Double-1>', self._on_drawings_double_click)
        self.drawings_tree.bind('<Button-3>', self._on_drawings_right_click)
        # תפריט סטטוס ציור
        self._drawing_status_menu = tk.Menu(self.drawings_tree, tearoff=0)
        for st in ("טרם נשלח","נשלח","הוחזר"):
            self._drawing_status_menu.add_command(label=st, command=lambda s=st: self._change_selected_drawing_status(s))
        # Stats bar
        self.drawings_stats_var = tk.StringVar(value="אין נתונים")
        tk.Label(tab, textvariable=self.drawings_stats_var, bg='#34495e', fg='white', anchor='w', padx=10, font=('Arial',10)).pack(fill='x', side='bottom')
        self._populate_drawings_tree(); self._update_drawings_stats()

    def _show_drawings_manager_tab(self):
        """מעבר לטאב מנהל הציורים"""
        # איתור אינדקס הטאב לפי טקסט
        for i in range(len(self.notebook.tabs())):
            if self.notebook.tab(i, 'text') == "מנהל ציורים":
                self.notebook.select(i)
                break

    def _populate_drawings_tree(self):
        if not hasattr(self, 'drawings_tree'):
            return
        for item in self.drawings_tree.get_children():
            self.drawings_tree.delete(item)
        for record in self.data_processor.drawings_data:
            products_count = len(record.get('מוצרים', []))
            total_quantity = record.get('סך כמויות', 0)
            self.drawings_tree.insert('', 'end', values=(
                record.get('id',''),
                record.get('שם הקובץ',''),
                record.get('תאריך יצירה',''),
                products_count,
                f"{total_quantity:.1f}" if isinstance(total_quantity, (int,float)) else total_quantity,
                record.get('status','נשלח')
            ))

    def _update_drawings_stats(self):
        total_drawings = len(self.data_processor.drawings_data)
        total_quantity = sum(r.get('סך כמויות', 0) for r in self.data_processor.drawings_data)
        self.drawings_stats_var.set(f"סך הכל: {total_drawings} ציורים | סך כמויות: {total_quantity:.1f}")

    def _refresh_drawings_tree(self):
        if hasattr(self.data_processor, 'refresh_drawings_data'):
            try:
                self.data_processor.refresh_drawings_data()
            except Exception:
                pass
        self._populate_drawings_tree(); self._update_drawings_stats()

    def _on_drawings_double_click(self, event):
        item_id = self.drawings_tree.focus()
        if not item_id:
            return
        vals = self.drawings_tree.item(item_id, 'values')
        if not vals:
            return
        try:
            rec_id = int(vals[0])
        except Exception:
            return
        record = self.data_processor.get_drawing_by_id(rec_id) if hasattr(self.data_processor, 'get_drawing_by_id') else None
        if not record:
            return
        self._show_drawing_details(record)

    def _on_drawings_right_click(self, event):
        row_id = self.drawings_tree.identify_row(event.y)
        if row_id:
            self.drawings_tree.selection_set(row_id)
        sel = self.drawings_tree.selection()
        if not sel:
            return
        menu = tk.Menu(self.drawings_tree, tearoff=0)
        menu.add_command(label="📋 הצג פרטים", command=lambda: self._on_drawings_double_click(None))
        menu.add_cascade(label="סטטוס", menu=self._drawing_status_menu)
        menu.add_separator()
        menu.add_command(label="🗑️ מחק", command=self._delete_selected_drawing_tab)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _change_selected_drawing_status(self, new_status):
        sel = self.drawings_tree.selection()
        if not sel:
            return
        vals = self.drawings_tree.item(sel[0], 'values')
        if not vals:
            return
        try:
            rec_id = int(vals[0])
        except Exception:
            return
        if hasattr(self.data_processor, 'update_drawing_status') and self.data_processor.update_drawing_status(rec_id, new_status):
            # update row display
            new_vals = list(vals)
            new_vals[-1] = new_status
            self.drawings_tree.item(sel[0], values=new_vals)

    def _show_drawing_details(self, record):
        top = tk.Toplevel(self.root)
        top.title(f"פרטי ציור - {record.get('שם הקובץ','')}")
        top.geometry('900x700')
        top.configure(bg='#f0f0f0')
        tk.Label(top, text=f"פרטי ציור: {record.get('שם הקובץ','')}", font=('Arial',14,'bold'), bg='#f0f0f0').pack(pady=10)
        info = tk.LabelFrame(top, text="מידע כללי", bg='#f0f0f0')
        info.pack(fill='x', padx=12, pady=6)
        txt = (
            f"ID: {record.get('id','')}\n"
            f"תאריך יצירה: {record.get('תאריך יצירה','')}\n"
            f"מספר מוצרים: {len(record.get('מוצרים', []))}\n"
            f"סך הכמויות: {record.get('סך כמויות',0)}"
        )
        tk.Label(info, text=txt, bg='#f0f0f0', justify='left', anchor='w').pack(fill='x', padx=8, pady=6)
        tk.Label(top, text="פירוט מוצרים ומידות:", font=('Arial',12,'bold'), bg='#f0f0f0').pack(anchor='w', padx=12, pady=(6,2))
        st = scrolledtext.ScrolledText(top, height=20, font=('Courier New',10))
        st.pack(fill='both', expand=True, padx=12, pady=4)
        for product in record.get('מוצרים', []):
            st.insert(tk.END, f"\n📦 {product.get('שם המוצר','')}\n")
            st.insert(tk.END, "="*60 + "\n")
            total_prod_q = 0
            for size_info in product.get('מידות', []):
                size = size_info.get('מידה','')
                quantity = size_info.get('כמות',0)
                note = size_info.get('הערה','')
                total_prod_q += quantity
                st.insert(tk.END, f"   מידה {size:>8}: {quantity:>8} - {note}\n")
            st.insert(tk.END, f"\nסך עבור מוצר זה: {total_prod_q}\n")
            st.insert(tk.END, "-"*60 + "\n")
        st.config(state='disabled')
        tk.Button(top, text="סגור", command=top.destroy, bg='#95a5a6', fg='white', font=('Arial',11,'bold'), width=12).pack(pady=10)

    def _delete_selected_drawing_tab(self):
        sel = self.drawings_tree.selection()
        if not sel:
            return
        vals = self.drawings_tree.item(sel[0], 'values')
        if not vals:
            return
        rec_id = vals[0]
        file_name = vals[1]
        if not messagebox.askyesno("אישור מחיקה", f"למחוק את הציור:\n{file_name}? פעולה זו אינה הפיכה"):
            return
        # שימוש במתודה קיימת ב-data_processor אם קיימת
        deleted = False
        if hasattr(self.data_processor, 'delete_drawing'):
            try:
                deleted = self.data_processor.delete_drawing(rec_id)
            except Exception:
                deleted = False
        if not deleted:
            # נפילה אחורה – מחיקה ידנית
            before = len(self.data_processor.drawings_data)
            self.data_processor.drawings_data = [r for r in self.data_processor.drawings_data if str(r.get('id')) != str(rec_id)]
            if len(self.data_processor.drawings_data) < before and hasattr(self.data_processor, 'save_drawings_data'):
                try:
                    self.data_processor.save_drawings_data(); deleted = True
                except Exception:
                    pass
        if deleted:
            self._refresh_drawings_tree()
            messagebox.showinfo("הצלחה", "נמחק בהצלחה")
        else:
            messagebox.showerror("שגיאה", "המחיקה נכשלה")

    def _clear_all_drawings_tab(self):
        if not self.data_processor.drawings_data:
            messagebox.showinfo("מידע", "אין ציורים למחיקה")
            return
        if not messagebox.askyesno("אישור מחיקה", f"למחוק את כל {len(self.data_processor.drawings_data)} הציורים? הפעולה לא ניתנת לשחזור"):
            return
        cleared = False
        if hasattr(self.data_processor, 'clear_all_drawings'):
            try:
                cleared = self.data_processor.clear_all_drawings()
            except Exception:
                cleared = False
        if cleared:
            self._refresh_drawings_tree(); messagebox.showinfo("הצלחה", "כל הציורים נמחקו")
        else:
            messagebox.showerror("שגיאה", "מחיקה נכשלה")

    def _export_drawings_to_excel_tab(self):
        if not self.data_processor.drawings_data:
            messagebox.showwarning("אזהרה", "אין ציורים לייצוא")
            return
        file_path = filedialog.asksaveasfilename(title="ייצא ציורים לאקסל", defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx"),("All files","*.*")])
        if not file_path:
            return
        try:
            self.data_processor.export_drawings_to_excel(file_path)
            messagebox.showinfo("הצלחה", f"הציורים יוצאו אל:\n{file_path}")
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
    
    # Utility Methods
    def _update_status(self, message):
        """עדכון שורת הסטטוס"""
        self.status_label.config(text=message)
        self.root.update()
    
    def _log_message(self, message):
        """הוספת הודעה לאזור התוצאות"""
        self.results_text.insert(tk.END, message + "\n")
        self.results_text.see(tk.END)
        self.root.update()
    
    def _clear_results(self):
        """ניקוי אזור התוצאות"""
        self.results_text.delete(1.0, tk.END)
    
    def _clear_all(self):
        """ניקוי כל הנתונים"""
        self.rib_file = ""
        self.current_results = []
        self.rib_label.config(text="לא נבחר קובץ")
        self._clear_results()
        
        # טעינה מחדש של קובץ מוצרים אם קיים
        if self.settings.get("app.auto_load_products", True):
            products_file = self.settings.get("app.products_file", "קובץ מוצרים.xlsx")
            if os.path.exists(products_file):
                self.products_file = os.path.abspath(products_file)
                self.products_label.config(text=os.path.basename(products_file))
                self._update_status(f"נטען קובץ מוצרים: {os.path.basename(products_file)}")
            else:
                self.products_file = ""
                self.products_label.config(text="לא נבחר קובץ")
                self._update_status("מוכן לעבודה")
        else:
            self._update_status("מוכן לעבודה")
