import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from threading import Thread

class ConverterTabMixin:
    """Mixin המממש את טאב הממיר וכל הפונקציונליות התומכת בו."""
    # ===== Converter Tab =====
    def _create_converter_tab(self):
        """יצירת הטאב הראשי (קיים היסטורית)."""
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="ממיר קבצים")
        self._build_converter_tab_content(tab)

    def _build_converter_tab_content(self, container: tk.Widget):
        """בניית תוכן הממיר בתוך קונטיינר (לשימוש גם כמיני-טאב בתוך מנהל ציורים)."""
        for builder in (
            self._create_files_section,
            self._create_options_section,
            self._create_action_buttons,
            self._create_results_section
        ):
            orig_root = getattr(self, 'root', None)
            try:
                self.root = container
                builder()
            finally:
                if orig_root is not None:
                    self.root = orig_root

    # Sub-sections
    def _create_files_section(self):
        files_frame = ttk.LabelFrame(self.root, text="בחירת קבצים", padding=15)
        files_frame.pack(fill="x", padx=20, pady=10)
        # RIB file
        rib_frame = tk.Frame(files_frame); rib_frame.pack(fill="x", pady=8)
        tk.Label(rib_frame, text="קובץ אופטיטקס:", font=('Arial', 10, 'bold'), width=15, anchor="w").pack(side="left")
        self.rib_label = tk.Label(rib_frame, text="לא נבחר קובץ", bg="white", relief="sunken", width=60, anchor="w", padx=5)
        self.rib_label.pack(side="left", padx=10)
        tk.Button(rib_frame, text="📁 בחר קובץ", command=self._select_rib_file, bg='#3498db', fg='white', font=('Arial', 9, 'bold'), width=12).pack(side="right")
        # Products file
        products_frame = tk.Frame(files_frame); products_frame.pack(fill="x", pady=8)
        tk.Label(products_frame, text="קובץ מוצרים:", font=('Arial', 10, 'bold'), width=15, anchor="w").pack(side="left")
        self.products_label = tk.Label(products_frame, text="לא נבחר קובץ", bg="white", relief="sunken", width=60, anchor="w", padx=5)
        self.products_label.pack(side="left", padx=10)
        tk.Button(products_frame, text="📁 בחר קובץ", command=self._select_products_file, bg='#3498db', fg='white', font=('Arial', 9, 'bold'), width=12).pack(side="right")

    def _create_options_section(self):
        options_frame = ttk.LabelFrame(self.root, text="אפשרויות", padding=15)
        options_frame.pack(fill="x", padx=20, pady=10)
        processing_frame = tk.Frame(options_frame); processing_frame.pack(fill="x", pady=5)
        self.tubular_var = tk.BooleanVar(value=True)
        tk.Checkbutton(processing_frame, text="טיפול אוטומטי ב-Layout Tubular (חלוקה ב-2)", variable=self.tubular_var, font=('Arial', 10)).pack(anchor="w")
        self.only_positive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(processing_frame, text="הצג רק מידות עם כמות גדולה מ-0", variable=self.only_positive_var, font=('Arial', 10)).pack(anchor="w")
        # Fabric type selection
        fabric_type_frame = tk.Frame(options_frame); fabric_type_frame.pack(fill="x", pady=5)
        tk.Label(fabric_type_frame, text="סוג בד:", font=('Arial', 10, 'bold'), width=15, anchor='w').pack(side='left')
        self.fabric_type_options = ["בחר סוג בד", "פלנל לבן", "טריקו לבן", "פלנל מודפס", "טריקו מודפס"]
        self.fabric_type_var = tk.StringVar(value=self.fabric_type_options[0])
        self.fabric_type_combo = ttk.Combobox(
            fabric_type_frame,
            textvariable=self.fabric_type_var,
            values=self.fabric_type_options,
            state='readonly',
            width=20
        )
        self.fabric_type_combo.pack(side='left', padx=5)
        # Supplier recipient selection (למי נשלח הציור)
        supplier_frame = tk.Frame(options_frame); supplier_frame.pack(fill='x', pady=5)
        tk.Label(supplier_frame, text="נמען (ספק):", font=('Arial',10,'bold'), width=15, anchor='w').pack(side='left')
        self.recipient_supplier_var = tk.StringVar()
        self.recipient_supplier_combo = ttk.Combobox(supplier_frame, textvariable=self.recipient_supplier_var, state='readonly', width=30)
        self.recipient_supplier_combo.pack(side='left', padx=5)
        # כפתור רענון שמות ספקים
        tk.Button(supplier_frame, text="↺", width=3, command=self._refresh_converter_suppliers, bg='#3498db', fg='white').pack(side='left', padx=2)
        try:
            # אתחול ראשוני
            self._refresh_converter_suppliers()
        except Exception:
            pass
        # עדכון כפתור הוספה בעת בחירת נמען
        try:
            self.recipient_supplier_combo.bind('<<ComboboxSelected>>', lambda e: self._update_add_local_button_state())
        except Exception:
            pass

    def _create_action_buttons(self):
        buttons_frame = tk.Frame(self.root, bg='#f0f0f0'); buttons_frame.pack(fill="x", padx=20, pady=15)
        row1 = tk.Frame(buttons_frame, bg='#f0f0f0'); row1.pack(fill="x", pady=5)
        tk.Button(row1, text="🔍 נתח קבצים", command=self._analyze_files, bg='#27ae60', fg='white', font=('Arial', 11, 'bold'), height=2, width=15).pack(side="left", padx=5)
        tk.Button(row1, text="💾 שמור כ-Excel", command=self._save_excel, bg='#e67e22', fg='white', font=('Arial', 11, 'bold'), height=2, width=15).pack(side="left", padx=5)
        tk.Button(row1, text="️ נקה הכל", command=self._clear_all, bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'), height=2, width=15).pack(side="right", padx=5)

    def _create_results_section(self):
        results_frame = ttk.LabelFrame(self.root, text="תוצאות וסטטוס", padding=10)
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        # --- Analysis summary info (updates after run) ---
        info_frame = tk.Frame(results_frame, bg='#f7f9fa'); info_frame.pack(fill='x', pady=(0,6))
        self.analysis_info_var = tk.StringVar(value="הרץ ניתוח להצגת נתוני הציור (Tubular, סוג בד, קובץ וכו')")
        tk.Label(info_frame, textvariable=self.analysis_info_var, anchor='e', justify='right', bg='#f7f9fa', fg='#2c3e50', font=('Arial',10,'bold')).pack(fill='x')
        # --- Results table ---
        table_container = tk.Frame(results_frame)
        table_container.pack(fill='both', expand=True)
        cols = ('model','size','quantity')  # order right->left visually in RTL
        self.results_tree = ttk.Treeview(table_container, columns=cols, show='headings', height=14)
        headers = {'model':'דגם','size':'מידה','quantity':'כמות'}
        widths = {'model':200,'size':90,'quantity':90}
        for c in cols:
            self.results_tree.heading(c, text=headers[c])
            # align right for Hebrew readability
            self.results_tree.column(c, width=widths[c], anchor='e', stretch=False)
        vs = ttk.Scrollbar(table_container, orient='vertical', command=self.results_tree.yview)
        self.results_tree.configure(yscroll=vs.set)
        self.results_tree.grid(row=0,column=0,sticky='nsew'); vs.grid(row=0,column=1,sticky='ns')
        table_container.grid_rowconfigure(0,weight=1); table_container.grid_columnconfigure(0,weight=1)
        # --- Log area (small) to preserve existing _log_message usage ---
        self.results_text = scrolledtext.ScrolledText(results_frame, height=6, font=('Consolas', 10), wrap=tk.WORD, bg='#f0f3f5', fg='#2c3e50')
        self.results_text.pack(fill='x', expand=False, pady=(8,0))

    # File Selection
    def _select_rib_file(self):
        file_path = filedialog.askopenfilename(title="בחר קובץ אופטיטקס אקסל אקספורט", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            self.rib_file = file_path
            self.rib_label.config(text=os.path.basename(file_path))
            self._update_status(f"נבחר קובץ אופטיטקס: {os.path.basename(file_path)}")

    def _select_products_file(self):
        file_path = filedialog.askopenfilename(title="בחר קובץ רשימת מוצרים", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            self.products_file = file_path
            self.products_label.config(text=os.path.basename(file_path))
            self._update_status(f"נבחר קובץ מוצרים: {os.path.basename(file_path)}")

    # Analysis
    def _analyze_files(self):
        if not self.rib_file or not self.products_file:
            messagebox.showerror("שגיאה", "אנא בחר את שני הקבצים")
            return
        self._clear_results()
        self._update_status("מנתח קבצים...")
        Thread(target=self._analyze_files_thread, daemon=True).start()

    def _analyze_files_thread(self):
        try:
            self._log_message("=== התחלת ניתוח ===")
            self._log_message("טוען מיפוי מוצרים...")
            if not self.file_analyzer.load_products_mapping(self.products_file):
                raise Exception("שגיאה בטעינת קובץ מוצרים")
            products_count = len(self.file_analyzer.product_mapping)
            self._log_message(f"✅ נטען מיפוי עבור {products_count} מוצרים")
            self._log_message("מנתח קובץ אופטיטקס...")
            results = self.file_analyzer.analyze_file(self.rib_file, self.tubular_var.get(), self.only_positive_var.get())
            if not results:
                self._log_message("❌ לא נמצאו נתונים מתאימים")
                self._update_status("לא נמצאו נתונים")
                return
            self.current_results = self.file_analyzer.sort_results()
            summary = self.file_analyzer.get_analysis_summary()
            self._log_message(f"✅ נוצרה טבלה עם {summary['total_records']} רשומות")
            found_products = self.file_analyzer.get_products_found()
            if found_products:
                self._log_message("\n📦 מוצרים שנמצאו:")
                for file_name, product_name in found_products:
                    self._log_message(f"   {file_name} → {product_name}")
            self._display_detailed_results()
            self._display_statistics(summary)
            self._update_status("הניתוח הושלם בהצלחה!")
        except Exception as e:
            error_msg = f"❌ שגיאה בניתוח: {str(e)}"
            self._log_message(error_msg)
            self._update_status("שגיאה בניתוח")
            messagebox.showerror("שגיאה", str(e))

    def _display_detailed_results(self):
        # Populate the results table instead of verbose text lines
        try:
            for iid in self.results_tree.get_children():
                self.results_tree.delete(iid)
        except Exception:
            pass
        for row in self.current_results:
            self.results_tree.insert('', 'end', values=(row.get('שם המוצר',''), row.get('מידה',''), row.get('כמות','')))
        # Add a concise header note to log area
        self._log_message("=== תוצאות הניתוח (תצוגת טבלה) ===")
        if getattr(self.file_analyzer, 'is_tubular', False):
            self._log_message("(Layout Tubular) הכמויות בטבלה לאחר חלוקה ב-2 מהמקור")

    def _display_statistics(self, summary):
        # Update header info
        try:
            fabric_type = self.fabric_type_var.get() if hasattr(self,'fabric_type_var') else '—'
        except Exception:
            fabric_type = '—'
        file_name = os.path.basename(self.rib_file) if getattr(self,'rib_file','') else '—'
        tubular_txt = 'Tubular (חולק ב-2)' if summary.get('is_tubular') else 'רגיל'
        info = f"סוג בד: {fabric_type} | Layout: {tubular_txt} | קובץ: {file_name} | מוצרים: {summary.get('unique_products')} | מידות: {summary.get('unique_sizes')} | רשומות: {summary.get('total_records')} | סך כמות: {summary.get('total_quantity'): .1f}"
        # הוספת Marker Width/Length אם קיימים
        mw = summary.get('marker_width')
        ml = summary.get('marker_length')
        extra_marker = []
        if mw is not None:
            extra_marker.append(f"Marker Width: {mw:.2f}")
        if ml is not None:
            extra_marker.append(f"Marker Length: {ml:.2f}")
        if extra_marker:
            info += " | " + " | ".join(extra_marker)
        self.analysis_info_var.set(info)
        # Minimal log section
        self._log_message("=== סיכום ניתוח ===")
        if summary.get('is_tubular'):
            self._log_message("(Tubular) הכמויות בטבלה הן לאחר חלוקה ב-2")

    # Export & Local Table
    def _save_excel(self):
        if not self.current_results:
            messagebox.showwarning("אזהרה", "אין נתונים לשמירה. אנא בצע ניתוח תחילה.")
            return
        file_path = filedialog.asksaveasfilename(title="שמור כ-Excel", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            try:
                self.data_processor.export_to_excel(self.current_results, file_path)
                self._log_message(f"📁 הקובץ נשמר בהצלחה: {file_path}")
                self._update_status(f"נשמר: {os.path.basename(file_path)}")
                messagebox.showinfo("הצלחה", "הקובץ נשמר בהצלחה!")
            except Exception as e:
                messagebox.showerror("שגיאה", str(e))

    def _add_to_local_table(self):
        if not self.current_results:
            messagebox.showwarning("אזהרה", "אין נתונים להוספה. אנא בצע ניתוח תחילה.")
            return
        try:
            fabric_type = self.fabric_type_var.get() if hasattr(self, 'fabric_type_var') else ""
            if not fabric_type or (hasattr(self, 'fabric_type_options') and fabric_type == self.fabric_type_options[0]):
                messagebox.showwarning("אזהרה", "אנא בחר סוג בד לפני ההוספה לטבלה המקומית")
                return
            try:
                recipient = self.recipient_supplier_var.get().strip()
            except Exception:
                recipient = ''
            if not recipient:
                messagebox.showwarning("אזהרה", "יש לבחור נמען (ספק) לפני ההוספה לטבלה המקומית")
                return
            record_id = self.data_processor.add_to_local_table(self.current_results, self.rib_file, fabric_type=fabric_type)
            # שמירת הנמען (אם יש יכולת ב- data_processor בעתיד; לעת עתה נשמור בלוג בלבד)
            self._log_message(f"\n✅ הציור נוסף לטבלה המקומית!")
            self._log_message(f"ID רשומה חדשה: {record_id}")
            file_name = os.path.splitext(os.path.basename(self.rib_file))[0] if self.rib_file else 'לא ידוע'
            total_quantity = sum(r['כמות'] for r in self.current_results)
            self._log_message(f"שם הקובץ: {file_name}")
            self._log_message(f"סך כמויות: {total_quantity}")
            if recipient:
                self._log_message(f"נמען (ספק): {recipient}")
            self._update_status("נוסף לטבלה המקומית")
            # Refresh drawings manager tab if exists
            if hasattr(self, '_refresh_drawings_tree'):
                try:
                    self._refresh_drawings_tree()
                except Exception:
                    pass
            messagebox.showinfo("הצלחה", f"הציור נוסף בהצלחה לטבלה המקומית!\nID: {record_id}")
        except Exception as e:
            error_msg = str(e)
            self._log_message(f"❌ שגיאה בהוספה: {error_msg}")
            messagebox.showerror("שגיאה", error_msg)

    def _refresh_converter_suppliers(self):
        """רענון רשימת הספקים עבור קומבובוקס הנמען בממיר."""
        try:
            if not hasattr(self, 'recipient_supplier_combo'):
                return
            names = []
            if hasattr(self, '_get_supplier_names'):
                names = self._get_supplier_names()
            self.recipient_supplier_combo['values'] = names
            # אם הערך הנוכחי כבר לא קיים – ננקה
            cur = self.recipient_supplier_var.get().strip()
            if cur and cur not in names:
                self.recipient_supplier_var.set('')
            self._update_add_local_button_state()
        except Exception:
            pass

    def _update_add_local_button_state(self):
        try:
            if not hasattr(self, 'add_to_local_btn'):
                return
            recipient = ''
            try:
                recipient = self.recipient_supplier_var.get().strip()
            except Exception:
                recipient = ''
            desired = 'normal' if recipient else 'disabled'
            if str(self.add_to_local_btn['state']) != desired:
                self.add_to_local_btn.configure(state=desired)
        except Exception:
            pass

    # Clearing
    def _clear_results(self):
        # Clear table and log
        if hasattr(self, 'results_tree'):
            try:
                self.results_tree.delete(*self.results_tree.get_children())
            except Exception:
                pass
        if hasattr(self, 'results_text'):
            self.results_text.delete(1.0, tk.END)

    def _clear_all(self):
        self.rib_file = ""; self.current_results = []
        self.rib_label.config(text="לא נבחר קובץ")
        self._clear_results()
        if self.settings.get("app.auto_load_products", True):
            products_file = self.settings.get("app.products_file", "קובץ מוצרים.xlsx")
            if os.path.exists(products_file):
                self.products_file = os.path.abspath(products_file)
                self.products_label.config(text=os.path.basename(products_file))
                self._update_status(f"נטען קובץ מוצרים: {os.path.basename(products_file)}")
            else:
                self.products_file = ""; self.products_label.config(text="לא נבחר קובץ"); self._update_status("מוכן לעבודה")
        else:
            self._update_status("מוכן לעבודה")
        # Reset fabric type selection if exists
        if hasattr(self, 'fabric_type_var') and hasattr(self, 'fabric_type_options'):
            self.fabric_type_var.set(self.fabric_type_options[0])
