import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from threading import Thread
from . import theme

class ConverterTabMixin:
    """Mixin המממש את טאב הממיר וכל הפונקציונליות התומכת בו."""
    # ===== Converter Tab =====
    def _create_converter_tab(self):
        """יצירת הטאב הראשי (קיים היסטורית)."""
        tab = tk.Frame(self.notebook, bg=theme.PAGE_BG)
        self.notebook.add(tab, text="ממיר קבצים")
        self._build_converter_tab_content(tab)

    def _build_converter_tab_content(self, container: tk.Widget):
        """בניית תוכן הממיר בתוך קונטיינר (לשימוש גם כמיני-טאב בתוך מנהל ציורים)."""
        # שמירת הרפרנס ל-main window לצורך עדכוני UI מ-threads
        if not hasattr(self, '_main_window'):
            self._main_window = getattr(self, 'root', None)
        
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
        tk.Label(rib_frame, text="קובץ אופטיטקס:", font=(theme.FONT_FAMILY, 10, 'bold'), width=15, anchor="w").pack(side="left")
        self.rib_label = tk.Label(rib_frame, text="לא נבחר קובץ", bg="white", relief="sunken", width=60, anchor="w", padx=5)
        self.rib_label.pack(side="left", padx=10)
        tk.Button(rib_frame, text="📁 בחר קובץ", command=self._select_rib_file, bg=theme.PRIMARY, fg='white', font=(theme.FONT_FAMILY, 9, 'bold'), width=12).pack(side="right")

    def _create_options_section(self):
        options_frame = ttk.LabelFrame(self.root, text="אפשרויות", padding=10)
        options_frame.pack(fill="x", padx=20, pady=5)
        
        # יצירת עמודות - 3 עמודות
        columns_frame = tk.Frame(options_frame)
        columns_frame.pack(fill="x")
        
        # עמודה ראשונה
        col1 = tk.Frame(columns_frame)
        col1.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # עמודה שנייה  
        col2 = tk.Frame(columns_frame)
        col2.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # עמודה שלישית
        col3 = tk.Frame(columns_frame)
        col3.pack(side="left", fill="both", expand=True)
        
        # עמודה 1: עיבוד וסוג בד
        processing_frame = tk.Frame(col1); processing_frame.pack(fill="x", pady=2)
        self.tubular_var = tk.BooleanVar(value=True)
        tk.Checkbutton(processing_frame, text="טיפול אוטומטי ב-Layout Tubular (חלוקה ב-2)", variable=self.tubular_var, font=(theme.FONT_FAMILY, 9)).pack(anchor="w")
        self.only_positive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(processing_frame, text="הצג רק מידות עם כמות גדולה מ-0", variable=self.only_positive_var, font=(theme.FONT_FAMILY, 9)).pack(anchor="w")
        
        # Fabric type selection
        fabric_type_frame = tk.Frame(col1); fabric_type_frame.pack(fill="x", pady=2)
        tk.Label(fabric_type_frame, text="סוג בד:", font=(theme.FONT_FAMILY, 9, 'bold'), width=12, anchor='w').pack(side='left')
        self.fabric_type_options = ["בחר סוג בד", "פלנל לבן", "טריקו לבן", "פלנל מודפס", "טריקו מודפס"]
        self.fabric_type_var = tk.StringVar(value=self.fabric_type_options[0])
        self.fabric_type_combo = ttk.Combobox(
            fabric_type_frame,
            textvariable=self.fabric_type_var,
            values=self.fabric_type_options,
            state='readonly',
            width=15
        )
        self.fabric_type_combo.pack(side='left', padx=5)
        
        # עמודה 2: ספק ושכבות
        supplier_frame = tk.Frame(col2); supplier_frame.pack(fill="x", pady=2)
        tk.Label(supplier_frame, text="נמען (ספק):", font=(theme.FONT_FAMILY,9,'bold'), width=12, anchor='w').pack(side='left')
        self.recipient_supplier_var = tk.StringVar()
        self.recipient_supplier_combo = ttk.Combobox(supplier_frame, textvariable=self.recipient_supplier_var, state='readonly', width=20)
        self.recipient_supplier_combo.pack(side='left', padx=5)
        tk.Button(supplier_frame, text="↺", width=2, command=self._refresh_converter_suppliers, bg=theme.PRIMARY, fg='white').pack(side='left', padx=2)
        
        # כמות שכבות משוערת
        layers_frame = tk.Frame(col2); layers_frame.pack(fill="x", pady=2)
        tk.Label(layers_frame, text="כמות שכבות:", font=(theme.FONT_FAMILY,9,'bold'), width=12, anchor='w').pack(side='left')
        self.estimated_layers_var = tk.StringVar(value='200')
        self.estimated_layers_entry = tk.Entry(layers_frame, textvariable=self.estimated_layers_var, width=8, font=(theme.FONT_FAMILY, 9))
        self.estimated_layers_entry.pack(side='left', padx=5)
        tk.Label(layers_frame, text="(ברירת מחדל: 200)", font=(theme.FONT_FAMILY, 8), fg=theme.SUBTEXT).pack(side='left', padx=2)
        
        # עמודה 3: משקלים
        # משקל בד למטר
        fabric_weight_frame = tk.Frame(col3); fabric_weight_frame.pack(fill="x", pady=2)
        tk.Label(fabric_weight_frame, text="משקל למטר:", font=(theme.FONT_FAMILY,9,'bold'), width=12, anchor='w').pack(side='left')
        self.fabric_weight_per_meter_var = tk.StringVar(value='400')
        self.fabric_weight_per_meter_entry = tk.Entry(fabric_weight_frame, textvariable=self.fabric_weight_per_meter_var, width=8, font=(theme.FONT_FAMILY, 9))
        self.fabric_weight_per_meter_entry.pack(side='left', padx=5)
        tk.Label(fabric_weight_frame, text="גרם", font=(theme.FONT_FAMILY, 8), fg=theme.SUBTEXT).pack(side='left', padx=2)
        
        # משקל כולל (יוצג אחרי ניתוח)
        total_weight_frame = tk.Frame(col3); total_weight_frame.pack(fill="x", pady=2)
        tk.Label(total_weight_frame, text="משקל כולל:", font=(theme.FONT_FAMILY,9,'bold'), width=12, anchor='w').pack(side='left')
        self.total_fabric_weight_var = tk.StringVar(value='לא מחושב')
        self.total_fabric_weight_label = tk.Label(total_weight_frame, textvariable=self.total_fabric_weight_var, font=(theme.FONT_FAMILY, 9, 'bold'), fg=theme.DARK, bg=theme.PANEL_BG, relief='sunken', width=12, anchor='w', padx=3)
        self.total_fabric_weight_label.pack(side='left', padx=5)
        tk.Label(total_weight_frame, text="ק\"ג", font=(theme.FONT_FAMILY, 8), fg=theme.SUBTEXT).pack(side='left', padx=2)
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
        buttons_frame = tk.Frame(self.root, bg=theme.PAGE_BG); buttons_frame.pack(fill="x", padx=20, pady=8)
        row1 = tk.Frame(buttons_frame, bg=theme.PAGE_BG); row1.pack(fill="x", pady=3)
        tk.Button(row1, text="🔍 נתח קבצים", command=self._analyze_files, bg=theme.SUCCESS, fg='white', font=(theme.FONT_FAMILY, 10, 'bold'), height=1, width=15).pack(side="left", padx=3)
        tk.Button(row1, text="💾 שמור כ-Excel", command=self._save_excel, bg=theme.WARNING, fg='white', font=(theme.FONT_FAMILY, 10, 'bold'), height=1, width=15).pack(side="left", padx=3)
        tk.Button(row1, text="️ נקה הכל", command=self._clear_all, bg=theme.DANGER, fg='white', font=(theme.FONT_FAMILY, 10, 'bold'), height=1, width=15).pack(side="right", padx=3)
        # Row 2: add to local drawings table (disabled until recipient selected)
        row2 = tk.Frame(buttons_frame, bg=theme.PAGE_BG); row2.pack(fill='x', pady=(0,3))
        self.add_to_local_btn = tk.Button(row2, text="➕ הוסף לטבלה מקומית", command=self._add_to_local_table, bg=theme.PRIMARY_DARK, fg='white', font=(theme.FONT_FAMILY, 10, 'bold'), height=1, width=20, state='disabled')
        self.add_to_local_btn.pack(side='left', padx=3)
        # נסה לעדכן את מצב הכפתור לפי בחירת הנמען (אם כבר נטען קומבובוקס)
        try:
            self._update_add_local_button_state()
        except Exception:
            pass

    def _create_results_section(self):
        results_frame = ttk.LabelFrame(self.root, text="תוצאות וסטטוס", padding=10)
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        # --- Analysis summary info (updates after run) ---
        info_frame = tk.Frame(results_frame, bg=theme.PAGE_BG); info_frame.pack(fill='x', pady=(0,6))
        self.analysis_info_var = tk.StringVar(value="הרץ ניתוח להצגת נתוני הציור (Tubular, סוג בד, קובץ וכו')")
        tk.Label(info_frame, textvariable=self.analysis_info_var, anchor='e', justify='right', bg=theme.PAGE_BG, fg=theme.DARK, font=(theme.FONT_FAMILY,10,'bold')).pack(fill='x')
        # שורת מידע מודגשת למידות הציור (רוחב/אורך) כדי שיהיו ברורות לעין
        self.marker_info_var = tk.StringVar(value="")
        tk.Label(info_frame, textvariable=self.marker_info_var, anchor='e', justify='right', bg=theme.PANEL_BG, fg=theme.DARK, font=(theme.FONT_FAMILY,11,'bold')).pack(fill='x', pady=(4,0))
        # --- Results table ---
        table_container = tk.Frame(results_frame)
        table_container.pack(fill='both', expand=True, pady=(5,0))
        
        # Create frame for treeview and scrollbar
        tree_frame = tk.Frame(table_container)
        tree_frame.pack(fill='both', expand=True)
        
        cols = ('model','size','quantity')  # order right->left visually in RTL
        self.converter_results_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=22)
        headers = {'model':'דגם','size':'מידה','quantity':'כמות'}
        widths = {'model':200,'size':90,'quantity':90}
        for c in cols:
            self.converter_results_tree.heading(c, text=headers[c])
            # align right for Hebrew readability
            self.converter_results_tree.column(c, width=widths[c], anchor='e', stretch=False)
        
        # Vertical scrollbar
        vs = ttk.Scrollbar(tree_frame, orient='vertical', command=self.converter_results_tree.yview)
        self.converter_results_tree.configure(yscroll=vs.set)
        
        # Horizontal scrollbar
        hs = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.converter_results_tree.xview)
        self.converter_results_tree.configure(xscroll=hs.set)
        
        # Grid layout with proper weights
        self.converter_results_tree.grid(row=0, column=0, sticky='nsew')
        vs.grid(row=0, column=1, sticky='ns')
        hs.grid(row=1, column=0, sticky='ew')
        
        # Configure grid weights
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        # --- Log area (small) to preserve existing _log_message usage ---
        self.results_text = scrolledtext.ScrolledText(results_frame, height=3, font=('Consolas', 9), wrap=tk.WORD, bg=theme.PAGE_BG, fg=theme.DARK)
        self.results_text.pack(fill='x', expand=False, pady=(8,0))

    # File Selection
    def _select_rib_file(self):
        file_path = filedialog.askopenfilename(
            title="בחר קובץ אופטיטקס אקסל אקספורט",
            initialdir=r"C:\Users\levyn\OneDrive\שולחן העבודה\optitex excell reports",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.rib_file = file_path
            self.rib_label.config(text=os.path.basename(file_path))
            self._update_status(f"נבחר קובץ אופטיטקס: {os.path.basename(file_path)}")

    # Analysis
    def _analyze_files(self):
        if not self.rib_file:
            messagebox.showerror("שגיאה", "יש לבחור קובץ אופטיטקס")
            return
        self._clear_results()
        self._update_status("מנתח קבצים...")
        Thread(target=self._analyze_files_thread, daemon=True).start()

    def _safe_log(self, message):
        """כתיבת הודעה ללוג באופן בטוח ל-threads (דרך main thread)"""
        main_window = getattr(self, '_main_window', None) or getattr(self, 'root', None)
        if main_window and hasattr(main_window, 'after'):
            main_window.after(0, lambda msg=message: self._log_message(msg))
        else:
            self._log_message(message)
    
    def _safe_update_status(self, message):
        """עדכון סטטוס באופן בטוח ל-threads (דרך main thread)"""
        main_window = getattr(self, '_main_window', None) or getattr(self, 'root', None)
        if main_window and hasattr(main_window, 'after'):
            main_window.after(0, lambda msg=message: self._update_status(msg))
        else:
            self._update_status(message)

    def _analyze_files_thread(self):
        try:
            self._safe_log("=== התחלת ניתוח ===")
            # יצירת מיפוי מהמילון הפנימי של הטאב 'מיפוי מוצרים'
            self._safe_log("טוען מיפוי מוצרים מהטאב...")
            mapping_rows = getattr(self, '_product_mapping_rows', [])
            internal_map = {}
            internal_unit_quantities = {}
            for r in mapping_rows:
                fn = r.get('file name'); pn = r.get('product name')
                if fn and pn:
                    internal_map[fn] = pn
                    # Get unit quantity (default 1)
                    try:
                        uq = int(r.get('unit quantity', 1))
                        if uq <= 0:
                            uq = 1
                    except (ValueError, TypeError):
                        uq = 1
                    internal_unit_quantities[fn] = uq
            self.file_analyzer.product_mapping = internal_map
            self.file_analyzer.product_unit_quantities = internal_unit_quantities
            if not internal_map:
                self._safe_log("⚠️ אין נתוני מיפוי (הטאב ריק)")
            products_count = len(self.file_analyzer.product_mapping)
            self._safe_log(f"✅ נטען מיפוי עבור {products_count} מוצרים")
            self._safe_log("מנתח קובץ אופטיטקס...")
            results = self.file_analyzer.analyze_file(self.rib_file, self.tubular_var.get(), self.only_positive_var.get())
            if not results:
                self._safe_log("❌ לא נמצאו נתונים מתאימים")
                self._safe_update_status("לא נמצאו נתונים")
                return
            self.current_results = self.file_analyzer.sort_results()
            summary = self.file_analyzer.get_analysis_summary()
            self._safe_log(f"✅ נוצרה טבלה עם {summary['total_records']} רשומות")
            found_products = self.file_analyzer.get_products_found()
            if found_products:
                self._safe_log("\n📦 מוצרים שנמצאו:")
                for file_name, product_name in found_products:
                    self._safe_log(f"   {file_name} → {product_name}")
            # עדכון UI חייב להיעשות ב-main thread - שמירת המשתנים מחוץ ל-lambda
            # כדי למנוע בעיות של late binding
            summary_to_display = summary
            # קבלת הרפרנס ל-main window (לא ל-container הזמני)
            main_window = getattr(self, '_main_window', None) or getattr(self, 'root', None)
            if main_window and hasattr(main_window, 'after'):
                main_window.after(0, lambda: self._display_detailed_results())
                main_window.after(0, lambda s=summary_to_display: self._display_statistics(s))
                main_window.after(0, lambda s=summary_to_display: self._calculate_total_fabric_weight(s))
                main_window.after(0, lambda: self._update_status("הניתוח הושלם בהצלחה!"))
            else:
                # fallback - קריאה ישירה (אם כבר במקרה ב-main thread)
                self._display_detailed_results()
                self._display_statistics(summary)
                self._calculate_total_fabric_weight(summary)
                self._update_status("הניתוח הושלם בהצלחה!")
        except Exception as e:
            error_msg = f"❌ שגיאה בניתוח: {str(e)}"
            self._safe_log(error_msg)
            self._safe_update_status("שגיאה בניתוח")
            # messagebox גם צריך להיות ב-main thread
            main_window = getattr(self, '_main_window', None) or getattr(self, 'root', None)
            if main_window and hasattr(main_window, 'after'):
                main_window.after(0, lambda err=str(e): messagebox.showerror("שגיאה", err))
            else:
                messagebox.showerror("שגיאה", str(e))

    def _display_detailed_results(self):
        """Populate the results table with analysis data."""
        if not hasattr(self, 'converter_results_tree'):
            self._log_message("שגיאה: טבלת התוצאות לא אותחלה.")
            return
        
        if not hasattr(self, 'current_results') or not self.current_results:
            self._log_message("שגיאה: אין נתונים להצגה בטבלה.")
            return
        
        # Clear existing items
        for iid in self.converter_results_tree.get_children():
            self.converter_results_tree.delete(iid)
        
        # Add new items
        for row in self.current_results:
            product = row.get('שם המוצר', '')
            size = row.get('מידה', '')
            quantity = row.get('כמות', '')
            self.converter_results_tree.insert('', 'end', values=(product, size, quantity))
        
        # Add a concise header note to log area
        self._log_message(f"=== תוצאות הניתוח ({len(self.current_results)} רשומות) ===")
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
        # הוספת נתוני מידות ציור (Marker) אם קיימים
        mw = summary.get('marker_width')
        ml = summary.get('marker_length')
        extra_marker = []
        if mw is not None:
            extra_marker.append(f"רוחב ציור: {mw:.2f} ס\"מ")
        if ml is not None:
            # המרת סנטימטרים למטרים
            ml_meters = ml / 100
            extra_marker.append(f"אורך ציור: {ml_meters:.2f} מטרים")
        if extra_marker:
            info += " | " + " | ".join(extra_marker)
        self.analysis_info_var.set(info)
        # הצגה מודגשת בשורה נפרדת עם יחידות מתאימות
        try:
            if hasattr(self, 'marker_info_var'):
                if (mw is not None) or (ml is not None):
                    parts = []
                    if ml is not None:
                        # המרת סנטימטרים למטרים
                        ml_meters = ml / 100
                        parts.append(f"אורך ציור: {ml_meters:.2f} מטרים")
                    if mw is not None:
                        parts.append(f"רוחב ציור: {mw:.2f} ס\"מ")
                    self.marker_info_var.set(" | ".join(parts))
                else:
                    self.marker_info_var.set("")
        except Exception:
            pass
        # Minimal log section
        self._log_message("=== סיכום ניתוח ===")
        if summary.get('is_tubular'):
            self._log_message("(Tubular) הכמויות בטבלה הן לאחר חלוקה ב-2")
        # רישום ביומן גם של מידות הציור
        if mw is not None:
            self._log_message(f"רוחב ציור: {mw:.2f} ס\"מ")
        if ml is not None:
            # המרת סנטימטרים למטרים
            ml_meters = ml / 100
            self._log_message(f"אורך ציור: {ml_meters:.2f} מטרים")

    def _calculate_total_fabric_weight(self, summary):
        """חישוב משקל בד כולל על בסיס אורך הציור, משקל למטר וכמות שכבות."""
        try:
            # קבלת משקל בד למטר
            try:
                weight_per_meter = float(self.fabric_weight_per_meter_var.get())
            except (ValueError, AttributeError):
                weight_per_meter = 400.0  # ברירת מחדל
            
            # קבלת אורך הציור
            marker_length = summary.get('marker_length')
            if marker_length is None:
                self.total_fabric_weight_var.set('אורך לא זמין')
                return
            
            # קבלת כמות שכבות
            try:
                estimated_layers = int(self.estimated_layers_var.get())
                if estimated_layers <= 0:
                    estimated_layers = 200  # ברירת מחדל
            except (ValueError, AttributeError):
                estimated_layers = 200  # ברירת מחדל
            
            # חישוב המשקל הכולל
            # משקל כולל = משקל למטר × אורך לשכבה × מספר שכבות
            # אורך לשכבה = אורך הציור (בס"מ) / 100 (להמרה למטר)
            length_per_layer_meters = marker_length / 100.0
            total_weight_grams = weight_per_meter * length_per_layer_meters * estimated_layers
            
            # המרה לק"ג
            total_weight_kg = total_weight_grams / 1000.0
            
            # הצגת התוצאה
            self.total_fabric_weight_var.set(f"{total_weight_kg:.2f}")
            
            # הוספה ללוג
            self._log_message(f"\n📏 חישוב משקל בד:")
            self._log_message(f"   משקל למטר: {weight_per_meter} גרם")
            self._log_message(f"   אורך ציור: {marker_length:.2f} ס\"מ")
            self._log_message(f"   אורך לשכבה: {length_per_layer_meters:.2f} מטר")
            self._log_message(f"   כמות שכבות: {estimated_layers}")
            self._log_message(f"   משקל כולל: {total_weight_kg:.2f} ק\"ג")
            
        except Exception as e:
            self.total_fabric_weight_var.set('שגיאה בחישוב')
            self._log_message(f"❌ שגיאה בחישוב משקל: {str(e)}")

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
            # קבלת כמות שכבות משוערת
            try:
                estimated_layers = int(self.estimated_layers_var.get())
                if estimated_layers <= 0:
                    estimated_layers = 200  # ברירת מחדל
            except (ValueError, AttributeError):
                estimated_layers = 200  # ברירת מחדל
            
            # קבלת נתוני מידות ציור מהניתוח
            marker_width = None
            marker_length = None
            efficiency = None
            if hasattr(self.file_analyzer, 'marker_width'):
                marker_width = self.file_analyzer.marker_width
            if hasattr(self.file_analyzer, 'marker_length'):
                marker_length = self.file_analyzer.marker_length
            if hasattr(self.file_analyzer, 'efficiency'):
                efficiency = self.file_analyzer.efficiency
            
            record_id = self.data_processor.add_to_local_table(
                self.current_results, 
                self.rib_file, 
                fabric_type=fabric_type, 
                recipient_supplier=recipient,
                estimated_layers=estimated_layers,
                marker_width=marker_width,
                marker_length=marker_length,
                efficiency=efficiency
            )
            # שמירת הנמען (אם יש יכולת ב- data_processor בעתיד; לעת עתה נשמור בלוג בלבד)
            self._log_message(f"\n✅ הציור נוסף לטבלה המקומית!")
            self._log_message(f"ID רשומה חדשה: {record_id}")
            file_name = os.path.splitext(os.path.basename(self.rib_file))[0] if self.rib_file else 'לא ידוע'
            total_quantity = sum(r['כמות'] for r in self.current_results)
            self._log_message(f"שם הקובץ: {file_name}")
            self._log_message(f"סך כמויות: {total_quantity}")
            if recipient:
                self._log_message(f"נמען (ספק): {recipient}")
            self._log_message(f"כמות שכבות משוערת: {estimated_layers}")
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
                self.converter_results_tree.delete(*self.converter_results_tree.get_children())
            except Exception:
                pass
        if hasattr(self, 'results_text'):
            self.results_text.delete(1.0, tk.END)
        # Clear total weight calculation
        if hasattr(self, 'total_fabric_weight_var'):
            self.total_fabric_weight_var.set('לא מחושב')

    def _clear_all(self):
        self.rib_file = ""; self.current_results = []
        self.rib_label.config(text="לא נבחר קובץ")
        self._clear_results()
        self._update_status("מוכן לעבודה")
        # Reset fabric type selection if exists
        if hasattr(self, 'fabric_type_var') and hasattr(self, 'fabric_type_options'):
            self.fabric_type_var.set(self.fabric_type_options[0])
