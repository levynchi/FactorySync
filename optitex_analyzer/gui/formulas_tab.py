"""Formulas and calculations tab for weight and measurements calculations."""
import tkinter as tk
from tkinter import ttk, messagebox

class FormulasTabMixin:
    """Mixin for formulas and calculations tab."""
    
    def _create_formulas_tab(self):
        """Create the formulas and calculations tab with sub-tabs."""
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="נוסחאות וחישובים")
        
        # Title
        title_label = tk.Label(
            tab, 
            text="נוסחאות וחישובים", 
            font=('Arial', 16, 'bold'), 
            bg='#f7f9fa', 
            fg='#2c3e50'
        )
        title_label.pack(pady=(10, 20))
        
        # Create inner notebook for sub-tabs
        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Sub-tabs
        weight_calc_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        fabric_weight_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        product_cost_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        tetra_cost_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        all_over_print_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        store_price_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        fabric_rolls_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        sqm_cost_tab = tk.Frame(inner_nb, bg='#f7f9fa')
        
        inner_nb.add(weight_calc_tab, text="חישובי משקל כללי")
        inner_nb.add(fabric_weight_tab, text="משקל בד לפריטים בציור")
        inner_nb.add(product_cost_tab, text="חישוב שמיכות")
        inner_nb.add(tetra_cost_tab, text="חישוב טטרות")
        inner_nb.add(all_over_print_tab, text="בגדי אול אובר")
        inner_nb.add(store_price_tab, text="מחיר לצרכן חנויות")
        inner_nb.add(fabric_rolls_tab, text="חישוב גלילי בד")
        inner_nb.add(sqm_cost_tab, text="עלות למ\"ר")
        
        # Build content for each sub-tab
        self._build_general_weight_content(weight_calc_tab)
        self._build_fabric_weight_content(fabric_weight_tab)
        self._build_product_cost_content(product_cost_tab)
        self._build_tetra_cost_content(tetra_cost_tab)
        self._build_all_over_print_content(all_over_print_tab)
        self._build_store_price_content(store_price_tab)
        self._build_fabric_rolls_content(fabric_rolls_tab)
        self._build_sqm_cost_content(sqm_cost_tab)
    
    def _build_general_weight_content(self, container):
        """Build the general weight calculations content."""
        
        # Main formula frame
        formula_frame = ttk.LabelFrame(container, text="נוסחה כללית לחישוב משקל", padding=20)
        formula_frame.pack(fill='x', padx=20, pady=10)
        
        # Formula display
        formula_text = "משקל כולל = אורך (במטר) × משקל למטר × מספר שכבות"
        formula_label = tk.Label(
            formula_frame,
            text=formula_text,
            font=('Arial', 14, 'bold'),
            bg='#f7f9fa',
            fg='#2c3e50',
            justify='center'
        )
        formula_label.pack(pady=10)
        
        # Instructions
        instructions_text = "הוראות: הזן 2 ערכים ולחץ על הכפתור המתאים לחישוב הערך השלישי"
        instructions_label = tk.Label(
            formula_frame,
            text=instructions_text,
            font=('Arial', 10),
            bg='#f7f9fa',
            fg='#7f8c8d',
            justify='center'
        )
        instructions_label.pack(pady=5)
        
        # Input fields frame
        inputs_frame = ttk.LabelFrame(container, text="חישוב משקל", padding=20)
        inputs_frame.pack(fill='x', padx=20, pady=10)
        
        # Input fields
        tk.Label(inputs_frame, text="אורך (במטר):", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.length_var = tk.StringVar()
        length_entry = tk.Entry(inputs_frame, textvariable=self.length_var, width=15, font=('Arial', 10))
        length_entry.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(inputs_frame, text="משקל למטר (ק״ג):", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.weight_per_meter_var = tk.StringVar()
        weight_per_meter_entry = tk.Entry(inputs_frame, textvariable=self.weight_per_meter_var, width=15, font=('Arial', 10))
        weight_per_meter_entry.grid(row=0, column=3, padx=5, pady=5)
        
        tk.Label(inputs_frame, text="מספר שכבות:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.layers_var = tk.StringVar()
        layers_entry = tk.Entry(inputs_frame, textvariable=self.layers_var, width=15, font=('Arial', 10))
        layers_entry.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(inputs_frame, text="משקל כולל (ק״ג):", font=('Arial', 10, 'bold')).grid(row=1, column=2, sticky='w', padx=5, pady=5)
        self.total_weight_var = tk.StringVar()
        total_weight_entry = tk.Entry(inputs_frame, textvariable=self.total_weight_var, width=15, font=('Arial', 10))
        total_weight_entry.grid(row=1, column=3, padx=5, pady=5)
        
        # Calculate buttons
        calculate_btn = tk.Button(
            inputs_frame,
            text="חשב משקל כולל",
            command=self._calculate_total_weight,
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        )
        calculate_btn.grid(row=2, column=0, padx=5, pady=5)
        
        calculate_layers_btn = tk.Button(
            inputs_frame,
            text="חשב שכבות",
            command=self._calculate_layers,
            bg='#e67e22',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        )
        calculate_layers_btn.grid(row=2, column=1, padx=5, pady=5)
        
        calculate_weight_per_meter_btn = tk.Button(
            inputs_frame,
            text="חשב משקל למטר",
            command=self._calculate_weight_per_meter,
            bg='#9b59b6',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        )
        calculate_weight_per_meter_btn.grid(row=2, column=2, padx=5, pady=5)
        
        # Result display
        result_frame = ttk.LabelFrame(container, text="תוצאת החישוב", padding=20)
        result_frame.pack(fill='x', padx=20, pady=10)
        
        self.result_var = tk.StringVar(value="הזן ערכים לחישוב")
        result_label = tk.Label(
            result_frame,
            textvariable=self.result_var,
            font=('Arial', 12, 'bold'),
            bg='#f7f9fa',
            fg='#27ae60',
            justify='center'
        )
        result_label.pack(pady=10)
        
        # Clear button
        clear_btn = tk.Button(
            container,
            text="נקה הכל",
            command=self._clear_formula_inputs,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        )
        clear_btn.pack(pady=20)
    
    def _calculate_total_weight(self):
        """Calculate total weight using the formula."""
        try:
            length = float(self.length_var.get() or 0)
            weight_per_meter = float(self.weight_per_meter_var.get() or 0)
            layers = float(self.layers_var.get() or 0)
            
            if length <= 0 or weight_per_meter <= 0 or layers <= 0:
                self.result_var.set("אנא הזן ערכים חיוביים")
                return
            
            total_weight = length * weight_per_meter * layers
            self.total_weight_var.set(f"{total_weight:.2f}")
            
            result_text = f"משקל כולל: {total_weight:.2f} ק״ג"
            self.result_var.set(result_text)
            
        except ValueError:
            self.result_var.set("אנא הזן מספרים תקינים")
        except Exception as e:
            self.result_var.set(f"שגיאה בחישוב: {str(e)}")
    
    def _calculate_layers(self):
        """Calculate number of layers."""
        try:
            length = float(self.length_var.get() or 0)
            weight_per_meter = float(self.weight_per_meter_var.get() or 0)
            total_weight = float(self.total_weight_var.get() or 0)
            
            if length <= 0 or weight_per_meter <= 0 or total_weight <= 0:
                self.result_var.set("אנא הזן ערכים חיוביים")
                return
            
            layers = total_weight / (length * weight_per_meter)
            self.layers_var.set(f"{layers:.2f}")
            
            result_text = f"מספר שכבות: {layers:.2f}"
            self.result_var.set(result_text)
            
        except ValueError:
            self.result_var.set("אנא הזן מספרים תקינים")
        except Exception as e:
            self.result_var.set(f"שגיאה בחישוב: {str(e)}")
    
    def _calculate_weight_per_meter(self):
        """Calculate weight per meter."""
        try:
            length = float(self.length_var.get() or 0)
            layers = float(self.layers_var.get() or 0)
            total_weight = float(self.total_weight_var.get() or 0)
            
            if length <= 0 or layers <= 0 or total_weight <= 0:
                self.result_var.set("אנא הזן ערכים חיוביים")
                return
            
            weight_per_meter = total_weight / (length * layers)
            self.weight_per_meter_var.set(f"{weight_per_meter:.2f}")
            
            result_text = f"משקל למטר: {weight_per_meter:.2f} ק״ג"
            self.result_var.set(result_text)
            
        except ValueError:
            self.result_var.set("אנא הזן מספרים תקינים")
        except Exception as e:
            self.result_var.set(f"שגיאה בחישוב: {str(e)}")
    
    def _clear_formula_inputs(self):
        """Clear all input fields."""
        self.length_var.set("")
        self.weight_per_meter_var.set("")
        self.layers_var.set("")
        self.total_weight_var.set("")
        self.result_var.set("הזן ערכים לחישוב")
    
    def _build_fabric_weight_content(self, container):
        """Build the fabric weight calculation for drawings content."""
        
        # Create scrollable container
        canvas = tk.Canvas(container, bg='#f7f9fa', highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#f7f9fa')
        
        # Create window and store the window ID
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # Configure scroll region when frame changes
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Make the frame expand to canvas width
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # Bind mousewheel to scroll
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Filters frame at the top
        filters_frame = ttk.LabelFrame(scrollable_frame, text="סינון ציורים", padding=10)
        filters_frame.pack(fill='x', padx=20, pady=(10, 5))
        
        # Row 1 - Supplier and Fabric Type filters
        tk.Label(filters_frame, text="ספק:", font=('Arial', 9, 'bold')).grid(row=0, column=0, sticky='e', padx=5, pady=3)
        self.filter_supplier_var = tk.StringVar(value="הכל")
        self.filter_supplier_combo = ttk.Combobox(filters_frame, textvariable=self.filter_supplier_var, 
                                                  state='readonly', width=20)
        self.filter_supplier_combo.grid(row=0, column=1, padx=5, pady=3, sticky='w')
        
        tk.Label(filters_frame, text="סוג בד:", font=('Arial', 9, 'bold')).grid(row=0, column=2, sticky='e', padx=5, pady=3)
        self.filter_fabric_var = tk.StringVar(value="הכל")
        self.filter_fabric_combo = ttk.Combobox(filters_frame, textvariable=self.filter_fabric_var, 
                                                state='readonly', width=20)
        self.filter_fabric_combo.grid(row=0, column=3, padx=5, pady=3, sticky='w')
        
        # Row 2 - Product name filter
        tk.Label(filters_frame, text="שם מוצר:", font=('Arial', 9, 'bold')).grid(row=1, column=0, sticky='e', padx=5, pady=3)
        self.filter_product_var = tk.StringVar(value="הכל")
        self.filter_product_combo = ttk.Combobox(filters_frame, textvariable=self.filter_product_var, 
                                                 state='readonly', width=20)
        self.filter_product_combo.grid(row=1, column=1, padx=5, pady=3, sticky='w')
        
        # Filter buttons
        filter_btn = tk.Button(filters_frame, text="🔍 סנן", command=self._apply_drawing_filters,
                              bg='#27ae60', fg='white', font=('Arial', 9, 'bold'))
        filter_btn.grid(row=1, column=2, padx=5, pady=3)
        
        clear_filter_btn = tk.Button(filters_frame, text="🗑️ נקה סינון", command=self._clear_drawing_filters,
                                    bg='#95a5a6', fg='white', font=('Arial', 9, 'bold'))
        clear_filter_btn.grid(row=1, column=3, padx=5, pady=3, sticky='w')
        
        # Combined frame for three columns: drawing selection, info, and instructions
        main_frame = tk.Frame(scrollable_frame)
        main_frame.pack(fill='x', padx=20, pady=5)
        
        # Left column - Drawing selection frame
        drawing_frame = ttk.LabelFrame(main_frame, text="בחירת ציור", padding=15)
        drawing_frame.pack(side='left', fill='both', expand=True, padx=(0, 7))
        
        # Drawing selection
        tk.Label(drawing_frame, text="בחר ציור:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.selected_drawing_var = tk.StringVar()
        self.drawing_combobox = ttk.Combobox(drawing_frame, textvariable=self.selected_drawing_var, 
                                           state='readonly', width=35, justify='right')
        self.drawing_combobox.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.drawing_combobox.bind('<<ComboboxSelected>>', self._on_drawing_selected)
        
        # Drawing count label
        self.drawing_count_var = tk.StringVar(value="")
        tk.Label(drawing_frame, textvariable=self.drawing_count_var, font=('Arial', 8), fg='#7f8c8d').grid(
            row=1, column=0, columnspan=2, padx=5, pady=2, sticky='w')
        
        # Load drawings button
        load_btn = tk.Button(drawing_frame, text="🔄 טען ציורים", command=self._load_drawings_list,
                           bg='#3498db', fg='white', font=('Arial', 9, 'bold'))
        load_btn.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        
        # Configure grid weights for drawing frame
        drawing_frame.grid_columnconfigure(1, weight=1)
        
        # Middle column - Drawing info frame
        info_frame = ttk.LabelFrame(main_frame, text="פרטי הציור", padding=15)
        info_frame.pack(side='left', fill='both', expand=True, padx=(7, 7))
        
        self.drawing_info_text = tk.Text(info_frame, height=6, width=30, font=('Arial', 9), 
                                       state='disabled', wrap='word')
        info_scrollbar = ttk.Scrollbar(info_frame, orient='vertical', command=self.drawing_info_text.yview)
        self.drawing_info_text.configure(yscrollcommand=info_scrollbar.set)
        self.drawing_info_text.pack(side='left', fill='both', expand=True)
        info_scrollbar.pack(side='right', fill='y')
        
        # Right column - Instructions frame
        instructions_frame = ttk.LabelFrame(main_frame, text="הוראות שימוש", padding=15)
        instructions_frame.pack(side='left', fill='both', expand=True, padx=(7, 0))
        
        instructions_text = ("בחר ציור מהרשימה לחישוב חלוקת משקל בד לפי שטח רבוע של הפריטים.\n\n"
                           "עבור ציורים בסטטוס 'נחתך' - כמות השכבות והמשקל ימולאו אוטומטית.\n\n"
                           "הנוסחה:\n"
                           "%ᵢ = (Aᵢ/ΣA) × 100\n"
                           "Gramsᵢ = W × (%ᵢ/100)\n\n"
                           "כאשר:\n"
                           "W = משקל השכבה\n"
                           "Aᵢ = שטח רבוע למידה i\n"
                           "ΣA = סכום כל השטחים")
        instructions_label = tk.Label(
            instructions_frame,
            text=instructions_text,
            font=('Arial', 8),
            fg='#7f8c8d',
            justify='right',
            wraplength=250
        )
        instructions_label.pack(pady=5, fill='both', expand=True)
        
        # Weight input frame
        weight_frame = ttk.LabelFrame(scrollable_frame, text="נתוני הפריסה", padding=15)
        weight_frame.pack(fill='x', padx=20, pady=10)
        
        # Layer count row
        tk.Label(weight_frame, text="כמות שכבות שנחתכו:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.layers_count_var = tk.StringVar()
        layers_entry = tk.Entry(weight_frame, textvariable=self.layers_count_var, width=15, font=('Arial', 10))
        layers_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Weight row
        tk.Label(weight_frame, text="משקל השכבה הכולל (בגרמים):", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.layer_weight_var = tk.StringVar()
        weight_entry = tk.Entry(weight_frame, textvariable=self.layer_weight_var, width=15, font=('Arial', 10))
        weight_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Price per kg row
        tk.Label(weight_frame, text="מחיר ל-1 ק\"ג בד (₪):", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.fabric_weight_price_var = tk.StringVar()
        price_entry = tk.Entry(weight_frame, textvariable=self.fabric_weight_price_var, width=15, font=('Arial', 10))
        price_entry.grid(row=2, column=1, padx=5, pady=5)
        
        # Calculate button
        calc_btn = tk.Button(weight_frame, text="חשב חלוקת משקל", command=self._calculate_weight_distribution,
                           bg='#27ae60', fg='white', font=('Arial', 10, 'bold'))
        calc_btn.grid(row=2, column=2, padx=20, pady=5)
        
        # Export button
        export_btn = tk.Button(weight_frame, text="ייצא ל-Excel", command=self._export_weight_results,
                             bg='#2c3e50', fg='white', font=('Arial', 10, 'bold'))
        export_btn.grid(row=2, column=3, padx=5, pady=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(scrollable_frame, text="תוצאות החישוב", padding=15)
        results_frame.pack(fill='x', padx=20, pady=10)
        
        # Results table - increased height
        cols = ('product_name', 'size', 'quantity', 'square_area', 'percentage', 'weight_total', 'weight_per_unit', 'cost_total', 'cost_per_unit')
        self.results_tree = ttk.Treeview(results_frame, columns=cols, show='headings', height=20)
        
        headers = {
            'product_name': 'שם המוצר',
            'size': 'מידה',
            'quantity': 'כמות',
            'square_area': 'שטח רבוע (מ״ר)',
            'percentage': 'אחוז (%)',
            'weight_total': 'משקל כולל (גרמים)',
            'weight_per_unit': 'משקל ליחידה (גרמים)',
            'cost_total': 'עלות כוללת (₪)',
            'cost_per_unit': 'עלות ליחידה (₪)'
        }
        
        widths = {
            'product_name': 120,
            'size': 50,
            'quantity': 50,
            'square_area': 90,
            'percentage': 65,
            'weight_total': 100,
            'weight_per_unit': 110,
            'cost_total': 85,
            'cost_per_unit': 90
        }
        
        for col in cols:
            self.results_tree.heading(col, text=headers[col])
            self.results_tree.column(col, width=widths[col], anchor='center')
        
        # Scrollbar for results table
        results_scrollbar = ttk.Scrollbar(results_frame, orient='vertical', command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=results_scrollbar.set)
        self.results_tree.pack(side='left', fill='both', expand=True)
        results_scrollbar.pack(side='right', fill='y')
        
        # Summary frame
        summary_frame = ttk.LabelFrame(scrollable_frame, text="סיכום", padding=10)
        summary_frame.pack(fill='x', padx=20, pady=5)
        
        self.summary_var = tk.StringVar(value="בחר ציור והזן משקל לחישוב")
        summary_label = tk.Label(summary_frame, textvariable=self.summary_var, 
                               font=('Arial', 11, 'bold'), fg='#2c3e50')
        summary_label.pack(pady=5)
        
        # Initialize
        self._load_drawings_list()
    
    def _load_drawings_list(self):
        """Load the list of drawings from the database and populate filter options."""
        try:
            # Get drawings from data processor
            drawings = getattr(self.data_processor, 'drawings_data', [])
            
            # Build filter options from all drawings
            suppliers = set(["הכל"])
            fabric_types = set(["הכל"])
            product_names = set(["הכל"])
            
            for drawing in drawings:
                # Get supplier
                supplier = drawing.get('נמען (ספק)', drawing.get('ספק', ''))
                if supplier:
                    suppliers.add(supplier)
                
                # Get fabric type
                fabric = drawing.get('סוג בד', '')
                if fabric:
                    fabric_types.add(fabric)
                
                # Get product names from this drawing
                products = drawing.get('מוצרים', [])
                for product in products:
                    product_name = product.get('שם המוצר', '')
                    if product_name:
                        product_names.add(product_name)
            
            # Update filter comboboxes
            self.filter_supplier_combo['values'] = sorted(list(suppliers), key=lambda x: (x != "הכל", x))
            self.filter_fabric_combo['values'] = sorted(list(fabric_types), key=lambda x: (x != "הכל", x))
            self.filter_product_combo['values'] = sorted(list(product_names), key=lambda x: (x != "הכל", x))
            
            # Apply filters and load drawings
            self._apply_drawing_filters()
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בטעינת רשימת הציורים: {str(e)}")
    
    def _apply_drawing_filters(self):
        """Apply filters and update the drawings list."""
        try:
            drawings = getattr(self.data_processor, 'drawings_data', [])
            
            # Get filter values
            supplier_filter = self.filter_supplier_var.get()
            fabric_filter = self.filter_fabric_var.get()
            product_filter = self.filter_product_var.get()
            
            # Filter drawings
            filtered_drawings = []
            for drawing in drawings:
                # Check supplier filter
                if supplier_filter != "הכל":
                    supplier = drawing.get('נמען (ספק)', drawing.get('ספק', ''))
                    if supplier != supplier_filter:
                        continue
                
                # Check fabric type filter
                if fabric_filter != "הכל":
                    fabric = drawing.get('סוג בד', '')
                    if fabric != fabric_filter:
                        continue
                
                # Check product name filter
                if product_filter != "הכל":
                    products = drawing.get('מוצרים', [])
                    product_names = [p.get('שם המוצר', '') for p in products]
                    if product_filter not in product_names:
                        continue
                
                filtered_drawings.append(drawing)
            
            # Build display list with ID + file name
            drawing_display_names = []
            self.drawings_dict = {}
            
            for drawing in filtered_drawings:
                drawing_id = drawing.get('id', '?')
                file_name = drawing.get('שם הקובץ', 'לא ידוע')
                display_name = f"ID: {drawing_id} - {file_name}"
                drawing_display_names.append(display_name)
                self.drawings_dict[display_name] = drawing
            
            # Update combobox
            self.drawing_combobox['values'] = drawing_display_names
            
            # Update count label
            total_count = len(drawings)
            filtered_count = len(filtered_drawings)
            self.drawing_count_var.set(f"מציג {filtered_count} מתוך {total_count} ציורים")
            
            if drawing_display_names:
                self.drawing_combobox.set(drawing_display_names[0])
                self._on_drawing_selected()
            else:
                self.drawing_combobox.set('')
                self.drawing_info_text.config(state='normal')
                self.drawing_info_text.delete(1.0, tk.END)
                self.drawing_info_text.insert(tk.END, "לא נמצאו ציורים התואמים לסינון")
                self.drawing_info_text.config(state='disabled')
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בסינון ציורים: {str(e)}")
    
    def _clear_drawing_filters(self):
        """Clear all filters and reload drawings."""
        self.filter_supplier_var.set("הכל")
        self.filter_fabric_var.set("הכל")
        self.filter_product_var.set("הכל")
        self._apply_drawing_filters()
    
    def _on_drawing_selected(self, event=None):
        """Handle drawing selection."""
        try:
            selected_name = self.selected_drawing_var.get()
            if not selected_name or selected_name not in self.drawings_dict:
                return
                
            drawing = self.drawings_dict[selected_name]
            
            # Display drawing info
            self.drawing_info_text.config(state='normal')
            self.drawing_info_text.delete(1.0, tk.END)
            
            info_text = f"שם הקובץ: {drawing.get('שם הקובץ', 'לא ידוע')}\n"
            info_text += f"תאריך יצירה: {drawing.get('תאריך יצירה', 'לא ידוע')}\n"
            info_text += f"ID: {drawing.get('id', 'לא ידוע')}\n"
            
            # Add status information
            status = drawing.get('status', 'לא ידוע')
            info_text += f"סטטוס: {status}\n"
            
            # Add layers and weight info if available
            layers = drawing.get('שכבות') or drawing.get('כמות שכבות משוערת')
            if layers is not None:
                layer_type = "שכבות" if drawing.get('שכבות') else "שכבות משוערת"
                info_text += f"כמות {layer_type}: {layers}\n"
            
            total_weight = drawing.get('משקל כולל') or drawing.get('משקל_כולל_נגזר')
            if total_weight is not None:
                weight_type = "משקל כולל" if drawing.get('משקל כולל') else "משקל כולל נגזר"
                info_text += f"{weight_type}: {total_weight} ק\"ג\n"
            
            info_text += "\n"
            
            products = drawing.get('מוצרים', [])
            info_text += f"מוצרים בציור ({len(products)} סוגים):\n"
            
            for product in products:
                product_name = product.get('שם המוצר', 'לא ידוע')
                sizes = product.get('מידות', [])
                info_text += f"• {product_name}: {len(sizes)} מידות\n"
                
                for size_info in sizes:
                    size = size_info.get('מידה', 'לא ידוע')
                    quantity = size_info.get('כמות', 0)
                    
                    # Try to find square area for this item
                    products_catalog = getattr(self.data_processor, 'products_catalog', [])
                    square_area = None
                    for catalog_item in products_catalog:
                        if (catalog_item.get('name', '') == product_name and 
                            catalog_item.get('size', '') == size):
                            square_area = catalog_item.get('square_area', 0.0)
                            break
                    
                    if square_area and square_area > 0:
                        info_text += f"  - מידה {size}: {quantity} יחידות (שטח רבוע: {square_area:.6f} מ״ר)\n"
                    else:
                        info_text += f"  - מידה {size}: {quantity} יחידות (⚠️ שטח רבוע לא נמצא בקטלוג)\n"
            
            self.drawing_info_text.insert(tk.END, info_text)
            self.drawing_info_text.config(state='disabled')
            
            # Auto-fill fields if drawing is cut and has data
            if status == "נחתך":
                # Fill layers count
                if layers is not None:
                    self.layers_count_var.set(str(layers))
                else:
                    self.layers_count_var.set("")
                
                # Fill weight if available - calculate weight per layer in grams
                # total_weight is in KG, we need grams per single layer
                if total_weight is not None and layers is not None:
                    try:
                        total_weight_val = float(total_weight)
                        layers_val = int(layers)
                        if layers_val > 0:
                            # Convert KG to grams and divide by number of layers
                            weight_per_layer_grams = (total_weight_val * 1000) / layers_val
                            self.layer_weight_var.set(f"{weight_per_layer_grams:.2f}")
                            layer_desc = "שכבות" if drawing.get('שכבות') else "שכבות משוערת"
                            self.summary_var.set(f"ציור נחתך זוהה! {layers} {layer_desc}, משקל לשכבה: {weight_per_layer_grams:.2f} גרם (חושב מ-{total_weight_val:.2f} ק\"ג)")
                        else:
                            self.layer_weight_var.set("")
                            self.summary_var.set("ציור נחתך זוהה! הזן נתונים נוספים לחישוב")
                    except (ValueError, TypeError):
                        self.layer_weight_var.set("")
                        self.summary_var.set("ציור נחתך זוהה! הזן נתונים נוספים לחישוב")
                elif total_weight is not None:
                    # If we have weight but no layers, can't calculate per-layer weight
                    self.layer_weight_var.set("")
                    weight_desc = "משקל כולל" if drawing.get('משקל כולל') else "משקל נגזר"
                    self.summary_var.set(f"ציור נחתך זוהה! {weight_desc}: {total_weight} ק\"ג - הזן כמות שכבות לחישוב")
                else:
                    self.layer_weight_var.set("")
                    if layers is not None:
                        layer_desc = "שכבות" if drawing.get('שכבות') else "שכבות משוערת"
                        self.summary_var.set(f"ציור נחתך זוהה! כמות {layer_desc} מולאה אוטומטית: {layers}")
                    else:
                        self.summary_var.set("ציור נחתך זוהה! הזן נתונים נוספים לחישוב")
            else:
                # Clear fields for non-cut drawings
                self.layers_count_var.set("")
                self.layer_weight_var.set("")
                self.summary_var.set("הזן כמות שכבות ומשקל השכבה וחץ על 'חשב חלוקת משקל'")
            
            # Clear previous results
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בטעינת פרטי הציור: {str(e)}")
    
    def _calculate_weight_distribution(self):
        """Calculate weight distribution based on square areas."""
        try:
            selected_name = self.selected_drawing_var.get()
            if not selected_name or selected_name not in self.drawings_dict:
                messagebox.showwarning("אזהרה", "אנא בחר ציור תחילה")
                return
            
            weight_str = self.layer_weight_var.get().strip()
            if not weight_str:
                messagebox.showwarning("אזהרה", "אנא הזן משקל השכבה")
                return
            
            try:
                total_weight = float(weight_str)
                if total_weight <= 0:
                    messagebox.showwarning("אזהרה", "משקל השכבה חייב להיות מספר חיובי")
                    return
            except ValueError:
                messagebox.showwarning("אזהרה", "אנא הזן משקל תקין")
                return
            
            # Get price per kg (optional)
            price_per_kg = 0.0
            price_str = self.fabric_weight_price_var.get().strip()
            if price_str:
                try:
                    price_per_kg = float(price_str)
                    if price_per_kg < 0:
                        price_per_kg = 0.0
                except ValueError:
                    price_per_kg = 0.0
            
            drawing = self.drawings_dict[selected_name]
            products = drawing.get('מוצרים', [])
            
            # Get products catalog for square area lookup
            products_catalog = getattr(self.data_processor, 'products_catalog', [])
            
            # Build lookup dictionary for square areas
            square_area_lookup = {}
            for catalog_item in products_catalog:
                product_name = catalog_item.get('name', '')
                size = catalog_item.get('size', '')
                square_area = catalog_item.get('square_area', 0.0)
                key = f"{product_name}_{size}"
                square_area_lookup[key] = square_area
            
            # Calculate total square area and collect items
            total_square_area = 0.0
            calculation_items = []
            
            for product in products:
                product_name = product.get('שם המוצר', '')
                sizes = product.get('מידות', [])
                
                for size_info in sizes:
                    size = size_info.get('מידה', '')
                    quantity = size_info.get('כמות', 0)
                    
                    # Look up square area
                    lookup_key = f"{product_name}_{size}"
                    square_area = square_area_lookup.get(lookup_key, 0.0)
                    
                    if square_area > 0:
                        total_area_for_item = square_area * quantity
                        total_square_area += total_area_for_item
                        
                        calculation_items.append({
                            'product_name': product_name,
                            'size': size,
                            'quantity': quantity,
                            'square_area': square_area,
                            'total_area': total_area_for_item
                        })
            
            if total_square_area == 0:
                messagebox.showwarning("אזהרה", "לא נמצאו נתוני שטח רבוע עבור הפריטים בציור.\nוודא שהפריטים קיימים בקטלוג המוצרים עם ערכי שטח רבוע.")
                return
            
            # Clear previous results
            for item in self.results_tree.get_children():
                self.results_tree.delete(item)
            
            # Calculate and display results
            total_calculated_weight = 0.0
            total_cost = 0.0
            
            for item in calculation_items:
                percentage = (item['total_area'] / total_square_area) * 100
                weight_total = total_weight * (percentage / 100)
                weight_per_unit = weight_total / item['quantity'] if item['quantity'] > 0 else 0
                total_calculated_weight += weight_total
                
                # Calculate cost: (weight in grams / 1000) * price per kg
                cost_total = (weight_total / 1000) * price_per_kg if price_per_kg > 0 else 0
                cost_per_unit = (weight_per_unit / 1000) * price_per_kg if price_per_kg > 0 else 0
                total_cost += cost_total
                
                # Insert into table
                self.results_tree.insert('', 'end', values=(
                    item['product_name'],
                    item['size'],
                    item['quantity'],
                    f"{item['square_area']:.6f}",
                    f"{percentage:.2f}%",
                    f"{weight_total:.2f}",
                    f"{weight_per_unit:.2f}",
                    f"{cost_total:.2f}" if price_per_kg > 0 else "--",
                    f"{cost_per_unit:.2f}" if price_per_kg > 0 else "--"
                ))
            
            # Update summary
            if price_per_kg > 0:
                summary_text = f"סה״כ שטח רבוע: {total_square_area:.6f} מ״ר | משקל כולל: {total_weight:.2f} גרמים | עלות כוללת: {total_cost:.2f} ₪ | חולק ל-{len(calculation_items)} פריטים"
            else:
                summary_text = f"סה״כ שטח רבוע: {total_square_area:.6f} מ״ר | משקל כולל: {total_weight:.2f} גרמים | חולק ל-{len(calculation_items)} פריטים"
            self.summary_var.set(summary_text)
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בחישוב חלוקת המשקל: {str(e)}")
    
    def _export_weight_results(self):
        """Export weight calculation results to Excel."""
        try:
            # Check if there are results to export
            if not hasattr(self, 'results_tree') or not self.results_tree.get_children():
                messagebox.showwarning("אזהרה", "אין תוצאות לייצוא. חשב חלוקת משקל תחילה.")
                return
            
            # Get file path from user
            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                title="ייצוא תוצאות חלוקת משקל",
                defaultextension='.xlsx',
                filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')]
            )
            
            if not file_path:
                return
            
            # Collect data from results tree
            results_data = []
            for item in self.results_tree.get_children():
                values = self.results_tree.item(item, 'values')
                results_data.append({
                    'שם המוצר': values[0],
                    'מידה': values[1],
                    'כמות': int(values[2]),
                    'שטח רבוע (מ״ר)': float(values[3]),
                    'אחוז (%)': values[4],
                    'משקל כולל (גרמים)': float(values[5].replace(',', '')),
                    'משקל ליחידה (גרמים)': float(values[6].replace(',', ''))
                })
            
            # Create DataFrame and export
            import pandas as pd
            df = pd.DataFrame(results_data)
            
            # Add summary information
            selected_drawing = self.selected_drawing_var.get()
            total_weight = self.layer_weight_var.get()
            layers_count = self.layers_count_var.get()
            
            # Create Excel writer with multiple sheets
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Main results sheet
                df.to_excel(writer, sheet_name='תוצאות חלוקת משקל', index=False)
                
                # Summary sheet
                summary_data = {
                    'פרמטר': ['ציור נבחר', 'כמות שכבות', 'משקל השכבה הכולל (גרמים)', 'סה״כ פריטים', 'סה״כ יחידות', 'סה״כ שטח רבוע (מ״ר)', 'סה״כ משקל מחושב (גרמים)'],
                    'ערך': [
                        selected_drawing,
                        layers_count if layers_count else 'לא צוין',
                        total_weight,
                        len(results_data),
                        sum(item['כמות'] for item in results_data),
                        sum(item['שטח רבוע (מ״ר)'] for item in results_data),
                        sum(item['משקל כולל (גרמים)'] for item in results_data)
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='סיכום', index=False)
            
            messagebox.showinfo("הצלחה", f"התוצאות יוצאו בהצלחה לקובץ:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בייצוא התוצאות: {str(e)}")
    
    def _build_product_cost_content(self, container):
        """Build the product cost calculation content."""
        
        # Title and description
        title_frame = tk.Frame(container, bg='#f7f9fa')
        title_frame.pack(fill='x', padx=20, pady=(10, 5))
        
        tk.Label(
            title_frame,
            text="חישוב עלות שמיכות",
            font=('Arial', 14, 'bold'),
            bg='#f7f9fa',
            fg='#2c3e50'
        ).pack()
        
        tk.Label(
            title_frame,
            text="חישוב עלות יצור של פריטי טקסטיל - שמיכות, בגדי גוף, מגבות וכו'",
            font=('Arial', 9),
            bg='#f7f9fa',
            fg='#7f8c8d'
        ).pack()
        
        # Main input frame
        input_frame = ttk.LabelFrame(container, text="נתוני הפריט", padding=20)
        input_frame.pack(fill='x', padx=20, pady=10)
        
        # Configure grid columns
        input_frame.grid_columnconfigure(1, weight=1)
        input_frame.grid_columnconfigure(3, weight=1)
        
        row = 0
        
        # Fabric price per kg
        tk.Label(input_frame, text="מחיר הבד ל-1 ק״ג:", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.fabric_price_per_kg_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.fabric_price_per_kg_var, width=15, font=('Arial', 10)).grid(
            row=row, column=1, padx=5, pady=5, sticky='w')
        
        # Fabric weight per meter
        tk.Label(input_frame, text="משקל 1 מטר רץ בד (גרמים):", font=('Arial', 10, 'bold')).grid(
            row=row, column=2, sticky='w', padx=5, pady=5)
        self.fabric_weight_per_meter_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.fabric_weight_per_meter_var, width=15, font=('Arial', 10)).grid(
            row=row, column=3, padx=5, pady=5, sticky='w')
        
        row += 1
        
        # Roll width
        tk.Label(input_frame, text="רוחב הגליל (ס״מ):", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.roll_width_cm_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.roll_width_cm_var, width=15, font=('Arial', 10)).grid(
            row=row, column=1, padx=5, pady=5, sticky='w')
        
        # Printing price per meter
        tk.Label(input_frame, text="מחיר הדפסה למטר רץ:", font=('Arial', 10, 'bold')).grid(
            row=row, column=2, sticky='w', padx=5, pady=5)
        self.printing_price_per_meter_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.printing_price_per_meter_var, width=15, font=('Arial', 10)).grid(
            row=row, column=3, padx=5, pady=5, sticky='w')
        
        row += 1
        
        # Item width
        tk.Label(input_frame, text="רוחב הפריט (ס״מ):", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.item_width_cm_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.item_width_cm_var, width=15, font=('Arial', 10)).grid(
            row=row, column=1, padx=5, pady=5, sticky='w')
        
        # Item length
        tk.Label(input_frame, text="אורך הפריט (ס״מ):", font=('Arial', 10, 'bold')).grid(
            row=row, column=2, sticky='w', padx=5, pady=5)
        self.item_length_cm_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.item_length_cm_var, width=15, font=('Arial', 10)).grid(
            row=row, column=3, padx=5, pady=5, sticky='w')
        
        row += 1
        
        # Number of layers
        tk.Label(input_frame, text="מספר שכבות בד בפריט:", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.num_layers_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.num_layers_var, width=15, font=('Arial', 10)).grid(
            row=row, column=1, padx=5, pady=5, sticky='w')
        
        # Printed layers
        tk.Label(input_frame, text="כמה שכבות מודפסות (0, 1 או 2):", font=('Arial', 10, 'bold')).grid(
            row=row, column=2, sticky='w', padx=5, pady=5)
        self.printed_layers_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.printed_layers_var, width=15, font=('Arial', 10)).grid(
            row=row, column=3, padx=5, pady=5, sticky='w')
        
        row += 1
        
        # Waste percentage
        tk.Label(input_frame, text="אחוז בזבוז/התכווצות (%):", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.waste_percentage_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.waste_percentage_var, width=15, font=('Arial', 10)).grid(
            row=row, column=1, padx=5, pady=5, sticky='w')
        
        row += 1
        
        # Additional costs frame
        costs_frame = ttk.LabelFrame(container, text="עלויות נוספות", padding=20)
        costs_frame.pack(fill='x', padx=20, pady=10)
        
        # Configure grid columns
        costs_frame.grid_columnconfigure(1, weight=1)
        costs_frame.grid_columnconfigure(3, weight=1)
        
        # Sewing cost
        tk.Label(costs_frame, text="עלות תפירה ליחידה:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=5, pady=5)
        self.sewing_cost_per_unit_var = tk.StringVar()
        tk.Entry(costs_frame, textvariable=self.sewing_cost_per_unit_var, width=15, font=('Arial', 10)).grid(
            row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Cutting cost
        tk.Label(costs_frame, text="עלות גזירה ליחידה:", font=('Arial', 10, 'bold')).grid(
            row=0, column=2, sticky='w', padx=5, pady=5)
        self.cutting_cost_per_unit_var = tk.StringVar()
        tk.Entry(costs_frame, textvariable=self.cutting_cost_per_unit_var, width=15, font=('Arial', 10)).grid(
            row=0, column=3, padx=5, pady=5, sticky='w')
        
        # Filling cost
        tk.Label(costs_frame, text="עלות מילוי ליחידה:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky='w', padx=5, pady=5)
        self.filling_cost_per_unit_var = tk.StringVar()
        tk.Entry(costs_frame, textvariable=self.filling_cost_per_unit_var, width=15, font=('Arial', 10)).grid(
            row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Buttons frame
        buttons_frame = tk.Frame(container, bg='#f7f9fa')
        buttons_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Button(
            buttons_frame,
            text="חשב עלות",
            command=self._calculate_product_cost,
            bg='#27ae60',
            fg='white',
            font=('Arial', 11, 'bold'),
            width=20,
            height=2
        ).pack(side='left', padx=5)
        
        tk.Button(
            buttons_frame,
            text="נקה הכל",
            command=self._clear_product_cost_inputs,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 11, 'bold'),
            width=15,
            height=2
        ).pack(side='left', padx=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(container, text="תוצאות החישוב", padding=20)
        results_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Results display
        self.product_cost_results_text = tk.Text(
            results_frame,
            height=15,
            font=('Courier New', 10),
            wrap=tk.WORD,
            state='disabled',
            bg='#ffffff'
        )
        self.product_cost_results_text.pack(fill='both', expand=True)
        
        # Configure text tags for formatting
        self.product_cost_results_text.tag_configure('header', font=('Arial', 11, 'bold'), foreground='#2c3e50')
        self.product_cost_results_text.tag_configure('label', font=('Courier New', 10), foreground='#34495e')
        self.product_cost_results_text.tag_configure('value', font=('Courier New', 10, 'bold'), foreground='#2980b9')
        self.product_cost_results_text.tag_configure('total', font=('Arial', 13, 'bold'), foreground='#27ae60')
        self.product_cost_results_text.tag_configure('separator', foreground='#7f8c8d')
        
        # Initial message
        self._clear_product_cost_inputs()
    
    def _calculate_product_cost(self):
        """Calculate product manufacturing cost."""
        try:
            # Get all input values
            fabric_price_per_kg = float(self.fabric_price_per_kg_var.get() or 0)
            fabric_weight_per_meter = float(self.fabric_weight_per_meter_var.get() or 0)
            roll_width_cm = float(self.roll_width_cm_var.get() or 0)
            item_width_cm = float(self.item_width_cm_var.get() or 0)
            item_length_cm = float(self.item_length_cm_var.get() or 0)
            printing_price_per_meter = float(self.printing_price_per_meter_var.get() or 0)
            num_layers = float(self.num_layers_var.get() or 0)
            printed_layers = float(self.printed_layers_var.get() or 0)
            waste_percentage = float(self.waste_percentage_var.get() or 0)
            sewing_cost_per_unit = float(self.sewing_cost_per_unit_var.get() or 0)
            cutting_cost_per_unit = float(self.cutting_cost_per_unit_var.get() or 0)
            filling_cost_per_unit = float(self.filling_cost_per_unit_var.get() or 0)
            
            # Validate inputs
            if fabric_price_per_kg <= 0:
                messagebox.showwarning("אזהרה", "מחיר הבד חייב להיות גדול מ-0")
                return
            
            if fabric_weight_per_meter <= 0:
                messagebox.showwarning("אזהרה", "משקל הבד למטר חייב להיות גדול מ-0")
                return
            
            if roll_width_cm <= 0:
                messagebox.showwarning("אזהרה", "רוחב הגליל חייב להיות גדול מ-0")
                return
            
            if item_width_cm <= 0 or item_length_cm <= 0:
                messagebox.showwarning("אזהרה", "מידות הפריט חייבות להיות גדולות מ-0")
                return
            
            if item_width_cm > roll_width_cm:
                messagebox.showwarning("אזהרה", "רוחב הפריט גדול מרוחב הגליל")
                return
            
            if num_layers <= 0:
                messagebox.showwarning("אזהרה", "מספר השכבות חייב להיות גדול מ-0")
                return
            
            if printed_layers < 0 or printed_layers > 2:
                messagebox.showwarning("אזהרה", "מספר שכבות מודפסות חייב להיות 0, 1 או 2")
                return
            
            if printed_layers > num_layers:
                messagebox.showwarning("אזהרה", "מספר שכבות מודפסות לא יכול להיות גדול ממספר השכבות הכולל")
                return
            
            if waste_percentage < 0:
                messagebox.showwarning("אזהרה", "אחוז הבזבוז לא יכול להיות שלילי")
                return
            
            # Perform calculations
            import math
            
            # Units per width of roll
            units_per_width = math.floor(roll_width_cm / item_width_cm)
            
            # Meters per unit (including waste)
            meters_per_unit = (item_length_cm / 100) * (1 + waste_percentage / 100)
            
            # Fabric cost per meter
            fabric_cost_per_meter = (fabric_price_per_kg * fabric_weight_per_meter) / 1000
            
            # Fabric cost per unit
            fabric_cost_per_unit = fabric_cost_per_meter * meters_per_unit * num_layers
            
            # Printing cost per unit
            if units_per_width > 0 and printed_layers > 0:
                printing_cost_per_unit = (printing_price_per_meter * meters_per_unit * printed_layers) / units_per_width
            else:
                printing_cost_per_unit = 0
            
            # Total cost per unit
            total_cost_per_unit = fabric_cost_per_unit + printing_cost_per_unit + sewing_cost_per_unit + cutting_cost_per_unit + filling_cost_per_unit
            
            # Display results
            self.product_cost_results_text.config(state='normal')
            self.product_cost_results_text.delete(1.0, tk.END)
            
            # Header
            self.product_cost_results_text.insert(tk.END, "תוצאות חישוב עלות יצור מוצר\n", 'header')
            self.product_cost_results_text.insert(tk.END, "=" * 70 + "\n\n", 'separator')
            
            # Calculation results
            self.product_cost_results_text.insert(tk.END, "נתוני ייצור:\n", 'header')
            self.product_cost_results_text.insert(tk.END, f"  יחידות לרוחב הגליל:                ", 'label')
            self.product_cost_results_text.insert(tk.END, f"{units_per_width}\n", 'value')
            
            self.product_cost_results_text.insert(tk.END, f"  מטר רץ ליחידה (כולל בזבוז):         ", 'label')
            self.product_cost_results_text.insert(tk.END, f"{meters_per_unit:.4f} מ'\n", 'value')
            
            self.product_cost_results_text.insert(tk.END, f"  עלות בד למטר רץ:                     ", 'label')
            self.product_cost_results_text.insert(tk.END, f"{fabric_cost_per_meter:.4f} ₪\n\n", 'value')
            
            self.product_cost_results_text.insert(tk.END, "-" * 70 + "\n\n", 'separator')
            
            # Cost breakdown
            self.product_cost_results_text.insert(tk.END, "פירוט עלויות ליחידה:\n", 'header')
            self.product_cost_results_text.insert(tk.END, f"  עלות בד ליחידה:                      ", 'label')
            self.product_cost_results_text.insert(tk.END, f"{fabric_cost_per_unit:.4f} ₪\n", 'value')
            
            self.product_cost_results_text.insert(tk.END, f"  עלות הדפסה ליחידה:                   ", 'label')
            self.product_cost_results_text.insert(tk.END, f"{printing_cost_per_unit:.4f} ₪\n", 'value')
            
            self.product_cost_results_text.insert(tk.END, f"  עלות תפירה ליחידה:                   ", 'label')
            self.product_cost_results_text.insert(tk.END, f"{sewing_cost_per_unit:.4f} ₪\n", 'value')
            
            self.product_cost_results_text.insert(tk.END, f"  עלות גזירה ליחידה:                   ", 'label')
            self.product_cost_results_text.insert(tk.END, f"{cutting_cost_per_unit:.4f} ₪\n", 'value')
            
            self.product_cost_results_text.insert(tk.END, f"  עלות מילוי ליחידה:                   ", 'label')
            self.product_cost_results_text.insert(tk.END, f"{filling_cost_per_unit:.4f} ₪\n\n", 'value')
            
            self.product_cost_results_text.insert(tk.END, "=" * 70 + "\n\n", 'separator')
            
            # Total cost
            self.product_cost_results_text.insert(tk.END, "עלות כוללת ליחידה: ", 'total')
            self.product_cost_results_text.insert(tk.END, f"{total_cost_per_unit:.4f} ₪\n\n", 'total')
            
            self.product_cost_results_text.insert(tk.END, "=" * 70 + "\n", 'separator')
            
            self.product_cost_results_text.config(state='disabled')
            
        except ValueError:
            messagebox.showerror("שגיאה", "אנא הזן ערכים נומריים תקינים בכל השדות")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בחישוב: {str(e)}")
    
    def _clear_product_cost_inputs(self):
        """Clear all product cost input fields and results."""
        # Clear all input variables
        self.fabric_price_per_kg_var.set("")
        self.fabric_weight_per_meter_var.set("")
        self.roll_width_cm_var.set("")
        self.item_width_cm_var.set("")
        self.item_length_cm_var.set("")
        self.printing_price_per_meter_var.set("")
        self.num_layers_var.set("")
        self.printed_layers_var.set("")
        self.waste_percentage_var.set("")
        self.sewing_cost_per_unit_var.set("")
        self.cutting_cost_per_unit_var.set("")
        self.filling_cost_per_unit_var.set("")
        
        # Clear results display
        self.product_cost_results_text.config(state='normal')
        self.product_cost_results_text.delete(1.0, tk.END)
        self.product_cost_results_text.insert(
            tk.END,
            "\n\n\n          הזן את נתוני הפריט ולחץ על 'חשב עלות' לקבלת תוצאות\n\n\n",
            'header'
        )
        self.product_cost_results_text.config(state='disabled')
    
    def _build_tetra_cost_content(self, container):
        """Build the tetra cost calculation content."""
        
        # Title and description
        title_frame = tk.Frame(container, bg='#f7f9fa')
        title_frame.pack(fill='x', padx=20, pady=(10, 5))
        
        tk.Label(
            title_frame,
            text="חישוב עלות טטרות",
            font=('Arial', 14, 'bold'),
            bg='#f7f9fa',
            fg='#2c3e50'
        ).pack()
        
        tk.Label(
            title_frame,
            text="חישוב עלות יצור טטרות לפי מחיר למטר רץ",
            font=('Arial', 9),
            bg='#f7f9fa',
            fg='#7f8c8d'
        ).pack()
        
        # Main input frame
        input_frame = ttk.LabelFrame(container, text="נתוני הפריט", padding=20)
        input_frame.pack(fill='x', padx=20, pady=10)
        
        # Configure grid columns
        input_frame.grid_columnconfigure(1, weight=1)
        input_frame.grid_columnconfigure(3, weight=1)
        
        row = 0
        
        # Fabric price per meter
        tk.Label(input_frame, text="מחיר בד למטר רץ:", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.tetra_fabric_price_per_meter_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.tetra_fabric_price_per_meter_var, width=15, font=('Arial', 10)).grid(
            row=row, column=1, padx=5, pady=5, sticky='w')
        
        # Roll width
        tk.Label(input_frame, text="רוחב הגליל (ס״מ):", font=('Arial', 10, 'bold')).grid(
            row=row, column=2, sticky='w', padx=5, pady=5)
        self.tetra_roll_width_cm_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.tetra_roll_width_cm_var, width=15, font=('Arial', 10)).grid(
            row=row, column=3, padx=5, pady=5, sticky='w')
        
        row += 1
        
        # Printing price per meter
        tk.Label(input_frame, text="מחיר הדפסה למטר רץ:", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.tetra_printing_price_per_meter_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.tetra_printing_price_per_meter_var, width=15, font=('Arial', 10)).grid(
            row=row, column=1, padx=5, pady=5, sticky='w')
        
        # Item width
        tk.Label(input_frame, text="רוחב הפריט (ס״מ):", font=('Arial', 10, 'bold')).grid(
            row=row, column=2, sticky='w', padx=5, pady=5)
        self.tetra_item_width_cm_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.tetra_item_width_cm_var, width=15, font=('Arial', 10)).grid(
            row=row, column=3, padx=5, pady=5, sticky='w')
        
        row += 1
        
        # Item length
        tk.Label(input_frame, text="אורך הפריט (ס״מ):", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.tetra_item_length_cm_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.tetra_item_length_cm_var, width=15, font=('Arial', 10)).grid(
            row=row, column=1, padx=5, pady=5, sticky='w')
        
        # Number of layers
        tk.Label(input_frame, text="מספר שכבות בד בפריט:", font=('Arial', 10, 'bold')).grid(
            row=row, column=2, sticky='w', padx=5, pady=5)
        self.tetra_num_layers_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.tetra_num_layers_var, width=15, font=('Arial', 10)).grid(
            row=row, column=3, padx=5, pady=5, sticky='w')
        
        row += 1
        
        # Printed layers
        tk.Label(input_frame, text="כמה שכבות מודפסות (0, 1 או 2):", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky='w', padx=5, pady=5)
        self.tetra_printed_layers_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.tetra_printed_layers_var, width=15, font=('Arial', 10)).grid(
            row=row, column=1, padx=5, pady=5, sticky='w')
        
        # Waste percentage
        tk.Label(input_frame, text="אחוז בזבוז/התכווצות (%):", font=('Arial', 10, 'bold')).grid(
            row=row, column=2, sticky='w', padx=5, pady=5)
        self.tetra_waste_percentage_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.tetra_waste_percentage_var, width=15, font=('Arial', 10)).grid(
            row=row, column=3, padx=5, pady=5, sticky='w')
        
        row += 1
        
        # Additional costs frame
        costs_frame = ttk.LabelFrame(container, text="עלויות נוספות", padding=20)
        costs_frame.pack(fill='x', padx=20, pady=10)
        
        # Configure grid columns
        costs_frame.grid_columnconfigure(1, weight=1)
        costs_frame.grid_columnconfigure(3, weight=1)
        
        # Sewing cost
        tk.Label(costs_frame, text="עלות תפירה ליחידה:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=5, pady=5)
        self.tetra_sewing_cost_per_unit_var = tk.StringVar()
        tk.Entry(costs_frame, textvariable=self.tetra_sewing_cost_per_unit_var, width=15, font=('Arial', 10)).grid(
            row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Cutting cost
        tk.Label(costs_frame, text="עלות גזירה ליחידה:", font=('Arial', 10, 'bold')).grid(
            row=0, column=2, sticky='w', padx=5, pady=5)
        self.tetra_cutting_cost_per_unit_var = tk.StringVar()
        tk.Entry(costs_frame, textvariable=self.tetra_cutting_cost_per_unit_var, width=15, font=('Arial', 10)).grid(
            row=0, column=3, padx=5, pady=5, sticky='w')
        
        # Buttons frame
        buttons_frame = tk.Frame(container, bg='#f7f9fa')
        buttons_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Button(
            buttons_frame,
            text="חשב עלות",
            command=self._calculate_tetra_cost,
            bg='#27ae60',
            fg='white',
            font=('Arial', 11, 'bold'),
            width=20,
            height=2
        ).pack(side='left', padx=5)
        
        tk.Button(
            buttons_frame,
            text="נקה הכל",
            command=self._clear_tetra_cost_inputs,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 11, 'bold'),
            width=15,
            height=2
        ).pack(side='left', padx=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(container, text="תוצאות החישוב", padding=20)
        results_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Results display
        self.tetra_cost_results_text = tk.Text(
            results_frame,
            height=15,
            font=('Courier New', 10),
            wrap=tk.WORD,
            state='disabled',
            bg='#ffffff'
        )
        self.tetra_cost_results_text.pack(fill='both', expand=True)
        
        # Configure text tags for formatting
        self.tetra_cost_results_text.tag_configure('header', font=('Arial', 11, 'bold'), foreground='#2c3e50')
        self.tetra_cost_results_text.tag_configure('label', font=('Courier New', 10), foreground='#34495e')
        self.tetra_cost_results_text.tag_configure('value', font=('Courier New', 10, 'bold'), foreground='#2980b9')
        self.tetra_cost_results_text.tag_configure('total', font=('Arial', 13, 'bold'), foreground='#27ae60')
        self.tetra_cost_results_text.tag_configure('separator', foreground='#7f8c8d')
        
        # Initial message
        self._clear_tetra_cost_inputs()
    
    def _calculate_tetra_cost(self):
        """Calculate tetra manufacturing cost."""
        try:
            # Get all input values
            fabric_price_per_meter = float(self.tetra_fabric_price_per_meter_var.get() or 0)
            roll_width_cm = float(self.tetra_roll_width_cm_var.get() or 0)
            item_width_cm = float(self.tetra_item_width_cm_var.get() or 0)
            item_length_cm = float(self.tetra_item_length_cm_var.get() or 0)
            printing_price_per_meter = float(self.tetra_printing_price_per_meter_var.get() or 0)
            num_layers = float(self.tetra_num_layers_var.get() or 0)
            printed_layers = float(self.tetra_printed_layers_var.get() or 0)
            waste_percentage = float(self.tetra_waste_percentage_var.get() or 0)
            sewing_cost_per_unit = float(self.tetra_sewing_cost_per_unit_var.get() or 0)
            cutting_cost_per_unit = float(self.tetra_cutting_cost_per_unit_var.get() or 0)
            
            # Validate inputs
            if fabric_price_per_meter <= 0:
                messagebox.showwarning("אזהרה", "מחיר הבד חייב להיות גדול מ-0")
                return
            
            if roll_width_cm <= 0:
                messagebox.showwarning("אזהרה", "רוחב הגליל חייב להיות גדול מ-0")
                return
            
            if item_width_cm <= 0 or item_length_cm <= 0:
                messagebox.showwarning("אזהרה", "מידות הפריט חייבות להיות גדולות מ-0")
                return
            
            if item_width_cm > roll_width_cm:
                messagebox.showwarning("אזהרה", "רוחב הפריט גדול מרוחב הגליל")
                return
            
            if num_layers <= 0:
                messagebox.showwarning("אזהרה", "מספר השכבות חייב להיות גדול מ-0")
                return
            
            if printed_layers < 0 or printed_layers > 2:
                messagebox.showwarning("אזהרה", "מספר שכבות מודפסות חייב להיות 0, 1 או 2")
                return
            
            if printed_layers > num_layers:
                messagebox.showwarning("אזהרה", "מספר שכבות מודפסות לא יכול להיות גדול ממספר השכבות הכולל")
                return
            
            if waste_percentage < 0:
                messagebox.showwarning("אזהרה", "אחוז הבזבוז לא יכול להיות שלילי")
                return
            
            # Perform calculations
            import math
            
            # Units per width of roll
            units_per_width = math.floor(roll_width_cm / item_width_cm)
            
            # Meters per unit (including waste)
            meters_per_unit = (item_length_cm / 100) * (1 + waste_percentage / 100)
            
            # Fabric cost per unit (based on price per meter)
            fabric_cost_per_unit = fabric_price_per_meter * meters_per_unit * num_layers
            
            # Printing cost per unit
            if units_per_width > 0 and printed_layers > 0:
                printing_cost_per_unit = (printing_price_per_meter * meters_per_unit * printed_layers) / units_per_width
            else:
                printing_cost_per_unit = 0
            
            # Total cost per unit
            total_cost_per_unit = fabric_cost_per_unit + printing_cost_per_unit + sewing_cost_per_unit + cutting_cost_per_unit
            
            # Display results
            self.tetra_cost_results_text.config(state='normal')
            self.tetra_cost_results_text.delete(1.0, tk.END)
            
            # Header
            self.tetra_cost_results_text.insert(tk.END, "תוצאות חישוב עלות טטרה\n", 'header')
            self.tetra_cost_results_text.insert(tk.END, "=" * 70 + "\n\n", 'separator')
            
            # Calculation results
            self.tetra_cost_results_text.insert(tk.END, "נתוני ייצור:\n", 'header')
            self.tetra_cost_results_text.insert(tk.END, f"  יחידות לרוחב הגליל:                ", 'label')
            self.tetra_cost_results_text.insert(tk.END, f"{units_per_width}\n", 'value')
            
            self.tetra_cost_results_text.insert(tk.END, f"  מטר רץ ליחידה (כולל בזבוז):         ", 'label')
            self.tetra_cost_results_text.insert(tk.END, f"{meters_per_unit:.4f} מ'\n\n", 'value')
            
            self.tetra_cost_results_text.insert(tk.END, "-" * 70 + "\n\n", 'separator')
            
            # Cost breakdown
            self.tetra_cost_results_text.insert(tk.END, "פירוט עלויות ליחידה:\n", 'header')
            self.tetra_cost_results_text.insert(tk.END, f"  עלות בד ליחידה:                      ", 'label')
            self.tetra_cost_results_text.insert(tk.END, f"{fabric_cost_per_unit:.4f} ₪\n", 'value')
            
            self.tetra_cost_results_text.insert(tk.END, f"  עלות הדפסה ליחידה:                   ", 'label')
            self.tetra_cost_results_text.insert(tk.END, f"{printing_cost_per_unit:.4f} ₪\n", 'value')
            
            self.tetra_cost_results_text.insert(tk.END, f"  עלות תפירה ליחידה:                   ", 'label')
            self.tetra_cost_results_text.insert(tk.END, f"{sewing_cost_per_unit:.4f} ₪\n", 'value')
            
            self.tetra_cost_results_text.insert(tk.END, f"  עלות גזירה ליחידה:                   ", 'label')
            self.tetra_cost_results_text.insert(tk.END, f"{cutting_cost_per_unit:.4f} ₪\n\n", 'value')
            
            self.tetra_cost_results_text.insert(tk.END, "=" * 70 + "\n\n", 'separator')
            
            # Total cost
            self.tetra_cost_results_text.insert(tk.END, "עלות כוללת ליחידה: ", 'total')
            self.tetra_cost_results_text.insert(tk.END, f"{total_cost_per_unit:.4f} ₪\n\n", 'total')
            
            self.tetra_cost_results_text.insert(tk.END, "=" * 70 + "\n", 'separator')
            
            self.tetra_cost_results_text.config(state='disabled')
            
        except ValueError:
            messagebox.showerror("שגיאה", "אנא הזן ערכים נומריים תקינים בכל השדות")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בחישוב: {str(e)}")
    
    def _clear_tetra_cost_inputs(self):
        """Clear all tetra cost input fields and results."""
        # Clear all input variables
        self.tetra_fabric_price_per_meter_var.set("")
        self.tetra_roll_width_cm_var.set("")
        self.tetra_printing_price_per_meter_var.set("")
        self.tetra_item_width_cm_var.set("")
        self.tetra_item_length_cm_var.set("")
        self.tetra_num_layers_var.set("")
        self.tetra_printed_layers_var.set("")
        self.tetra_waste_percentage_var.set("")
        self.tetra_sewing_cost_per_unit_var.set("")
        self.tetra_cutting_cost_per_unit_var.set("")
        
        # Clear results display
        self.tetra_cost_results_text.config(state='normal')
        self.tetra_cost_results_text.delete(1.0, tk.END)
        self.tetra_cost_results_text.insert(
            tk.END,
            "\n\n\n          הזן את נתוני הפריט ולחץ על 'חשב עלות' לקבלת תוצאות\n\n\n",
            'header'
        )
        self.tetra_cost_results_text.config(state='disabled')
    
    def _build_all_over_print_content(self, container):
        """Build the all over print cost calculation content."""
        
        # Title and description
        title_frame = tk.Frame(container, bg='#f7f9fa')
        title_frame.pack(fill='x', padx=20, pady=(10, 5))
        
        tk.Label(
            title_frame,
            text="חישוב עלות הדפסת בגדי אול אובר",
            font=('Arial', 14, 'bold'),
            bg='#f7f9fa',
            fg='#2c3e50'
        ).pack()
        
        tk.Label(
            title_frame,
            text="חישוב עלות הדפסה לפי שטח רבוע של הפריט",
            font=('Arial', 9),
            bg='#f7f9fa',
            fg='#7f8c8d'
        ).pack()
        
        # Product selection frame
        selection_frame = ttk.LabelFrame(container, text="בחירת פריט", padding=20)
        selection_frame.pack(fill='x', padx=20, pady=10)
        
        # Configure grid columns
        selection_frame.grid_columnconfigure(1, weight=1)
        selection_frame.grid_columnconfigure(3, weight=1)
        
        # Product name selection
        tk.Label(selection_frame, text="שם דגם:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=5, pady=5)
        self.aop_product_name_var = tk.StringVar()
        self.aop_product_name_combo = ttk.Combobox(
            selection_frame, 
            textvariable=self.aop_product_name_var,
            state='readonly',
            width=30,
            font=('Arial', 10)
        )
        self.aop_product_name_combo.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.aop_product_name_combo.bind('<<ComboboxSelected>>', self._on_product_name_selected)
        
        # Product size selection
        tk.Label(selection_frame, text="מידה:", font=('Arial', 10, 'bold')).grid(
            row=0, column=2, sticky='w', padx=5, pady=5)
        self.aop_product_size_var = tk.StringVar()
        self.aop_product_size_combo = ttk.Combobox(
            selection_frame,
            textvariable=self.aop_product_size_var,
            state='readonly',
            width=20,
            font=('Arial', 10)
        )
        self.aop_product_size_combo.grid(row=0, column=3, padx=5, pady=5, sticky='ew')
        self.aop_product_size_combo.bind('<<ComboboxSelected>>', self._on_product_size_selected)
        
        # Square area display
        tk.Label(selection_frame, text="שטח רבוע:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky='w', padx=5, pady=5)
        self.aop_square_area_label = tk.Label(
            selection_frame,
            text="בחר פריט",
            font=('Arial', 10),
            fg='#7f8c8d'
        )
        self.aop_square_area_label.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        # Printing data frame
        printing_frame = ttk.LabelFrame(container, text="נתוני הדפסה", padding=20)
        printing_frame.pack(fill='x', padx=20, pady=10)
        
        # Configure grid columns
        printing_frame.grid_columnconfigure(1, weight=1)
        printing_frame.grid_columnconfigure(3, weight=1)
        
        # Printing cost per square meter
        tk.Label(printing_frame, text="עלות הדפסה למטר רבוע (₪/מ״ר):", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=5, pady=5)
        self.aop_printing_cost_per_sqm_var = tk.StringVar()
        tk.Entry(printing_frame, textvariable=self.aop_printing_cost_per_sqm_var, width=15, font=('Arial', 10)).grid(
            row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Waste percentage
        tk.Label(printing_frame, text="אחוז פחת/בזבוז (%):", font=('Arial', 10, 'bold')).grid(
            row=0, column=2, sticky='w', padx=5, pady=5)
        self.aop_waste_percentage_var = tk.StringVar()
        tk.Entry(printing_frame, textvariable=self.aop_waste_percentage_var, width=15, font=('Arial', 10)).grid(
            row=0, column=3, padx=5, pady=5, sticky='w')
        
        # Buttons frame
        buttons_frame = tk.Frame(container, bg='#f7f9fa')
        buttons_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Button(
            buttons_frame,
            text="חשב עלות",
            command=self._calculate_all_over_print_cost,
            bg='#27ae60',
            fg='white',
            font=('Arial', 11, 'bold'),
            width=20,
            height=2
        ).pack(side='left', padx=5)
        
        tk.Button(
            buttons_frame,
            text="נקה הכל",
            command=self._clear_all_over_print_inputs,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 11, 'bold'),
            width=15,
            height=2
        ).pack(side='left', padx=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(container, text="תוצאות החישוב", padding=20)
        results_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Results display
        self.aop_results_text = tk.Text(
            results_frame,
            height=12,
            font=('Courier New', 10),
            wrap=tk.WORD,
            state='disabled',
            bg='#ffffff'
        )
        self.aop_results_text.pack(fill='both', expand=True)
        
        # Configure text tags for formatting
        self.aop_results_text.tag_configure('header', font=('Arial', 11, 'bold'), foreground='#2c3e50')
        self.aop_results_text.tag_configure('label', font=('Courier New', 10), foreground='#34495e')
        self.aop_results_text.tag_configure('value', font=('Courier New', 10, 'bold'), foreground='#2980b9')
        self.aop_results_text.tag_configure('total', font=('Arial', 13, 'bold'), foreground='#27ae60')
        self.aop_results_text.tag_configure('separator', foreground='#7f8c8d')
        
        # Initialize
        self.aop_selected_product = None
        self._load_product_names()
        self._clear_all_over_print_inputs()
    
    def _load_product_names(self):
        """Load unique product names from catalog."""
        try:
            products_catalog = getattr(self.data_processor, 'products_catalog', [])
            
            # Get unique product names
            product_names = sorted(list(set(
                product.get('name', '') 
                for product in products_catalog 
                if product.get('name')
            )))
            
            self.aop_product_name_combo['values'] = product_names
            
            if product_names:
                self.aop_product_name_combo.set(product_names[0])
                self._on_product_name_selected()
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בטעינת רשימת המוצרים: {str(e)}")
    
    def _on_product_name_selected(self, event=None):
        """Handle product name selection."""
        try:
            selected_name = self.aop_product_name_var.get()
            if not selected_name:
                return
            
            products_catalog = getattr(self.data_processor, 'products_catalog', [])
            
            # Get all sizes for selected product name
            sizes = sorted(list(set(
                product.get('size', '')
                for product in products_catalog
                if product.get('name') == selected_name and product.get('size')
            )))
            
            self.aop_product_size_combo['values'] = sizes
            
            # Clear previous selection
            self.aop_product_size_var.set('')
            self.aop_selected_product = None
            self.aop_square_area_label.config(text="בחר מידה")
            
            if sizes:
                self.aop_product_size_combo.set(sizes[0])
                self._on_product_size_selected()
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בטעינת המידות: {str(e)}")
    
    def _on_product_size_selected(self, event=None):
        """Handle product size selection."""
        try:
            selected_name = self.aop_product_name_var.get()
            selected_size = self.aop_product_size_var.get()
            
            if not selected_name or not selected_size:
                return
            
            products_catalog = getattr(self.data_processor, 'products_catalog', [])
            
            # Find the exact product
            for product in products_catalog:
                if (product.get('name') == selected_name and 
                    product.get('size') == selected_size):
                    self.aop_selected_product = product
                    square_area = product.get('square_area', 0)
                    self.aop_square_area_label.config(
                        text=f"{square_area:.6f} מ״ר",
                        fg='#27ae60',
                        font=('Arial', 10, 'bold')
                    )
                    break
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בטעינת פרטי המוצר: {str(e)}")
    
    def _calculate_all_over_print_cost(self):
        """Calculate all over print cost."""
        try:
            # Validate product selection
            if not self.aop_selected_product:
                messagebox.showwarning("אזהרה", "אנא בחר דגם ומידה")
                return
            
            # Get input values
            printing_cost_per_sqm_str = self.aop_printing_cost_per_sqm_var.get().strip()
            waste_percentage_str = self.aop_waste_percentage_var.get().strip()
            
            if not printing_cost_per_sqm_str:
                messagebox.showwarning("אזהרה", "אנא הזן עלות הדפסה למטר רבוע")
                return
            
            try:
                printing_cost_per_sqm = float(printing_cost_per_sqm_str)
                waste_percentage = float(waste_percentage_str) if waste_percentage_str else 0
            except ValueError:
                messagebox.showwarning("אזהרה", "אנא הזן ערכים נומריים תקינים")
                return
            
            # Validate inputs
            if printing_cost_per_sqm <= 0:
                messagebox.showwarning("אזהרה", "עלות הדפסה חייבת להיות גדולה מ-0")
                return
            
            if waste_percentage < 0:
                messagebox.showwarning("אזהרה", "אחוז הפחת לא יכול להיות שלילי")
                return
            
            # Get product data
            product_name = self.aop_selected_product.get('name', '')
            product_size = self.aop_selected_product.get('size', '')
            square_area = self.aop_selected_product.get('square_area', 0)
            
            # Perform calculations
            actual_square_area = square_area * (1 + waste_percentage / 100)
            total_printing_cost = actual_square_area * printing_cost_per_sqm
            
            # Display results
            self.aop_results_text.config(state='normal')
            self.aop_results_text.delete(1.0, tk.END)
            
            # Header
            self.aop_results_text.insert(tk.END, "תוצאות חישוב עלות הדפסה - אול אובר\n", 'header')
            self.aop_results_text.insert(tk.END, "=" * 70 + "\n\n", 'separator')
            
            # Product details
            self.aop_results_text.insert(tk.END, "פרטי הפריט:\n", 'header')
            self.aop_results_text.insert(tk.END, f"  שם דגם:                               ", 'label')
            self.aop_results_text.insert(tk.END, f"{product_name}\n", 'value')
            
            self.aop_results_text.insert(tk.END, f"  מידה:                                  ", 'label')
            self.aop_results_text.insert(tk.END, f"{product_size}\n", 'value')
            
            self.aop_results_text.insert(tk.END, f"  שטח רבוע בסיסי:                       ", 'label')
            self.aop_results_text.insert(tk.END, f"{square_area:.6f} מ״ר\n\n", 'value')
            
            self.aop_results_text.insert(tk.END, "-" * 70 + "\n\n", 'separator')
            
            # Calculation details
            self.aop_results_text.insert(tk.END, "נתוני חישוב:\n", 'header')
            self.aop_results_text.insert(tk.END, f"  אחוז פחת:                              ", 'label')
            self.aop_results_text.insert(tk.END, f"{waste_percentage:.2f}%\n", 'value')
            
            self.aop_results_text.insert(tk.END, f"  שטח בפועל (כולל פחת):                 ", 'label')
            self.aop_results_text.insert(tk.END, f"{actual_square_area:.6f} מ״ר\n", 'value')
            
            self.aop_results_text.insert(tk.END, f"  עלות הדפסה למטר רבוע:                 ", 'label')
            self.aop_results_text.insert(tk.END, f"{printing_cost_per_sqm:.2f} ₪/מ״ר\n\n", 'value')
            
            self.aop_results_text.insert(tk.END, "=" * 70 + "\n\n", 'separator')
            
            # Total cost
            self.aop_results_text.insert(tk.END, "עלות הדפסה כוללת: ", 'total')
            self.aop_results_text.insert(tk.END, f"{total_printing_cost:.2f} ₪\n\n", 'total')
            
            self.aop_results_text.insert(tk.END, "=" * 70 + "\n", 'separator')
            
            self.aop_results_text.config(state='disabled')
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בחישוב: {str(e)}")
    
    def _clear_all_over_print_inputs(self):
        """Clear all all over print input fields and results."""
        # Clear input variables
        self.aop_printing_cost_per_sqm_var.set("")
        self.aop_waste_percentage_var.set("")
        
        # Reset product selection display
        if hasattr(self, 'aop_square_area_label'):
            self.aop_square_area_label.config(text="בחר פריט", fg='#7f8c8d', font=('Arial', 10))
        
        # Clear results display
        self.aop_results_text.config(state='normal')
        self.aop_results_text.delete(1.0, tk.END)
        self.aop_results_text.insert(
            tk.END,
            "\n\n\n          בחר פריט והזן נתוני הדפסה לחישוב עלות\n\n\n",
            'header'
        )
        self.aop_results_text.config(state='disabled')
    
    def _build_store_price_content(self, container):
        """Build the store price calculator content."""
        
        # Main formula frame
        formula_frame = ttk.LabelFrame(container, text="נוסחת חישוב מחיר לצרכן", padding=20)
        formula_frame.pack(fill='x', padx=20, pady=10)
        
        # Formula display
        formula_text = "מחיר לצרכן = (מחיר לחנות × 1.17) × 2"
        formula_label = tk.Label(
            formula_frame,
            text=formula_text,
            font=('Arial', 14, 'bold'),
            bg='#f7f9fa',
            fg='#2c3e50',
            justify='center'
        )
        formula_label.pack(pady=10)
        
        # Formula breakdown
        breakdown_text = "שלב 1: מחיר כולל מע״מ = מחיר × 1.17\nשלב 2: מחיר לצרכן = מחיר כולל מע״מ × 2 (הוספת 100%)"
        breakdown_label = tk.Label(
            formula_frame,
            text=breakdown_text,
            font=('Arial', 10),
            bg='#f7f9fa',
            fg='#7f8c8d',
            justify='center'
        )
        breakdown_label.pack(pady=5)
        
        # Input fields frame
        inputs_frame = ttk.LabelFrame(container, text="הזן מחיר", padding=20)
        inputs_frame.pack(fill='x', padx=20, pady=10)
        
        # Store price input
        tk.Label(
            inputs_frame, 
            text="מחיר לחנות (לפני מע״מ):", 
            font=('Arial', 12, 'bold')
        ).grid(row=0, column=0, sticky='e', padx=10, pady=10)
        
        self.store_price_var = tk.StringVar()
        store_price_entry = tk.Entry(
            inputs_frame, 
            textvariable=self.store_price_var, 
            width=20, 
            font=('Arial', 12)
        )
        store_price_entry.grid(row=0, column=1, padx=10, pady=10)
        
        tk.Label(
            inputs_frame, 
            text="ש״ח", 
            font=('Arial', 12)
        ).grid(row=0, column=2, sticky='w', padx=5, pady=10)
        
        # Calculate button
        calculate_btn = tk.Button(
            inputs_frame,
            text="חשב מחיר לצרכן",
            command=self._calculate_store_price,
            bg='#27ae60',
            fg='white',
            font=('Arial', 12, 'bold'),
            width=20,
            cursor='hand2'
        )
        calculate_btn.grid(row=1, column=0, columnspan=3, pady=15)
        
        # Results frame
        results_frame = ttk.LabelFrame(container, text="תוצאות חישוב", padding=20)
        results_frame.pack(fill='x', padx=20, pady=10)
        
        # Price with VAT result
        tk.Label(
            results_frame, 
            text="מחיר כולל מע״מ:", 
            font=('Arial', 11, 'bold')
        ).grid(row=0, column=0, sticky='e', padx=10, pady=8)
        
        self.price_with_vat_var = tk.StringVar(value="--")
        price_with_vat_label = tk.Label(
            results_frame,
            textvariable=self.price_with_vat_var,
            font=('Arial', 14, 'bold'),
            fg='#3498db',
            bg='#f7f9fa'
        )
        price_with_vat_label.grid(row=0, column=1, padx=10, pady=8)
        
        tk.Label(
            results_frame, 
            text="ש״ח", 
            font=('Arial', 11)
        ).grid(row=0, column=2, sticky='w', padx=5, pady=8)
        
        # Consumer price result
        tk.Label(
            results_frame, 
            text="מחיר לצרכן:", 
            font=('Arial', 12, 'bold')
        ).grid(row=1, column=0, sticky='e', padx=10, pady=8)
        
        self.consumer_price_var = tk.StringVar(value="--")
        consumer_price_label = tk.Label(
            results_frame,
            textvariable=self.consumer_price_var,
            font=('Arial', 16, 'bold'),
            fg='#27ae60',
            bg='#f7f9fa'
        )
        consumer_price_label.grid(row=1, column=1, padx=10, pady=8)
        
        tk.Label(
            results_frame, 
            text="ש״ח", 
            font=('Arial', 12, 'bold')
        ).grid(row=1, column=2, sticky='w', padx=5, pady=8)
        
        # Example frame
        example_frame = ttk.LabelFrame(container, text="דוגמה", padding=15)
        example_frame.pack(fill='x', padx=20, pady=10)
        
        example_text = """
📌 דוגמה לחישוב:

מחיר לחנות: 26 ש״ח
↓
מחיר כולל מע״מ: 26 × 1.17 = 30.42 ש״ח
↓
מחיר לצרכן: 30.42 × 2 = 60.84 ש״ח
        """
        
        example_label = tk.Label(
            example_frame,
            text=example_text,
            font=('Arial', 10),
            bg='#f7f9fa',
            fg='#34495e',
            justify='right'
        )
        example_label.pack()
        
        # Clear button
        clear_btn = tk.Button(
            container,
            text="נקה",
            command=self._clear_store_price_inputs,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        )
        clear_btn.pack(pady=20)
    
    def _calculate_store_price(self):
        """Calculate consumer price from store price."""
        try:
            store_price = float(self.store_price_var.get() or 0)
            
            if store_price <= 0:
                messagebox.showwarning("אזהרה", "אנא הזן מחיר חיובי")
                return
            
            # Calculate price with VAT (17%)
            price_with_vat = store_price * 1.17
            
            # Calculate consumer price (100% markup = x2)
            consumer_price = price_with_vat * 2
            
            # Update result fields
            self.price_with_vat_var.set(f"{price_with_vat:.2f}")
            self.consumer_price_var.set(f"{consumer_price:.2f}")
            
        except ValueError:
            messagebox.showerror("שגיאה", "אנא הזן מספר תקין")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בחישוב: {str(e)}")
    
    def _clear_store_price_inputs(self):
        """Clear all store price calculator inputs and results."""
        self.store_price_var.set("")
        self.price_with_vat_var.set("--")
        self.consumer_price_var.set("--")
    
    def _build_fabric_rolls_content(self, container):
        """Build the fabric rolls calculator content."""
        
        # Title and description
        title_frame = tk.Frame(container, bg='#f7f9fa')
        title_frame.pack(fill='x', padx=20, pady=(10, 5))
        
        tk.Label(
            title_frame,
            text="חישוב גלילי בד לגיזרה",
            font=('Arial', 14, 'bold'),
            bg='#f7f9fa',
            fg='#2c3e50'
        ).pack()
        
        tk.Label(
            title_frame,
            text="חישוב כמות גלילי הבד הנדרשים לגיזרה - כולל קליברציה מעבודה קודמת",
            font=('Arial', 9),
            bg='#f7f9fa',
            fg='#7f8c8d'
        ).pack()
        
        # ============ STEPS CONTAINER (Side by Side) ============
        steps_container = tk.Frame(container, bg='#f7f9fa')
        steps_container.pack(fill='x', padx=20, pady=10)
        steps_container.grid_columnconfigure(0, weight=1)
        steps_container.grid_columnconfigure(1, weight=1)
        
        # ============ STEP 1: Calibration (Right Side) ============
        calibration_frame = ttk.LabelFrame(steps_container, text="שלב 1: קליברציה - חישוב אורך גליל", padding=10)
        calibration_frame.grid(row=0, column=1, sticky='nsew', padx=(5, 0))
        
        # Row 0: Previous drawing length
        tk.Label(calibration_frame, text="אורך ציור קודם (מטר):", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        self.rolls_prev_drawing_length_var = tk.StringVar()
        tk.Entry(calibration_frame, textvariable=self.rolls_prev_drawing_length_var, width=10, font=('Arial', 10)).grid(
            row=0, column=1, padx=5, pady=5, sticky='w')
        
        # Row 1: Previous layers
        tk.Label(calibration_frame, text="כמות שכבות:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky='e', padx=5, pady=5)
        self.rolls_prev_layers_var = tk.StringVar()
        tk.Entry(calibration_frame, textvariable=self.rolls_prev_layers_var, width=10, font=('Arial', 10)).grid(
            row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Row 2: Rolls used
        tk.Label(calibration_frame, text="כמות גלילים שנצרכו:", font=('Arial', 10, 'bold')).grid(
            row=2, column=0, sticky='e', padx=5, pady=5)
        self.rolls_prev_rolls_used_var = tk.StringVar()
        tk.Entry(calibration_frame, textvariable=self.rolls_prev_rolls_used_var, width=10, font=('Arial', 10)).grid(
            row=2, column=1, padx=5, pady=5, sticky='w')
        
        # Row 3: Calculate button
        calc_roll_btn = tk.Button(
            calibration_frame,
            text="חשב אורך גליל",
            command=self._calculate_roll_length,
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        )
        calc_roll_btn.grid(row=3, column=0, columnspan=2, padx=5, pady=10)
        
        # Row 4: Roll length result
        result_frame = tk.Frame(calibration_frame)
        result_frame.grid(row=4, column=0, columnspan=2, pady=5)
        tk.Label(result_frame, text="אורך ממוצע לגליל:", font=('Arial', 10, 'bold')).pack(side='right', padx=2)
        self.rolls_avg_roll_length_var = tk.StringVar(value="--")
        self.rolls_avg_roll_length_label = tk.Label(
            result_frame,
            textvariable=self.rolls_avg_roll_length_var,
            font=('Arial', 12, 'bold'),
            fg='#27ae60'
        )
        self.rolls_avg_roll_length_label.pack(side='right', padx=2)
        tk.Label(result_frame, text="מטר", font=('Arial', 10)).pack(side='right', padx=2)
        
        # Row 5: Manual entry
        manual_frame = tk.Frame(calibration_frame)
        manual_frame.grid(row=5, column=0, columnspan=2, pady=5)
        tk.Label(manual_frame, text="או הזן ידנית:", font=('Arial', 9), fg='#7f8c8d').pack(side='right', padx=2)
        self.rolls_manual_roll_length_var = tk.StringVar()
        manual_entry = tk.Entry(manual_frame, textvariable=self.rolls_manual_roll_length_var, width=8, font=('Arial', 10))
        manual_entry.pack(side='right', padx=2)
        use_manual_btn = tk.Button(
            manual_frame,
            text="השתמש בערך ידני",
            command=self._use_manual_roll_length,
            bg='#95a5a6',
            fg='white',
            font=('Arial', 9, 'bold')
        )
        use_manual_btn.pack(side='right', padx=2)
        
        # ============ STEP 2: New Job Calculation (Left Side) ============
        new_job_frame = ttk.LabelFrame(steps_container, text="שלב 2: חישוב לעבודה חדשה", padding=10)
        new_job_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        
        # Row 0: Drawing selection
        tk.Label(new_job_frame, text="בחר ציור (אופציונלי):", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='e', padx=5, pady=5)
        
        drawing_select_frame = tk.Frame(new_job_frame)
        drawing_select_frame.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        self.rolls_drawing_var = tk.StringVar()
        self.rolls_drawing_combo = ttk.Combobox(
            drawing_select_frame,
            textvariable=self.rolls_drawing_var,
            state='readonly',
            width=20,
            justify='right'
        )
        self.rolls_drawing_combo.pack(side='right', padx=2)
        self.rolls_drawing_combo.bind('<<ComboboxSelected>>', self._on_rolls_drawing_selected)
        
        load_btn = tk.Button(
            drawing_select_frame,
            text="טען ציורים",
            command=self._load_rolls_drawings_list,
            bg='#3498db',
            fg='white',
            font=('Arial', 9, 'bold')
        )
        load_btn.pack(side='right', padx=2)
        
        # Row 1: New drawing length
        tk.Label(new_job_frame, text="אורך ציור חדש (מטר):", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky='e', padx=5, pady=8)
        self.rolls_new_drawing_length_var = tk.StringVar()
        tk.Entry(new_job_frame, textvariable=self.rolls_new_drawing_length_var, width=12, font=('Arial', 10)).grid(
            row=1, column=1, padx=5, pady=8, sticky='w')
        
        # Row 2: New layers count
        tk.Label(new_job_frame, text="כמות שכבות:", font=('Arial', 10, 'bold')).grid(
            row=2, column=0, sticky='e', padx=5, pady=8)
        self.rolls_new_layers_var = tk.StringVar()
        tk.Entry(new_job_frame, textvariable=self.rolls_new_layers_var, width=12, font=('Arial', 10)).grid(
            row=2, column=1, padx=5, pady=8, sticky='w')
        
        # Buttons frame
        buttons_frame = tk.Frame(container, bg='#f7f9fa')
        buttons_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Button(
            buttons_frame,
            text="חשב גלילים נדרשים",
            command=self._calculate_fabric_rolls,
            bg='#27ae60',
            fg='white',
            font=('Arial', 11, 'bold'),
            width=20,
            height=2
        ).pack(side='left', padx=5)
        
        tk.Button(
            buttons_frame,
            text="נקה הכל",
            command=self._clear_fabric_rolls_inputs,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 11, 'bold'),
            width=15,
            height=2
        ).pack(side='left', padx=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(container, text="תוצאות החישוב", padding=20)
        results_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Results display
        self.rolls_results_text = tk.Text(
            results_frame,
            height=12,
            font=('Courier New', 10),
            wrap=tk.WORD,
            state='disabled',
            bg='#ffffff'
        )
        self.rolls_results_text.pack(fill='both', expand=True)
        
        # Configure text tags for formatting
        self.rolls_results_text.tag_configure('header', font=('Arial', 11, 'bold'), foreground='#2c3e50')
        self.rolls_results_text.tag_configure('label', font=('Courier New', 10), foreground='#34495e')
        self.rolls_results_text.tag_configure('value', font=('Courier New', 10, 'bold'), foreground='#2980b9')
        self.rolls_results_text.tag_configure('total', font=('Arial', 14, 'bold'), foreground='#27ae60')
        self.rolls_results_text.tag_configure('warning', font=('Arial', 10, 'bold'), foreground='#e67e22')
        self.rolls_results_text.tag_configure('separator', foreground='#7f8c8d')
        
        # Initialize
        self.rolls_drawings_dict = {}
        self.calculated_roll_length = None
        self._load_rolls_drawings_list()
        self._clear_fabric_rolls_inputs()
    
    def _calculate_roll_length(self):
        """Calculate average roll length from previous job data."""
        try:
            prev_length_str = self.rolls_prev_drawing_length_var.get().strip()
            prev_layers_str = self.rolls_prev_layers_var.get().strip()
            prev_rolls_str = self.rolls_prev_rolls_used_var.get().strip()
            
            if not prev_length_str or not prev_layers_str or not prev_rolls_str:
                messagebox.showwarning("אזהרה", "אנא הזן את כל נתוני העבודה הקודמת")
                return
            
            try:
                prev_length = float(prev_length_str)
                prev_layers = float(prev_layers_str)
                prev_rolls = float(prev_rolls_str)
            except ValueError:
                messagebox.showwarning("אזהרה", "אנא הזן ערכים נומריים תקינים")
                return
            
            if prev_length <= 0 or prev_layers <= 0 or prev_rolls <= 0:
                messagebox.showwarning("אזהרה", "כל הערכים חייבים להיות גדולים מ-0")
                return
            
            # Calculate: total fabric used / number of rolls = avg length per roll
            total_fabric = prev_length * prev_layers
            avg_roll_length = total_fabric / prev_rolls
            
            self.calculated_roll_length = avg_roll_length
            self.rolls_avg_roll_length_var.set(f"{avg_roll_length:.2f}")
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בחישוב: {str(e)}")
    
    def _use_manual_roll_length(self):
        """Use manually entered roll length."""
        try:
            manual_str = self.rolls_manual_roll_length_var.get().strip()
            if not manual_str:
                messagebox.showwarning("אזהרה", "אנא הזן אורך גליל ידני")
                return
            
            try:
                manual_length = float(manual_str)
            except ValueError:
                messagebox.showwarning("אזהרה", "אנא הזן ערך נומרי תקין")
                return
            
            if manual_length <= 0:
                messagebox.showwarning("אזהרה", "אורך הגליל חייב להיות גדול מ-0")
                return
            
            self.calculated_roll_length = manual_length
            self.rolls_avg_roll_length_var.set(f"{manual_length:.2f}")
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה: {str(e)}")
    
    def _load_rolls_drawings_list(self):
        """Load the list of drawings for the fabric rolls calculator."""
        try:
            drawings = getattr(self.data_processor, 'drawings_data', [])
            drawing_names = []
            self.rolls_drawings_dict = {}
            
            for drawing in drawings:
                drawing_name = drawing.get('שם הקובץ', f"ציור {drawing.get('id', 'לא ידוע')}")
                drawing_names.append(drawing_name)
                self.rolls_drawings_dict[drawing_name] = drawing
            
            self.rolls_drawing_combo['values'] = [''] + drawing_names
            self.rolls_drawing_combo.set('')
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בטעינת רשימת הציורים: {str(e)}")
    
    def _on_rolls_drawing_selected(self, event=None):
        """Handle drawing selection for fabric rolls calculator."""
        try:
            selected_name = self.rolls_drawing_var.get()
            if not selected_name or selected_name not in self.rolls_drawings_dict:
                return
            
            drawing = self.rolls_drawings_dict[selected_name]
            
            # Get drawing length - try multiple possible field names
            drawing_length = drawing.get('אורך', None) or drawing.get('אורך הציור', None) or drawing.get('length', None)
            if drawing_length is not None:
                self.rolls_new_drawing_length_var.set(str(drawing_length))
            
            # Get layers count - try multiple possible field names
            layers = drawing.get('שכבות', None) or drawing.get('כמות שכבות משוערת', None) or drawing.get('layers', None)
            if layers is not None:
                self.rolls_new_layers_var.set(str(layers))
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בטעינת נתוני הציור: {str(e)}")
    
    def _calculate_fabric_rolls(self):
        """Calculate the number of fabric rolls needed for new job."""
        try:
            # Check if roll length is set
            if self.calculated_roll_length is None or self.calculated_roll_length <= 0:
                messagebox.showwarning("אזהרה", "אנא חשב או הזן אורך גליל תחילה (שלב 1)")
                return
            
            # Get new job input values
            new_length_str = self.rolls_new_drawing_length_var.get().strip()
            new_layers_str = self.rolls_new_layers_var.get().strip()
            
            if not new_length_str:
                messagebox.showwarning("אזהרה", "אנא הזן אורך ציור חדש")
                return
            
            if not new_layers_str:
                messagebox.showwarning("אזהרה", "אנא הזן כמות שכבות")
                return
            
            try:
                new_length = float(new_length_str)
                new_layers = float(new_layers_str)
            except ValueError:
                messagebox.showwarning("אזהרה", "אנא הזן ערכים נומריים תקינים")
                return
            
            if new_length <= 0 or new_layers <= 0:
                messagebox.showwarning("אזהרה", "כל הערכים חייבים להיות גדולים מ-0")
                return
            
            # Perform calculation
            import math
            
            roll_length = self.calculated_roll_length
            total_fabric_needed = new_length * new_layers
            exact_rolls = total_fabric_needed / roll_length
            whole_rolls = math.ceil(exact_rolls)
            
            # Calculate leftover
            leftover = (whole_rolls * roll_length) - total_fabric_needed
            leftover_percentage = (leftover / (whole_rolls * roll_length)) * 100 if whole_rolls > 0 else 0
            
            # Display results
            self.rolls_results_text.config(state='normal')
            self.rolls_results_text.delete(1.0, tk.END)
            
            # Header
            self.rolls_results_text.insert(tk.END, "תוצאות חישוב גלילי בד\n", 'header')
            self.rolls_results_text.insert(tk.END, "=" * 70 + "\n\n", 'separator')
            
            # Roll length used
            self.rolls_results_text.insert(tk.END, "נתוני גליל:\n", 'header')
            self.rolls_results_text.insert(tk.END, f"  אורך ממוצע לגליל:                    ", 'label')
            self.rolls_results_text.insert(tk.END, f"{roll_length:.2f} מטר\n\n", 'value')
            
            # New job data
            self.rolls_results_text.insert(tk.END, "נתוני עבודה חדשה:\n", 'header')
            self.rolls_results_text.insert(tk.END, f"  אורך הציור:                          ", 'label')
            self.rolls_results_text.insert(tk.END, f"{new_length:.2f} מטר\n", 'value')
            
            self.rolls_results_text.insert(tk.END, f"  כמות שכבות:                          ", 'label')
            self.rolls_results_text.insert(tk.END, f"{new_layers:.0f}\n\n", 'value')
            
            self.rolls_results_text.insert(tk.END, "-" * 70 + "\n\n", 'separator')
            
            # Calculation details
            self.rolls_results_text.insert(tk.END, "חישוב:\n", 'header')
            self.rolls_results_text.insert(tk.END, f"  סה״כ בד נדרש:                        ", 'label')
            self.rolls_results_text.insert(tk.END, f"{total_fabric_needed:.2f} מטר\n", 'value')
            
            self.rolls_results_text.insert(tk.END, f"  ({new_length:.2f} × {new_layers:.0f} = {total_fabric_needed:.2f})\n\n", 'label')
            
            self.rolls_results_text.insert(tk.END, f"  מספר גלילים מדויק:                   ", 'label')
            self.rolls_results_text.insert(tk.END, f"{exact_rolls:.2f}\n\n", 'value')
            
            self.rolls_results_text.insert(tk.END, "=" * 70 + "\n\n", 'separator')
            
            # Main result
            self.rolls_results_text.insert(tk.END, f"מספר גלילים נדרש: {whole_rolls}\n\n", 'total')
            
            # Exact rolls info
            self.rolls_results_text.insert(tk.END, f"(מספר מדויק: {exact_rolls:.2f} גלילים)\n\n", 'warning')
            
            # Leftover info
            if leftover > 0 and whole_rolls > exact_rolls:
                self.rolls_results_text.insert(tk.END, f"עודף צפוי: {leftover:.2f} מטר ({leftover_percentage:.1f}%)\n", 'label')
            
            self.rolls_results_text.insert(tk.END, "\n" + "=" * 70 + "\n", 'separator')
            
            self.rolls_results_text.config(state='disabled')
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בחישוב: {str(e)}")
    
    def _clear_fabric_rolls_inputs(self):
        """Clear all fabric rolls calculator inputs and results."""
        # Clear calibration inputs
        if hasattr(self, 'rolls_prev_drawing_length_var'):
            self.rolls_prev_drawing_length_var.set("")
        if hasattr(self, 'rolls_prev_layers_var'):
            self.rolls_prev_layers_var.set("")
        if hasattr(self, 'rolls_prev_rolls_used_var'):
            self.rolls_prev_rolls_used_var.set("")
        if hasattr(self, 'rolls_manual_roll_length_var'):
            self.rolls_manual_roll_length_var.set("")
        if hasattr(self, 'rolls_avg_roll_length_var'):
            self.rolls_avg_roll_length_var.set("--")
        
        # Clear new job inputs
        if hasattr(self, 'rolls_new_drawing_length_var'):
            self.rolls_new_drawing_length_var.set("")
        if hasattr(self, 'rolls_new_layers_var'):
            self.rolls_new_layers_var.set("")
        
        # Clear drawing selection
        if hasattr(self, 'rolls_drawing_combo'):
            self.rolls_drawing_combo.set('')
        if hasattr(self, 'rolls_drawing_var'):
            self.rolls_drawing_var.set('')
        
        # Reset calculated roll length
        self.calculated_roll_length = None
        
        # Clear results display
        if hasattr(self, 'rolls_results_text'):
            self.rolls_results_text.config(state='normal')
            self.rolls_results_text.delete(1.0, tk.END)
            self.rolls_results_text.insert(
                tk.END,
                "\n\n          שלב 1: הזן נתוני עבודה קודמת לחישוב אורך גליל\n"
                "          שלב 2: הזן נתוני עבודה חדשה לחישוב כמות הגלילים\n\n",
                'header'
            )
            self.rolls_results_text.config(state='disabled')
    
    # =====================================================================
    # Square Meter Cost Tab - חישוב עלות למ"ר
    # =====================================================================
    
    def _build_sqm_cost_content(self, container):
        """Build the square meter cost calculation content with 2-column layout."""
        
        # Initialize selected drawings list
        self.selected_sqm_drawings = []
        
        # Main 2-column container
        main_frame = tk.Frame(container, bg='#f7f9fa')
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Configure grid weights for 2 columns
        main_frame.columnconfigure(0, weight=1)  # Left column
        main_frame.columnconfigure(1, weight=1)  # Right column
        main_frame.rowconfigure(0, weight=1)
        
        # ========== LEFT COLUMN (Input) ==========
        left_frame = tk.Frame(main_frame, bg='#f7f9fa')
        left_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        
        # Instructions
        instructions_frame = ttk.LabelFrame(left_frame, text="הסבר", padding=10)
        instructions_frame.pack(fill='x', pady=(0, 10))
        
        instructions_text = (
            "1. בחר ציורים שנחתכו מהרשימה\n"
            "2. לחץ 'הוסף לחישוב' להוספה לרשימה\n"
            "3. הזן מחיר ל-1 ק\"ג בד\n"
            "4. לחץ 'חשב ממוצע' לקבלת עלות ממוצעת למ\"ר"
        )
        tk.Label(
            instructions_frame,
            text=instructions_text,
            font=('Arial', 9),
            bg='#f7f9fa',
            fg='#2c3e50',
            justify='right',
            anchor='e'
        ).pack(fill='x')
        
        # Available drawings frame
        available_frame = ttk.LabelFrame(left_frame, text="ציורים זמינים (נחתכו)", padding=10)
        available_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Listbox for available drawings (multi-select)
        listbox_frame = tk.Frame(available_frame)
        listbox_frame.pack(fill='both', expand=True)
        
        self.sqm_drawings_listbox = tk.Listbox(
            listbox_frame,
            selectmode=tk.EXTENDED,
            font=('Arial', 9),
            height=8,
            exportselection=False
        )
        listbox_scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=self.sqm_drawings_listbox.yview)
        self.sqm_drawings_listbox.configure(yscrollcommand=listbox_scrollbar.set)
        
        self.sqm_drawings_listbox.pack(side='left', fill='both', expand=True)
        listbox_scrollbar.pack(side='right', fill='y')
        
        # Buttons for add/refresh
        btn_frame1 = tk.Frame(available_frame, bg='#f7f9fa')
        btn_frame1.pack(fill='x', pady=(5, 0))
        
        tk.Button(
            btn_frame1,
            text="➕ הוסף לחישוב",
            command=self._add_drawings_to_selection,
            bg='#27ae60',
            fg='white',
            font=('Arial', 9, 'bold')
        ).pack(side='right', padx=2)
        
        tk.Button(
            btn_frame1,
            text="🔄 רענן",
            command=self._load_sqm_cut_drawings,
            bg='#3498db',
            fg='white',
            font=('Arial', 9, 'bold')
        ).pack(side='right', padx=2)
        
        # Selected drawings frame
        selected_frame = ttk.LabelFrame(left_frame, text="ציורים נבחרים לחישוב", padding=10)
        selected_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Treeview for selected drawings
        tree_frame = tk.Frame(selected_frame)
        tree_frame.pack(fill='both', expand=True)
        
        columns = ('id', 'fabric', 'gsm', 'cost_sqm')
        self.sqm_selected_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=6)
        
        self.sqm_selected_tree.heading('id', text='ID')
        self.sqm_selected_tree.heading('fabric', text='סוג בד')
        self.sqm_selected_tree.heading('gsm', text='משקל למ"ר')
        self.sqm_selected_tree.heading('cost_sqm', text='עלות למ"ר')
        
        self.sqm_selected_tree.column('id', width=50, anchor='center')
        self.sqm_selected_tree.column('fabric', width=100, anchor='center')
        self.sqm_selected_tree.column('gsm', width=80, anchor='center')
        self.sqm_selected_tree.column('cost_sqm', width=80, anchor='center')
        
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.sqm_selected_tree.yview)
        self.sqm_selected_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        self.sqm_selected_tree.pack(side='left', fill='both', expand=True)
        tree_scrollbar.pack(side='right', fill='y')
        
        # Remove button
        tk.Button(
            selected_frame,
            text="🗑️ הסר נבחרים",
            command=self._remove_selected_drawings,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 9, 'bold')
        ).pack(anchor='e', pady=(5, 0))
        
        # Price input frame
        input_frame = ttk.LabelFrame(left_frame, text="קלט מחיר", padding=10)
        input_frame.pack(fill='x', pady=(0, 10))
        
        price_row = tk.Frame(input_frame, bg='#f7f9fa')
        price_row.pack(fill='x')
        
        tk.Label(price_row, text="מחיר ל-1 ק\"ג בד (₪):", font=('Arial', 10, 'bold'), bg='#f7f9fa').pack(side='right', padx=5)
        self.sqm_price_per_kg_var = tk.StringVar()
        tk.Entry(price_row, textvariable=self.sqm_price_per_kg_var, width=12, font=('Arial', 11)).pack(side='right', padx=5)
        
        # Buttons row
        btn_row = tk.Frame(input_frame, bg='#f7f9fa')
        btn_row.pack(fill='x', pady=(10, 0))
        
        tk.Button(
            btn_row,
            text="📊 חשב ממוצע עלות למ\"ר",
            command=self._calculate_sqm_cost,
            bg='#27ae60',
            fg='white',
            font=('Arial', 11, 'bold'),
            padx=15,
            pady=5
        ).pack(side='right', padx=5)
        
        tk.Button(
            btn_row,
            text="🗑️ נקה הכל",
            command=self._clear_sqm_cost_inputs,
            bg='#95a5a6',
            fg='white',
            font=('Arial', 9, 'bold')
        ).pack(side='right', padx=5)
        
        # ========== RIGHT COLUMN (Results) ==========
        right_frame = tk.Frame(main_frame, bg='#f7f9fa')
        right_frame.grid(row=0, column=1, sticky='nsew', padx=(5, 0))
        
        results_frame = ttk.LabelFrame(right_frame, text="תוצאות החישוב", padding=10)
        results_frame.pack(fill='both', expand=True)
        
        # Results text widget
        self.sqm_results_text = tk.Text(
            results_frame,
            font=('Consolas', 10),
            bg='#2c3e50',
            fg='#ecf0f1',
            insertbackground='white',
            state='disabled',
            wrap='word'
        )
        results_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.sqm_results_text.yview)
        self.sqm_results_text.configure(yscrollcommand=results_scrollbar.set)
        
        self.sqm_results_text.pack(side='left', fill='both', expand=True)
        results_scrollbar.pack(side='right', fill='y')
        
        # Configure tags for results
        self.sqm_results_text.tag_configure('header', foreground='#f39c12', font=('Consolas', 11, 'bold'))
        self.sqm_results_text.tag_configure('subheader', foreground='#3498db', font=('Consolas', 10, 'bold'))
        self.sqm_results_text.tag_configure('label', foreground='#bdc3c7', font=('Consolas', 9))
        self.sqm_results_text.tag_configure('value', foreground='#2ecc71', font=('Consolas', 10, 'bold'))
        self.sqm_results_text.tag_configure('total', foreground='#e74c3c', font=('Consolas', 14, 'bold'))
        self.sqm_results_text.tag_configure('average', foreground='#f1c40f', font=('Consolas', 16, 'bold'))
        self.sqm_results_text.tag_configure('separator', foreground='#7f8c8d')
        
        # Initial welcome message
        self.sqm_results_text.config(state='normal')
        self.sqm_results_text.insert(tk.END, "\n\n    בחר ציורים והזן מחיר לק\"ג\n", 'header')
        self.sqm_results_text.insert(tk.END, "    לחישוב ממוצע עלות למ\"ר\n\n", 'header')
        self.sqm_results_text.config(state='disabled')
        
        # Load cut drawings
        self._load_sqm_cut_drawings()
    
    def _add_drawings_to_selection(self):
        """Add selected drawings from listbox to the calculation list."""
        try:
            selected_indices = self.sqm_drawings_listbox.curselection()
            if not selected_indices:
                messagebox.showinfo("הודעה", "נא לבחור ציורים מהרשימה")
                return
            
            price_per_kg_str = self.sqm_price_per_kg_var.get()
            if not price_per_kg_str:
                messagebox.showerror("שגיאה", "נא להזין מחיר לק\"ג לפני הוספת ציורים")
                return
            
            try:
                price_per_kg = float(price_per_kg_str)
                if price_per_kg <= 0:
                    messagebox.showerror("שגיאה", "המחיר חייב להיות גדול מ-0")
                    return
            except ValueError:
                messagebox.showerror("שגיאה", "נא להזין מחיר תקין")
                return
            
            for idx in selected_indices:
                display_text = self.sqm_drawings_listbox.get(idx)
                drawing_id = self.sqm_drawing_id_map.get(display_text)
                
                # Check if already added
                if any(d['id'] == drawing_id for d in self.selected_sqm_drawings):
                    continue
                
                # Find the drawing data
                drawing = None
                for d in self.data_processor.drawings_data:
                    if d.get('id') == drawing_id:
                        drawing = d
                        break
                
                if drawing:
                    # Get GSM (weight per sqm)
                    gsm = drawing.get('משקל כולל', drawing.get('משקל_כולל', 0))
                    try:
                        gsm = float(gsm) if gsm else 0
                    except (ValueError, TypeError):
                        gsm = 0
                    
                    # Calculate cost per sqm
                    cost_per_sqm = (gsm / 1000) * price_per_kg if gsm > 0 else 0
                    
                    # Add to list
                    self.selected_sqm_drawings.append({
                        'id': drawing_id,
                        'fabric': drawing.get('סוג בד', '--'),
                        'gsm': gsm,
                        'cost_per_sqm': cost_per_sqm,
                        'drawing': drawing
                    })
            
            # Update the treeview
            self._update_selected_drawings_tree()
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בהוספת ציורים: {str(e)}")
    
    def _update_selected_drawings_tree(self):
        """Update the selected drawings treeview."""
        # Clear existing items
        for item in self.sqm_selected_tree.get_children():
            self.sqm_selected_tree.delete(item)
        
        # Add all selected drawings
        for d in self.selected_sqm_drawings:
            self.sqm_selected_tree.insert('', 'end', values=(
                d['id'],
                d['fabric'],
                f"{d['gsm']:.1f}",
                f"{d['cost_per_sqm']:.2f} ₪"
            ))
    
    def _remove_selected_drawings(self):
        """Remove selected drawings from the calculation list."""
        selected_items = self.sqm_selected_tree.selection()
        if not selected_items:
            messagebox.showinfo("הודעה", "נא לבחור ציורים להסרה")
            return
        
        # Get IDs to remove (convert to string for comparison)
        ids_to_remove = []
        for item in selected_items:
            values = self.sqm_selected_tree.item(item, 'values')
            if values:
                ids_to_remove.append(str(values[0]))
        
        # Remove from list (compare as strings)
        self.selected_sqm_drawings = [d for d in self.selected_sqm_drawings if str(d['id']) not in ids_to_remove]
        
        # Update treeview
        self._update_selected_drawings_tree()
    
    def _load_sqm_cut_drawings(self):
        """Load cut drawings that have actual layers into the listbox."""
        try:
            drawings = self.data_processor.drawings_data
            cut_drawings = []
            
            for d in drawings:
                # Check if drawing is cut (status = "נחתך") and has layers
                status = d.get('status', d.get('סטטוס', ''))
                layers = d.get('שכבות', 0)
                
                if status == 'נחתך' and layers and int(layers) > 0:
                    drawing_id = d.get('id', '')
                    file_name = d.get('שם הקובץ', '')
                    fabric_type = d.get('סוג בד', '')
                    gsm = d.get('משקל כולל', d.get('משקל_כולל', 0))
                    try:
                        gsm = float(gsm) if gsm else 0
                    except (ValueError, TypeError):
                        gsm = 0
                    display_text = f"ID: {drawing_id} - {file_name} ({fabric_type}) - {gsm:.0f}g/m²"
                    cut_drawings.append((drawing_id, display_text))
            
            # Sort by ID
            cut_drawings.sort(key=lambda x: x[0])
            
            # Clear and update listbox
            self.sqm_drawings_listbox.delete(0, tk.END)
            for _, display_text in cut_drawings:
                self.sqm_drawings_listbox.insert(tk.END, display_text)
            
            # Store mapping
            self.sqm_drawing_id_map = {item[1]: item[0] for item in cut_drawings}
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בטעינת ציורים: {str(e)}")
    
    
    def _calculate_sqm_cost(self):
        """Calculate average cost per square meter for selected drawings."""
        try:
            # Validate selected drawings
            if not self.selected_sqm_drawings:
                messagebox.showerror("שגיאה", "נא להוסיף ציורים לחישוב")
                return
            
            price_per_kg_str = self.sqm_price_per_kg_var.get()
            if not price_per_kg_str:
                messagebox.showerror("שגיאה", "נא להזין מחיר לק\"ג")
                return
            
            try:
                price_per_kg = float(price_per_kg_str)
                if price_per_kg <= 0:
                    messagebox.showerror("שגיאה", "המחיר חייב להיות גדול מ-0")
                    return
            except ValueError:
                messagebox.showerror("שגיאה", "נא להזין מחיר תקין")
                return
            
            # Recalculate costs for all drawings with current price
            total_cost_per_sqm = 0
            valid_drawings = 0
            
            for d in self.selected_sqm_drawings:
                gsm = d['gsm']
                if gsm > 0:
                    d['cost_per_sqm'] = (gsm / 1000) * price_per_kg
                    total_cost_per_sqm += d['cost_per_sqm']
                    valid_drawings += 1
            
            # Update treeview with new costs
            self._update_selected_drawings_tree()
            
            if valid_drawings == 0:
                messagebox.showerror("שגיאה", "אין ציורים עם משקל תקין")
                return
            
            # Calculate average
            average_cost_per_sqm = total_cost_per_sqm / valid_drawings
            
            # Display results
            self.sqm_results_text.config(state='normal')
            self.sqm_results_text.delete(1.0, tk.END)
            
            self.sqm_results_text.insert(tk.END, "=" * 45 + "\n", 'separator')
            self.sqm_results_text.insert(tk.END, "   תוצאות חישוב עלות למ\"ר\n", 'header')
            self.sqm_results_text.insert(tk.END, "=" * 45 + "\n\n", 'separator')
            
            self.sqm_results_text.insert(tk.END, f"מחיר לק\"ג: {price_per_kg:.2f} ₪\n", 'value')
            self.sqm_results_text.insert(tk.END, f"כמות ציורים: {valid_drawings}\n\n", 'value')
            
            self.sqm_results_text.insert(tk.END, "-" * 45 + "\n", 'separator')
            self.sqm_results_text.insert(tk.END, "פירוט לכל ציור:\n", 'subheader')
            self.sqm_results_text.insert(tk.END, "-" * 45 + "\n\n", 'separator')
            
            # Show each drawing's calculation
            for i, d in enumerate(self.selected_sqm_drawings, 1):
                drawing = d.get('drawing', {})
                drawing_id = d['id']
                fabric = d['fabric']
                gsm = d['gsm']
                cost = d['cost_per_sqm']
                
                self.sqm_results_text.insert(tk.END, f"ציור #{i} (ID: {drawing_id})\n", 'subheader')
                self.sqm_results_text.insert(tk.END, f"  סוג בד: ", 'label')
                self.sqm_results_text.insert(tk.END, f"{fabric}\n", 'value')
                self.sqm_results_text.insert(tk.END, f"  משקל למ\"ר: ", 'label')
                self.sqm_results_text.insert(tk.END, f"{gsm:.1f} גרם\n", 'value')
                self.sqm_results_text.insert(tk.END, f"  חישוב: ", 'label')
                self.sqm_results_text.insert(tk.END, f"({gsm:.1f}/1000) × {price_per_kg:.2f}\n", 'label')
                self.sqm_results_text.insert(tk.END, f"  עלות למ\"ר: ", 'label')
                self.sqm_results_text.insert(tk.END, f"{cost:.2f} ₪\n\n", 'total')
            
            self.sqm_results_text.insert(tk.END, "=" * 45 + "\n", 'separator')
            self.sqm_results_text.insert(tk.END, "\n", 'separator')
            
            # Average result - highlighted
            self.sqm_results_text.insert(tk.END, "  ממוצע עלות למ\"ר:\n", 'header')
            self.sqm_results_text.insert(tk.END, f"  {average_cost_per_sqm:.2f} ₪\n\n", 'average')
            
            self.sqm_results_text.insert(tk.END, "=" * 45 + "\n", 'separator')
            
            self.sqm_results_text.config(state='disabled')
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בחישוב: {str(e)}")
    
    def _clear_sqm_cost_inputs(self):
        """Clear all SQM cost calculation inputs and results."""
        # Clear selected drawings list
        self.selected_sqm_drawings = []
        
        # Clear listbox selection
        if hasattr(self, 'sqm_drawings_listbox'):
            self.sqm_drawings_listbox.selection_clear(0, tk.END)
        
        # Clear selected drawings treeview
        if hasattr(self, 'sqm_selected_tree'):
            for item in self.sqm_selected_tree.get_children():
                self.sqm_selected_tree.delete(item)
        
        # Clear price input
        if hasattr(self, 'sqm_price_per_kg_var'):
            self.sqm_price_per_kg_var.set('')
        
        # Clear results
        if hasattr(self, 'sqm_results_text'):
            self.sqm_results_text.config(state='normal')
            self.sqm_results_text.delete(1.0, tk.END)
            self.sqm_results_text.insert(tk.END, "\n\n    בחר ציורים והזן מחיר לק\"ג\n", 'header')
            self.sqm_results_text.insert(tk.END, "    לחישוב ממוצע עלות למ\"ר\n\n", 'header')
            self.sqm_results_text.config(state='disabled')