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
        
        inner_nb.add(weight_calc_tab, text="חישובי משקל כללי")
        inner_nb.add(fabric_weight_tab, text="משקל בד לפריטים בציור")
        inner_nb.add(product_cost_tab, text="חישוב שמיכות")
        inner_nb.add(tetra_cost_tab, text="חישוב טטרות")
        inner_nb.add(all_over_print_tab, text="בגדי אול אובר")
        inner_nb.add(store_price_tab, text="מחיר לצרכן חנויות")
        
        # Build content for each sub-tab
        self._build_general_weight_content(weight_calc_tab)
        self._build_fabric_weight_content(fabric_weight_tab)
        self._build_product_cost_content(product_cost_tab)
        self._build_tetra_cost_content(tetra_cost_tab)
        self._build_all_over_print_content(all_over_print_tab)
        self._build_store_price_content(store_price_tab)
    
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
        # Combined frame for three columns: drawing selection, info, and instructions
        main_frame = tk.Frame(container)
        main_frame.pack(fill='x', padx=20, pady=10)
        
        # Left column - Drawing selection frame
        drawing_frame = ttk.LabelFrame(main_frame, text="בחירת ציור", padding=15)
        drawing_frame.pack(side='left', fill='both', expand=True, padx=(0, 7))
        
        # Drawing selection
        tk.Label(drawing_frame, text="בחר ציור:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.selected_drawing_var = tk.StringVar()
        self.drawing_combobox = ttk.Combobox(drawing_frame, textvariable=self.selected_drawing_var, 
                                           state='readonly', width=25, justify='right')
        self.drawing_combobox.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.drawing_combobox.bind('<<ComboboxSelected>>', self._on_drawing_selected)
        
        # Load drawings button
        load_btn = tk.Button(drawing_frame, text="טען ציורים", command=self._load_drawings_list,
                           bg='#3498db', fg='white', font=('Arial', 9, 'bold'))
        load_btn.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky='ew')
        
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
        weight_frame = ttk.LabelFrame(container, text="נתוני הפריסה", padding=15)
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
        
        # Calculate button
        calc_btn = tk.Button(weight_frame, text="חשב חלוקת משקל", command=self._calculate_weight_distribution,
                           bg='#27ae60', fg='white', font=('Arial', 10, 'bold'))
        calc_btn.grid(row=1, column=2, padx=20, pady=5)
        
        # Export button
        export_btn = tk.Button(weight_frame, text="ייצא ל-Excel", command=self._export_weight_results,
                             bg='#2c3e50', fg='white', font=('Arial', 10, 'bold'))
        export_btn.grid(row=1, column=3, padx=5, pady=5)
        
        # Results frame
        results_frame = ttk.LabelFrame(container, text="תוצאות החישוב", padding=15)
        results_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Results table
        cols = ('product_name', 'size', 'quantity', 'square_area', 'percentage', 'weight_total', 'weight_per_unit')
        self.results_tree = ttk.Treeview(results_frame, columns=cols, show='headings', height=12)
        
        headers = {
            'product_name': 'שם המוצר',
            'size': 'מידה',
            'quantity': 'כמות',
            'square_area': 'שטח רבוע (מ״ר)',
            'percentage': 'אחוז (%)',
            'weight_total': 'משקל כולל (גרמים)',
            'weight_per_unit': 'משקל ליחידה (גרמים)'
        }
        
        widths = {
            'product_name': 160,
            'size': 60,
            'quantity': 60,
            'square_area': 120,
            'percentage': 80,
            'weight_total': 120,
            'weight_per_unit': 140
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
        summary_frame = ttk.LabelFrame(container, text="סיכום", padding=10)
        summary_frame.pack(fill='x', padx=20, pady=5)
        
        self.summary_var = tk.StringVar(value="בחר ציור והזן משקל לחישוב")
        summary_label = tk.Label(summary_frame, textvariable=self.summary_var, 
                               font=('Arial', 11, 'bold'), fg='#2c3e50')
        summary_label.pack(pady=5)
        
        # Initialize
        self._load_drawings_list()
    
    def _load_drawings_list(self):
        """Load the list of drawings from the database."""
        try:
            # Get drawings from data processor
            drawings = getattr(self.data_processor, 'drawings_data', [])
            drawing_names = []
            self.drawings_dict = {}
            
            for drawing in drawings:
                drawing_name = drawing.get('שם הקובץ', f"ציור {drawing.get('id', 'לא ידוע')}")
                drawing_names.append(drawing_name)
                self.drawings_dict[drawing_name] = drawing
            
            self.drawing_combobox['values'] = drawing_names
            
            if drawing_names:
                self.drawing_combobox.set(drawing_names[0])
                self._on_drawing_selected()
            else:
                self.drawing_info_text.config(state='normal')
                self.drawing_info_text.delete(1.0, tk.END)
                self.drawing_info_text.insert(tk.END, "לא נמצאו ציורים במסד הנתונים")
                self.drawing_info_text.config(state='disabled')
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בטעינת רשימת הציורים: {str(e)}")
    
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
                info_text += f"{weight_type}: {total_weight} גרם\n"
            
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
                
                # Fill weight if available
                if total_weight is not None:
                    self.layer_weight_var.set(str(total_weight))
                    # If both fields are filled, show message about automatic calculation
                    if layers is not None:
                        layer_desc = "שכבות" if drawing.get('שכבות') else "שכבות משוערת"
                        weight_desc = "משקל כולל" if drawing.get('משקל כולל') else "משקל נגזר"
                        self.summary_var.set(f"ציור נחתך זוהה! נתונים מולאו אוטומטית: {layers} {layer_desc}, {total_weight} גרם ({weight_desc})")
                    else:
                        weight_desc = "משקל כולל" if drawing.get('משקל כולל') else "משקל נגזר"
                        self.summary_var.set(f"ציור נחתך זוהה! {weight_desc} מולא אוטומטית: {total_weight} גרם")
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
            
            for item in calculation_items:
                percentage = (item['total_area'] / total_square_area) * 100
                weight_total = total_weight * (percentage / 100)
                weight_per_unit = weight_total / item['quantity'] if item['quantity'] > 0 else 0
                total_calculated_weight += weight_total
                
                # Insert into table
                self.results_tree.insert('', 'end', values=(
                    item['product_name'],
                    item['size'],
                    item['quantity'],
                    f"{item['square_area']:.6f}",
                    f"{percentage:.2f}%",
                    f"{weight_total:.2f}",
                    f"{weight_per_unit:.2f}"
                ))
            
            # Update summary
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
