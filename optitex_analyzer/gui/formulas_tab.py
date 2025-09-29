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
        
        inner_nb.add(weight_calc_tab, text="חישובי משקל כללי")
        inner_nb.add(fabric_weight_tab, text="משקל בד לפריטים בציור")
        
        # Build content for each sub-tab
        self._build_general_weight_content(weight_calc_tab)
        self._build_fabric_weight_content(fabric_weight_tab)
    
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
        # Instructions frame
        instructions_frame = ttk.LabelFrame(container, text="הוראות שימוש", padding=10)
        instructions_frame.pack(fill='x', padx=20, pady=5)
        
        instructions_text = ("בחר ציור מהרשימה לחישוב חלוקת משקל בד לפי שטח רבוע של הפריטים.\n"
                           "עבור ציורים בסטטוס 'נחתך' - כמות השכבות והמשקל ימולאו אוטומטית אם זמינים במסד הנתונים.\n\n"
                           "הנוסחה: %ᵢ = (Aᵢ/ΣA) × 100, Gramsᵢ = W × (%ᵢ/100)\n"
                           "כאשר W = משקל השכבה, Aᵢ = שטח רבוע למידה i, ΣA = סכום כל השטחים")
        instructions_label = tk.Label(
            instructions_frame,
            text=instructions_text,
            font=('Arial', 9),
            fg='#7f8c8d',
            justify='right',
            wraplength=600
        )
        instructions_label.pack(pady=5)
        
        # Drawing selection frame
        drawing_frame = ttk.LabelFrame(container, text="בחירת ציור", padding=15)
        drawing_frame.pack(fill='x', padx=20, pady=10)
        
        # Drawing selection
        tk.Label(drawing_frame, text="בחר ציור:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.selected_drawing_var = tk.StringVar()
        self.drawing_combobox = ttk.Combobox(drawing_frame, textvariable=self.selected_drawing_var, 
                                           state='readonly', width=40, justify='right')
        self.drawing_combobox.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.drawing_combobox.bind('<<ComboboxSelected>>', self._on_drawing_selected)
        
        # Load drawings button
        load_btn = tk.Button(drawing_frame, text="טען ציורים", command=self._load_drawings_list,
                           bg='#3498db', fg='white', font=('Arial', 9, 'bold'))
        load_btn.grid(row=0, column=2, padx=10, pady=5)
        
        # Drawing info frame
        info_frame = ttk.LabelFrame(container, text="פרטי הציור", padding=15)
        info_frame.pack(fill='x', padx=20, pady=10)
        
        self.drawing_info_text = tk.Text(info_frame, height=6, width=80, font=('Arial', 9), 
                                       state='disabled', wrap='word')
        info_scrollbar = ttk.Scrollbar(info_frame, orient='vertical', command=self.drawing_info_text.yview)
        self.drawing_info_text.configure(yscrollcommand=info_scrollbar.set)
        self.drawing_info_text.pack(side='left', fill='both', expand=True)
        info_scrollbar.pack(side='right', fill='y')
        
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
