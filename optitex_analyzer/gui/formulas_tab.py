"""Formulas and calculations tab for weight and measurements calculations."""
import tkinter as tk
from tkinter import ttk, messagebox

class FormulasTabMixin:
    """Mixin for formulas and calculations tab."""
    
    def _create_formulas_tab(self):
        """Create the formulas and calculations tab."""
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="נוסחאות וחישובים")
        self._build_formulas_content(tab)
    
    def _build_formulas_content(self, container):
        """Build the formulas and calculations content."""
        # Title
        title_label = tk.Label(
            container, 
            text="נוסחאות וחישובים", 
            font=('Arial', 16, 'bold'), 
            bg='#f7f9fa', 
            fg='#2c3e50'
        )
        title_label.pack(pady=(10, 20))
        
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
