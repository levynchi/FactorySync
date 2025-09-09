"""Shipping costs and fabrics tab for managing shipping expenses and fabric costs."""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import json
import os
from .shipping_companies_tab import ShippingCompaniesTabMixin

class ShippingCostsTabMixin(ShippingCompaniesTabMixin):
    """Mixin for shipping costs and fabrics management tab."""
    
    def _create_shipping_costs_tab(self):
        """Create the shipping costs and fabrics tab."""
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="עלויות משלוחים ובדים")
        
        # Create inner notebook for sub-tabs
        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=4, pady=4)
        
        # Shipping costs sub-tab
        shipping_costs_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(shipping_costs_page, text="עלויות משלוחים")
        self._build_shipping_costs_content(shipping_costs_page)
        
        # Shipping companies sub-tab
        shipping_companies_page = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(shipping_companies_page, text="חברות עמילות/שילוח")
        self._build_shipping_companies_content(shipping_companies_page)
    
    def _build_shipping_costs_content(self, container):
        """Build the shipping costs and fabrics content."""
        # Title
        title_label = tk.Label(
            container, 
            text="ניהול עלויות משלוחים ובדים", 
            font=('Arial', 16, 'bold'), 
            bg='#f7f9fa', 
            fg='#2c3e50'
        )
        title_label.pack(pady=(10, 20))
        
        # Input form frame
        input_frame = ttk.LabelFrame(container, text="הוספת משלוח חדש", padding=20)
        input_frame.pack(fill='x', padx=20, pady=10)
        
        # Row 1 - Name and Date
        tk.Label(input_frame, text="שם:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.shipping_name_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.shipping_name_var, width=20, font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="תאריך משלוח:", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.shipping_date_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
        tk.Entry(input_frame, textvariable=self.shipping_date_var, width=15, font=('Arial', 10)).grid(row=0, column=3, padx=5, pady=5)
        
        # Row 2 - Cub and Total Weight
        tk.Label(input_frame, text="Cub:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.cub_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.cub_var, width=20, font=('Arial', 10)).grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="משקל כולל:", font=('Arial', 10, 'bold')).grid(row=1, column=2, sticky='w', padx=5, pady=5)
        self.total_weight_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.total_weight_var, width=15, font=('Arial', 10)).grid(row=1, column=3, padx=5, pady=5)
        
        # Row 3 - Fabric cost and Quantity
        tk.Label(input_frame, text="עלות בד לק״ג:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.fabric_cost_per_kg_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.fabric_cost_per_kg_var, width=20, font=('Arial', 10)).grid(row=2, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="כמות גלילים:", font=('Arial', 10, 'bold')).grid(row=2, column=2, sticky='w', padx=5, pady=5)
        self.rolls_quantity_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.rolls_quantity_var, width=15, font=('Arial', 10)).grid(row=2, column=3, padx=5, pady=5)
        
        # Row 4 - Shipping costs
        tk.Label(input_frame, text="עלות משלוח סופית (ללא מע״מ):", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky='w', padx=5, pady=5)
        self.final_shipping_cost_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.final_shipping_cost_var, width=20, font=('Arial', 10)).grid(row=3, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="משלוח פנימי:", font=('Arial', 10, 'bold')).grid(row=3, column=2, sticky='w', padx=5, pady=5)
        self.domestic_shipping_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.domestic_shipping_var, width=15, font=('Arial', 10)).grid(row=3, column=3, padx=5, pady=5)
        
        # Row 5 - Documents
        tk.Label(input_frame, text="מסמכים לשילוח:", font=('Arial', 10, 'bold')).grid(row=4, column=0, sticky='w', padx=5, pady=5)
        self.documents_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.documents_var, width=50, font=('Arial', 10)).grid(row=4, column=1, columnspan=3, padx=5, pady=5)
        
        # Row 6 - Packing List Upload
        tk.Label(input_frame, text="PACKING LIST:", font=('Arial', 10, 'bold')).grid(row=5, column=0, sticky='w', padx=5, pady=5)
        self.packing_list_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.packing_list_var, width=30, font=('Arial', 10), state='readonly').grid(row=5, column=1, padx=5, pady=5)
        tk.Button(input_frame, text="📁 בחר קובץ", command=self._select_packing_list_file, bg='#3498db', fg='white', font=('Arial', 9)).grid(row=5, column=2, padx=5, pady=5)
        tk.Button(input_frame, text="🗑 נקה", command=self._clear_packing_list, bg='#e74c3c', fg='white', font=('Arial', 9)).grid(row=5, column=3, padx=5, pady=5)
        
        # Row 7 - Payment Request Upload
        tk.Label(input_frame, text="דרישת תשלום:", font=('Arial', 10, 'bold')).grid(row=6, column=0, sticky='w', padx=5, pady=5)
        self.payment_request_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.payment_request_var, width=30, font=('Arial', 10), state='readonly').grid(row=6, column=1, padx=5, pady=5)
        tk.Button(input_frame, text="📁 בחר קובץ", command=self._select_payment_request_file, bg='#3498db', fg='white', font=('Arial', 9)).grid(row=6, column=2, padx=5, pady=5)
        tk.Button(input_frame, text="🗑 נקה", command=self._clear_payment_request, bg='#e74c3c', fg='white', font=('Arial', 9)).grid(row=6, column=3, padx=5, pady=5)
        
        # Buttons
        buttons_frame = tk.Frame(input_frame)
        buttons_frame.grid(row=7, column=0, columnspan=4, pady=10)
        
        tk.Button(
            buttons_frame,
            text="הוסף משלוח",
            command=self._add_shipping_record,
            bg='#27ae60',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        ).pack(side='left', padx=5)
        
        tk.Button(
            buttons_frame,
            text="נקה שדות",
            command=self._clear_shipping_inputs,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        ).pack(side='left', padx=5)
        
        tk.Button(
            buttons_frame,
            text="ייצא לאקסל",
            command=self._export_to_excel,
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        ).pack(side='left', padx=5)
        
        # Data table frame
        table_frame = ttk.LabelFrame(container, text="רשימת משלוחים", padding=10)
        table_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Treeview for data display
        columns = ('name', 'date', 'cub', 'total_weight', 'fabric_cost_per_kg', 'rolls_quantity', 
                  'final_shipping_cost', 'domestic_shipping', 'final_cost_incl_domestic', 
                  'total_price_per_kg', 'total_price_per_cubic', 'fabric_shipping_cost_percent', 'documents', 'packing_list', 'payment_request')
        
        headers = {
            'name': 'שם',
            'date': 'תאריך משלוח',
            'cub': 'Cub',
            'total_weight': 'משקל כולל',
            'fabric_cost_per_kg': 'עלות בד לק״ג',
            'rolls_quantity': 'כמות גלילים',
            'final_shipping_cost': 'עלות משלוח סופית (ללא מע״מ)',
            'domestic_shipping': 'משלוח פנימי',
            'final_cost_incl_domestic': 'עלות סופית כולל משלוח פנימי',
            'total_price_per_kg': 'מחיר כולל לק״ג',
            'total_price_per_cubic': 'מחיר כולל למטר מעוקב (ללא מע״מ)',
            'fabric_shipping_cost_percent': 'אחוז עלות משלוח בד',
            'documents': 'מסמכים לשילוח',
            'packing_list': 'PACKING LIST',
            'payment_request': 'דרישת תשלום'
        }
        
        self.shipping_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        for col in columns:
            self.shipping_tree.heading(col, text=headers[col])
            self.shipping_tree.column(col, width=100, anchor='center')
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=self.shipping_tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal', command=self.shipping_tree.xview)
        self.shipping_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack treeview and scrollbars
        self.shipping_tree.pack(side='left', fill='both', expand=True)
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        
        # Bind double-click to open files
        self.shipping_tree.bind('<Double-1>', self._on_row_double_click)
        
        # Load existing data
        self._load_shipping_data()
        
        # Load CSV data if exists
        self._load_csv_data()
    
    def _add_shipping_record(self):
        """Add a new shipping record."""
        try:
            # Get input values
            name = self.shipping_name_var.get().strip()
            date = self.shipping_date_var.get().strip()
            cub = self.cub_var.get().strip()
            total_weight = self.total_weight_var.get().strip()
            fabric_cost_per_kg = self.fabric_cost_per_kg_var.get().strip()
            rolls_quantity = self.rolls_quantity_var.get().strip()
            final_shipping_cost = self.final_shipping_cost_var.get().strip()
            domestic_shipping = self.domestic_shipping_var.get().strip()
            documents = self.documents_var.get().strip()
            packing_list_file = self.packing_list_var.get().strip()
            payment_request_file = self.payment_request_var.get().strip()
            
            # Validate required fields
            if not name or not date:
                messagebox.showerror("שגיאה", "שם ותאריך הם שדות חובה")
                return
            
            # Convert to numbers for calculations
            cub_val = float(cub) if cub else 0
            total_weight_val = float(total_weight) if total_weight else 0
            fabric_cost_per_kg_val = float(fabric_cost_per_kg) if fabric_cost_per_kg else 0
            rolls_quantity_val = int(rolls_quantity) if rolls_quantity else 0
            final_shipping_cost_val = float(final_shipping_cost) if final_shipping_cost else 0
            domestic_shipping_val = float(domestic_shipping) if domestic_shipping else 0
            
            # Calculate derived values
            final_cost_incl_domestic = final_shipping_cost_val + domestic_shipping_val
            
            if total_weight_val > 0:
                total_price_per_kg = final_cost_incl_domestic / total_weight_val
            else:
                total_price_per_kg = 0
            
            if cub_val > 0:
                total_price_per_cubic = final_cost_incl_domestic / cub_val
            else:
                total_price_per_cubic = 0
            
            if fabric_cost_per_kg_val > 0 and total_price_per_kg > 0:
                fabric_shipping_cost_percent = (total_price_per_kg / fabric_cost_per_kg_val) * 100
            else:
                fabric_shipping_cost_percent = 0
            
            # Create record
            record = {
                'name': name,
                'date': date,
                'cub': cub_val,
                'total_weight': total_weight_val,
                'fabric_cost_per_kg': fabric_cost_per_kg_val,
                'rolls_quantity': rolls_quantity_val,
                'final_shipping_cost': final_shipping_cost_val,
                'domestic_shipping': domestic_shipping_val,
                'final_cost_incl_domestic': final_cost_incl_domestic,
                'total_price_per_kg': total_price_per_kg,
                'total_price_per_cubic': total_price_per_cubic,
                'fabric_shipping_cost_percent': fabric_shipping_cost_percent,
                'documents': documents,
                'packing_list': packing_list_file,
                'payment_request': payment_request_file
            }
            
            # Add to treeview
            values = (
                name, date, f"{cub_val:.2f}", f"{total_weight_val:.1f}", f"{fabric_cost_per_kg_val:.2f}",
                str(rolls_quantity_val), f"{final_shipping_cost_val:.0f}", f"{domestic_shipping_val:.0f}",
                f"{final_cost_incl_domestic:.0f}", f"{total_price_per_kg:.2f}", f"{total_price_per_cubic:.0f}",
                f"{fabric_shipping_cost_percent:.0f}%", documents, packing_list_file, payment_request_file
            )
            
            self.shipping_tree.insert('', 'end', values=values)
            
            # Save to file
            self._save_shipping_data()
            
            # Clear inputs
            self._clear_shipping_inputs()
            
            messagebox.showinfo("הצלחה", "המשלוח נוסף בהצלחה")
            
        except ValueError as e:
            messagebox.showerror("שגיאה", f"אנא הזן מספרים תקינים: {str(e)}")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בהוספת המשלוח: {str(e)}")
    
    def _clear_shipping_inputs(self):
        """Clear all input fields."""
        self.shipping_name_var.set("")
        self.shipping_date_var.set(datetime.now().strftime('%d/%m/%Y'))
        self.cub_var.set("")
        self.total_weight_var.set("")
        self.fabric_cost_per_kg_var.set("")
        self.rolls_quantity_var.set("")
        self.final_shipping_cost_var.set("")
        self.domestic_shipping_var.set("")
        self.documents_var.set("")
        self.packing_list_var.set("")
        self.payment_request_var.set("")
    
    def _load_shipping_data(self):
        """Load shipping data from file."""
        try:
            data_file = "shipping_costs.json"
            if os.path.exists(data_file):
                with open(data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                for record in data:
                    values = (
                        record.get('name', ''),
                        record.get('date', ''),
                        f"{record.get('cub', 0):.2f}",
                        f"{record.get('total_weight', 0):.1f}",
                        f"{record.get('fabric_cost_per_kg', 0):.2f}",
                        str(record.get('rolls_quantity', 0)),
                        f"{record.get('final_shipping_cost', 0):.0f}",
                        f"{record.get('domestic_shipping', 0):.0f}",
                        f"{record.get('final_cost_incl_domestic', 0):.0f}",
                        f"{record.get('total_price_per_kg', 0):.2f}",
                        f"{record.get('total_price_per_cubic', 0):.0f}",
                        f"{record.get('fabric_shipping_cost_percent', 0):.0f}%",
                        record.get('documents', ''),
                        record.get('packing_list', ''),
                        record.get('payment_request', '')
                    )
                    self.shipping_tree.insert('', 'end', values=values)
        except Exception as e:
            print(f"Error loading shipping data: {e}")
    
    def _load_csv_data(self):
        """Load shipping data from CSV file."""
        try:
            csv_file = "עלויות של משלוחים ובדים-Grid view.csv"
            if os.path.exists(csv_file):
                import csv
                
                with open(csv_file, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    
                    for row in reader:
                        # Extract data from CSV row - handle BOM in first column
                        name = row.get('\ufeffName', row.get('Name', '')).strip()
                        date = row.get('תאריך משלוח', '').strip()
                        cub = row.get('Cub', '').strip()
                        total_weight = row.get('Total weight', '').strip()
                        fabric_cost_per_kg = row.get('fabric cost per 1 kg', '').strip()
                        rolls_quantity = row.get('כמות גלילים', '').strip()
                        final_shipping_cost = row.get('Final shipping cost excluding VAT', '').strip()
                        domestic_shipping = row.get('Domestic shipping', '').strip()
                        final_cost_incl_domestic = row.get('Final cost incl. domestic shipping', '').strip()
                        total_price_per_kg = row.get('total price per 1 kg', '').strip()
                        total_price_per_cubic = row.get('Total price per cubic meter excluding VAT', '').strip()
                        fabric_shipping_cost_percent = row.get('Fabric shipping cost % to be added.', '').strip()
                        documents = row.get('מסמכים לשילוח', '').strip()
                        
                        
                        # Skip empty rows
                        if not name and not date:
                            continue
                        
                        # Convert values to appropriate types
                        try:
                            cub_val = float(cub) if cub else 0
                            total_weight_val = float(total_weight) if total_weight else 0
                            fabric_cost_per_kg_val = float(fabric_cost_per_kg) if fabric_cost_per_kg else 0
                            rolls_quantity_val = int(rolls_quantity) if rolls_quantity else 0
                            final_shipping_cost_val = float(final_shipping_cost) if final_shipping_cost else 0
                            domestic_shipping_val = float(domestic_shipping) if domestic_shipping else 0
                            final_cost_incl_domestic_val = float(final_cost_incl_domestic) if final_cost_incl_domestic else 0
                            total_price_per_kg_val = float(total_price_per_kg) if total_price_per_kg else 0
                            total_price_per_cubic_val = float(total_price_per_cubic) if total_price_per_cubic else 0
                            fabric_shipping_cost_percent_val = float(fabric_shipping_cost_percent.replace('%', '')) if fabric_shipping_cost_percent else 0
                        except ValueError:
                            # Skip rows with invalid data
                            continue
                        
                        # Add to treeview
                        values = (
                            name,
                            date,
                            f"{cub_val:.2f}",
                            f"{total_weight_val:.1f}",
                            f"{fabric_cost_per_kg_val:.2f}",
                            str(rolls_quantity_val),
                            f"{final_shipping_cost_val:.0f}",
                            f"{domestic_shipping_val:.0f}",
                            f"{final_cost_incl_domestic_val:.0f}",
                            f"{total_price_per_kg_val:.2f}",
                            f"{total_price_per_cubic_val:.0f}",
                            f"{fabric_shipping_cost_percent_val:.0f}%",
                            documents,
                            '',  # Empty packing list for CSV data
                            ''   # Empty payment request for CSV data
                        )
                        
                        self.shipping_tree.insert('', 'end', values=values)
                
                print(f"Loaded {len(self.shipping_tree.get_children())} records from CSV file")
                
        except Exception as e:
            print(f"Error loading CSV data: {e}")
    
    def _save_shipping_data(self):
        """Save shipping data to file."""
        try:
            data = []
            for item in self.shipping_tree.get_children():
                values = self.shipping_tree.item(item)['values']
                record = {
                    'name': values[0],
                    'date': values[1],
                    'cub': float(values[2]) if values[2] else 0,
                    'total_weight': float(values[3]) if values[3] else 0,
                    'fabric_cost_per_kg': float(values[4]) if values[4] else 0,
                    'rolls_quantity': int(values[5]) if values[5] else 0,
                    'final_shipping_cost': float(values[6]) if values[6] else 0,
                    'domestic_shipping': float(values[7]) if values[7] else 0,
                    'final_cost_incl_domestic': float(values[8]) if values[8] else 0,
                    'total_price_per_kg': float(values[9]) if values[9] else 0,
                    'total_price_per_cubic': float(values[10]) if values[10] else 0,
                    'fabric_shipping_cost_percent': float(values[11].replace('%', '')) if values[11] else 0,
                    'documents': values[12],
                    'packing_list': values[13] if len(values) > 13 else '',
                    'payment_request': values[14] if len(values) > 14 else ''
                }
                data.append(record)
            
            with open("shipping_costs.json", 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving shipping data: {e}")
    
    def _export_to_excel(self):
        """Export shipping data to Excel."""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment
            
            wb = Workbook()
            ws = wb.active
            ws.title = "עלויות משלוחים ובדים"
            
            # Headers
            headers = ['שם', 'תאריך משלוח', 'Cub', 'משקל כולל', 'עלות בד לק״ג', 'כמות גלילים',
                      'עלות משלוח סופית (ללא מע״מ)', 'משלוח פנימי', 'עלות סופית כולל משלוח פנימי',
                      'מחיר כולל לק״ג', 'מחיר כולל למטר מעוקב (ללא מע״מ)', 'אחוז עלות משלוח בד', 'מסמכים לשילוח', 'PACKING LIST', 'דרישת תשלום']
            
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            
            # Data
            for row, item in enumerate(self.shipping_tree.get_children(), 2):
                values = self.shipping_tree.item(item)['values']
                for col, value in enumerate(values, 1):
                    ws.cell(row=row, column=col, value=value)
            
            # Auto-adjust column widths
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width
            
            # Save file
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="שמור קובץ אקסל"
            )
            
            if filename:
                wb.save(filename)
                messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה: {filename}")
            
        except ImportError:
            messagebox.showerror("שגיאה", "ספריית openpyxl לא מותקנת. אנא התקן אותה עם: pip install openpyxl")
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בייצוא לאקסל: {str(e)}")
    
    def _select_packing_list_file(self):
        """Select a packing list Excel file to upload."""
        file_path = filedialog.askopenfilename(
            title="בחר קובץ PACKING LIST",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            # Copy file to packing_lists directory
            try:
                import shutil
                import os
                
                # Create packing_lists directory if it doesn't exist
                packing_lists_dir = "packing_lists"
                if not os.path.exists(packing_lists_dir):
                    os.makedirs(packing_lists_dir)
                
                # Generate unique filename
                filename = os.path.basename(file_path)
                name, ext = os.path.splitext(filename)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                new_filename = f"{name}_{timestamp}{ext}"
                dest_path = os.path.join(packing_lists_dir, new_filename)
                
                # Copy file
                shutil.copy2(file_path, dest_path)
                
                # Update the display
                self.packing_list_var.set(new_filename)
                
                messagebox.showinfo("הצלחה", f"קובץ PACKING LIST נשמר: {new_filename}")
                
            except Exception as e:
                messagebox.showerror("שגיאה", f"שגיאה בשמירת הקובץ: {str(e)}")
    
    def _clear_packing_list(self):
        """Clear the selected packing list file."""
        self.packing_list_var.set("")
    
    def _on_row_double_click(self, event):
        """Handle double-click on any row to open files."""
        # Get the item that was clicked
        item = self.shipping_tree.selection()[0] if self.shipping_tree.selection() else None
        if not item:
            return
        
        # Get the values of the selected row
        values = self.shipping_tree.item(item, 'values')
        if len(values) < 15:  # Need at least 15 columns
            return
        
        packing_list_file = values[13] if len(values) > 13 else ""
        payment_request_file = values[14] if len(values) > 14 else ""
        
        # Check which files exist and ask user which one to open
        files_to_open = []
        
        if packing_list_file and packing_list_file.strip():
            files_to_open.append(("PACKING LIST", packing_list_file, "packing_lists"))
        
        if payment_request_file and payment_request_file.strip():
            files_to_open.append(("דרישת תשלום", payment_request_file, "payment_requests"))
        
        if not files_to_open:
            messagebox.showinfo("מידע", "אין קבצים לשורה זו")
            return
        
        if len(files_to_open) == 1:
            # Only one file, open it directly
            file_type, filename, directory = files_to_open[0]
            self._open_file(directory, filename)
        else:
            # Multiple files, ask user which one to open
            import tkinter as tk
            from tkinter import messagebox
            
            choice_window = tk.Toplevel()
            choice_window.title("בחר קובץ לפתיחה")
            choice_window.geometry("300x150")
            choice_window.transient()
            choice_window.grab_set()
            
            tk.Label(choice_window, text="איזה קובץ ברצונך לפתוח?", font=('Arial', 12, 'bold')).pack(pady=10)
            
            for i, (file_type, filename, directory) in enumerate(files_to_open):
                tk.Button(
                    choice_window, 
                    text=f"פתח {file_type}",
                    command=lambda d=directory, f=filename: [self._open_file(d, f), choice_window.destroy()],
                    width=20
                ).pack(pady=5)
            
            tk.Button(choice_window, text="ביטול", command=choice_window.destroy, width=20).pack(pady=5)
    
    def _open_file(self, directory, filename):
        """Open a file from the specified directory."""
        try:
            import subprocess
            import os
            
            file_path = os.path.join(directory, filename)
            if os.path.exists(file_path):
                # Try to open with default application
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                elif os.name == 'posix':  # macOS and Linux
                    subprocess.call(['open', file_path])
                else:
                    subprocess.call(['xdg-open', file_path])
            else:
                messagebox.showerror("שגיאה", f"קובץ לא נמצא: {file_path}")
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן לפתוח את הקובץ: {str(e)}")
    
    def _select_payment_request_file(self):
        """Select a payment request file to upload."""
        file_path = filedialog.askopenfilename(
            title="בחר קובץ דרישת תשלום",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("Excel files", "*.xlsx *.xls"),
                ("Image files", "*.jpg *.jpeg *.png"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            # Copy file to payment_requests directory
            try:
                import shutil
                import os
                
                # Create payment_requests directory if it doesn't exist
                payment_requests_dir = "payment_requests"
                if not os.path.exists(payment_requests_dir):
                    os.makedirs(payment_requests_dir)
                
                # Generate unique filename
                filename = os.path.basename(file_path)
                name, ext = os.path.splitext(filename)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                new_filename = f"{name}_{timestamp}{ext}"
                dest_path = os.path.join(payment_requests_dir, new_filename)
                
                # Copy file
                shutil.copy2(file_path, dest_path)
                
                # Update the display
                self.payment_request_var.set(new_filename)
                
                messagebox.showinfo("הצלחה", f"קובץ דרישת תשלום נשמר: {new_filename}")
                
            except Exception as e:
                messagebox.showerror("שגיאה", f"שגיאה בשמירת הקובץ: {str(e)}")
    
    def _clear_payment_request(self):
        """Clear the selected payment request file."""
        self.payment_request_var.set("")
    
    def _on_payment_request_double_click(self, event):
        """Handle double-click on payment request column to open file."""
        # Get the item that was clicked
        item = self.shipping_tree.selection()[0] if self.shipping_tree.selection() else None
        if not item:
            return
        
        # Get the values of the selected row
        values = self.shipping_tree.item(item, 'values')
        if len(values) < 15:  # payment_request is the last column (index 14)
            return
        
        payment_request_file = values[14]
        if not payment_request_file or payment_request_file.strip() == "":
            messagebox.showinfo("מידע", "אין קובץ דרישת תשלום לשורה זו")
            return
        
        # Try to open the file
        try:
            import subprocess
            import os
            
            file_path = os.path.join("payment_requests", payment_request_file)
            if os.path.exists(file_path):
                # Try to open with default application
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                elif os.name == 'posix':  # macOS and Linux
                    subprocess.call(['open', file_path])
                else:
                    subprocess.call(['xdg-open', file_path])
            else:
                messagebox.showerror("שגיאה", f"קובץ לא נמצא: {file_path}")
                
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן לפתוח את הקובץ: {str(e)}")
