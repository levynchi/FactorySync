"""Shipping companies management tab for managing shipping companies and contacts."""
import tkinter as tk
from tkinter import ttk, messagebox
import json
import os

class ShippingCompaniesTabMixin:
    """Mixin for shipping companies management tab."""
    
    def _build_shipping_companies_content(self, container):
        """Build the shipping companies management content."""
        # Title
        title_label = tk.Label(
            container, 
            text="ניהול חברות עמילות/שילוח", 
            font=('Arial', 16, 'bold'), 
            bg='#f7f9fa', 
            fg='#2c3e50'
        )
        title_label.pack(pady=(10, 20))
        
        # Input form frame
        input_frame = ttk.LabelFrame(container, text="הוספת חברת שילוח חדשה", padding=20)
        input_frame.pack(fill='x', padx=20, pady=10)
        
        # Company details
        tk.Label(input_frame, text="שם החברה:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.company_name_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.company_name_var, width=30, font=('Arial', 10)).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="איש קשר:", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.contact_person_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.contact_person_var, width=30, font=('Arial', 10)).grid(row=0, column=3, padx=5, pady=5)
        
        # Contact details
        tk.Label(input_frame, text="טלפון:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.phone_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.phone_var, width=30, font=('Arial', 10)).grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="מייל:", font=('Arial', 10, 'bold')).grid(row=1, column=2, sticky='w', padx=5, pady=5)
        self.email_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.email_var, width=30, font=('Arial', 10)).grid(row=1, column=3, padx=5, pady=5)
        
        # Additional details
        tk.Label(input_frame, text="כתובת:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.address_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.address_var, width=30, font=('Arial', 10)).grid(row=2, column=1, padx=5, pady=5)
        
        tk.Label(input_frame, text="הערות:", font=('Arial', 10, 'bold')).grid(row=2, column=2, sticky='w', padx=5, pady=5)
        self.notes_var = tk.StringVar()
        tk.Entry(input_frame, textvariable=self.notes_var, width=30, font=('Arial', 10)).grid(row=2, column=3, padx=5, pady=5)
        
        # Buttons
        buttons_frame = tk.Frame(input_frame)
        buttons_frame.grid(row=3, column=0, columnspan=4, pady=10)
        
        tk.Button(
            buttons_frame,
            text="הוסף חברה",
            command=self._add_shipping_company,
            bg='#27ae60',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        ).pack(side='left', padx=5)
        
        tk.Button(
            buttons_frame,
            text="עדכן חברה",
            command=self._update_shipping_company,
            bg='#f39c12',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        ).pack(side='left', padx=5)
        
        tk.Button(
            buttons_frame,
            text="מחק חברה",
            command=self._delete_shipping_company,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        ).pack(side='left', padx=5)
        
        tk.Button(
            buttons_frame,
            text="נקה שדות",
            command=self._clear_shipping_company_inputs,
            bg='#95a5a6',
            fg='white',
            font=('Arial', 10, 'bold'),
            width=15
        ).pack(side='left', padx=5)
        
        # Data table frame
        table_frame = ttk.LabelFrame(container, text="רשימת חברות שילוח", padding=10)
        table_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Treeview for data display
        columns = ('company_name', 'contact_person', 'phone', 'email', 'address', 'notes')
        
        headers = {
            'company_name': 'שם החברה',
            'contact_person': 'איש קשר',
            'phone': 'טלפון',
            'email': 'מייל',
            'address': 'כתובת',
            'notes': 'הערות'
        }
        
        self.shipping_companies_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        for col in columns:
            self.shipping_companies_tree.heading(col, text=headers[col])
            self.shipping_companies_tree.column(col, width=150, anchor='center')
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=self.shipping_companies_tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal', command=self.shipping_companies_tree.xview)
        self.shipping_companies_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack treeview and scrollbars
        self.shipping_companies_tree.pack(side='left', fill='both', expand=True)
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        
        # Bind selection event
        self.shipping_companies_tree.bind('<<TreeviewSelect>>', self._on_company_select)
        
        # Load existing data
        self._load_shipping_companies_data()
    
    def _add_shipping_company(self):
        """Add a new shipping company."""
        try:
            # Get input values
            company_name = self.company_name_var.get().strip()
            contact_person = self.contact_person_var.get().strip()
            phone = self.phone_var.get().strip()
            email = self.email_var.get().strip()
            address = self.address_var.get().strip()
            notes = self.notes_var.get().strip()
            
            # Validate required fields
            if not company_name:
                messagebox.showerror("שגיאה", "שם החברה הוא שדה חובה")
                return
            
            # Create record
            record = {
                'company_name': company_name,
                'contact_person': contact_person,
                'phone': phone,
                'email': email,
                'address': address,
                'notes': notes
            }
            
            # Add to treeview
            values = (company_name, contact_person, phone, email, address, notes)
            self.shipping_companies_tree.insert('', 'end', values=values)
            
            # Save to file
            self._save_shipping_companies_data()
            
            # Refresh shipping costs combobox if it exists
            if hasattr(self, 'refresh_shipping_companies_combobox'):
                self.refresh_shipping_companies_combobox()
            
            # Clear inputs
            self._clear_shipping_company_inputs()
            
            messagebox.showinfo("הצלחה", "החברה נוספה בהצלחה")
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בהוספת החברה: {str(e)}")
    
    def _update_shipping_company(self):
        """Update selected shipping company."""
        try:
            selected_item = self.shipping_companies_tree.selection()
            if not selected_item:
                messagebox.showwarning("אזהרה", "נא לבחור חברה לעדכון")
                return
            
            # Get input values
            company_name = self.company_name_var.get().strip()
            contact_person = self.contact_person_var.get().strip()
            phone = self.phone_var.get().strip()
            email = self.email_var.get().strip()
            address = self.address_var.get().strip()
            notes = self.notes_var.get().strip()
            
            # Validate required fields
            if not company_name:
                messagebox.showerror("שגיאה", "שם החברה הוא שדה חובה")
                return
            
            # Update treeview
            values = (company_name, contact_person, phone, email, address, notes)
            self.shipping_companies_tree.item(selected_item[0], values=values)
            
            # Save to file
            self._save_shipping_companies_data()
            
            # Refresh shipping costs combobox if it exists
            if hasattr(self, 'refresh_shipping_companies_combobox'):
                self.refresh_shipping_companies_combobox()
            
            # Clear inputs
            self._clear_shipping_company_inputs()
            
            messagebox.showinfo("הצלחה", "החברה עודכנה בהצלחה")
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה בעדכון החברה: {str(e)}")
    
    def _delete_shipping_company(self):
        """Delete selected shipping company."""
        try:
            selected_item = self.shipping_companies_tree.selection()
            if not selected_item:
                messagebox.showwarning("אזהרה", "נא לבחור חברה למחיקה")
                return
            
            # Confirm deletion
            if messagebox.askyesno("אישור מחיקה", "האם אתה בטוח שברצונך למחוק את החברה הנבחרת?"):
                self.shipping_companies_tree.delete(selected_item[0])
                self._save_shipping_companies_data()
                
                # Refresh shipping costs combobox if it exists
                if hasattr(self, 'refresh_shipping_companies_combobox'):
                    self.refresh_shipping_companies_combobox()
                
                self._clear_shipping_company_inputs()
                messagebox.showinfo("הצלחה", "החברה נמחקה בהצלחה")
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"שגיאה במחיקת החברה: {str(e)}")
    
    def _clear_shipping_company_inputs(self):
        """Clear all input fields."""
        self.company_name_var.set("")
        self.contact_person_var.set("")
        self.phone_var.set("")
        self.email_var.set("")
        self.address_var.set("")
        self.notes_var.set("")
    
    def _on_company_select(self, event):
        """Handle company selection to populate input fields."""
        selected_item = self.shipping_companies_tree.selection()
        if selected_item:
            values = self.shipping_companies_tree.item(selected_item[0], 'values')
            if values:
                self.company_name_var.set(values[0])
                self.contact_person_var.set(values[1])
                self.phone_var.set(values[2])
                self.email_var.set(values[3])
                self.address_var.set(values[4])
                self.notes_var.set(values[5])
    
    def _load_shipping_companies_data(self):
        """Load shipping companies data from file."""
        try:
            data_file = "shipping_companies.json"
            if os.path.exists(data_file):
                with open(data_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                for record in data:
                    values = (
                        record.get('company_name', ''),
                        record.get('contact_person', ''),
                        record.get('phone', ''),
                        record.get('email', ''),
                        record.get('address', ''),
                        record.get('notes', '')
                    )
                    self.shipping_companies_tree.insert('', 'end', values=values)
        except Exception as e:
            print(f"Error loading shipping companies data: {e}")
    
    def _save_shipping_companies_data(self):
        """Save shipping companies data to file."""
        try:
            data = []
            for item in self.shipping_companies_tree.get_children():
                values = self.shipping_companies_tree.item(item)['values']
                record = {
                    'company_name': values[0],
                    'contact_person': values[1],
                    'phone': values[2],
                    'email': values[3],
                    'address': values[4],
                    'notes': values[5]
                }
                data.append(record)
            
            with open("shipping_companies.json", 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving shipping companies data: {e}")
