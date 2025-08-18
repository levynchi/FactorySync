import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

from .methods import ProductsCatalogMethodsMixin

class ProductsCatalogTabMixin(ProductsCatalogMethodsMixin):
	"""Compose the Products Catalog feature tab with sub-pages."""
	def _create_products_catalog_tab(self):
		tab = tk.Frame(self.notebook, bg='#f7f9fa')
		self.notebook.add(tab, text="קטלוג מוצרים ופריטים")
		tk.Label(tab, text="ניהול קטלוג מוצרים ופריטים", font=('Arial', 16, 'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=4)

		inner_nb = ttk.Notebook(tab)
		inner_nb.pack(fill='both', expand=True, padx=6, pady=4)

		products_tab = tk.Frame(inner_nb, bg='#f7f9fa')
		accessories_tab = tk.Frame(inner_nb, bg='#f7f9fa')
		categories_tab = tk.Frame(inner_nb, bg='#f7f9fa')
		main_categories_tab = tk.Frame(inner_nb, bg='#f7f9fa')
		attributes_tab = tk.Frame(inner_nb, bg='#f7f9fa')

		inner_nb.add(products_tab, text="פריטים")
		inner_nb.add(accessories_tab, text="אביזרי תפירה")
		inner_nb.add(categories_tab, text="תת קטגוריות")
		inner_nb.add(main_categories_tab, text="קטגוריה ראשית")
		inner_nb.add(attributes_tab, text="תכונות מוצר")

		# Build each sub-section
		self._build_products_section(products_tab)
		self._build_accessories_section(accessories_tab)
		self._build_categories_section(categories_tab)
		self._build_main_categories_section(main_categories_tab)
		self._build_attributes_section(attributes_tab)
