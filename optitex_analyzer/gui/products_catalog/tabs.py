import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

from .methods import ProductsCatalogMethodsMixin
from .. import theme

class ProductsCatalogTabMixin(ProductsCatalogMethodsMixin):
	"""Compose the Products Catalog feature tab with sub-pages."""
	def _create_products_catalog_tab(self):
		tab = tk.Frame(self.notebook, bg=theme.PAGE_BG)
		self.notebook.add(tab, text="קטלוג מוצרים ופריטים")
		tk.Label(tab, text="ניהול קטלוג מוצרים ופריטים", font=(theme.FONT_FAMILY, 16, 'bold'), bg=theme.PAGE_BG, fg=theme.DARK).pack(pady=4)

		inner_nb = ttk.Notebook(tab)
		inner_nb.pack(fill='both', expand=True, padx=6, pady=4)

		products_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
		accessories_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
		label_inv_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
		categories_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
		main_categories_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
		attributes_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
		barcodes_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)
		cuts_tab = tk.Frame(inner_nb, bg=theme.PAGE_BG)

		inner_nb.add(products_tab, text="פריטים")
		inner_nb.add(accessories_tab, text="אביזרי תפירה")
		inner_nb.add(label_inv_tab, text="מלאי תוויות")
		inner_nb.add(categories_tab, text="תת קטגוריות")
		inner_nb.add(main_categories_tab, text="קטגוריה ראשית")
		inner_nb.add(attributes_tab, text="תכונות מוצר")
		inner_nb.add(barcodes_tab, text="ברקודים")
		inner_nb.add(cuts_tab, text="גזרות")

		# Build each sub-section
		self._build_products_section(products_tab)
		self._build_accessories_section(accessories_tab)
		self._build_label_inventory_section(label_inv_tab)
		self._build_categories_section(categories_tab)
		self._build_main_categories_section(main_categories_tab)
		self._build_attributes_section(attributes_tab)
		self._build_barcodes_section(barcodes_tab)
		self._build_cuts_section(cuts_tab)
