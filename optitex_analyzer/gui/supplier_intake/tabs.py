import tkinter as tk
from tkinter import ttk
from .methods import SupplierIntakeMethodsMixin

class SupplierIntakeTabMixin(SupplierIntakeMethodsMixin):
    """Compose the Supplier Intake tab (entry + saved list)."""
    def _create_supplier_intake_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="תעודת קליטה")
        tk.Label(tab, text="תעודת קליטה (הזנה ידנית)", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=8)

        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=4, pady=4)
        entry_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        list_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(entry_wrapper, text="קליטה")
        inner_nb.add(list_wrapper, text="קליטות שמורות")

        self._build_supplier_entry_tab(entry_wrapper)
        self._build_supplier_list_tab(list_wrapper)

    def _build_supplier_entry_tab(self, container: tk.Frame):
        from .entry_tab import build_entry_tab
        build_entry_tab(self, container)

    def _build_supplier_list_tab(self, container: tk.Frame):
        from .list_tab import build_list_tab
        build_list_tab(self, container)
