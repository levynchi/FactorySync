import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from .methods import DeliveryNoteMethodsMixin

class DeliveryNoteTabMixin(DeliveryNoteMethodsMixin):
    """Compose the Delivery Note tab by embedding the entry and list sub-tabs."""
    def _create_delivery_note_tab(self):
        tab = tk.Frame(self.notebook, bg='#f7f9fa')
        self.notebook.add(tab, text="תעודת משלוח")
        tk.Label(tab, text="תעודת משלוח (הזנה ידנית)", font=('Arial',16,'bold'), bg='#f7f9fa', fg='#2c3e50').pack(pady=4)

        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=4, pady=4)

        # Build subtabs via helpers
        entry_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        list_wrapper = tk.Frame(inner_nb, bg='#f7f9fa')
        inner_nb.add(entry_wrapper, text="קליטה")
        inner_nb.add(list_wrapper, text="תעודות שמורות")

        # Delegate to sub-tab builders
        self._build_delivery_entry_tab(entry_wrapper)
        self._build_delivery_list_tab(list_wrapper)

    # Subtab builders imported from separate modules at bottom to avoid circular imports
    def _build_delivery_entry_tab(self, container: tk.Frame):
        from .entry.entry_tab import build_entry_tab
        build_entry_tab(self, container)

    def _build_delivery_list_tab(self, container: tk.Frame):
        from .list.list_tab import build_list_tab
        build_list_tab(self, container)
