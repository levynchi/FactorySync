import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from .methods import DeliveryNoteMethodsMixin
from .. import theme


class DeliveryNoteTabMixin(DeliveryNoteMethodsMixin):
    """Compose the Delivery Note tab by embedding the entry and list sub-tabs."""

    def _create_delivery_note_tab(self):
        tab = tk.Frame(self.notebook, bg=theme.PAGE_BG)
        self.notebook.add(tab, text="תעודת משלוח")
        tk.Label(
            tab,
            text="תעודת משלוח (הזנה ידנית)",
            font=(theme.FONT_FAMILY, 16, 'bold'),
            bg=theme.PAGE_BG,
            fg=theme.DARK,
        ).pack(pady=4)

        inner_nb = ttk.Notebook(tab)
        inner_nb.pack(fill='both', expand=True, padx=4, pady=4)

        # Build subtabs via helpers
        entry_wrapper = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        list_wrapper = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        fabrics_send_wrapper = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        fabrics_send_list_wrapper = tk.Frame(inner_nb, bg=theme.PAGE_BG)
        inner_nb.add(entry_wrapper, text="קליטה")
        inner_nb.add(list_wrapper, text="תעודות שמורות")
        inner_nb.add(fabrics_send_wrapper, text="שליחת בדים")
        inner_nb.add(fabrics_send_list_wrapper, text="שליחות בדים שמורות")

        # Delegate to sub-tab builders
        self._build_delivery_entry_tab(entry_wrapper)
        self._build_delivery_list_tab(list_wrapper)
        self._build_fabrics_send_tab(fabrics_send_wrapper)
        self._build_fabrics_send_list_tab(fabrics_send_list_wrapper)

    # Subtab builders imported from separate modules at bottom to avoid circular imports
    def _build_delivery_entry_tab(self, container: tk.Frame):
        from .entry_tab import build_entry_tab
        build_entry_tab(self, container)

    def _build_delivery_list_tab(self, container: tk.Frame):
        from .list_tab import build_list_tab
        build_list_tab(self, container)

    def _build_fabrics_send_tab(self, container: tk.Frame):
        from .fabrics_send_tab import build_fabrics_send_tab
        build_fabrics_send_tab(self, container)

    def _build_fabrics_send_list_tab(self, container: tk.Frame):
        from .fabrics_send_list_tab import build_fabrics_send_list_tab
        build_fabrics_send_list_tab(self, container)
