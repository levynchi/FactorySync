"""Main application window orchestrating separate tab mixins."""
import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from .converter_tab import ConverterTabMixin  # still used indirectly for embedded converter inside drawings manager
from .returned_drawing_tab import ReturnedDrawingTabMixin
from .fabrics_inventory_tab import FabricsInventoryTabMixin
from .supplier_intake import SupplierIntakeTabMixin
from .delivery_note import DeliveryNoteTabMixin
from .products_catalog import ProductsCatalogTabMixin
from .drawings_manager_tab import DrawingsManagerTabMixin
from .suppliers_tab import SuppliersTabMixin
from .shipments_tab import ShipmentsTabMixin
from .products_balance_tab import ProductsBalanceTabMixin
from .business_details_tab import BusinessDetailsTabMixin
from .stickers_tab import StickersTabMixin


class MainWindow(
    ConverterTabMixin,
    ReturnedDrawingTabMixin,
    FabricsInventoryTabMixin,
    SupplierIntakeTabMixin,
    DeliveryNoteTabMixin,
    ProductsCatalogTabMixin,
    DrawingsManagerTabMixin,
    SuppliersTabMixin,
    ShipmentsTabMixin,
    ProductsBalanceTabMixin,
    BusinessDetailsTabMixin,
    StickersTabMixin,
):
    def __init__(self, root, settings_manager, file_analyzer, data_processor):
        """Initialize the main window, assemble all tab mixins and shared UI."""
        # Core dependencies
        self.root = root
        self.settings = settings_manager
        self.file_analyzer = file_analyzer
        self.data_processor = data_processor

        # ----- Window geometry & basic appearance -----
        self.root.title("FactorySync - ממיר אופטיטקס")
        try:
            desired_geom = self.settings.get("app.window_size", "1400x900")
            if not isinstance(desired_geom, str):
                desired_geom = "1400x900"
        except Exception:
            desired_geom = "1400x900"

        def _safe_apply_geometry(g: str) -> str:
            import re
            scr_w = self.root.winfo_screenwidth()
            scr_h = self.root.winfo_screenheight()
            m = re.match(r"^(\d+)x(\d+)([+-]\d+)?([+-]\d+)?$", (g or "").strip())
            if not m:
                return "1400x900+50+50"
            w = max(600, min(int(m.group(1)), scr_w))
            h = max(400, min(int(m.group(2)), scr_h))
            x, y = 50, 50
            if m.group(3) and m.group(4):
                try:
                    x = int(m.group(3)); y = int(m.group(4))
                except ValueError:
                    x, y = 50, 50
            if x < 0 or x > scr_w - 100: x = 50
            if y < 0 or y > scr_h - 100: y = 50
            return f"{w}x{h}+{x}+{y}"

        safe_geom = _safe_apply_geometry(desired_geom)
        try:
            self.root.geometry(safe_geom)
        except Exception:
            self.root.geometry("1400x900+50+50")
        self.root.update_idletasks()
        self.root.deiconify()
        self.root.lift()
        try:
            self.root.attributes('-topmost', True)
            self.root.after(500, lambda: self.root.attributes('-topmost', False))
        except Exception:
            pass

        # Global context menus (right-click) for text fields
        try:
            self._setup_right_click_text_menus()
        except Exception:
            pass

        # ----- State -----
        self.rib_file = ""
        self.products_file = self.settings.get("app.products_file", "")
        if self.products_file and not os.path.exists(self.products_file):
            self.products_file = ""
        self.current_results = []
        self.drawings_manager_window = None

        # ----- Notebook & Tabs -----
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        # Create each tab from its mixin
    # Removed standalone converter tab: converter now embedded as sub-tab inside 'מנהל ציורים'
    # Returned (cut) drawings now embedded inside 'מנהל ציורים' tab
        # Software management parent tab (contains Business Details)
        try:
            self._create_software_management_tab()
        except Exception:
            pass
        self._create_fabrics_inventory_tab()
        self._create_supplier_intake_tab()
        # Delivery note tab (duplicate logic for separate process)
        try:
            self._create_delivery_note_tab()
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", f"טעינת טאב 'תעודת משלוח' נכשלה: {e}")
            except Exception:
                pass
        self._create_products_catalog_tab()
        self._create_drawings_manager_tab()
        self._create_suppliers_tab()
        # Stickers tab
        try:
            self._create_stickers_tab()
        except Exception:
            pass
        # Shipments summary tab
        try:
            self._create_shipments_tab()
        except Exception:
            pass
        # Products balance tab
        try:
            self._create_products_balance_tab()
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", f"טעינת טאב 'מאזן מוצרים ופריטים' נכשלה: {e}")
            except Exception:
                pass

        # Apply per-tab colors on the main notebook
        try:
            self._apply_tab_colors()
        except Exception:
            pass

        # ----- Footer / Status -----
        self._create_status_bar()
        self._load_initial_settings()

    def _setup_right_click_text_menus(self):
        """הפעלת תפריט קליק ימני גלובלי לכל שדות הטקסט (Entry/Text) עם 'הדבק'.

        כולל גם פעולות שימושיות נוספות: גזור/העתק/בחר הכל. עובד על Windows.
        """
        import tkinter as tk
        self._rc_menu = tk.Menu(self.root, tearoff=0)
        # Use closures to send events to the currently targeted widget
        def do(event_name: str):
            try:
                if hasattr(self, '_rc_target') and self._rc_target:
                    self._rc_target.event_generate(event_name)
            except Exception:
                pass
        self._rc_menu.add_command(label="גזור", command=lambda: do("<<Cut>>"))
        self._rc_menu.add_command(label="העתק", command=lambda: do("<<Copy>>"))
        self._rc_menu.add_command(label="הדבק", command=lambda: do("<<Paste>>"))
        self._rc_menu.add_separator()
        self._rc_menu.add_command(label="בחר הכל", command=lambda: do("<<SelectAll>>"))

        def show_menu(event):
            try:
                self._rc_target = event.widget
                self._rc_menu.tk_popup(event.x_root, event.y_root)
            finally:
                try:
                    self._rc_menu.grab_release()
                except Exception:
                    pass

        # Bind to classic Tk widgets and ttk counterparts
        try:
            self.root.bind_class("Entry", "<Button-3>", show_menu)
        except Exception:
            pass
        try:
            self.root.bind_class("TEntry", "<Button-3>", show_menu)
        except Exception:
            pass
        try:
            self.root.bind_class("Text", "<Button-3>", show_menu)
        except Exception:
            pass
    
    def _create_status_bar(self):
        """יצירת שורת הסטטוס"""
        self.status_label = tk.Label(
            self.root,
            text="מוכן לעבודה",
            bg='#34495e',
            fg='white',
            anchor='w',
            padx=15,
            font=('Arial', 10)
        )
        self.status_label.pack(fill="x", side="bottom")
    
    def _load_initial_settings(self):
        """טעינת הגדרות ראשוניות"""
        # טעינה אוטומטית של קובץ מוצרים
        if self.settings.get("app.auto_load_products", True):
            products_file = self.settings.get("app.products_file", "קובץ מוצרים.xlsx")
            if os.path.exists(products_file) and hasattr(self, 'products_label'):
                self.products_file = os.path.abspath(products_file)
                self.products_label.config(text=os.path.basename(products_file))
                self._update_status(f"נטען קובץ מוצרים: {os.path.basename(products_file)}")
        # רענון רשימת ספקים עבור הקומבו בטאבים (אם כבר נוצרו)
        try:
            self._refresh_all_supplier_name_combos()
        except Exception:
            pass
    
    # Utility Methods

    # (Moved)
    
    # Utility Methods
    def _update_status(self, message):
        self.status_label.config(text=message)
        self.root.update()
    
    def _log_message(self, message):
        """Append a log line to results_text in RTL (Hebrew) orientation.

        שימוש בסימן RLM (\u200f) + תג עם justify='right' כדי לוודא יישור מימין לשמאל גם בטקסט מעורב.
        """
        try:
            if not hasattr(self.results_text, 'rtl_tag_configured'):
                try:
                    self.results_text.tag_configure('rtl', justify='right')
                    self.results_text.rtl_tag_configured = True  # type: ignore[attr-defined]
                except Exception:
                    pass
            rlm = '\u200f'
            self.results_text.insert(tk.END, rlm + str(message) + "\n", 'rtl')
            self.results_text.see(tk.END)
            self.root.update()
        except Exception:
            # Fallback without RTL formatting
            self.results_text.insert(tk.END, str(message) + "\n")
            self.results_text.see(tk.END)
            self.root.update()
    
    # clear handled in mixin

    def _apply_tab_colors(self):
        """Assign a distinct background/accent color to each top-level tab in the main notebook."""
        try:
            style = ttk.Style(self.root)
            try:
                # Use a theme that allows background customization
                style.theme_use('clam')
            except Exception:
                pass

            palette = [
                '#fef3c7',  # amber-100
                '#e0f2fe',  # sky-100
                '#e9d5ff',  # violet-200
                '#dcfce7',  # green-100
                '#ffe4e6',  # rose-100
                '#ede9fe',  # indigo-100
                '#f5d0fe',  # fuchsia-200
                '#d1fae5',  # emerald-100
                '#fee2e2',  # red-100
                '#e2e8f0',  # slate-200
            ]

            for i, tab_id in enumerate(self.notebook.tabs()):
                color = palette[i % len(palette)]
                # Color the tab's root frame
                try:
                    tab_frame = self.root.nametowidget(tab_id)
                    tab_frame.configure(bg=color)
                except Exception:
                    continue
                # Add a thin header bar as an accent
                try:
                    children = tab_frame.winfo_children()
                    header = tk.Frame(tab_frame, bg=color, height=6)
                    if children:
                        header.pack(side='top', fill='x', before=children[0])
                    else:
                        header.pack(side='top', fill='x')
                    header.pack_propagate(False)
                except Exception:
                    pass
        except Exception:
            pass
    
    # clear_all handled in mixin
    # ---- Suppliers helpers shared for tabs ----
    def _get_supplier_names(self):
        try:
            names = sorted({ (rec.get('business_name') or rec.get('name') or '').strip() for rec in getattr(self.data_processor,'suppliers',[]) if (rec.get('business_name') or rec.get('name')) })
            # Fallback: derive from drawings_data recipients ('נמען') if no explicit suppliers
            if not names:
                drawings = getattr(self.data_processor, 'drawings_data', []) or []
                names = sorted({ (r.get('נמען') or '').strip() for r in drawings if r.get('נמען') })
            return names
        except Exception:
            return []

    def _refresh_all_supplier_name_combos(self):
        names = self._get_supplier_names()
        # קליטת סחורה
        if hasattr(self, 'supplier_name_combo'):
            try:
                self.supplier_name_combo['values'] = names
                if self.supplier_name_var.get() and self.supplier_name_var.get() not in names:
                    self.supplier_name_var.set('')
            except Exception:
                pass
        # תעודת משלוח
        if hasattr(self, 'dn_supplier_name_combo'):
            try:
                self.dn_supplier_name_combo['values'] = names
                if self.dn_supplier_name_var.get() and self.dn_supplier_name_var.get() not in names:
                    self.dn_supplier_name_var.set('')
            except Exception:
                pass
        # מאזן מוצרים
        if hasattr(self, 'balance_supplier_combo'):
            try:
                self.balance_supplier_combo['values'] = names
                if self.balance_supplier_var.get() and self.balance_supplier_var.get() not in names:
                    self.balance_supplier_var.set('')
            except Exception:
                pass

    def _notify_suppliers_changed(self):
        """קריאה לאחר שינוי ברשימת הספקים לעדכון קומבואים."""
        self._refresh_all_supplier_name_combos()
