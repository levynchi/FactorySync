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
from .formulas_tab import FormulasTabMixin
from .shipping_costs_tab import ShippingCostsTabMixin
from .orders_tab import OrdersTabMixin
from .stickers_tab import StickersTabMixin
from .rivhit_tab import RivhitTabMixin
from . import theme


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
    FormulasTabMixin,
    ShippingCostsTabMixin,
    OrdersTabMixin,
    StickersTabMixin,
    RivhitTabMixin,
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

        # ----- Header (כמו header של אתר) -----
        self._create_header()

        # ----- Notebook & Tabs -----
        # פס הטאבים המובנה מוסתר; הניווט נעשה בסרגל דינמי שנשבר לשורות (נבנה בסוף)
        self.notebook = ttk.Notebook(self.root, style="Main.TNotebook")
        self.notebook.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        # Create each tab from its mixin
    # Removed standalone converter tab: converter now embedded as sub-tab inside 'מנהל ציורים'
    # Returned (cut) drawings now embedded inside 'מנהל ציורים' tab
        # Software management tab (contains 'פרטי עסק' and 'גיבויים')
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
        # Formulas and calculations tab
        try:
            self._create_formulas_tab()
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", f"טעינת טאב 'נוסחאות וחישובים' נכשלה: {e}")
            except Exception:
                pass
        # Shipping costs and fabrics tab
        try:
            self._create_shipping_costs_tab()
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", f"טעינת טאב 'עלויות משלוחים ובדים' נכשלה: {e}")
            except Exception:
                pass
        # Orders management tab
        try:
            self._create_orders_tab()
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", f"טעינת טאב 'יצירת הזמנה' נכשלה: {e}")
            except Exception:
                pass

        # Stickers/Labels printing tab
        try:
            self._create_stickers_tab()
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", f"טעינת טאב 'מדבקות' נכשלה: {e}")
            except Exception:
                pass

        # Rivhit products tab
        try:
            self._create_rivhit_tab()
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", f"טעינת טאב 'ריווחית' נכשלה: {e}")
            except Exception:
                pass

        # ----- Navigation bar (dynamic, wraps to rows) -----
        self._create_nav_bar()

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
    
    def _create_header(self):
        """כותרת עליונה בסגנון אתר: לוגו + שם התוכנה על רקע כהה."""
        header = tk.Frame(self.root, bg=theme.DARK, height=54)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        title = tk.Label(
            header,
            text="FactorySync",
            bg=theme.DARK,
            fg=theme.CARD_BG,
            font=(theme.FONT_FAMILY, 16, "bold"),
        )
        title.pack(side="right", padx=(0, 18))

        subtitle = tk.Label(
            header,
            text="ניהול ייצור, מלאי ומשלוחים",
            bg=theme.DARK,
            fg=theme.MUTED,
            font=(theme.FONT_FAMILY, 10),
        )
        subtitle.pack(side="right", padx=(0, 12), pady=(6, 0))

        # לוגו (אם קיים)
        try:
            from PIL import Image, ImageTk
            logo_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'assets', 'labels', 'logo.png')
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                img.thumbnail((40, 40), Image.LANCZOS)
                self._header_logo_img = ImageTk.PhotoImage(img)
                tk.Label(header, image=self._header_logo_img, bg=theme.DARK).pack(side="right", padx=(0, 10))
        except Exception:
            pass

        # פס הדגשה בצבע ראשי מתחת ל-header
        tk.Frame(self.root, bg=theme.PRIMARY, height=3).pack(fill="x", side="top")

    def _create_nav_bar(self):
        """סרגל ניווט דינמי במקום פס הטאבים המובנה.

        הכפתורים נשברים אוטומטית לשורות לפי רוחב החלון, כך ששום כיתוב לא נחתך.
        """
        self._nav_bar = tk.Frame(self.root, bg=theme.DARK)
        try:
            self._nav_bar.pack(fill="x", side="top", before=self.notebook)
        except Exception:
            self._nav_bar.pack(fill="x", side="top")

        self._nav_buttons = []
        self._nav_gap_x = 6
        self._nav_gap_y = 6

        def on_click(idx):
            try:
                self.notebook.select(idx)
            except Exception:
                pass

        for i, tab_id in enumerate(self.notebook.tabs()):
            try:
                text = self.notebook.tab(tab_id, 'text')
            except Exception:
                text = ''
            btn = tk.Label(
                self._nav_bar,
                text=text,
                font=(theme.FONT_FAMILY, 10, 'bold'),
                bg=theme.DARK,
                fg='#cbd5e1',
                padx=12,
                pady=5,
                cursor='hand2',
            )
            btn.bind('<Button-1>', lambda e, idx=i: on_click(idx))

            def on_enter(e, b=btn):
                if getattr(b, '_nav_active', False):
                    return
                b.configure(bg=theme.DARK_2, fg='#ffffff')

            def on_leave(e, b=btn):
                if getattr(b, '_nav_active', False):
                    return
                b.configure(bg=theme.DARK, fg='#cbd5e1')

            btn.bind('<Enter>', on_enter)
            btn.bind('<Leave>', on_leave)
            self._nav_buttons.append(btn)

        self._nav_last_width = 0
        self._nav_bar.bind('<Configure>', self._relayout_nav_bar)
        try:
            self.notebook.bind('<<NotebookTabChanged>>', self._update_nav_selection, add='+')
        except Exception:
            pass
        self.root.after_idle(lambda: (self._relayout_nav_bar(), self._update_nav_selection()))

    def _relayout_nav_bar(self, event=None):
        """פריסת כפתורי הניווט מימין לשמאל עם שבירה אוטומטית לשורות."""
        try:
            bar = self._nav_bar
            width = bar.winfo_width()
            if width <= 1:
                return
            if event is not None and width == self._nav_last_width:
                return
            self._nav_last_width = width

            gap_x, gap_y = self._nav_gap_x, self._nav_gap_y
            pad = 8
            avail = width - 2 * pad
            x = pad  # מרחק מהקצה הימני
            row = 0
            row_h = 0
            for btn in self._nav_buttons:
                bw = btn.winfo_reqwidth()
                bh = btn.winfo_reqheight()
                row_h = max(row_h, bh)
                if x > pad and (x + bw) > avail + pad:
                    row += 1
                    x = pad
                btn.place(relx=1.0, x=-(x + bw), y=row * (row_h + gap_y) + gap_y, width=bw, height=bh)
                x += bw + gap_x
            total_h = (row + 1) * (row_h + gap_y) + gap_y
            bar.configure(height=total_h)
        except Exception:
            pass

    def _update_nav_selection(self, event=None):
        """הדגשת הטאב הפעיל בסרגל הניווט."""
        try:
            current = self.notebook.index(self.notebook.select())
        except Exception:
            return
        for i, btn in enumerate(self._nav_buttons):
            try:
                if i == current:
                    btn._nav_active = True
                    btn.configure(bg=theme.PRIMARY, fg='#ffffff')
                else:
                    btn._nav_active = False
                    btn.configure(bg=theme.DARK, fg='#cbd5e1')
            except Exception:
                pass

    def _create_status_bar(self):
        """יצירת שורת הסטטוס"""
        self.status_label = tk.Label(
            self.root,
            text="מוכן לעבודה",
            bg=theme.DARK,
            fg='#e2e8f0',
            anchor='w',
            padx=15,
            pady=4,
            font=(theme.FONT_FAMILY, 9)
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
        
        # טעינת נתוני הזמנות ולקוחות
        try:
            self._load_orders_from_file()
            self._load_customers_from_file()
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
        # Fabrics send supplier combo
        if hasattr(self, 'fs_supplier_combo'):
            try:
                self.fs_supplier_combo['values'] = names
                if getattr(self, 'fs_supplier_var', None) and self.fs_supplier_var.get() and self.fs_supplier_var.get() not in names:
                    self.fs_supplier_var.set('')
            except Exception:
                pass

    def _notify_suppliers_changed(self):
        """קריאה לאחר שינוי ברשימת הספקים לעדכון קומבואים."""
        self._refresh_all_supplier_name_combos()
