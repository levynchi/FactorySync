import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class BusinessDetailsTabMixin:
    """Tab for managing business/company details (name, VAT id, logo, address, etc.)."""

    def _create_business_details_tab(self):
        """Create a standalone 'Business Details' tab on the main notebook (legacy placement)."""
        tab = tk.Frame(self.notebook, bg="#f7f9fa")
        self.notebook.add(tab, text="פרטי עסק")
        self._build_business_details_panel(tab)

    def _build_business_details_panel(self, parent: tk.Widget):
        """Build the Business Details UI into the given parent container.

        Used by the standalone tab and by the 'ניהול תוכנה' wrapper tab.
        """
        title = tk.Label(parent, text="פרטי העסק", font=("Arial", 16, "bold"), bg="#f7f9fa", fg="#2c3e50")
        title.pack(pady=(10, 6))

        body = tk.Frame(parent, bg="#f7f9fa")
        body.pack(fill="both", expand=True, padx=14, pady=8)

        # state vars
        self.bd_business_name = tk.StringVar()
        self.bd_vat_id = tk.StringVar()  # עוסק מורשה / ח.פ.
        self.bd_business_type = tk.StringVar(value="עוסק מורשה")
        self.bd_address = tk.StringVar()
        self.bd_city = tk.StringVar()
        self.bd_zip = tk.StringVar()
        self.bd_phone = tk.StringVar()
        self.bd_email = tk.StringVar()
        self.bd_website = tk.StringVar()
        self.bd_contact = tk.StringVar()
        self.bd_logo_path = tk.StringVar()
        self._business_logo_img = None  # keep reference

        # layout: two columns (form on right, logo on left)
        form = tk.Frame(body, bg="#f7f9fa")
        form.pack(side="right", fill="both", expand=True)

        logo_frame = tk.Frame(body, bg="#eef3f7", bd=1, relief="solid")
        logo_frame.pack(side="left", padx=(0, 16), pady=4)

        # Form fields
        def add_row(row, label_text, var, width=32):
            lbl = tk.Label(form, text=label_text, bg="#f7f9fa")
            lbl.grid(row=row, column=1, sticky="e", padx=(8, 4), pady=4)
            ent = tk.Entry(form, textvariable=var, width=width, justify="right")
            ent.grid(row=row, column=0, sticky="we", padx=(0, 4), pady=4)
            return ent

        form.grid_columnconfigure(0, weight=1)
        r = 0
        add_row(r, "שם העסק:", self.bd_business_name); r += 1
        # Type + VAT id in one row
        tk.Label(form, text="סוג עוסק:", bg="#f7f9fa").grid(row=r, column=1, sticky="e", padx=(8, 4), pady=4)
        type_combo = ttk.Combobox(form, textvariable=self.bd_business_type, values=["עוסק מורשה", "חברה בע""מ", "שותפות", "אחר"], state="readonly", width=14)
        type_combo.grid(row=r, column=0, sticky="w", padx=(0, 4), pady=4)
        r += 1
        add_row(r, "מס' עוסק/ח.פ:", self.bd_vat_id); r += 1
        add_row(r, "כתובת:", self.bd_address); r += 1
        add_row(r, "עיר:", self.bd_city); r += 1
        add_row(r, "מיקוד:", self.bd_zip); r += 1
        add_row(r, "טלפון:", self.bd_phone); r += 1
        add_row(r, "דוא""ל:", self.bd_email); r += 1
        add_row(r, "אתר:", self.bd_website); r += 1
        add_row(r, "איש קשר:", self.bd_contact); r += 1

        # Buttons
        btns = tk.Frame(form, bg="#f7f9fa")
        btns.grid(row=r, column=0, columnspan=2, sticky="we", pady=(8, 4))
        tk.Button(btns, text="💾 שמור", bg="#27ae60", fg="white", command=self._bd_save).pack(side="right", padx=(8, 0))
        tk.Button(btns, text="איפוס", command=self._bd_load_from_settings).pack(side="right")

        # Logo area
        tk.Label(logo_frame, text="לוגו העסק", font=("Arial", 12, "bold"), bg="#eef3f7").pack(padx=10, pady=(10, 6))
        self.bd_logo_canvas = tk.Label(logo_frame, bg="#ffffff", width=38, height=12, relief="sunken", bd=1, anchor="center")
        self.bd_logo_canvas.pack(padx=10, pady=(0, 8))

        pick_row = tk.Frame(logo_frame, bg="#eef3f7")
        pick_row.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(pick_row, text="בחר לוגו…", command=self._bd_pick_logo).pack(side="right")
        tk.Button(pick_row, text="הסר", command=self._bd_clear_logo).pack(side="right", padx=(6, 0))

        path_row = tk.Frame(logo_frame, bg="#eef3f7")
        path_row.pack(fill="x", padx=10, pady=(0, 12))
        tk.Label(path_row, text="נתיב:", bg="#eef3f7").pack(side="right")
        tk.Entry(path_row, textvariable=self.bd_logo_path, width=34, justify="right").pack(side="right", padx=(6, 0))

        # Load values
        self._bd_load_from_settings()

    def _create_software_management_tab(self):
        """Create a parent tab 'ניהול תוכנה' and place 'פרטי עסק' as a sub-tab within it."""
        parent_tab = tk.Frame(self.notebook, bg="#f7f9fa")
        self.notebook.add(parent_tab, text="ניהול תוכנה")

        # Inner notebook for management pages
        inner_nb = ttk.Notebook(parent_tab)
        inner_nb.pack(fill="both", expand=True)
        self.software_mgmt_notebook = inner_nb

        # Business Details sub-tab
        bd_tab = tk.Frame(inner_nb, bg="#f7f9fa")
        inner_nb.add(bd_tab, text="פרטי עסק")
        self._build_business_details_panel(bd_tab)

    # ---- Logo helpers ----
    def _bd_pick_logo(self):
        fn = filedialog.askopenfilename(
            title="בחר קובץ לוגו",
            filetypes=[("תמונות", "*.png;*.gif;*.jpg;*.jpeg;*.bmp"), ("כל הקבצים", "*.*")]
        )
        if not fn:
            return
        self.bd_logo_path.set(fn)
        self._bd_update_logo_preview()

    def _bd_clear_logo(self):
        self.bd_logo_path.set("")
        self._bd_update_logo_preview()

    def _bd_update_logo_preview(self):
        # Show image if possible; use subsample to shrink if very large
        self._business_logo_img = None
        try:
            path = (self.bd_logo_path.get() or "").strip()
            if not path or not os.path.exists(path):
                self.bd_logo_canvas.config(image="", text="אין תצוגה")
                return
            # Prefer PNG/GIF via tkinter.PhotoImage
            ext = os.path.splitext(path)[1].lower()
            if ext in (".png", ".gif"):
                img = tk.PhotoImage(file=path)
                # shrink if needed
                max_w, max_h = 300, 160
                w, h = img.width(), img.height()
                fx = max(1, int(w / max_w))
                fy = max(1, int(h / max_h))
                if fx > 1 or fy > 1:
                    img = img.subsample(fx, fy)
                self._business_logo_img = img
                self.bd_logo_canvas.config(image=img, text="")
            else:
                # Try Pillow if available for JPEGs, otherwise no preview
                try:
                    from PIL import Image, ImageTk  # type: ignore
                    im = Image.open(path)
                    im.thumbnail((300, 160))
                    img = ImageTk.PhotoImage(im)
                    self._business_logo_img = img
                    self.bd_logo_canvas.config(image=img, text="")
                except Exception:
                    self.bd_logo_canvas.config(image="", text="לא ניתן להציג תצוגה מקדימה")
        except Exception:
            self.bd_logo_canvas.config(image="", text="שגיאה בתצוגה")

    # ---- Settings IO ----
    def _bd_load_from_settings(self):
        s = getattr(self, "settings", None)
        if not s:
            return
        get = s.get
        self.bd_business_name.set(get("business.name", ""))
        self.bd_business_type.set(get("business.type", "עוסק מורשה") or "עוסק מורשה")
        self.bd_vat_id.set(get("business.vat_id", ""))
        self.bd_address.set(get("business.address", ""))
        self.bd_city.set(get("business.city", ""))
        self.bd_zip.set(get("business.zip", ""))
        self.bd_phone.set(get("business.phone", ""))
        self.bd_email.set(get("business.email", ""))
        self.bd_website.set(get("business.website", ""))
        self.bd_contact.set(get("business.contact", ""))
        self.bd_logo_path.set(get("business.logo_path", ""))
        self._bd_update_logo_preview()

    def _bd_save(self):
        s = getattr(self, "settings", None)
        if not s:
            return
        # minimal validation
        name = (self.bd_business_name.get() or "").strip()
        if not name:
            try:
                messagebox.showwarning("חסר שם", "נא להזין שם עסק")
            except Exception:
                pass
            return
        # Persist each field under business.* keys
        ok = True
        ok &= s.set("business.name", name)
        ok &= s.set("business.type", (self.bd_business_type.get() or "").strip())
        ok &= s.set("business.vat_id", (self.bd_vat_id.get() or "").strip())
        ok &= s.set("business.address", (self.bd_address.get() or "").strip())
        ok &= s.set("business.city", (self.bd_city.get() or "").strip())
        ok &= s.set("business.zip", (self.bd_zip.get() or "").strip())
        ok &= s.set("business.phone", (self.bd_phone.get() or "").strip())
        ok &= s.set("business.email", (self.bd_email.get() or "").strip())
        ok &= s.set("business.website", (self.bd_website.get() or "").strip())
        ok &= s.set("business.contact", (self.bd_contact.get() or "").strip())
        ok &= s.set("business.logo_path", (self.bd_logo_path.get() or "").strip())
        try:
            if ok:
                messagebox.showinfo("נשמר", "פרטי העסק נשמרו" )
            else:
                messagebox.showerror("שגיאה", "שמירת פרטי העסק נכשלה")
        except Exception:
            pass
