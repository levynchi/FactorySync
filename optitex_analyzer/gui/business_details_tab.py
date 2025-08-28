import os
import zipfile
from datetime import datetime
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

        # track controls to enable/disable
        self._bd_entries = []
        self._bd_controls_to_toggle = []

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
            self._bd_entries.append(ent)
            return ent

        form.grid_columnconfigure(0, weight=1)
        r = 0
        add_row(r, "שם העסק:", self.bd_business_name); r += 1
        # Type + VAT id in one row
        tk.Label(form, text="סוג עוסק:", bg="#f7f9fa").grid(row=r, column=1, sticky="e", padx=(8, 4), pady=4)
        type_combo = ttk.Combobox(form, textvariable=self.bd_business_type, values=["עוסק מורשה", "חברה בע""מ", "שותפות", "אחר"], state="readonly", width=14)
        type_combo.grid(row=r, column=0, sticky="w", padx=(0, 4), pady=4)
        self._bd_controls_to_toggle.append(type_combo)
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
        self.bd_save_btn = tk.Button(btns, text="💾 שמור", bg="#27ae60", fg="white", command=self._bd_save)
        self.bd_save_btn.pack(side="right", padx=(8, 0))
        self.bd_reset_btn = tk.Button(btns, text="איפוס", command=self._bd_load_from_settings)
        self.bd_reset_btn.pack(side="right")
        # Unlock editing button (always enabled)
        self.bd_unlock_btn = tk.Button(btns, text="פתח לשינוי פרטי העסק", command=self._bd_prompt_enable_editing)
        self.bd_unlock_btn.pack(side="left")

        # Logo area
        tk.Label(logo_frame, text="לוגו העסק", font=("Arial", 12, "bold"), bg="#eef3f7").pack(padx=10, pady=(10, 6))
        self.bd_logo_canvas = tk.Label(logo_frame, bg="#ffffff", width=38, height=12, relief="sunken", bd=1, anchor="center")
        self.bd_logo_canvas.pack(padx=10, pady=(0, 8))

        pick_row = tk.Frame(logo_frame, bg="#eef3f7")
        pick_row.pack(fill="x", padx=10, pady=(0, 10))
        self.bd_pick_logo_btn = tk.Button(pick_row, text="בחר לוגו…", command=self._bd_pick_logo)
        self.bd_pick_logo_btn.pack(side="right")
        self.bd_clear_logo_btn = tk.Button(pick_row, text="הסר", command=self._bd_clear_logo)
        self.bd_clear_logo_btn.pack(side="right", padx=(6, 0))

        path_row = tk.Frame(logo_frame, bg="#eef3f7")
        path_row.pack(fill="x", padx=10, pady=(0, 12))
        tk.Label(path_row, text="נתיב:", bg="#eef3f7").pack(side="right")
        self.bd_logo_entry = tk.Entry(path_row, textvariable=self.bd_logo_path, width=34, justify="right")
        self.bd_logo_entry.pack(side="right", padx=(6, 0))
        self._bd_entries.append(self.bd_logo_entry)
        # also toggle logo buttons
        self._bd_controls_to_toggle.extend([self.bd_pick_logo_btn, self.bd_clear_logo_btn, self.bd_save_btn, self.bd_reset_btn])

        # Load values
        self._bd_load_from_settings()
        # Lock editing by default
        self._bd_set_editable(False)

    def _bd_set_editable(self, enabled: bool):
        """Enable/disable editing of Business Details fields and related buttons."""
        # Entries (tk.Entry)
        for ent in getattr(self, '_bd_entries', []):
            try:
                if enabled:
                    ent.configure(state='normal')
                else:
                    ent.configure(state='disabled', disabledbackground='#e9ecef', disabledforeground='#7a7a7a')
            except Exception:
                pass
        # Combo + buttons
        for w in getattr(self, '_bd_controls_to_toggle', []):
            try:
                if isinstance(w, ttk.Combobox):
                    w.configure(state=('readonly' if enabled else 'disabled'))
                else:
                    w.configure(state=('normal' if enabled else 'disabled'))
            except Exception:
                pass

    def _bd_prompt_enable_editing(self):
        try:
            if messagebox.askyesno("אישור", "האם אתה רוצה לשנות את פרטי העסק?"):
                self._bd_set_editable(True)
        except Exception:
            # Fallback: enable without prompt if messagebox fails
            self._bd_set_editable(True)

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

        # Backups sub-tab
        self._create_backups_tab(inner_nb)

    # ---- Backups Tab ----
    def _create_backups_tab(self, inner_nb: ttk.Notebook):
        tab = tk.Frame(inner_nb, bg="#f7f9fa")
        inner_nb.add(tab, text="גיבויים")

        title = tk.Label(tab, text="גיבוי כל נתוני התוכנה", font=("Arial", 16, "bold"), bg="#f7f9fa", fg="#2c3e50")
        title.pack(pady=(10, 6))

        body = tk.Frame(tab, bg="#f7f9fa")
        body.pack(fill="both", expand=True, padx=12, pady=8)

        # Controls
        ctrl = tk.Frame(body, bg="#f7f9fa")
        ctrl.pack(fill="x", pady=(0, 8))
        tk.Button(ctrl, text="צור גיבוי עכשיו", bg="#2980b9", fg="white", command=self._run_full_backup).pack(side="right", padx=(8, 0))
        tk.Button(ctrl, text="שחזר מגיבוי…", command=self._restore_from_backup).pack(side="right", padx=(8, 0))
        tk.Button(ctrl, text="פתח תיקיית גיבויים", command=self._open_backups_folder).pack(side="right", padx=(8, 0))
        tk.Button(ctrl, text="רענן רשימה", command=self._refresh_backups_list).pack(side="right")

        self.backup_status_label = tk.Label(body, text="", bg="#f7f9fa", fg="#2c3e50", anchor="e", justify="right")
        self.backup_status_label.pack(fill="x", pady=(0, 6))

        # Backups list
        list_frame = tk.Frame(body, bg="#f7f9fa")
        list_frame.pack(fill="both", expand=True)

        columns = ("name", "size", "date")
        self.backups_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        self.backups_tree.heading("name", text="שם קובץ")
        self.backups_tree.heading("size", text="גודל")
        self.backups_tree.heading("date", text="תאריך")
        self.backups_tree.column("name", anchor="e", width=460)
        self.backups_tree.column("size", anchor="center", width=120)
        self.backups_tree.column("date", anchor="center", width=180)
        self.backups_tree.pack(side="right", fill="both", expand=True)

        yscroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.backups_tree.yview)
        self.backups_tree.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side="left", fill="y")

        # Ensure backups dir exists and load list
        try:
            os.makedirs(self._get_backups_dir(), exist_ok=True)
        except Exception:
            pass
        self._refresh_backups_list()

    def _get_root_dir(self) -> str:
        # project root (folder containing main.py)
        return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))

    def _get_backups_dir(self) -> str:
        return os.path.join(self._get_root_dir(), "backups")

    def _format_size(self, bytes_val: int) -> str:
        try:
            for unit in ["B", "KB", "MB", "GB"]:
                if bytes_val < 1024.0:
                    return f"{bytes_val:3.1f} {unit}"
                bytes_val /= 1024.0
            return f"{bytes_val:.1f} TB"
        except Exception:
            return str(bytes_val)

    def _refresh_backups_list(self):
        try:
            dir_path = self._get_backups_dir()
            items = []
            if os.path.isdir(dir_path):
                for name in os.listdir(dir_path):
                    if not name.lower().endswith(".zip"):
                        continue
                    fp = os.path.join(dir_path, name)
                    try:
                        st = os.stat(fp)
                        size = self._format_size(st.st_size)
                        dt = datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")
                        items.append((name, size, dt))
                    except Exception:
                        pass
            # sort newest first by date string (already formatted), sort by name as fallback
            items.sort(key=lambda x: x[2], reverse=True)
            # update tree
            for iid in self.backups_tree.get_children():
                self.backups_tree.delete(iid)
            if not items:
                self.backups_tree.insert("", "end", values=("— אין גיבויים —", "", ""))
            else:
                for row in items:
                    self.backups_tree.insert("", "end", values=row)
        except Exception:
            pass

    def _open_backups_folder(self):
        try:
            path = self._get_backups_dir()
            os.makedirs(path, exist_ok=True)
            # Open in OS file explorer
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                import subprocess, sys
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", f"פתיחת התיקייה נכשלה: {e}")
            except Exception:
                pass

    def _run_full_backup(self):
        root_dir = self._get_root_dir()
        backups_dir = self._get_backups_dir()
        try:
            os.makedirs(backups_dir, exist_ok=True)
        except Exception:
            pass

        # Build file list: include data files, exclude code/virtual env/backups
        include_exts = {".json", ".xlsx", ".csv", ".txt"}
        exclude_dirs = {"optitex_analyzer", "src", ".git", "__pycache__", "backups", ".venv", "venv", "legacy"}

        files_to_zip = []
        for base, dirs, files in os.walk(root_dir):
            rel_base = os.path.relpath(base, root_dir)
            # Skip excluded dirs at any depth
            parts = set(rel_base.split(os.sep)) if rel_base != "." else set()
            if parts & exclude_dirs:
                # prune traversal
                dirs[:] = []
                continue
            # Always include everything under 'exports'
            if os.path.basename(base) == "exports" or "exports" in parts:
                for f in files:
                    files_to_zip.append(os.path.join(base, f))
                continue
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in include_exts:
                    files_to_zip.append(os.path.join(base, f))

        if not files_to_zip:
            try:
                messagebox.showwarning("אין נתונים", "לא נמצאו קבצים לגיבוי")
            except Exception:
                pass
            return

        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_name = f"backup_{ts}.zip"
        backup_path = os.path.join(backups_dir, backup_name)

        ok = True
        try:
            with zipfile.ZipFile(backup_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for abs_path in files_to_zip:
                    try:
                        arcname = os.path.relpath(abs_path, root_dir)
                        zf.write(abs_path, arcname)
                    except Exception:
                        ok = False
                        # continue with other files
                        continue
        except Exception:
            ok = False

        if ok:
            msg = f"נוצר גיבוי: {backup_name}"
            try:
                messagebox.showinfo("גיבוי הושלם", msg)
            except Exception:
                pass
            if hasattr(self, "backup_status_label"):
                try:
                    self.backup_status_label.config(text=msg)
                except Exception:
                    pass
            self._refresh_backups_list()
        else:
            try:
                messagebox.showwarning("הושלם חלקית", "הגיבוי נוצר אך ייתכן שחלק מהקבצים לא נכללו.")
            except Exception:
                pass

    def _restore_from_backup(self):
        try:
            initial_dir = self._get_backups_dir()
        except Exception:
            initial_dir = None
        fp = filedialog.askopenfilename(
            title="בחר קובץ גיבוי לשחזור",
            filetypes=[("קובץ גיבוי (ZIP)", "*.zip"), ("כל הקבצים", "*.*")],
            initialdir=initial_dir or os.getcwd(),
        )
        if not fp:
            return
        try:
            size_mb = os.path.getsize(fp) / (1024*1024)
        except Exception:
            size_mb = 0.0
        if not messagebox.askyesno("אישור שחזור", f"לשחזר את הגיבוי הבא?\n\n{os.path.basename(fp)} ({size_mb:.1f} MB)\n\nפעולה זו תחליף קבצי נתונים קיימים." ):
            return

        # Auto backup current state before restore
        try:
            self._run_full_backup()
        except Exception:
            pass

        root_dir = self._get_root_dir()
        ok = True
        err = None
        try:
            with zipfile.ZipFile(fp, 'r') as zf:
                for member in zf.infolist():
                    # prevent directory traversal
                    member_name = member.filename
                    # skip absolute or parent-traversal paths
                    if os.path.isabs(member_name):
                        continue
                    norm_target = os.path.normpath(os.path.join(root_dir, member_name))
                    if not norm_target.startswith(os.path.abspath(root_dir)):
                        continue
                    if member.is_dir():
                        os.makedirs(norm_target, exist_ok=True)
                        continue
                    os.makedirs(os.path.dirname(norm_target), exist_ok=True)
                    with zf.open(member, 'r') as src, open(norm_target, 'wb') as dst:
                        dst.write(src.read())
        except Exception as e:
            ok = False
            err = e

        if ok:
            try:
                messagebox.showinfo("שחזור הושלם", "השחזור הושלם בהצלחה. מומלץ להפעיל את התוכנה מחדש.")
            except Exception:
                pass
            if hasattr(self, "backup_status_label"):
                try:
                    self.backup_status_label.config(text=f"שוחזר מגיבוי: {os.path.basename(fp)}")
                except Exception:
                    pass
            self._refresh_backups_list()
        else:
            try:
                messagebox.showerror("שגיאת שחזור", f"השחזור נכשל: {err}")
            except Exception:
                pass

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
        # Persist immediately so it loads after restart
        try:
            if hasattr(self, 'settings') and self.settings:
                self.settings.set("business.logo_path", fn)
        except Exception:
            pass

    def _bd_clear_logo(self):
        self.bd_logo_path.set("")
        self._bd_update_logo_preview()
        # Persist removal immediately
        try:
            if hasattr(self, 'settings') and self.settings:
                self.settings.set("business.logo_path", "")
        except Exception:
            pass

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
                # After saving, return fields to locked (read-only) state
                try:
                    self._bd_set_editable(False)
                except Exception:
                    pass
            else:
                messagebox.showerror("שגיאה", "שמירת פרטי העסק נכשלה")
        except Exception:
            pass
