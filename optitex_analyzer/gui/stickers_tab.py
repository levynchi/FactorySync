"""Stickers tab: input row (product name, qty per package, size, fabric type) + table with persistence."""
import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


class StickersTabMixin:
    def _create_stickers_tab(self):
        """Create the 'מדבקות' tab UI and load persisted data."""
        # Notebook must be defined by MainWindow
        self.stickers_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.stickers_tab, text="מדבקות")

        container = self.stickers_tab

        # Header
        tk.Label(
            container,
            text="ניהול מדבקות למוצרים",
            font=('Arial', 14, 'bold')
        ).pack(pady=(10, 5))

        # Input row
        frm = ttk.LabelFrame(container, text="שורת קליטה", padding=10)
        frm.pack(fill="x", padx=15, pady=10)

        self._stk_main_category_var = tk.StringVar()
        self._stk_product_var = tk.StringVar()
        self._stk_qty_var = tk.StringVar()
        self._stk_size_var = tk.StringVar()
        self._stk_fabric_var = tk.StringVar()

        # Main category (default: בגדים)
        tk.Label(frm, text="קטגוריה ראשית:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        try:
            cats = self._stk_get_main_categories()
        except Exception:
            cats = []
        if 'בגדים' not in cats:
            cats = ['בגדים'] + [c for c in cats if c != 'בגדים']
        if not self._stk_main_category_var.get():
            self._stk_main_category_var.set('בגדים')
        self._stk_main_category_cb = ttk.Combobox(frm, textvariable=self._stk_main_category_var, values=cats, width=18, state='readonly')
        self._stk_main_category_cb.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        try:
            self._stk_main_category_cb.bind('<<ComboboxSelected>>', lambda e: self._stk_on_main_category_change())
        except Exception:
            pass

        # Product name (filtered by selected main category)
        tk.Label(frm, text="שם המוצר:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self._stk_product_cb = ttk.Combobox(frm, textvariable=self._stk_product_var, values=[], width=30)
        self._stk_product_cb.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        # Initialize product options for default category
        try:
            self._stk_refresh_product_options()
        except Exception:
            pass

        # Qty per package
        tk.Label(frm, text="כמות באריזה:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        tk.Entry(frm, textvariable=self._stk_qty_var, width=10).grid(row=0, column=5, padx=5, pady=5, sticky="w")

        # Size
        tk.Label(frm, text="מידה:").grid(row=0, column=6, padx=5, pady=5, sticky="e")
        tk.Entry(frm, textvariable=self._stk_size_var, width=15).grid(row=0, column=7, padx=5, pady=5, sticky="w")

        # Fabric type
        tk.Label(frm, text="סוג הבד:").grid(row=0, column=8, padx=5, pady=5, sticky="e")
        fabric_types = self._load_fabric_types_for_stickers()
        self._stk_fabric_cb = ttk.Combobox(frm, textvariable=self._stk_fabric_var, values=fabric_types, width=20)
        self._stk_fabric_cb.grid(row=0, column=9, padx=5, pady=5, sticky="w")

        # Buttons
        btns = tk.Frame(frm)
        btns.grid(row=0, column=10, padx=10, pady=5, sticky="w")
        tk.Button(btns, text="➕ הוסף", bg="#27ae60", fg="white", command=self._stk_add).pack(side="left", padx=4)
        tk.Button(btns, text="🧹 נקה", bg="#95a5a6", fg="white", command=self._stk_clear_inputs).pack(side="left", padx=4)

        # Table
        tbl_frame = ttk.LabelFrame(container, text="טבלת מדבקות", padding=10)
        tbl_frame.pack(fill="both", expand=True, padx=15, pady=10)

        cols = ("שם המוצר", "כמות באריזה", "מידה", "סוג הבד")
        self._stk_tree = ttk.Treeview(tbl_frame, columns=cols, show="headings")
        for c in cols:
            self._stk_tree.heading(c, text=c)
            self._stk_tree.column(c, width=(240 if c == "שם המוצר" else 140), anchor="center")

        vs = ttk.Scrollbar(tbl_frame, orient="vertical", command=self._stk_tree.yview)
        self._stk_tree.configure(yscrollcommand=vs.set)
        self._stk_tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        tbl_frame.grid_rowconfigure(0, weight=1)
        tbl_frame.grid_columnconfigure(0, weight=1)

        actions = tk.Frame(container)
        actions.pack(fill="x", padx=15, pady=(0, 10))
        tk.Button(actions, text="🗑️ מחק נבחר", bg="#e67e22", fg="white", command=self._stk_delete_selected).pack(side="left", padx=5)
        tk.Button(actions, text="💾 שמור ל-Excel", bg="#3498db", fg="white", command=self._stk_export_excel).pack(side="left", padx=5)

        # Data
        self._stickers_data = []
        self._stk_load()
        self._stk_refresh()

    # ---------- Helpers ----------
    def _stickers_file_path(self) -> str:
        return os.path.join(os.getcwd(), "stickers_data.json")

    def _load_fabric_types_for_stickers(self):
        try:
            # Try project root json
            p = os.path.join(os.getcwd(), "fabric_types.json")
            if os.path.exists(p):
                with open(p, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    if data and isinstance(data[0], dict):
                        keys = list(data[0].keys())
                        for k in ("name", "שם", "type", "סוג"):
                            if k in keys:
                                return [str(d.get(k, "")) for d in data]
                        return [str(d.get(keys[0], "")) for d in data]
                    return [str(x) for x in data]
        except Exception:
            pass
        return []

    def _stk_get_main_categories(self):
        """Return list of main category names from data processor or json file."""
        # Try data_processor
        try:
            mcs = getattr(self, 'data_processor', None)
            if mcs is not None:
                lst = getattr(self.data_processor, 'main_categories', []) or []
                names = []
                for c in lst:
                    name = (c.get('name') or c.get('שם') or '').strip()
                    if name:
                        names.append(name)
                if names:
                    return sorted(set(names), key=lambda x: (x!='בגדים', x))
        except Exception:
            pass
        # Fallback: read from main_categories.json if exists
        try:
            p = os.path.join(os.getcwd(), 'main_categories.json')
            if os.path.exists(p):
                with open(p, 'r', encoding='utf-8') as f:
                    data = json.load(f) or []
                names = []
                for c in data:
                    if isinstance(c, dict):
                        nm = (c.get('name') or c.get('שם') or '').strip()
                        if nm:
                            names.append(nm)
                    elif isinstance(c, str) and c.strip():
                        names.append(c.strip())
                if names:
                    return sorted(set(names), key=lambda x: (x!='בגדים', x))
        except Exception:
            pass
        return ['בגדים']

    def _stk_get_products_for_category(self, main_category: str):
        """Return product names filtered by main_category from catalog; fallback to all."""
        names = []
        try:
            model_names = getattr(self, 'data_processor', None)
            if model_names is not None:
                lst = getattr(self.data_processor, 'product_model_names', []) or []
                if lst:
                    for r in lst:
                        n = (r.get('name') or '').strip()
                        mc = (r.get('main_category') or 'בגדים').strip()
                        if n and (not main_category or mc == main_category):
                            names.append(n)
            if not names:
                catalog = getattr(self, 'data_processor', None)
                if catalog is not None:
                    cl = getattr(self.data_processor, 'products_catalog', []) or []
                    for r in cl:
                        n = (r.get('name') or '').strip()
                        mc = (r.get('main_category') or 'בגדים').strip()
                        if n and (not main_category or mc == main_category):
                            names.append(n)
        except Exception:
            names = []
        # Dedup + sort
        seen = set(); names = [x for x in names if not (x in seen or seen.add(x))]
        return sorted(names)

    def _stk_refresh_product_options(self):
        cat = (self._stk_main_category_var.get() or '').strip()
        try:
            options = self._stk_get_products_for_category(cat)
        except Exception:
            options = []
        try:
            self._stk_product_cb['values'] = options
        except Exception:
            pass

    def _stk_on_main_category_change(self):
        try:
            self._stk_product_var.set('')
        except Exception:
            pass
        self._stk_refresh_product_options()

    def _stk_validate(self):
        name = (self._stk_product_var.get() or "").strip()
        qty = (self._stk_qty_var.get() or "").strip()
        size = (self._stk_size_var.get() or "").strip()
        fabric = (self._stk_fabric_var.get() or "").strip()
        # Enforce selection from the filtered list
        try:
            allowed = list(self._stk_product_cb['values'] or [])
        except Exception:
            allowed = []
        if not name:
            try:
                messagebox.showwarning("אזהרה", "אנא הזן שם מוצר")
            except Exception:
                pass
            return None
        if allowed and name not in allowed:
            try:
                messagebox.showwarning("אזהרה", "יש לבחור מוצר מהרשימה לפי הקטגוריה הראשית")
            except Exception:
                pass
            return None
        if not qty:
            try:
                messagebox.showwarning("אזהרה", "אנא הזן כמות באריזה")
            except Exception:
                pass
            return None
        try:
            qty_num = int(qty)
            if qty_num <= 0:
                raise ValueError()
        except Exception:
            try:
                messagebox.showwarning("אזהרה", "כמות באריזה חייבת להיות מספר שלם חיובי")
            except Exception:
                pass
            return None
        if not size:
            try:
                messagebox.showwarning("אזהרה", "אנא הזן מידה")
            except Exception:
                pass
            return None
        if not fabric:
            try:
                messagebox.showwarning("אזהרה", "אנא בחר סוג בד")
            except Exception:
                pass
            return None
        return {"שם המוצר": name, "כמות באריזה": qty_num, "מידה": size, "סוג הבד": fabric}

    def _stk_add(self):
        row = self._stk_validate()
        if not row:
            return
        self._stickers_data.append(row)
        self._stk_refresh()
        self._stk_save()
        self._stk_clear_inputs()
        try:
            self._update_status("שורה נוספה לטבלת המדבקות")
        except Exception:
            pass

    def _stk_clear_inputs(self):
        self._stk_product_var.set("")
        self._stk_qty_var.set("")
        self._stk_size_var.set("")
        self._stk_fabric_var.set("")

    def _stk_refresh(self):
        for it in self._stk_tree.get_children():
            self._stk_tree.delete(it)
        for r in self._stickers_data:
            self._stk_tree.insert("", "end", values=(
                r.get("שם המוצר", ""),
                r.get("כמות באריזה", ""),
                r.get("מידה", ""),
                r.get("סוג הבד", ""),
            ))

    def _stk_delete_selected(self):
        sel = self._stk_tree.selection()
        if not sel:
            try:
                messagebox.showwarning("אזהרה", "אנא בחר שורה למחיקה")
            except Exception:
                pass
            return
        try:
            if not messagebox.askyesno("אישור", "למחוק את השורה הנבחרת?"):
                return
        except Exception:
            pass
        item = self._stk_tree.item(sel[0])
        vals = item.get('values', [])
        if vals:
            target = {"שם המוצר": vals[0], "כמות באריזה": vals[1], "מידה": vals[2], "סוג הבד": vals[3]}
            new_data, deleted = [], False
            for r in self._stickers_data:
                if (not deleted and r.get("שם המוצר") == target["שם המוצר"] and
                    r.get("כמות באריזה") == target["כמות באריזה"] and
                    r.get("מידה") == target["מידה"] and
                    r.get("סוג הבד") == target["סוג הבד"]):
                    deleted = True
                    continue
                new_data.append(r)
            self._stickers_data = new_data
            self._stk_refresh()
            self._stk_save()
            try:
                self._update_status("שורה נמחקה מטבלת המדבקות")
            except Exception:
                pass

    def _stk_export_excel(self):
        if not self._stickers_data:
            try:
                messagebox.showwarning("אזהרה", "אין נתונים לייצוא")
            except Exception:
                pass
            return
        try:
            df = pd.DataFrame(self._stickers_data)
            file_path = filedialog.asksaveasfilename(
                title="שמור טבלת מדבקות",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if file_path:
                if file_path.lower().endswith('.csv'):
                    df.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(file_path, index=False)
                try:
                    messagebox.showinfo("הצלחה", f"הטבלה נשמרה בהצלחה:\n{file_path}")
                except Exception:
                    pass
        except Exception as e:
            try:
                messagebox.showerror("שגיאה", f"שגיאה בייצוא: {e}")
            except Exception:
                pass

    def _stk_load(self):
        try:
            p = self._stickers_file_path()
            if os.path.exists(p):
                with open(p, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    self._stickers_data = data
        except Exception:
            self._stickers_data = []

    def _stk_save(self):
        try:
            p = self._stickers_file_path()
            with open(p, 'w', encoding='utf-8') as f:
                json.dump(self._stickers_data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
