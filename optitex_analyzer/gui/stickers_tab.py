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

        self._stk_product_var = tk.StringVar()
        self._stk_qty_var = tk.StringVar()
        self._stk_size_var = tk.StringVar()
        self._stk_fabric_var = tk.StringVar()

        # Product name
        tk.Label(frm, text="שם המוצר:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        tk.Entry(frm, textvariable=self._stk_product_var, width=30).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Qty per package
        tk.Label(frm, text="כמות באריזה:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        tk.Entry(frm, textvariable=self._stk_qty_var, width=10).grid(row=0, column=3, padx=5, pady=5, sticky="w")

        # Size
        tk.Label(frm, text="מידה:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        tk.Entry(frm, textvariable=self._stk_size_var, width=15).grid(row=0, column=5, padx=5, pady=5, sticky="w")

        # Fabric type
        tk.Label(frm, text="סוג הבד:").grid(row=0, column=6, padx=5, pady=5, sticky="e")
        fabric_types = self._load_fabric_types_for_stickers()
        self._stk_fabric_cb = ttk.Combobox(frm, textvariable=self._stk_fabric_var, values=fabric_types, width=20)
        self._stk_fabric_cb.grid(row=0, column=7, padx=5, pady=5, sticky="w")

        # Buttons
        btns = tk.Frame(frm)
        btns.grid(row=0, column=8, padx=10, pady=5, sticky="w")
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

    def _stk_validate(self):
        name = (self._stk_product_var.get() or "").strip()
        qty = (self._stk_qty_var.get() or "").strip()
        size = (self._stk_size_var.get() or "").strip()
        fabric = (self._stk_fabric_var.get() or "").strip()
        if not name:
            try:
                messagebox.showwarning("אזהרה", "אנא הזן שם מוצר")
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
