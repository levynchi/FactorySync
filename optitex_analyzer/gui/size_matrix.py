"""SizeMatrixFrame: per-size quantity boxes for fast multi-size line entry."""
import tkinter as tk
from tkinter import ttk
from . import theme


class SizeMatrixFrame(ttk.Frame):
    """שורת מידות עם תיבת כמות לכל מידה.

    כשאין מידות (מוצר חופשי / לא נבחר מוצר) מוצג מצב חלופי:
    שדה מידה חופשי + שדה כמות בודד, כדי לשמר את היכולת הקיימת.
    """

    MAX_PER_ROW = 10

    def __init__(self, parent, *, allow_free_entry: bool = True, hint: str = 'בחר מוצר כדי להציג מידות'):
        super().__init__(parent)
        self._allow_free_entry = allow_free_entry
        self._hint = hint
        self._sizes = []
        self._qty_vars = {}       # size -> StringVar
        self._entries = []        # ordered entries for keyboard navigation
        self._free_size_var = tk.StringVar()
        self._free_qty_var = tk.StringVar()
        self._build_empty_state()

    # ---- public API ----
    def set_sizes(self, sizes):
        """בנייה מחדש של תיבות הכמות לפי רשימת מידות (מנקה ערכים קודמים)."""
        sizes = [str(s).strip() for s in (sizes or []) if str(s).strip()]
        self._sizes = sizes
        self._qty_vars = {}
        self._entries = []
        for w in self.winfo_children():
            w.destroy()
        if not sizes:
            self._build_empty_state()
            return
        for idx, size in enumerate(sizes):
            row = (idx // self.MAX_PER_ROW) * 2
            # פריסה מימין לשמאל: העמודה הגבוהה ביותר ראשונה
            col = self.MAX_PER_ROW - 1 - (idx % self.MAX_PER_ROW)
            tk.Label(self, text=size, font=(theme.FONT_FAMILY, 9, 'bold')).grid(row=row, column=col, padx=3, pady=(2, 0))
            var = tk.StringVar()
            ent = tk.Entry(self, textvariable=var, width=5, justify='center')
            ent.grid(row=row + 1, column=col, padx=3, pady=(0, 3))
            self._bind_navigation(ent)
            self._qty_vars[size] = var
            self._entries.append(ent)

    def get_quantities(self):
        """מחזיר {מידה: כמות} רק עבור מידות עם כמות חיובית שלמה."""
        out = {}
        if self._sizes:
            for size, var in self._qty_vars.items():
                qty = self._parse_qty(var.get())
                if qty > 0:
                    out[size] = qty
        elif self._allow_free_entry:
            size = self._free_size_var.get().strip()
            qty = self._parse_qty(self._free_qty_var.get())
            if qty > 0:
                out[size] = qty
        return out

    def clear_quantities(self):
        for var in self._qty_vars.values():
            var.set('')
        self._free_size_var.set('')
        self._free_qty_var.set('')

    def set_quantity(self, size, qty):
        """ממלא תיבת כמות למידה קיימת (או מידה חופשית אם אין רשימה)."""
        size = str(size or '').strip()
        val = '' if qty in (None, '') else str(qty)
        if size in self._qty_vars:
            self._qty_vars[size].set(val)
        elif self._allow_free_entry and not self._sizes:
            self._free_size_var.set(size)
            self._free_qty_var.set(val)

    @property
    def has_sizes(self) -> bool:
        return bool(self._sizes)

    def focus_first(self):
        if self._entries:
            self._entries[0].focus_set()

    # ---- internals ----
    @staticmethod
    def _parse_qty(raw) -> int:
        try:
            qty = int(str(raw).strip())
            return qty if qty > 0 else 0
        except Exception:
            return 0

    def _build_empty_state(self):
        if self._allow_free_entry:
            tk.Label(self, text='מידה:', font=(theme.FONT_FAMILY, 9)).grid(row=0, column=3, padx=3)
            free_size = tk.Entry(self, textvariable=self._free_size_var, width=10, justify='center')
            free_size.grid(row=0, column=2, padx=3)
            tk.Label(self, text='כמות:', font=(theme.FONT_FAMILY, 9)).grid(row=0, column=1, padx=3)
            free_qty = tk.Entry(self, textvariable=self._free_qty_var, width=6, justify='center')
            free_qty.grid(row=0, column=0, padx=3)
            self._entries = [free_size, free_qty]
            self._bind_navigation(free_size)
            self._bind_navigation(free_qty)
        else:
            tk.Label(self, text=self._hint, fg=theme.SUBTEXT, font=(theme.FONT_FAMILY, 9)).grid(row=0, column=0, padx=4, pady=4)
            self._entries = []

    def _bind_navigation(self, entry):
        entry.bind('<Return>', self._focus_next)
        entry.bind('<KP_Enter>', self._focus_next)

    def _focus_next(self, event):
        try:
            idx = self._entries.index(event.widget)
        except ValueError:
            return
        if idx + 1 < len(self._entries):
            nxt = self._entries[idx + 1]
            nxt.focus_set()
            nxt.select_range(0, 'end')
        return 'break'
