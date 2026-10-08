"""טאב יבוא — היסטוריית משלוחים ותשלומים מטורקיה."""
from __future__ import annotations

import os
import webbrowser
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

from . import theme

KINDS = (
    "תשלום ספק",
    "תשלום סחורה דרך משלח",
    "תשלום שילוח/עמילות",
    "משלוח יצא",
    "משלוח הגיע",
    "חשבונית",
    "פרופורמה",
    "מסמך",
)

CURRENCIES = ("", "USD", "ILS", "EUR", "TRY")
CERTAINTIES = ("ודאי", "סביר", "חלקי")

KIND_TAG = {
    "תשלום ספק": "pay_usd",
    "תשלום סחורה דרך משלח": "pay_fwd",
    "תשלום שילוח/עמילות": "pay_ils",
    "משלוח יצא": "ship_out",
    "משלוח הגיע": "ship_in",
    "חשבונית": "invoice",
    "פרופורמה": "proforma",
    "מסמך": "doc",
}

SKIP_USD_NOTES = ("הוחזרה", "נכשלה", "לא נספר בסה")
EXCEL_PATH = os.path.join("exports", "turkey_shipments_payments_timeline.xlsx")
COLS = ("date", "kind", "party", "description", "amount", "currency", "ils", "qty", "certainty", "files")
HEADERS = {
    "date": "תאריך",
    "kind": "סוג",
    "party": "גורם",
    "description": "תיאור",
    "amount": "סכום",
    "currency": "מטבע",
    "ils": "₪",
    "qty": "כמות",
    "certainty": "ודאות",
    "files": "קבצים",
}
WIDTHS = {
    "date": 95,
    "kind": 150,
    "party": 150,
    "description": 420,
    "amount": 100,
    "currency": 60,
    "ils": 90,
    "qty": 140,
    "certainty": 70,
    "files": 60,
}


def _parse_iso(date_str: str):
    try:
        return datetime.strptime((date_str or "")[:10], "%Y-%m-%d")
    except ValueError:
        return None


def _fmt_date(date_str: str) -> str:
    dt = _parse_iso(date_str)
    return dt.strftime("%d/%m/%Y") if dt else (date_str or "")


def _parse_display_date(text: str) -> str:
    raw = (text or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return raw


def _fmt_num(val) -> str:
    if val is None or val == "":
        return ""
    try:
        n = float(val)
    except (TypeError, ValueError):
        return str(val)
    if abs(n - round(n)) < 1e-9:
        return f"{int(round(n)):,}"
    return f"{n:,.2f}"


def _to_float(val):
    if val is None or val == "":
        return None
    try:
        return float(str(val).replace(",", "").replace(" ", ""))
    except (TypeError, ValueError):
        return None


def _counts_usd(ev: dict) -> bool:
    if ev.get("currency") != "USD":
        return False
    if ev.get("kind") not in ("תשלום ספק", "תשלום סחורה דרך משלח"):
        return False
    if not ev.get("amount"):
        return False
    notes = ev.get("notes") or ""
    return not any(s in notes for s in SKIP_USD_NOTES)


def _open_path(path: str) -> None:
    if not path:
        raise FileNotFoundError("אין נתיב")
    full = path if os.path.isabs(path) else os.path.abspath(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    os.startfile(full)


class ImportHistoryTabMixin:
    """Mixin לטאב יבוא."""

    def _create_import_history_tab(self):
        tab = tk.Frame(self.notebook, bg=theme.PAGE_BG)
        self.notebook.add(tab, text="יבוא")

        tk.Label(
            tab,
            text="היסטוריית יבוא מטורקיה — משלוחים, תשלומים ומסמכים",
            font=theme.FONT_TITLE,
            bg=theme.PAGE_BG,
            fg=theme.DARK,
        ).pack(pady=(10, 4))

        self._imp_summary_frame = tk.Frame(tab, bg=theme.PAGE_BG)
        self._imp_summary_frame.pack(fill="x", padx=16, pady=(0, 8))
        self._imp_summary_labels = []
        for _ in range(4):
            card = tk.Frame(self._imp_summary_frame, bg=theme.CARD_BG, highlightbackground=theme.BORDER, highlightthickness=1)
            card.pack(side="right", expand=True, fill="x", padx=6)
            val = tk.Label(card, text="—", font=(theme.FONT_FAMILY, 14, "bold"), bg=theme.CARD_BG, fg=theme.DARK)
            val.pack(pady=(10, 0))
            cap = tk.Label(card, text="", font=theme.FONT_SMALL, bg=theme.CARD_BG, fg=theme.SUBTEXT)
            cap.pack(pady=(0, 10))
            self._imp_summary_labels.append((val, cap))

        filt = tk.Frame(tab, bg=theme.PAGE_BG)
        filt.pack(fill="x", padx=16, pady=(0, 6))

        tk.Label(filt, text="חיפוש:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).pack(side="right", padx=(8, 4))
        self.imp_search_var = tk.StringVar()
        search = ttk.Entry(filt, textvariable=self.imp_search_var, width=28)
        search.pack(side="right", padx=4)
        search.bind("<KeyRelease>", lambda _e: self._load_import_history_tree())

        tk.Label(filt, text="סוג:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).pack(side="right", padx=(12, 4))
        self.imp_kind_var = tk.StringVar(value="הכל")
        self.imp_kind_combo = ttk.Combobox(
            filt, textvariable=self.imp_kind_var, values=("הכל",) + KINDS, width=22, state="readonly"
        )
        self.imp_kind_combo.pack(side="right", padx=4)
        self.imp_kind_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_import_history_tree())

        tk.Label(filt, text="גורם:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).pack(side="right", padx=(12, 4))
        self.imp_party_var = tk.StringVar(value="הכל")
        self.imp_party_combo = ttk.Combobox(filt, textvariable=self.imp_party_var, values=("הכל",), width=22, state="readonly")
        self.imp_party_combo.pack(side="right", padx=4)
        self.imp_party_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_import_history_tree())

        tk.Label(filt, text="שנה:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).pack(side="right", padx=(12, 4))
        self.imp_year_var = tk.StringVar(value="הכל")
        self.imp_year_combo = ttk.Combobox(filt, textvariable=self.imp_year_var, values=("הכל",), width=8, state="readonly")
        self.imp_year_combo.pack(side="right", padx=4)
        self.imp_year_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_import_history_tree())

        theme.make_button(filt, "נקה", kind="secondary", command=self._clear_import_filters).pack(side="left", padx=4)

        actions = tk.Frame(tab, bg=theme.PAGE_BG)
        actions.pack(fill="x", padx=16, pady=(0, 6))
        theme.make_button(actions, "פרטים / קבצים", kind="primary", command=self._open_selected_import_event).pack(side="right", padx=4)
        theme.make_button(actions, "הוסף אירוע", kind="success", command=lambda: self._import_event_form()).pack(side="right", padx=4)
        theme.make_button(actions, "ערוך", kind="purple", command=self._edit_selected_import_event).pack(side="right", padx=4)
        theme.make_button(actions, "מחק", kind="danger", command=self._delete_selected_import_event).pack(side="right", padx=4)
        theme.make_button(actions, "פתח אקסל", kind="secondary", command=self._open_import_excel).pack(side="left", padx=4)
        theme.make_button(actions, "רענן", kind="secondary", command=self._reload_import_history).pack(side="left", padx=4)

        table = ttk.LabelFrame(tab, text="ציר זמן", padding=6)
        table.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        wrap = tk.Frame(table, bg=theme.PAGE_BG)
        wrap.pack(fill="both", expand=True)
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)

        self.import_tree = theme.make_treeview(wrap, columns=COLS, show="headings")
        for col in COLS:
            self.import_tree.heading(col, text=HEADERS[col])
            anchor = "w" if col in ("description", "party", "kind", "qty") else "center"
            self.import_tree.column(col, width=WIDTHS[col], minwidth=50, anchor=anchor, stretch=(col == "description"))
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.import_tree.yview)
        hs = ttk.Scrollbar(wrap, orient="horizontal", command=self.import_tree.xview)
        self.import_tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.import_tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        self.import_tree.bind("<Double-1>", lambda _e: self._open_selected_import_event())

        self.import_tree.tag_configure("pay_usd", background="#ecfdf5")
        self.import_tree.tag_configure("pay_fwd", background="#f0fdf4")
        self.import_tree.tag_configure("pay_ils", background="#fff7ed")
        self.import_tree.tag_configure("ship_out", background="#eff6ff")
        self.import_tree.tag_configure("ship_in", background="#e0f2fe")
        self.import_tree.tag_configure("invoice", background="#f5f3ff")
        self.import_tree.tag_configure("proforma", background="#fae8ff")
        self.import_tree.tag_configure("doc", background="#f8fafc")

        self._refresh_import_filter_choices()
        self._load_import_history_tree()

    def _import_events(self):
        return list(getattr(self.data_processor, "import_history", []) or [])

    def _reload_import_history(self):
        try:
            self.data_processor.import_history = self.data_processor.load_import_history()
        except Exception as e:
            messagebox.showerror("שגיאה", str(e))
            return
        self._refresh_import_filter_choices()
        self._load_import_history_tree()
        if hasattr(self, "_update_status"):
            self._update_status("נטענה היסטוריית יבוא")

    def _refresh_import_filter_choices(self):
        events = self._import_events()
        parties = sorted({(e.get("party") or "").strip() for e in events if (e.get("party") or "").strip()})
        years = sorted({(e.get("date") or "")[:4] for e in events if (e.get("date") or "")[:4].isdigit()}, reverse=True)
        if hasattr(self, "imp_party_combo"):
            current = self.imp_party_var.get() or "הכל"
            self.imp_party_combo["values"] = ("הכל",) + tuple(parties)
            if current not in self.imp_party_combo["values"]:
                self.imp_party_var.set("הכל")
        if hasattr(self, "imp_year_combo"):
            current = self.imp_year_var.get() or "הכל"
            self.imp_year_combo["values"] = ("הכל",) + tuple(years)
            if current not in self.imp_year_combo["values"]:
                self.imp_year_var.set("הכל")

    def _clear_import_filters(self):
        self.imp_search_var.set("")
        self.imp_kind_var.set("הכל")
        self.imp_party_var.set("הכל")
        self.imp_year_var.set("הכל")
        self._load_import_history_tree()

    def _filtered_import_events(self):
        events = self._import_events()
        kind = self.imp_kind_var.get()
        party = self.imp_party_var.get()
        year = self.imp_year_var.get()
        q = (self.imp_search_var.get() or "").strip().lower()
        out = []
        for e in events:
            if kind and kind != "הכל" and e.get("kind") != kind:
                continue
            if party and party != "הכל" and e.get("party") != party:
                continue
            if year and year != "הכל" and not (e.get("date") or "").startswith(year):
                continue
            if q:
                blob = " ".join(
                    str(e.get(k) or "")
                    for k in ("date", "kind", "party", "description", "source", "notes", "qty", "amount")
                ).lower()
                if q not in blob:
                    continue
            out.append(e)
        out.sort(key=lambda e: (e.get("date") or "", e.get("id") or 0), reverse=True)
        return out

    def _update_import_summary(self):
        events = self._import_events()
        usd = sum(float(e["amount"]) for e in events if _counts_usd(e))
        ils = 0.0
        for e in events:
            if e.get("kind") == "תשלום שילוח/עמילות" and (e.get("ils") or e.get("currency") == "ILS"):
                try:
                    ils += float(e.get("ils") or e.get("amount") or 0)
                except (TypeError, ValueError):
                    pass
        ships = sum(1 for e in events if e.get("kind") in ("משלוח יצא", "משלוח הגיע"))
        meta = getattr(self.data_processor, "import_history_meta", {}) or {}
        akdem = meta.get("akdem_claimed_balance_usd")
        cards = [
            (f"{usd:,.0f} $", "תשלומים לספקים (USD נטו)"),
            (f"{ils:,.0f} ₪", "שילוח / עמילות"),
            (str(ships), "משלוחים (יצא + הגיע)"),
            (f"{float(akdem):,.0f} $" if akdem else "—", "יתרת Akdem פתוחה"),
        ]
        for (val, cap), (text, title) in zip(self._imp_summary_labels, cards):
            val.config(text=text)
            cap.config(text=title)

    def _load_import_history_tree(self):
        if not hasattr(self, "import_tree"):
            return
        for item in self.import_tree.get_children():
            self.import_tree.delete(item)
        rows = self._filtered_import_events()
        for e in rows:
            n_files = len(e.get("attachments") or [])
            tag = KIND_TAG.get(e.get("kind") or "", "doc")
            self.import_tree.insert(
                "",
                "end",
                iid=str(e.get("id")),
                values=(
                    _fmt_date(e.get("date") or ""),
                    e.get("kind") or "",
                    e.get("party") or "",
                    e.get("description") or "",
                    _fmt_num(e.get("amount")),
                    e.get("currency") or "",
                    _fmt_num(e.get("ils")),
                    e.get("qty") or "",
                    e.get("certainty") or "",
                    str(n_files) if n_files else "",
                ),
                tags=(tag,),
            )
        self._update_import_summary()

    def _selected_import_event_id(self):
        sel = self.import_tree.selection() if hasattr(self, "import_tree") else ()
        if not sel:
            return None
        try:
            return int(sel[0])
        except (TypeError, ValueError):
            return None

    def _open_selected_import_event(self):
        event_id = self._selected_import_event_id()
        if event_id is None:
            messagebox.showinfo("יבוא", "יש לבחור שורה")
            return
        self._show_import_event_dialog(event_id)

    def _edit_selected_import_event(self):
        event_id = self._selected_import_event_id()
        if event_id is None:
            messagebox.showinfo("יבוא", "יש לבחור שורה")
            return
        ev = self.data_processor.get_import_event(event_id)
        if not ev:
            messagebox.showerror("שגיאה", "אירוע לא נמצא")
            return
        self._import_event_form(ev)

    def _delete_selected_import_event(self):
        event_id = self._selected_import_event_id()
        if event_id is None:
            messagebox.showinfo("יבוא", "יש לבחור שורה")
            return
        ev = self.data_processor.get_import_event(event_id)
        label = ""
        if ev:
            label = f"{_fmt_date(ev.get('date') or '')} — {ev.get('kind') or ''}"
        if not messagebox.askyesno("מחיקה", f"למחוק את האירוע {label}?"):
            return
        if self.data_processor.delete_import_event(event_id):
            self._refresh_import_filter_choices()
            self._load_import_history_tree()
        else:
            messagebox.showerror("שגיאה", "לא ניתן למחוק את האירוע")

    def _open_import_excel(self):
        try:
            _open_path(EXCEL_PATH)
        except Exception as e:
            messagebox.showerror("שגיאה", f"לא ניתן לפתוח את האקסל: {e}")

    def _related_import_events(self, event: dict):
        party = (event.get("party") or "").strip()
        dt = _parse_iso(event.get("date") or "")
        if not party or not dt:
            return []
        event_id = event.get("id")
        window = timedelta(days=45)
        related = []
        for other in self._import_events():
            if other.get("id") == event_id:
                continue
            if (other.get("party") or "").strip() != party:
                continue
            odt = _parse_iso(other.get("date") or "")
            if not odt or abs((odt - dt).days) > window.days:
                continue
            related.append(other)
        related.sort(key=lambda e: e.get("date") or "")
        return related

    def _show_import_event_dialog(self, event_id: int):
        ev = self.data_processor.get_import_event(event_id)
        if not ev:
            messagebox.showerror("שגיאה", "אירוע לא נמצא")
            return

        dialog = tk.Toplevel()
        dialog.title("פרטי אירוע יבוא")
        dialog.geometry("860x640")
        dialog.minsize(720, 520)
        dialog.transient(self.root if hasattr(self, "root") else None)
        dialog.configure(bg=theme.PAGE_BG)
        try:
            dialog.grab_set()
        except Exception:
            pass

        main = tk.Frame(dialog, bg=theme.PAGE_BG)
        main.pack(fill="both", expand=True, padx=16, pady=12)

        header = tk.Label(main, text="", font=theme.FONT_SUBTITLE, bg=theme.PAGE_BG, fg=theme.DARK, anchor="e", justify="right")
        header.pack(fill="x")
        desc = tk.Label(main, text="", font=theme.FONT_BODY, bg=theme.PAGE_BG, fg=theme.TEXT, wraplength=800, justify="right", anchor="e")
        desc.pack(fill="x", pady=(4, 2))
        meta = tk.Label(main, text="", font=theme.FONT_SMALL, bg=theme.PAGE_BG, fg=theme.SUBTEXT, wraplength=800, justify="right", anchor="e")
        meta.pack(fill="x", pady=(0, 8))

        files_box = ttk.LabelFrame(main, text="קבצים נלווים", padding=6)
        files_box.pack(fill="both", expand=True)
        files_wrap = tk.Frame(files_box)
        files_wrap.pack(fill="both", expand=True)
        files_tree = ttk.Treeview(files_wrap, columns=("label", "kind", "status"), show="headings", height=7)
        files_tree.heading("label", text="שם")
        files_tree.heading("kind", text="סוג")
        files_tree.heading("status", text="סטטוס")
        files_tree.column("label", width=480, anchor="w")
        files_tree.column("kind", width=80, anchor="center")
        files_tree.column("status", width=80, anchor="center")
        fvs = ttk.Scrollbar(files_wrap, orient="vertical", command=files_tree.yview)
        files_tree.configure(yscrollcommand=fvs.set)
        files_tree.pack(side="left", fill="both", expand=True)
        fvs.pack(side="right", fill="y")

        mail_box = ttk.LabelFrame(main, text="מיילים", padding=6)
        mail_box.pack(fill="x", pady=(8, 0))
        mail_btns = tk.Frame(mail_box, bg=theme.PAGE_BG)
        mail_btns.pack(fill="x")

        related_box = ttk.LabelFrame(main, text="אירועים קשורים (אותו גורם, ±45 יום)", padding=6)
        related_box.pack(fill="both", expand=True, pady=(8, 0))
        related_tree = ttk.Treeview(related_box, columns=("date", "kind", "desc"), show="headings", height=4)
        related_tree.heading("date", text="תאריך")
        related_tree.heading("kind", text="סוג")
        related_tree.heading("desc", text="תיאור")
        related_tree.column("date", width=100, anchor="center")
        related_tree.column("kind", width=150, anchor="center")
        related_tree.column("desc", width=520, anchor="w")
        related_tree.pack(fill="both", expand=True)

        state = {"event_id": event_id}

        def selected_attachment():
            sel = files_tree.selection()
            if not sel:
                return None
            current = self.data_processor.get_import_event(state["event_id"]) or {}
            idx = int(sel[0].replace("att", ""))
            atts = current.get("attachments") or []
            if 0 <= idx < len(atts):
                return atts[idx]
            return None

        def open_selected_file():
            att = selected_attachment()
            if not att:
                messagebox.showinfo("קבצים", "יש לבחור קובץ", parent=dialog)
                return
            try:
                _open_path(att.get("path") or "")
            except Exception as e:
                messagebox.showerror("שגיאה", str(e), parent=dialog)

        def open_folder():
            att = selected_attachment()
            folder = None
            if att:
                p = att.get("path") or ""
                folder = os.path.dirname(os.path.abspath(p))
            if not folder or not os.path.isdir(folder):
                folder = os.path.abspath(self.data_processor.import_event_documents_dir(state["event_id"]))
            try:
                os.makedirs(folder, exist_ok=True)
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("שגיאה", str(e), parent=dialog)

        def add_file():
            paths = filedialog.askopenfilenames(
                title="צירוף קובץ לאירוע",
                parent=dialog,
                filetypes=[
                    ("מסמכים", "*.pdf *.xlsx *.xls *.txt *.zip *.doc *.docx"),
                    ("תמונות", "*.png *.jpg *.jpeg *.webp"),
                    ("כל הקבצים", "*.*"),
                ],
            )
            if not paths:
                return
            errors = []
            for path in paths:
                try:
                    self.data_processor.add_import_attachment(state["event_id"], path)
                except Exception as e:
                    errors.append(f"{os.path.basename(path)}: {e}")
            refresh()
            self._load_import_history_tree()
            if errors:
                messagebox.showerror("שגיאה", "\n".join(errors), parent=dialog)

        def remove_file():
            att = selected_attachment()
            if not att:
                messagebox.showinfo("קבצים", "יש לבחור קובץ", parent=dialog)
                return
            label = att.get("label") or att.get("path") or "הקובץ"
            if not messagebox.askyesno("מחיקה", f"להסיר את {label}?", parent=dialog):
                return
            if self.data_processor.delete_import_attachment(state["event_id"], att.get("path") or ""):
                refresh()
                self._load_import_history_tree()
            else:
                messagebox.showerror("שגיאה", "לא ניתן להסיר את הקובץ", parent=dialog)

        files_tree.bind("<Double-1>", lambda _e: open_selected_file())

        file_btns = tk.Frame(files_box, bg=theme.PAGE_BG)
        file_btns.pack(fill="x", pady=(6, 0))
        theme.make_button(file_btns, "הוסף קובץ", kind="success", command=add_file).pack(side="right", padx=4)
        theme.make_button(file_btns, "פתח", kind="primary", command=open_selected_file).pack(side="right", padx=4)
        theme.make_button(file_btns, "פתח תיקייה", kind="secondary", command=open_folder).pack(side="right", padx=4)
        theme.make_button(file_btns, "הסר קובץ", kind="danger", command=remove_file).pack(side="right", padx=4)

        def open_gmail(gid: str):
            webbrowser.open(f"https://mail.google.com/mail/u/0/#all/{gid}")

        def open_related(_e=None):
            sel = related_tree.selection()
            if not sel:
                return
            try:
                rid = int(sel[0])
            except (TypeError, ValueError):
                return
            state["event_id"] = rid
            refresh()

        related_tree.bind("<Double-1>", open_related)

        def refresh():
            current = self.data_processor.get_import_event(state["event_id"])
            if not current:
                dialog.destroy()
                return
            header.config(
                text=f"{_fmt_date(current.get('date') or '')}  ·  {current.get('kind') or ''}  ·  {current.get('party') or ''}"
            )
            amount_bits = []
            if current.get("amount") not in (None, ""):
                amount_bits.append(f"{_fmt_num(current.get('amount'))} {current.get('currency') or ''}".strip())
            if current.get("ils"):
                amount_bits.append(f"{_fmt_num(current.get('ils'))} ₪")
            if current.get("qty"):
                amount_bits.append(current.get("qty"))
            extra = "  ·  ".join(amount_bits)
            body = current.get("description") or ""
            if extra:
                body = f"{body}\n{extra}" if body else extra
            desc.config(text=body)
            bits = [
                f"ודאות: {current.get('certainty') or ''}",
                f"מקור: {current.get('source') or ''}",
            ]
            if current.get("notes"):
                bits.append(f"הערות: {current.get('notes')}")
            meta.config(text="\n".join(bits))

            for item in files_tree.get_children():
                files_tree.delete(item)
            for i, att in enumerate(current.get("attachments") or []):
                path = att.get("path") or ""
                ext = os.path.splitext(path)[1].lstrip(".").upper() or "קובץ"
                exists = os.path.exists(path)
                files_tree.insert(
                    "",
                    "end",
                    iid=f"att{i}",
                    values=(att.get("label") or os.path.basename(path), ext, "קיים" if exists else "חסר"),
                )

            for child in mail_btns.winfo_children():
                child.destroy()
            threads = current.get("gmail_threads") or []
            if not threads:
                tk.Label(mail_btns, text="אין קישורי מייל לאירוע זה", bg=theme.PAGE_BG, fg=theme.SUBTEXT, font=theme.FONT_SMALL).pack(anchor="e")
            else:
                for th in threads:
                    gid = th.get("id") or ""
                    label = th.get("label") or gid
                    theme.make_button(
                        mail_btns,
                        label,
                        kind="primary",
                        command=lambda g=gid: open_gmail(g),
                    ).pack(side="right", padx=4, pady=2)

            for item in related_tree.get_children():
                related_tree.delete(item)
            for other in self._related_import_events(current):
                related_tree.insert(
                    "",
                    "end",
                    iid=str(other.get("id")),
                    values=(
                        _fmt_date(other.get("date") or ""),
                        other.get("kind") or "",
                        other.get("description") or "",
                    ),
                )

        footer = tk.Frame(main, bg=theme.PAGE_BG)
        footer.pack(fill="x", pady=(10, 0))
        theme.make_button(footer, "ערוך", kind="purple", command=lambda: self._import_event_form(
            self.data_processor.get_import_event(state["event_id"]), on_saved=refresh
        )).pack(side="right", padx=4)
        theme.make_button(footer, "סגור", kind="dark", command=dialog.destroy).pack(side="left", padx=4)
        refresh()

    def _import_event_form(self, event=None, on_saved=None):
        editing = event is not None
        dialog = tk.Toplevel()
        dialog.title("עריכת אירוע" if editing else "הוספת אירוע יבוא")
        dialog.geometry("560x520")
        dialog.minsize(500, 460)
        dialog.transient(self.root if hasattr(self, "root") else None)
        dialog.configure(bg=theme.PAGE_BG)
        try:
            dialog.grab_set()
        except Exception:
            pass

        form = tk.Frame(dialog, bg=theme.PAGE_BG)
        form.pack(fill="both", expand=True, padx=18, pady=14)

        def row(r, label, widget):
            tk.Label(form, text=label, font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG).grid(row=r, column=0, sticky="e", padx=6, pady=5)
            widget.grid(row=r, column=1, sticky="ew", padx=6, pady=5)

        form.grid_columnconfigure(1, weight=1)

        date_var = tk.StringVar(value=_fmt_date((event or {}).get("date") or "") or datetime.now().strftime("%d/%m/%Y"))
        kind_var = tk.StringVar(value=(event or {}).get("kind") or KINDS[0])
        parties = sorted({(e.get("party") or "").strip() for e in self._import_events() if (e.get("party") or "").strip()})
        party_var = tk.StringVar(value=(event or {}).get("party") or "")
        amount_var = tk.StringVar(value="" if (event or {}).get("amount") in (None, "") else str(event.get("amount")))
        currency_var = tk.StringVar(value=(event or {}).get("currency") or "USD")
        ils_var = tk.StringVar(value="" if (event or {}).get("ils") in (None, "") else str(event.get("ils")))
        qty_var = tk.StringVar(value=(event or {}).get("qty") or "")
        certainty_var = tk.StringVar(value=(event or {}).get("certainty") or "סביר")
        desc_var = tk.StringVar(value=(event or {}).get("description") or "")
        notes_var = tk.StringVar(value=(event or {}).get("notes") or "")

        row(0, "תאריך:", ttk.Entry(form, textvariable=date_var, width=18))
        kind_cb = ttk.Combobox(form, textvariable=kind_var, values=KINDS, state="readonly", width=28)
        row(1, "סוג:", kind_cb)
        party_cb = ttk.Combobox(form, textvariable=party_var, values=parties, width=28)
        row(2, "גורם:", party_cb)
        row(3, "תיאור:", ttk.Entry(form, textvariable=desc_var, width=40))
        row(4, "סכום:", ttk.Entry(form, textvariable=amount_var, width=18))
        curr_cb = ttk.Combobox(form, textvariable=currency_var, values=CURRENCIES, width=10)
        row(5, "מטבע:", curr_cb)
        row(6, "₪:", ttk.Entry(form, textvariable=ils_var, width=18))
        row(7, "כמות:", ttk.Entry(form, textvariable=qty_var, width=28))
        cert_cb = ttk.Combobox(form, textvariable=certainty_var, values=CERTAINTIES, state="readonly", width=12)
        row(8, "ודאות:", cert_cb)
        row(9, "הערות:", ttk.Entry(form, textvariable=notes_var, width=40))

        def save():
            iso = _parse_display_date(date_var.get())
            if not iso or not _parse_iso(iso):
                messagebox.showerror("שגיאה", "תאריך לא תקין (dd/mm/yyyy)", parent=dialog)
                return
            payload = {
                "date": iso,
                "kind": kind_var.get().strip() or "מסמך",
                "party": party_var.get().strip(),
                "description": desc_var.get().strip(),
                "amount": _to_float(amount_var.get()),
                "currency": currency_var.get().strip(),
                "ils": _to_float(ils_var.get()),
                "qty": qty_var.get().strip(),
                "certainty": certainty_var.get().strip() or "סביר",
                "notes": notes_var.get().strip(),
            }
            try:
                if editing:
                    self.data_processor.update_import_event(int(event["id"]), payload)
                else:
                    payload["source"] = "הוזן ידנית"
                    self.data_processor.add_import_event(payload)
            except Exception as e:
                messagebox.showerror("שגיאה", str(e), parent=dialog)
                return
            self._refresh_import_filter_choices()
            self._load_import_history_tree()
            dialog.destroy()
            if on_saved:
                try:
                    on_saved()
                except Exception:
                    pass

        btns = tk.Frame(form, bg=theme.PAGE_BG)
        btns.grid(row=10, column=0, columnspan=2, pady=16)
        theme.make_button(btns, "שמור", kind="success", command=save).pack(side="right", padx=6)
        theme.make_button(btns, "ביטול", kind="secondary", command=dialog.destroy).pack(side="right", padx=6)
