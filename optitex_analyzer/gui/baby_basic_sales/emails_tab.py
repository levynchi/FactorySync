"""טאב 'מיילים ומסמכים' של לקוח בייבי בייסיק — כרטיס לקוח, התכתבויות מ-Gmail וקבצים."""
from __future__ import annotations

import os
import webbrowser
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog, ttk
import tkinter as tk

from .. import theme

GMAIL_BASE = "https://mail.google.com/mail/u/0/#all/"
ALL = "הכל"


def _fmt_date(s: str) -> str:
    try:
        return datetime.strptime((s or "")[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return s or ""


def _open_path(path: str) -> None:
    if not path:
        raise FileNotFoundError("אין נתיב לקובץ (הקובץ לא נמצא מקומית — פתח דרך Gmail)")
    full = path if os.path.isabs(path) else os.path.abspath(path)
    if not os.path.exists(full):
        raise FileNotFoundError(full)
    os.startfile(full)


def _who(addr: str) -> str:
    a = (addr or "").lower()
    if "arye" in a:
        return "אריה"
    if "omer" in a:
        return "עומר"
    if "hadar" in a:
        return "הדר"
    if "cadena" in a:
        return "קדנה"
    if "printmarket" in a:
        return "פרינטמרקט"
    return addr or ""


class BabyBasicEmailsTabMixin:
    """Mixin: self.data_processor חייב להיות זמין."""

    # ------------------------------------------------------------------ build
    def _build_bb_emails_tab(self, parent):
        self._bb_email_selected_thread = None

        card = ttk.LabelFrame(parent, text="כרטיס לקוח", padding=8)
        card.pack(fill="x", padx=10, pady=(8, 4))
        self._bb_customer_card = tk.Frame(card, bg=theme.PAGE_BG)
        self._bb_customer_card.pack(fill="x")

        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        mails_page = tk.Frame(nb, bg=theme.PAGE_BG)
        docs_page = tk.Frame(nb, bg=theme.PAGE_BG)
        nb.add(mails_page, text="התכתבויות")
        nb.add(docs_page, text="כל המסמכים")
        self._bb_emails_nb = nb

        self._build_bb_mails_page(mails_page)
        self._build_bb_docs_page(docs_page)
        self._bb_refresh_emails()

    def _build_bb_mails_page(self, page):
        filt = tk.Frame(page, bg=theme.PAGE_BG)
        filt.pack(fill="x", pady=6)
        tk.Label(filt, text="קטגוריה:", bg=theme.PAGE_BG).pack(side="right", padx=(4, 2))
        self._bb_email_cat_var = tk.StringVar(value=ALL)
        self._bb_email_cat_combo = ttk.Combobox(filt, textvariable=self._bb_email_cat_var, state="readonly", width=22, justify="right")
        self._bb_email_cat_combo.pack(side="right", padx=4)
        self._bb_email_cat_combo.bind("<<ComboboxSelected>>", lambda e: self._bb_refresh_email_threads())
        tk.Label(filt, text="חיפוש:", bg=theme.PAGE_BG).pack(side="right", padx=(10, 2))
        self._bb_email_search_var = tk.StringVar()
        ent = ttk.Entry(filt, textvariable=self._bb_email_search_var, width=30, justify="right")
        ent.pack(side="right", padx=4)
        ent.bind("<KeyRelease>", lambda e: self._bb_refresh_email_threads())
        theme.make_button(filt, "הוסף רשומה ידנית", kind="success", command=self._bb_add_manual_thread).pack(side="left", padx=4)
        theme.make_button(filt, "רענן", kind="secondary", command=self._bb_refresh_emails).pack(side="left", padx=4)

        paned = ttk.PanedWindow(page, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # ----- detail (left in LTR coordinates == visually left; list on right for RTL)
        detail = tk.Frame(paned, bg=theme.PAGE_BG)
        listing = tk.Frame(paned, bg=theme.PAGE_BG)
        paned.add(detail, weight=3)
        paned.add(listing, weight=2)

        cols = ("date", "category", "subject", "msgs", "files")
        headers = {"date": "תאריך", "category": "קטגוריה", "subject": "נושא", "msgs": "הודעות", "files": "קבצים"}
        widths = {"date": 85, "category": 120, "subject": 300, "msgs": 55, "files": 55}
        self._bb_email_tree = theme.make_treeview(listing, columns=cols, show="headings")
        for c in cols:
            self._bb_email_tree.heading(c, text=headers[c])
            self._bb_email_tree.column(c, width=widths[c], anchor="center" if c != "subject" else "e")
        vs = ttk.Scrollbar(listing, orient="vertical", command=self._bb_email_tree.yview)
        self._bb_email_tree.configure(yscrollcommand=vs.set)
        self._bb_email_tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self._bb_email_tree.bind("<<TreeviewSelect>>", self._bb_on_email_select)
        self._bb_email_tree.bind("<Double-1>", lambda e: self._bb_open_thread_gmail())

        self._bb_email_subject = tk.Label(detail, text="בחר התכתבות מהרשימה", font=theme.FONT_SUBTITLE, bg=theme.PAGE_BG, fg=theme.DARK, anchor="e", justify="right", wraplength=640)
        self._bb_email_subject.pack(fill="x", padx=6, pady=(4, 0))
        self._bb_email_summary = tk.Label(detail, text="", bg=theme.PAGE_BG, fg=theme.SUBTEXT, anchor="e", justify="right", wraplength=640, font=theme.FONT_SMALL)
        self._bb_email_summary.pack(fill="x", padx=6, pady=(0, 4))

        body_paned = ttk.PanedWindow(detail, orient="vertical")
        body_paned.pack(fill="both", expand=True, padx=4)

        msgs_box = ttk.LabelFrame(body_paned, text="הודעות", padding=4)
        body_paned.add(msgs_box, weight=3)
        self._bb_email_text = tk.Text(msgs_box, wrap="word", font=theme.FONT_BODY, bg=theme.CARD_BG, relief="flat", padx=8, pady=6)
        tvs = ttk.Scrollbar(msgs_box, orient="vertical", command=self._bb_email_text.yview)
        self._bb_email_text.configure(yscrollcommand=tvs.set, state="disabled")
        self._bb_email_text.pack(side="left", fill="both", expand=True)
        tvs.pack(side="right", fill="y")
        self._bb_email_text.tag_configure("hdr_arye", foreground=theme.PRIMARY_DARK, font=theme.FONT_BODY_BOLD, justify="right")
        self._bb_email_text.tag_configure("hdr_omer", foreground=theme.PURPLE, font=theme.FONT_BODY_BOLD, justify="right")
        self._bb_email_text.tag_configure("hdr_other", foreground=theme.TEAL, font=theme.FONT_BODY_BOLD, justify="right")
        self._bb_email_text.tag_configure("body", justify="right", lmargin1=6, lmargin2=6, rmargin=6)
        self._bb_email_text.tag_configure("att", foreground=theme.SUBTEXT, font=theme.FONT_SMALL, justify="right")
        self._bb_email_text.tag_configure("sep", foreground=theme.BORDER, justify="right")

        files_box = ttk.LabelFrame(body_paned, text="קבצים מצורפים", padding=4)
        body_paned.add(files_box, weight=2)
        fcols = ("filename", "from", "date", "status")
        fheaders = {"filename": "קובץ", "from": "מאת", "date": "תאריך", "status": "סטטוס"}
        fwidths = {"filename": 360, "from": 80, "date": 85, "status": 90}
        fwrap = tk.Frame(files_box)
        fwrap.pack(fill="both", expand=True)
        self._bb_email_files_tree = theme.make_treeview(fwrap, columns=fcols, show="headings", height=6)
        for c in fcols:
            self._bb_email_files_tree.heading(c, text=fheaders[c])
            self._bb_email_files_tree.column(c, width=fwidths[c], anchor="e" if c == "filename" else "center")
        fvs = ttk.Scrollbar(fwrap, orient="vertical", command=self._bb_email_files_tree.yview)
        self._bb_email_files_tree.configure(yscrollcommand=fvs.set)
        self._bb_email_files_tree.pack(side="left", fill="both", expand=True)
        fvs.pack(side="right", fill="y")
        self._bb_email_files_tree.bind("<Double-1>", lambda e: self._bb_open_selected_email_file())
        self._bb_email_file_rows = {}

        fbtns = tk.Frame(files_box, bg=theme.PAGE_BG)
        fbtns.pack(fill="x", pady=(4, 0))
        theme.make_button(fbtns, "פתח קובץ", kind="primary", command=self._bb_open_selected_email_file).pack(side="right", padx=3)
        theme.make_button(fbtns, "פתח תיקייה", kind="secondary", command=self._bb_open_thread_folder).pack(side="right", padx=3)
        theme.make_button(fbtns, "הוסף קובץ", kind="success", command=self._bb_add_email_file).pack(side="right", padx=3)
        theme.make_button(fbtns, "הסר קובץ ידני", kind="danger", command=self._bb_remove_email_file).pack(side="right", padx=3)
        theme.make_button(fbtns, "פתח ב-Gmail", kind="primary", command=self._bb_open_thread_gmail).pack(side="left", padx=3)

        notes_box = ttk.LabelFrame(body_paned, text="הערות פנימיות", padding=4)
        body_paned.add(notes_box, weight=1)
        self._bb_email_notes = tk.Text(notes_box, height=3, wrap="word", font=theme.FONT_BODY, bg=theme.CARD_BG, relief="flat", padx=6, pady=4)
        self._bb_email_notes.pack(side="left", fill="both", expand=True)
        theme.make_button(notes_box, "שמור הערה", kind="success", command=self._bb_save_email_notes).pack(side="right", padx=4, anchor="n")

    def _build_bb_docs_page(self, page):
        filt = tk.Frame(page, bg=theme.PAGE_BG)
        filt.pack(fill="x", pady=6)
        tk.Label(filt, text="קטגוריה:", bg=theme.PAGE_BG).pack(side="right", padx=(4, 2))
        self._bb_docs_cat_var = tk.StringVar(value=ALL)
        self._bb_docs_cat_combo = ttk.Combobox(filt, textvariable=self._bb_docs_cat_var, state="readonly", width=22, justify="right")
        self._bb_docs_cat_combo.pack(side="right", padx=4)
        self._bb_docs_cat_combo.bind("<<ComboboxSelected>>", lambda e: self._bb_refresh_docs())
        tk.Label(filt, text="חיפוש:", bg=theme.PAGE_BG).pack(side="right", padx=(10, 2))
        self._bb_docs_search_var = tk.StringVar()
        ent = ttk.Entry(filt, textvariable=self._bb_docs_search_var, width=30, justify="right")
        ent.pack(side="right", padx=4)
        ent.bind("<KeyRelease>", lambda e: self._bb_refresh_docs())
        self._bb_docs_only_local = tk.BooleanVar(value=False)
        ttk.Checkbutton(filt, text="רק קבצים שקיימים מקומית", variable=self._bb_docs_only_local, command=self._bb_refresh_docs).pack(side="right", padx=8)

        cols = ("date", "filename", "subject", "category", "from", "status")
        headers = {"date": "תאריך", "filename": "קובץ", "subject": "התכתבות", "category": "קטגוריה", "from": "מאת", "status": "סטטוס"}
        widths = {"date": 85, "filename": 320, "subject": 320, "category": 120, "from": 80, "status": 80}
        wrap = tk.Frame(page, bg=theme.PAGE_BG)
        wrap.pack(fill="both", expand=True)
        self._bb_docs_tree = theme.make_treeview(wrap, columns=cols, show="headings")
        for c in cols:
            self._bb_docs_tree.heading(c, text=headers[c])
            self._bb_docs_tree.column(c, width=widths[c], anchor="e" if c in ("filename", "subject") else "center")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self._bb_docs_tree.yview)
        self._bb_docs_tree.configure(yscrollcommand=vs.set)
        self._bb_docs_tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self._bb_docs_tree.bind("<Double-1>", lambda e: self._bb_open_selected_doc())
        self._bb_docs_rows = {}

        btns = tk.Frame(page, bg=theme.PAGE_BG)
        btns.pack(fill="x", pady=(4, 0))
        theme.make_button(btns, "פתח קובץ", kind="primary", command=self._bb_open_selected_doc).pack(side="right", padx=3)
        theme.make_button(btns, "פתח תיקייה", kind="secondary", command=self._bb_open_selected_doc_folder).pack(side="right", padx=3)
        theme.make_button(btns, "עבור להתכתבות", kind="secondary", command=self._bb_goto_doc_thread).pack(side="right", padx=3)
        theme.make_button(btns, "פתח ב-Gmail", kind="primary", command=self._bb_open_selected_doc_gmail).pack(side="left", padx=3)
        theme.make_button(btns, "פתח תיקיית כל המסמכים", kind="secondary", command=self._bb_open_docs_root).pack(side="left", padx=3)

    # ------------------------------------------------------------------ refresh
    def _bb_refresh_emails(self):
        self._bb_refresh_customer_card()
        cats = [ALL] + list(self.data_processor.get_baby_basic_email_categories())
        for combo, var in ((self._bb_email_cat_combo, self._bb_email_cat_var), (self._bb_docs_cat_combo, self._bb_docs_cat_var)):
            combo["values"] = cats
            if var.get() not in cats:
                var.set(ALL)
        self._bb_refresh_email_threads()
        self._bb_refresh_docs()

    def _bb_refresh_customer_card(self):
        card = self._bb_customer_card
        for ch in card.winfo_children():
            ch.destroy()
        cust = self.data_processor.get_baby_basic_customer()
        threads = self.data_processor.get_baby_basic_email_threads()
        docs = self.data_processor.iter_baby_basic_documents()
        local_docs = sum(1 for d in docs if d.get("status") == "ok")
        last_date = max((t.get("date") or "" for t in threads), default="")
        first_date = min((t.get("date") or "" for t in threads if t.get("date")), default="")

        left = tk.Frame(card, bg=theme.PAGE_BG)
        right = tk.Frame(card, bg=theme.PAGE_BG)
        right.pack(side="right", fill="x", expand=True)
        left.pack(side="left", fill="y")

        rows = [
            ("לקוח", f"{cust.get('name', 'בייבי בייסיק')} — {cust.get('contact', '')}"),
            ("אימייל", ", ".join([cust.get("email", "")] + list(cust.get("extra_emails") or []))),
            ("אתר", cust.get("website", "")),
            ("קשר", cust.get("relationship", "")),
            ("תוויות", cust.get("label_sizes", "")),
            ("מידות", cust.get("size_labels", "")),
            ("סבל", cust.get("porter_phone", "")),
            ("סליקה", cust.get("rivhit_terminal", "")),
        ]
        for i, (k, v) in enumerate(rows):
            if not v:
                continue
            r = tk.Frame(right, bg=theme.PAGE_BG)
            r.pack(fill="x")
            tk.Label(r, text=f"{k}:", font=theme.FONT_BODY_BOLD, bg=theme.PAGE_BG, fg=theme.DARK, width=8, anchor="e").pack(side="right")
            tk.Label(r, text=v, bg=theme.PAGE_BG, fg=theme.TEXT, anchor="e", justify="right", wraplength=820).pack(side="right", padx=(0, 6))

        stats = (
            f"התכתבויות: {len(threads)}\n"
            f"קבצים: {len(docs)} (מקומית: {local_docs})\n"
            f"תקופה: {_fmt_date(first_date)} – {_fmt_date(last_date)}"
        )
        tk.Label(left, text=stats, bg=theme.PANEL_BG, fg=theme.DARK, justify="right", anchor="e", padx=10, pady=6, font=theme.FONT_BODY).pack(fill="x", pady=(0, 4))
        email = cust.get("email") or ""
        theme.make_button(left, "כל ההתכתבות ב-Gmail", kind="primary",
                          command=lambda: webbrowser.open(f"https://mail.google.com/mail/u/0/#search/{email}" if email else GMAIL_BASE)).pack(fill="x", pady=1)
        theme.make_button(left, "פתח תיקיית מסמכים", kind="secondary", command=self._bb_open_docs_root).pack(fill="x", pady=1)

    def _bb_visible_threads(self):
        cat = self._bb_email_cat_var.get()
        q = (self._bb_email_search_var.get() or "").strip().lower()
        out = []
        for t in self.data_processor.get_baby_basic_email_threads():
            if cat != ALL and (t.get("category") or "") != cat:
                continue
            if q:
                hay = " ".join([
                    t.get("subject") or "", t.get("summary") or "", t.get("notes") or "",
                    " ".join(m.get("body") or "" for m in t.get("messages") or []),
                    " ".join(a.get("filename") or "" for m in t.get("messages") or [] for a in m.get("attachments") or []),
                ]).lower()
                if q not in hay:
                    continue
            out.append(t)
        return out

    def _bb_refresh_email_threads(self):
        tree = self._bb_email_tree
        for item in tree.get_children():
            tree.delete(item)
        for t in self._bb_visible_threads():
            msgs = t.get("messages") or []
            nfiles = sum(len(m.get("attachments") or []) for m in msgs) + len(t.get("extra_attachments") or [])
            tree.insert("", "end", iid=str(t.get("id")), values=(
                _fmt_date(t.get("date") or ""), t.get("category") or "", t.get("subject") or "", len(msgs), nfiles,
            ))
        theme.stripe_tree(tree)
        sel = self._bb_email_selected_thread
        if sel and tree.exists(sel):
            tree.selection_set(sel)
            tree.see(sel)
        else:
            self._bb_show_thread(None)

    def _bb_on_email_select(self, _e=None):
        sel = self._bb_email_tree.selection()
        if not sel:
            return
        self._bb_email_selected_thread = sel[0]
        self._bb_show_thread(self.data_processor.get_baby_basic_email_thread(sel[0]))

    def _bb_show_thread(self, t):
        txt = self._bb_email_text
        txt.configure(state="normal")
        txt.delete("1.0", "end")
        ftree = self._bb_email_files_tree
        for item in ftree.get_children():
            ftree.delete(item)
        self._bb_email_file_rows = {}
        self._bb_email_notes.delete("1.0", "end")
        if not t:
            self._bb_email_subject.config(text="בחר התכתבות מהרשימה")
            self._bb_email_summary.config(text="")
            txt.configure(state="disabled")
            return
        self._bb_email_subject.config(text=t.get("subject") or "")
        self._bb_email_summary.config(text=f"{_fmt_date(t.get('date') or '')}  ·  {t.get('category') or ''}\n{t.get('summary') or ''}")
        for m in t.get("messages") or []:
            who = _who(m.get("from"))
            tag = "hdr_arye" if who == "אריה" else ("hdr_omer" if who == "עומר" else "hdr_other")
            txt.insert("end", f"{who}  ·  {_fmt_date(m.get('date') or '')}\n", tag)
            body = (m.get("body") or "").strip()
            if body:
                txt.insert("end", body + "\n", "body")
            atts = m.get("attachments") or []
            if atts:
                txt.insert("end", "📎 " + " · ".join(a.get("filename") or "" for a in atts) + "\n", "att")
            txt.insert("end", "─" * 60 + "\n", "sep")
        txt.configure(state="disabled")

        idx = 0
        for m in t.get("messages") or []:
            for a in m.get("attachments") or []:
                iid = f"a{idx}"
                idx += 1
                status = "קיים" if a.get("status") == "ok" and os.path.exists(a.get("path") or "") else "רק ב-Gmail"
                ftree.insert("", "end", iid=iid, values=(a.get("filename") or "", _who(m.get("from")), _fmt_date(m.get("date") or ""), status))
                self._bb_email_file_rows[iid] = {"path": a.get("path") or "", "manual": False}
        for a in t.get("extra_attachments") or []:
            iid = f"a{idx}"
            idx += 1
            exists = os.path.exists(a.get("path") or "")
            ftree.insert("", "end", iid=iid, values=(a.get("label") or os.path.basename(a.get("path") or ""), "ידני", _fmt_date(a.get("added_at") or ""), "קיים" if exists else "חסר"))
            self._bb_email_file_rows[iid] = {"path": a.get("path") or "", "manual": True}
        theme.stripe_tree(ftree)
        self._bb_email_notes.insert("1.0", t.get("notes") or "")

    # ------------------------------------------------------------------ actions (mails)
    def _bb_current_thread(self):
        tid = self._bb_email_selected_thread
        if not tid:
            messagebox.showinfo("בייבי בייסיק", "יש לבחור התכתבות")
            return None
        return self.data_processor.get_baby_basic_email_thread(tid)

    def _bb_selected_email_file(self):
        sel = self._bb_email_files_tree.selection()
        if not sel:
            messagebox.showinfo("קבצים", "יש לבחור קובץ")
            return None
        return self._bb_email_file_rows.get(sel[0])

    def _bb_open_selected_email_file(self):
        row = self._bb_selected_email_file()
        if not row:
            return
        try:
            _open_path(row.get("path") or "")
        except Exception as e:
            if messagebox.askyesno("קובץ לא זמין מקומית", f"{e}\n\nלפתוח את ההתכתבות ב-Gmail?"):
                self._bb_open_thread_gmail()

    def _bb_open_thread_folder(self):
        t = self._bb_current_thread()
        if not t:
            return
        folder = os.path.abspath(self.data_processor.baby_basic_thread_documents_dir(t.get("id")))
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

    def _bb_open_docs_root(self):
        folder = os.path.abspath(self.data_processor.baby_basic_documents_root)
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

    def _bb_open_thread_gmail(self):
        t = self._bb_current_thread()
        if not t:
            return
        url = t.get("gmail_url") or ""
        if not url and t.get("gmail_ids"):
            url = GMAIL_BASE + t["gmail_ids"][0]
        if not url:
            messagebox.showinfo("Gmail", "לרשומה ידנית אין קישור Gmail")
            return
        webbrowser.open(url)

    def _bb_add_email_file(self):
        t = self._bb_current_thread()
        if not t:
            return
        paths = filedialog.askopenfilenames(
            title="צירוף קובץ להתכתבות",
            filetypes=[("מסמכים", "*.pdf *.xlsx *.xls *.xlsm *.numbers *.txt *.docx *.zip"), ("תמונות", "*.png *.jpg *.jpeg *.webp"), ("כל הקבצים", "*.*")],
        )
        if not paths:
            return
        errors = []
        for p in paths:
            try:
                self.data_processor.add_baby_basic_attachment(t.get("id"), p)
            except Exception as e:
                errors.append(f"{os.path.basename(p)}: {e}")
        self._bb_refresh_emails()
        if errors:
            messagebox.showerror("שגיאה", "\n".join(errors))

    def _bb_remove_email_file(self):
        t = self._bb_current_thread()
        if not t:
            return
        row = self._bb_selected_email_file()
        if not row:
            return
        if not row.get("manual"):
            messagebox.showinfo("קבצים", "ניתן להסיר רק קבצים שצורפו ידנית. קבצים מ-Gmail נשארים כתיעוד.")
            return
        if not messagebox.askyesno("מחיקה", "להסיר את הקובץ מההתכתבות?"):
            return
        if self.data_processor.delete_baby_basic_attachment(t.get("id"), row.get("path") or ""):
            self._bb_refresh_emails()

    def _bb_save_email_notes(self):
        t = self._bb_current_thread()
        if not t:
            return
        notes = self._bb_email_notes.get("1.0", "end").strip()
        if self.data_processor.set_baby_basic_thread_notes(t.get("id"), notes):
            self._bb_refresh_email_threads()

    def _bb_add_manual_thread(self):
        subject = simpledialog.askstring("רשומה ידנית", "נושא:", parent=self._bb_email_tree.winfo_toplevel())
        if not subject:
            return
        cats = self.data_processor.get_baby_basic_email_categories()
        cat = simpledialog.askstring("רשומה ידנית", "קטגוריה:\n" + " / ".join(cats), initialvalue=cats[-1] if cats else "אחר", parent=self._bb_email_tree.winfo_toplevel()) or "אחר"
        date = simpledialog.askstring("רשומה ידנית", "תאריך (YYYY-MM-DD):", initialvalue=datetime.now().strftime("%Y-%m-%d"), parent=self._bb_email_tree.winfo_toplevel()) or ""
        body = simpledialog.askstring("רשומה ידנית", "תוכן / סיכום:", parent=self._bb_email_tree.winfo_toplevel()) or ""
        rec = self.data_processor.add_baby_basic_manual_thread(subject, cat, date, body)
        self._bb_email_selected_thread = rec.get("id")
        self._bb_refresh_emails()

    # ------------------------------------------------------------------ docs page
    def _bb_refresh_docs(self):
        tree = self._bb_docs_tree
        for item in tree.get_children():
            tree.delete(item)
        self._bb_docs_rows = {}
        cat = self._bb_docs_cat_var.get()
        q = (self._bb_docs_search_var.get() or "").strip().lower()
        only_local = self._bb_docs_only_local.get()
        for i, d in enumerate(self.data_processor.iter_baby_basic_documents()):
            if cat != ALL and d.get("category") != cat:
                continue
            if only_local and d.get("status") != "ok":
                continue
            if q and q not in f"{d.get('filename', '')} {d.get('subject', '')}".lower():
                continue
            iid = f"d{i}"
            status = "קיים" if d.get("status") == "ok" and os.path.exists(d.get("path") or "") else ("חסר" if d.get("manual") else "רק ב-Gmail")
            tree.insert("", "end", iid=iid, values=(
                _fmt_date(d.get("date") or ""), d.get("filename") or "", d.get("subject") or "", d.get("category") or "", _who(d.get("from")), status,
            ))
            self._bb_docs_rows[iid] = d
        theme.stripe_tree(tree)

    def _bb_selected_doc(self):
        sel = self._bb_docs_tree.selection()
        if not sel:
            messagebox.showinfo("מסמכים", "יש לבחור קובץ")
            return None
        return self._bb_docs_rows.get(sel[0])

    def _bb_open_selected_doc(self):
        d = self._bb_selected_doc()
        if not d:
            return
        try:
            _open_path(d.get("path") or "")
        except Exception as e:
            if messagebox.askyesno("קובץ לא זמין מקומית", f"{e}\n\nלפתוח את ההתכתבות ב-Gmail?"):
                self._bb_open_selected_doc_gmail()

    def _bb_open_selected_doc_folder(self):
        d = self._bb_selected_doc()
        if not d:
            return
        p = d.get("path") or ""
        folder = os.path.dirname(os.path.abspath(p)) if p else os.path.abspath(self.data_processor.baby_basic_thread_documents_dir(d.get("thread_id")))
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

    def _bb_open_selected_doc_gmail(self):
        d = self._bb_selected_doc()
        if not d:
            return
        t = self.data_processor.get_baby_basic_email_thread(d.get("thread_id"))
        url = (t or {}).get("gmail_url") or ""
        if not url:
            messagebox.showinfo("Gmail", "אין קישור Gmail לקובץ זה")
            return
        webbrowser.open(url)

    def _bb_goto_doc_thread(self):
        d = self._bb_selected_doc()
        if not d:
            return
        tid = str(d.get("thread_id"))
        self._bb_email_cat_var.set(ALL)
        self._bb_email_search_var.set("")
        self._bb_email_selected_thread = tid
        self._bb_refresh_email_threads()
        try:
            self._bb_emails_nb.select(0)
        except Exception:
            pass
