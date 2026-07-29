"""מודול עיצוב מרכזי - מקור אמת יחיד לצבעים, פונטים וסגנונות של התוכנה.

כל קבצי ה-GUI צריכים להשתמש בקבועים ובפונקציות מכאן במקום צבעים/פונטים קשיחים.
"""
import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont

# ---------- פלטת צבעים (בהשראת Tailwind - נראה כמו אתר מודרני) ----------
PAGE_BG = "#f4f6fa"      # רקע דף כללי
CARD_BG = "#ffffff"      # רקע כרטיס/פאנל
PANEL_BG = "#e8eef7"     # פאנל משני
BORDER = "#e2e8f0"       # קווי הפרדה

TEXT = "#1e293b"         # טקסט ראשי כהה
SUBTEXT = "#64748b"      # טקסט משני אפור
MUTED = "#94a3b8"        # טקסט חלש

DARK = "#1e293b"         # רקע כהה (header / status bar)
DARK_2 = "#334155"       # כהה משני

PRIMARY = "#3b82f6"      # כחול ראשי
PRIMARY_DARK = "#2563eb"
SUCCESS = "#059669"      # ירוק פעולה
SUCCESS_LIGHT = "#22c55e"
DANGER = "#ef4444"       # אדום מחיקה
DANGER_DARK = "#dc2626"
WARNING = "#ea580c"      # כתום אזהרה
AMBER = "#f59e0b"
PURPLE = "#7c3aed"
PURPLE_LIGHT = "#8b5cf6"
TEAL = "#0d9488"

FONT_FAMILY = "Segoe UI"

FONT_TITLE = (FONT_FAMILY, 16, "bold")
FONT_SUBTITLE = (FONT_FAMILY, 12, "bold")
FONT_BODY = (FONT_FAMILY, 10)
FONT_BODY_BOLD = (FONT_FAMILY, 10, "bold")
FONT_SMALL = (FONT_FAMILY, 9)

# מיפוי סוגי כפתורים -> (רקע, רקע בלחיצה)
_BUTTON_KINDS = {
    "primary": (PRIMARY, PRIMARY_DARK),
    "success": (SUCCESS, "#047857"),
    "danger": (DANGER, DANGER_DARK),
    "warning": (WARNING, "#c2410c"),
    "purple": (PURPLE, "#6d28d9"),
    "dark": (DARK_2, DARK),
    "secondary": (SUBTEXT, "#475569"),
}


def apply_theme(root):
    """החלת העיצוב הגלובלי - נקרא פעם אחת מ-main.py אחרי יצירת החלון."""
    # ביטול הצביעה האוטומטית של ttkbootstrap על ווידג'טים קלאסיים (tk.*):
    # אחרת הוא דורס צבעי bg/fg מפורשים שהוגדרו בקוד.
    try:
        from ttkbootstrap.style import Bootstyle
        Bootstyle.update_tk_widget_style = staticmethod(lambda *a, **k: None)
    except Exception:
        pass

    # פלטת בסיס לכל הווידג'טים הקלאסיים
    try:
        root.tk_setPalette(background=PAGE_BG, foreground=TEXT)
    except Exception:
        pass

    # פונטים מובנים של Tk (משפיע על כל ווידג'ט שלא הוגדר לו פונט מפורש)
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkTooltipFont"):
        try:
            tkfont.nametofont(name).configure(family=FONT_FAMILY, size=10)
        except Exception:
            pass

    # ברירות מחדל לווידג'טים קלאסיים (לא דורס ערכים שהוגדרו מפורשות בקוד)
    try:
        root.option_add("*Button.relief", "flat")
        root.option_add("*Button.borderWidth", 0)
        root.option_add("*Button.cursor", "hand2")
        root.option_add("*Button.padX", 10)
        root.option_add("*Button.padY", 4)
        # כפתור ללא צבע מפורש - אפור בהיר (כמו כפתור משני באתר)
        root.option_add("*Button.background", "#dbe2ea")
        root.option_add("*Button.foreground", TEXT)
        root.option_add("*Button.activeBackground", "#c6d0db")
        root.option_add("*Button.activeForeground", TEXT)
        root.option_add("*Button.disabledForeground", "#8fa0b3")
        root.option_add("*Label.background", PAGE_BG)
        root.option_add("*Frame.background", PAGE_BG)
        # שדות קלט לבנים וקריאים
        root.option_add("*Entry.background", "#ffffff")
        root.option_add("*Entry.relief", "solid")
        root.option_add("*Entry.borderWidth", 1)
        root.option_add("*Entry.highlightThickness", 1)
        root.option_add("*Entry.highlightColor", PRIMARY)
        root.option_add("*Entry.highlightBackground", BORDER)
        root.option_add("*Text.background", "#ffffff")
        root.option_add("*Text.relief", "solid")
        root.option_add("*Text.borderWidth", 1)
        root.option_add("*Listbox.background", "#ffffff")
        root.option_add("*Listbox.relief", "solid")
        root.option_add("*Listbox.borderWidth", 1)
        root.option_add("*Spinbox.background", "#ffffff")
    except Exception:
        pass

    # אפקט ריחוף (hover) לכפתורים קלאסיים - הבהרה/הכהיה קלה של הרקע
    _install_button_hover(root)

    # סגנונות ttk (מעל ערכת ttkbootstrap)
    try:
        style = ttk.Style()

        # Notebook - טאבים קומפקטיים כדי שלא ייחתכו במסכים צרים
        style.configure("TNotebook", background=PAGE_BG, borderwidth=0, tabmargins=(4, 4, 4, 0))
        style.configure(
            "TNotebook.Tab",
            font=(FONT_FAMILY, 10),
            padding=(12, 6),
        )

        # שמירת פריסת הטאב המקורית תחת סגנון 'Inner' (ל-notebooks פנימיים),
        # והסתרת פס הטאבים בסגנון הבסיס - כך שה-notebook הראשי (שמנוהל ע"י
        # סרגל ניווט דינמי) לא מצייר טאבים בכלל.
        # הערה: הסתרה עובדת רק על סגנון הבסיס; סגנון נגזר עם layout ריק נופל
        # חזרה לבסיס, ולכן ההיפוך הזה הכרחי.
        _orig_tab_layout = style.layout("TNotebook.Tab")
        style.layout("Inner.TNotebook.Tab", _orig_tab_layout)
        style.configure("Inner.TNotebook", background=PAGE_BG, borderwidth=0, tabmargins=(4, 4, 4, 0))
        style.layout("TNotebook.Tab", [])

        # כל notebook שנוצר בלי סגנון מפורש יקבל את הסגנון הפנימי (עם טאבים גלויים)
        _patch_inner_notebooks()

        # Treeview - טבלאות בסגנון אתר: שורות גבוהות, כותרות כהות
        style.configure(
            "Treeview",
            font=(FONT_FAMILY, 9),
            rowheight=28,
            background=CARD_BG,
            fieldbackground=CARD_BG,
            foreground=TEXT,
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            font=(FONT_FAMILY, 9, "bold"),
            background=DARK,
            foreground="#ffffff",
            relief="flat",
            padding=(6, 6),
        )
        style.map(
            "Treeview.Heading",
            background=[("active", DARK_2), ("pressed", DARK_2)],
        )
        style.map(
            "Treeview",
            background=[("selected", PRIMARY)],
            foreground=[("selected", "#ffffff")],
        )

        # שדות קלט
        style.configure("TEntry", padding=4)
        style.configure("TCombobox", padding=4)
        style.configure("TLabelframe", background=PAGE_BG)
        style.configure("TLabelframe.Label", background=PAGE_BG, foreground=SUBTEXT, font=(FONT_FAMILY, 10, "bold"))
        style.configure("TFrame", background=PAGE_BG)
        style.configure("TLabel", background=PAGE_BG, foreground=TEXT)
        style.configure("TButton", font=(FONT_FAMILY, 10))
    except Exception:
        pass


_inner_nb_patched = False


def _patch_inner_notebooks():
    """ברירת מחדל של style='Inner.TNotebook' לכל ttk.Notebook שנוצר בלי סגנון.

    ה-notebook הראשי מועבר עם style='Main.TNotebook' מפורש ולכן נשאר בלי
    פס טאבים (נופל לסגנון הבסיס המוסתר).
    """
    global _inner_nb_patched
    if _inner_nb_patched:
        return
    _inner_nb_patched = True
    orig_init = ttk.Notebook.__init__

    def patched_init(self, master=None, **kw):
        if "style" not in kw:
            kw["style"] = "Inner.TNotebook"
        orig_init(self, master, **kw)

    ttk.Notebook.__init__ = patched_init


def _install_button_hover(root):
    """אפקט hover גלובלי לכל tk.Button - כמו כפתור באתר."""
    def _shift(widget, factor):
        try:
            r, g, b = widget.winfo_rgb(widget.cget("background"))
            r, g, b = r // 256, g // 256, b // 256
            if factor > 0:  # הבהרה
                r = min(255, int(r + (255 - r) * factor))
                g = min(255, int(g + (255 - g) * factor))
                b = min(255, int(b + (255 - b) * factor))
            else:  # הכהיה
                r = max(0, int(r * (1 + factor)))
                g = max(0, int(g * (1 + factor)))
                b = max(0, int(b * (1 + factor)))
            return f"#{r:02x}{g:02x}{b:02x}"
        except Exception:
            return None

    def on_enter(event):
        w = event.widget
        try:
            if w.cget("state") == "disabled":
                return
            if not hasattr(w, "_theme_orig_bg"):
                w._theme_orig_bg = w.cget("background")
            hover = _shift(w, -0.12)
            if hover:
                w.configure(background=hover)
        except Exception:
            pass

    def on_leave(event):
        w = event.widget
        try:
            if hasattr(w, "_theme_orig_bg"):
                w.configure(background=w._theme_orig_bg)
                del w._theme_orig_bg
        except Exception:
            pass

    try:
        root.bind_class("Button", "<Enter>", on_enter, add="+")
        root.bind_class("Button", "<Leave>", on_leave, add="+")
    except Exception:
        pass


# ---------- פונקציות עזר ליצירת ווידג'טים מעוצבים ----------

def make_page(parent, **kw):
    """מסגרת דף עם רקע אחיד."""
    kw.setdefault("bg", PAGE_BG)
    return tk.Frame(parent, **kw)


def make_card(parent, **kw):
    """מסגרת 'כרטיס' לבנה עם מסגרת עדינה - כמו card באתר."""
    kw.setdefault("bg", CARD_BG)
    kw.setdefault("highlightbackground", BORDER)
    kw.setdefault("highlightthickness", 1)
    kw.setdefault("bd", 0)
    return tk.Frame(parent, **kw)


def make_title(parent, text, bg=PAGE_BG, **kw):
    """כותרת דף ראשית."""
    kw.setdefault("font", FONT_TITLE)
    kw.setdefault("fg", TEXT)
    return tk.Label(parent, text=text, bg=bg, **kw)


def make_subtitle(parent, text, bg=PAGE_BG, **kw):
    kw.setdefault("font", FONT_SUBTITLE)
    kw.setdefault("fg", SUBTEXT)
    return tk.Label(parent, text=text, bg=bg, **kw)


def make_button(parent, text, kind="primary", command=None, **kw):
    """כפתור מודרני שטוח עם צבע לפי סוג הפעולה."""
    bg, active = _BUTTON_KINDS.get(kind, _BUTTON_KINDS["primary"])
    kw.setdefault("font", FONT_BODY_BOLD)
    kw.setdefault("bg", bg)
    kw.setdefault("fg", "#ffffff")
    kw.setdefault("activebackground", active)
    kw.setdefault("activeforeground", "#ffffff")
    kw.setdefault("relief", "flat")
    kw.setdefault("bd", 0)
    kw.setdefault("cursor", "hand2")
    kw.setdefault("padx", 14)
    kw.setdefault("pady", 6)
    return tk.Button(parent, text=text, command=command, **kw)


def stripe_tree(tree, even_bg="#f8fafc", odd_bg=CARD_BG):
    """הוספת פסי זברה ל-Treeview קיים (לקרוא אחרי מילוי שורות)."""
    try:
        tree.tag_configure("theme_even", background=even_bg)
        tree.tag_configure("theme_odd", background=odd_bg)
        for i, iid in enumerate(tree.get_children("")):
            tags = [t for t in tree.item(iid, "tags") if t not in ("theme_even", "theme_odd")]
            tags.append("theme_even" if i % 2 == 0 else "theme_odd")
            tree.item(iid, tags=tuple(tags))
    except Exception:
        pass
