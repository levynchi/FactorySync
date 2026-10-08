"""Combobox that completes the first matching prefix while typing.

The typed prefix stays in the entry; the completed remainder is shown as a
blue "ghost" label to the left of the typed text (correct for Hebrew RTL),
and committed into the entry on Enter / Tab / focus-out.
Splitting the entry text itself into selected / unselected chunks would make
Tk draw the Hebrew runs left-to-right, so no selection is used.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont


_NAV_KEYS = {
    'Left', 'Right', 'Up', 'Down', 'Home', 'End', 'Prior', 'Next',
    'Tab', 'Return', 'KP_Enter', 'Escape',
    'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R', 'Caps_Lock',
}

GHOST_BG = '#bfdbfe'   # כחול בהיר — החלק שהושלם אוטומטית
GHOST_FG = '#1e3a8a'


class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, master=None, on_accept=None, on_complete=None, **kwargs):
        kwargs.setdefault('exportselection', False)
        kwargs.setdefault('justify', 'right')
        super().__init__(master, **kwargs)
        self._completion_list = []
        self._on_accept = on_accept
        self._on_complete = on_complete  # fired whenever a full list value is identified
        self._match = None  # full value suggested for the current typed text
        try:
            font = self.cget('font') or 'TkTextFont'
            self._font = tkfont.nametofont(font) if isinstance(font, str) and font in tkfont.names() else font
        except Exception:
            self._font = 'TkTextFont'
        self._ghost = tk.Label(
            self,
            text='',
            bg=GHOST_BG,
            fg=GHOST_FG,
            font=self._font,
            padx=2,
            pady=0,
            bd=0,
            anchor='e',
        )
        self.bind('<KeyRelease>', self._on_keyrelease)
        self.bind('<Return>', self._accept)
        self.bind('<KP_Enter>', self._accept)
        self.bind('<Tab>', self._accept)
        self.bind('<Escape>', self._dismiss)
        self.bind('<FocusOut>', self._commit_pending)
        self.bind('<<ComboboxSelected>>', self._clear_ghost)
        self.bind('<Configure>', lambda e: self._reposition_ghost())

    # ---------- public ----------
    def set_on_accept(self, callback):
        self._on_accept = callback

    def set_on_complete(self, callback):
        self._on_complete = callback

    def set_completion_list(self, values):
        seen = set()
        names = []
        for value in values or []:
            name = str(value or '').strip()
            if not name or name in seen:
                continue
            seen.add(name)
            names.append(name)
        self._completion_list = names
        self['values'] = names

    def get_value(self) -> str:
        """הערך האפקטיבי: ההשלמה המוצעת אם קיימת, אחרת הטקסט שהוקלד."""
        return self._match if self._match else self.get()

    def commit(self):
        """מכניס את ההשלמה לשדה (אם קיימת) ומסתיר את הרוח."""
        if self._match:
            match = self._match
            self._match = None
            self.delete(0, 'end')
            self.insert(0, match)
            self.icursor('end')
        self._hide_ghost()

    # ---------- internals ----------
    def _fire_complete(self):
        if self._on_complete is None:
            return
        try:
            self._on_complete(self)
        except Exception:
            pass

    def _find_match(self, prefix: str):
        needle = (prefix or '').lower()
        if not needle:
            return None
        return next((name for name in self._completion_list if name.lower().startswith(needle)), None)

    def _filter_values(self, prefix: str):
        prefix = (prefix or '').strip()
        if not prefix:
            self['values'] = self._completion_list
            return
        needle = prefix.lower()
        hits = [name for name in self._completion_list if name.lower().startswith(needle)]
        self['values'] = hits or self._completion_list

    def _on_keyrelease(self, event=None):
        keysym = getattr(event, 'keysym', '') if event is not None else ''
        if keysym in _NAV_KEYS:
            return
        typed = self.get()
        self._filter_values(typed)
        match = self._find_match(typed)
        if match and match != typed:
            self._match = match
            self._show_ghost(typed, match)
            self._fire_complete()
        else:
            self._match = None
            self._hide_ghost()
            if match:
                self._fire_complete()

    def _show_ghost(self, typed: str, match: str):
        remainder = match[len(typed):]
        if not remainder:
            self._hide_ghost()
            return
        self._ghost.configure(text=remainder)
        self._reposition_ghost()

    def _reposition_ghost(self):
        if not self._match or not self._ghost.cget('text'):
            return
        try:
            self.update_idletasks()
            box = self.bbox(0)
        except Exception:
            box = None
        if not box:
            self._ghost.place_forget()
            return
        x, y, w, h = box
        # הטקסט מיושר לימין; הרוח יושבת מיד משמאל לאות הראשונה (התחלת הטקסט הלוגי)
        self._ghost.place(x=max(0, x - 1), y=y, height=h, anchor='ne')
        self._ghost.lift()

    def _hide_ghost(self):
        try:
            self._ghost.place_forget()
        except Exception:
            pass

    def _clear_ghost(self, event=None):
        self._match = None
        self._hide_ghost()

    def _dismiss(self, event=None):
        self._clear_ghost()
        return 'break'

    def _commit_pending(self, event=None):
        self.commit()

    def _accept(self, event=None):
        self.commit()
        try:
            self.selection_clear()
        except Exception:
            pass
        try:
            self.event_generate('<<ComboboxSelected>>')
        except Exception:
            pass
        if self._on_accept is not None:
            try:
                self._on_accept(self)
            except Exception:
                pass
        return 'break'
