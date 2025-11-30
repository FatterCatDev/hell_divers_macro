from __future__ import annotations

"""Reusable dialogs for the Tkinter UI."""

import tkinter as tk

from hell_divers_macro.ui.theme import (
    BG,
    BUTTON_ACTIVE,
    BUTTON_BG,
    FG,
    apply_dark_theme,
    place_window_near,
    set_dark_titlebar,
)


class MacroSelectionDialog:
    """Listbox + tabs for selecting a MacroTemplate."""

    def __init__(self, parent: tk.Tk, title: str, templates) -> None:
        self.parent = parent
        self.templates = templates
        self.result = None
        self._current_selection = None

        categories: dict[str, list] = {}
        for tpl in self.templates:
            categories.setdefault(tpl.category, []).append(tpl)
        ordered_categories = list(categories.keys())
        visible: list = []

        self.top = tk.Toplevel(parent, bg=BG)
        self.top.title(title)
        self.top.transient(parent)

        tk.Label(self.top, text="Choose a macro template:").pack(anchor="w", pady=(8, 4), padx=10)

        search_var = tk.StringVar()
        search_frame = tk.Frame(self.top, bg=BG)
        search_frame.pack(fill=tk.X, padx=10, pady=(0, 8))
        tk.Label(search_frame, text="Search", bg=BG, fg=FG).pack(side=tk.LEFT, padx=(0, 6))
        search_entry = tk.Entry(search_frame, textvariable=search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tabs_frame = tk.Frame(self.top, bg=BG)
        tabs_frame.pack(fill=tk.X, padx=10, pady=(0, 6))

        list_frame = tk.Frame(self.top, bg=BG)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))

        self.listbox = tk.Listbox(list_frame, height=12)
        list_scroll = tk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=list_scroll.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        current_cat = {"val": ordered_categories[0] if ordered_categories else ""}
        visible = list(self.templates)

        def populate_list(cat: str | None, query: str) -> None:
            self.listbox.delete(0, tk.END)
            q = query.strip().lower()
            nonlocal visible
            if q:
                visible = [tpl for tpl in self.templates if q in tpl.name.lower()]
            elif cat:
                visible = list(categories.get(cat, []))
            else:
                visible = list(self.templates)
            for tpl in visible:
                self.listbox.insert(tk.END, tpl.name)
            self._current_selection = None

        def switch_cat(cat: str) -> None:
            current_cat["val"] = cat
            for btn in tab_buttons.values():
                btn.config(relief=tk.RAISED)
            tab_buttons[cat].config(relief=tk.SUNKEN)
            populate_list(cat, search_var.get())
            layout_tabs()

        tab_buttons_list: list[tuple[str, tk.Button]] = []
        tab_buttons: dict[str, tk.Button] = {}
        for cat in ordered_categories:
            btn = tk.Button(tabs_frame, text=cat, command=lambda c=cat: switch_cat(c))
            tab_buttons[cat] = btn
            tab_buttons_list.append((cat, btn))

        def layout_tabs(event=None) -> None:  # noqa: ANN001
            tabs_frame.update_idletasks()
            available = tabs_frame.winfo_width()
            if available <= 1:
                available = self.top.winfo_width() - 20
            x = 0
            row = 0
            col = 0
            for _, btn in tab_buttons_list:
                w = btn.winfo_reqwidth() + 6
                if col > 0 and x + w > available:
                    row += 1
                    col = 0
                    x = 0
                btn.grid(row=row, column=col, padx=(0, 6), pady=2, sticky="w")
                col += 1
                x += w

        tabs_frame.bind("<Configure>", layout_tabs)

        def handle_select(event=None) -> None:  # noqa: ANN001
            indices = self.listbox.curselection()
            if not indices:
                return
            if not visible or indices[0] >= len(visible):
                return
            tpl = visible[indices[0]]
            self._current_selection = tpl

        self.listbox.bind("<<ListboxSelect>>", handle_select)
        self.listbox.bind("<Double-Button-1>", lambda _: (handle_select(), self.ok()))

        def on_search(*args: str) -> None:  # noqa: ANN001
            populate_list(current_cat["val"], search_var.get())

        search_var.trace_add("write", on_search)

        btn_frame = tk.Frame(self.top, bg=BG)
        btn_frame.pack(fill=tk.X, padx=10, pady=(4, 10))
        tk.Button(btn_frame, text="OK", command=self.ok).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT)

        self.top.protocol("WM_DELETE_WINDOW", self.cancel)
        self.top.grab_set()
        place_window_near(self.top, parent)
        self.top.focus_set()
        apply_dark_theme(self.top)
        set_dark_titlebar(self.top)

        if ordered_categories:
            switch_cat(ordered_categories[0])

        self.top.update_idletasks()
        parent_width = parent.winfo_width()
        if parent_width > 0:
            reqw = self.top.winfo_reqwidth()
            reqh = self.top.winfo_reqheight()
            if reqw > parent_width:
                self.top.geometry(f"{parent_width}x{reqh}")
        parent.wait_window(self.top)

    def ok(self) -> None:
        self.result = self._current_selection
        self.top.destroy()

    def cancel(self) -> None:
        self.result = None
        self.top.destroy()


class TextEntryDialog:
    """Simple text input dialog."""

    def __init__(self, parent: tk.Tk, title: str, prompt: str, initial: str = "") -> None:
        self.result: str | None = None
        self.top = tk.Toplevel(parent, bg=BG)
        self.top.title(title)
        self.top.transient(parent)

        tk.Label(self.top, text=prompt).pack(anchor="w", padx=10, pady=(10, 4))
        self.entry_var = tk.StringVar(value=initial)
        entry = tk.Entry(self.top, textvariable=self.entry_var)
        entry.pack(fill=tk.X, padx=10)

        btn_frame = tk.Frame(self.top, bg=BG)
        btn_frame.pack(fill=tk.X, padx=10, pady=(10, 10))
        tk.Button(btn_frame, text="OK", command=self.ok).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT)

        self.top.protocol("WM_DELETE_WINDOW", self.cancel)
        self.top.bind("<Return>", lambda _: self.ok())
        self.top.bind("<Escape>", lambda _: self.cancel())
        self.top.grab_set()
        place_window_near(self.top, parent)
        entry.focus_set()
        apply_dark_theme(self.top)
        set_dark_titlebar(self.top)
        parent.wait_window(self.top)

    def ok(self) -> None:
        self.result = self.entry_var.get()
        self.top.destroy()

    def cancel(self) -> None:
        self.result = None
        self.top.destroy()
