from __future__ import annotations

"""
Gemeinsame UI-Hilfsfunktionen für Treeviews.

Bietet eine generische `make_tree`-Funktion als Ersatz für die bisherige
`App._make_tree`-Methode, damit Views ohne enge Kopplung an die App-Klasse
Treeviews erstellen können.
"""

from tkinter import ttk
import tkinter.font as tkfont

from app.constants import MDConstants


def make_tree(parent: ttk.Frame, cols: list[str], bind_sort, height: int = None) -> ttk.Treeview:
    """Erstellt einen Treeview mit Scrollbar und optionaler Sort-Bindung.

    Args:
        parent: Container-Frame
        cols: Spaltenüberschriften
        bind_sort: Callable(tree: ttk.Treeview, numeric_like: set|None) -> None
                   Wird verwendet, um die Sortierlogik anzubinden.
        height: Optional - Anzahl sichtbarer Zeilen (Standard: None = unbegrenzt)

    Returns:
        Konfigurierter `ttk.Treeview`.
    """
    frame = ttk.Frame(parent)
    frame.pack(fill="both", expand=True)

    tree_kwargs = {"columns": cols, "show": "headings"}
    if height is not None:
        tree_kwargs["height"] = height
    
    tree = ttk.Treeview(frame, **tree_kwargs)
    for c in cols:
        tree.heading(c, text=c)
        tree.column(c, width=180 if c != "Betreff" else 300, anchor="w")
    tree.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side="right", fill="y")

    # Standard: alle Spalten sortierbar, Logik kommt von der App
    if bind_sort:
        bind_sort(tree)

    return tree


def bind_treeview_sort(tree: ttk.Treeview, numeric_like=None) -> None:
    """Bindet klickbare Sortierung an alle Spalten eines Treeviews.

    numeric_like: optionale Menge von Spaltennamen, die bevorzugt numerisch
    sortiert werden sollen.
    """
    numeric_like = numeric_like or set()

    def _key_for(col: str, value: str):
        s = (value or "").strip()
        if col in numeric_like:
            try:
                if s.endswith('.0') and s.replace('.', '', 1).isdigit():
                    return int(float(s))
                return int(s)
            except Exception:
                try:
                    return float(s)
                except Exception:
                    return s.lower()
        try:
            if s.endswith('.0') and s.replace('.', '', 1).isdigit():
                return int(float(s))
            return int(s)
        except Exception:
            try:
                return float(s)
            except Exception:
                return s.lower()

    def _sort(col: str, descending: bool):
        data = [(tree.set(k, col), k) for k in tree.get_children("")]
        data.sort(key=lambda t: _key_for(col, t[0]), reverse=descending)
        for idx, (_, k) in enumerate(data):
            tree.move(k, "", idx)
        tree.heading(col, command=lambda _c=col: _sort(_c, not descending))

    for c in list(tree["columns"] or []):
        tree.heading(c, command=lambda _c=c: _sort(_c, False))


def autosize_tree_columns(
    tree: ttk.Treeview,
    min_width: int = MDConstants.TREE_MIN_WIDTH,
    max_width: int = MDConstants.TREE_MAX_WIDTH,
    padding: int = MDConstants.TREE_PADDING,
) -> None:
    """Passt Spaltenbreiten eines Treeviews an Inhalte und Header an."""
    try:
        font = tkfont.nametofont(tree.cget("font")) if tree.cget("font") else tkfont.nametofont("TkDefaultFont")
    except Exception:
        font = tkfont.nametofont("TkDefaultFont")

    columns = list(tree.cget("columns") or [])
    items = tree.get_children("")
    for col in columns:
        header_text = tree.heading(col).get("text", "")
        width = font.measure(header_text)

        try:
            col_index = columns.index(col)
        except ValueError:
            col_index = None

        if col_index is not None:
            for iid in items:
                values = tree.item(iid, "values")
                if col_index < len(values):
                    cell_text = str(values[col_index])
                    width = max(width, font.measure(cell_text))

        col_max = max_width
        if col in ("Datei", "Ziel", "Zielordner", "Betreff"):
            col_max = max(max_width, 520)

        final_width = max(min_width, min(width + padding, col_max))
        tree.column(col, width=final_width, stretch=True, anchor="w")


