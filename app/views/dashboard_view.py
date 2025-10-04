from __future__ import annotations

"""
View-Modul für den Tab "MD-Dashboard".

Minimal-invasive Extraktion des UI-Aufbaus. Verwendet weiterhin Callbacks und
Hilfsmethoden der App-Instanz (z. B. `on_refresh_dashboard`).
"""

from tkinter import ttk

from constants import MDConstants, ProcStatus, DashTag
from utils import create_info_button


def build_dashboard(parent: ttk.Frame, app) -> None:
    """Erstellt den Dashboard-Tab-Inhalt in `parent` und bindet UI-Elemente an `app`."""
    parent.grid_columnconfigure(0, weight=1)
    parent.grid_rowconfigure(3, weight=1)

    # Toolbar
    toolbar = ttk.Frame(parent)
    toolbar.grid(row=0, column=0, sticky="ew", padx=8, pady=8)

    from controllers.dashboard_controller import refresh_dashboard
    ttk.Button(toolbar, text="Aktualisieren", command=lambda: refresh_dashboard(app)).pack(side="left", padx=(0, 8))
    from controllers.dashboard_controller import manual_adjustment
    ttk.Button(toolbar, text="Manuelle Anpassung", command=lambda: manual_adjustment(app)).pack(side="left", padx=(0, 8))
    from controllers.dashboard_controller import export_dashboard
    ttk.Button(toolbar, text="Export CSV", command=lambda: export_dashboard(app)).pack(side="left")

    # Info-Button
    create_info_button(
        parent=toolbar,
        title="Info • MD-Dashboard",
        text=(
            "MD-Dashboard\n"
            "- 'Aktualisieren' lädt die aktuellen Tracking-Daten.\n"
            "- Filter nach Name (VG/MA) und Status möglich.\n"
            "- 'Manuelle Anpassung' erlaubt Statuskorrekturen je Eintrag.\n"
            "- 'Export CSV' speichert die aktuell gefilterte Ansicht."
        ),
        side="right",
    )

    # Filter
    filter_frame = ttk.Frame(parent)
    filter_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))

    ttk.Label(filter_frame, text="Filter:").pack(side="left", padx=(0, 8))

    ttk.Label(filter_frame, text="Name:").pack(side="left", padx=(0, 4))
    app.dash_name_search = ttk.Entry(filter_frame, width=15)
    app.dash_name_search.pack(side="left", padx=(0, 16))
    app.dash_name_search.bind("<KeyRelease>", lambda e: refresh_dashboard(app))

    ttk.Label(filter_frame, text="Status:").pack(side="left", padx=(0, 4))
    app.dash_status_filter = ttk.Combobox(
        filter_frame,
        width=12,
        values=[
            "",
            ProcStatus.AUSSTEHEND.value,
            ProcStatus.ERHALTEN.value,
            ProcStatus.PRUEFUNG_NOETIG.value,
            ProcStatus.ERUEBRIGT.value,
        ],
    )
    app.dash_status_filter.pack(side="left", padx=(0, 16))
    app.dash_status_filter.bind("<<ComboboxSelected>>", lambda e: refresh_dashboard(app))

    ttk.Button(filter_frame, text="Anwenden", command=lambda: refresh_dashboard(app)).pack(side="left")

    # Status-Überschrift
    ttk.Label(parent, text="Status-Übersicht:").grid(row=2, column=0, sticky="w", padx=8, pady=(0, 4))

    # Haupt-Treeview
    cols = [
        "Log-ID",
        "VG PN",
        "VG Name",
        "MA PN",
        "MA Name",
        "Dokument-Typ",
        "Erwartet",
        "Erhalten",
        "Status",
        "Status Grund",
        "Versendet am",
        "Zuletzt erinnert am",
    ]
    app.tree_dashboard = ttk.Treeview(parent, columns=cols, show="headings", height=15)

    # Status-Farben
    app.tree_dashboard.tag_configure(DashTag.AUSSTEHEND.value, background="#ffebee")
    app.tree_dashboard.tag_configure(DashTag.ERHALTEN.value, background="#e8f5e8")
    app.tree_dashboard.tag_configure(DashTag.ERUEBRIGT.value, background="#e8f5e8")
    app.tree_dashboard.tag_configure(DashTag.PRUEFUNG_NOETIG.value, background="#fff3e0")
    app.tree_dashboard.tag_configure(DashTag.OK.value, background="#e3f2fd")
    app.tree_dashboard.tag_configure(DashTag.MANUELL.value, background="#f3e5f5")

    for c in cols:
        app.tree_dashboard.heading(c, text=c)
        if c in ["Log-ID", "VG PN", "MA PN"]:
            app.tree_dashboard.column(c, width=100, anchor="w")
        elif c in ["Status", "Dokument-Typ"]:
            app.tree_dashboard.column(c, width=120, anchor="w")
        elif c in ["Erwartet", "Erhalten"]:
            app.tree_dashboard.column(c, width=80, anchor="w")
        elif c in ["Zuletzt erinnert am", "Versendet am"]:
            app.tree_dashboard.column(c, width=140, anchor="w")
        else:
            app.tree_dashboard.column(c, width=150, anchor="w")

    # Sortierbare Spaltenköpfe (lokale Sort-Funktion wie im Original)
    def _sort_tree_by_dashboard(tree: ttk.Treeview, col: str, descending: bool) -> None:
        data = [(tree.set(k, col), k) for k in tree.get_children("")]
        def _to_key(v: str):
            s = (v or "").strip()
            try:
                if s.endswith('.0') and s.replace('.', '', 1).isdigit():
                    return int(float(s))
                return int(s)
            except Exception:
                try:
                    return float(s)
                except Exception:
                    return s.lower()
        data.sort(key=lambda t: _to_key(t[0]), reverse=descending)
        for idx, (_, k) in enumerate(data):
            tree.move(k, "", idx)
        tree.heading(col, command=lambda _c=col: _sort_tree_by_dashboard(tree, _c, not descending))

    for c in cols:
        app.tree_dashboard.heading(c, command=lambda _c=c: _sort_tree_by_dashboard(app.tree_dashboard, _c, False))

    app.tree_dashboard.grid(row=3, column=0, sticky="nsew", padx=8, pady=(0, 8))

    # Scrollbar
    scrollbar_dash = ttk.Scrollbar(parent, orient="vertical", command=app.tree_dashboard.yview)
    scrollbar_dash.grid(row=3, column=1, sticky="ns")
    app.tree_dashboard.configure(yscrollcommand=scrollbar_dash.set)

    # Initial load
    from controllers.dashboard_controller import refresh_dashboard
    refresh_dashboard(app)


