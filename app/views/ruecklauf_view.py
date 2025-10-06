from __future__ import annotations

"""
View-Modul für den Tab "Maileingang verwalten (Rücklauf)".

Minimal-invasive Extraktion des UI-Aufbaus. Verwendet weiterhin Callbacks und
Hilfsmethoden der App-Instanz (z. B. `on_scan_real`, `_make_tree`).
"""

from pathlib import Path
import tkinter as tk
from tkinter import ttk

from constants import MDConstants
from data_loader import load_config
from utils import create_info_button
from views.ui_utils import make_tree, bind_treeview_sort, autosize_tree_columns


def build_ruecklauf(parent: ttk.Frame, app) -> None:
    """Erstellt den Rücklauf-Tab-Inhalt in `parent` und bindet UI-Elemente an `app`."""
    # Toolbar
    toolbar = ttk.Frame(parent)
    toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

    create_info_button(
        parent=toolbar,
        title="Info • Rücklauf",
        text=(
            "Rücklauf verarbeiten\n"
            "1) 'Posteingang scannen' prüft das Gruppenpostfach 'VD-GS HR'.\n"
            "2) MD-Anhänge (RB/AB/Feedback) werden erkannt und kopiert.\n"
            "3) Nur MD-Anhänge → Mail wird nach '12 Mitarbeitenden-Dialog' verschoben.\n"
            "4) Probezeit- oder fremde Anhänge → Mail bleibt im Posteingang (Prüfen erforderlich).\n"
            "5) Ohne MD-Anhänge → Übersprungen."
        ),
        side="right",
    )

    from controllers.ruecklauf_controller import scan_real
    ttk.Button(toolbar, text="Posteingang scannen", command=lambda: scan_real(app)).pack(side="left", padx=(0, 8))
    ttk.Label(toolbar, text="Ziel für neue Anhänge:").pack(side="left", padx=(16,4))
    # Standard-Ziel aus Konfiguration (<root>/ruecklauf/unverarbeitet)
    CFG = load_config()
    app.inbox_target_var = tk.StringVar(value=str((Path(__file__).parent / CFG["paths"]["ruecklauf"]["unverarbeitet"]).resolve()))
    entry_target = ttk.Entry(toolbar, textvariable=app.inbox_target_var, width=60)
    entry_target.pack(side="left", padx=(0,8))

    # Statuslabels
    app.inbox_status = ttk.Label(parent, text="Bereit.", foreground="gray")
    app.inbox_status.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 4))

    app.ruecklauf_status = ttk.Label(parent, text="Noch kein Scan durchgeführt.", foreground="gray")
    app.ruecklauf_status.grid(row=2, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

    # Notebook für Ergebnislisten
    app.ruecklauf_nb = ttk.Notebook(parent)
    app.ruecklauf_nb.grid(row=3, column=0, columnspan=6, sticky="nsew", padx=8, pady=8)

    # Tabs
    app.tab_ok = ttk.Frame(app.ruecklauf_nb)
    app.tab_pruefen = ttk.Frame(app.ruecklauf_nb)
    app.tab_skip = ttk.Frame(app.ruecklauf_nb)

    app.ruecklauf_nb.add(app.tab_ok, text="Kopiert & verschoben")
    app.ruecklauf_nb.add(app.tab_pruefen, text="Prüfen erforderlich")
    app.ruecklauf_nb.add(app.tab_skip, text="Übersprungen")

    # Trees je Tab
    ttk.Label(
        app.tab_ok,
        text=(
            "Nur MD-Anhänge gefunden: Dateien wurden gespeichert und die E-Mail in den Ordner '12 Mitarbeitenden-Dialog' verschoben."
        ),
        foreground="gray",
    ).pack(anchor="w", padx=8, pady=(8,0))
    app.tree_ok = make_tree(app.tab_ok, ["Datei", "Zielordner", "Absender", "Betreff"], bind_sort=lambda tree: bind_treeview_sort(tree))

    ttk.Label(
        app.tab_pruefen,
        text=(
            "Fremde Anhänge oder Sonderfälle gefunden: MD-Dateien wurden gespeichert; die E-Mail blieb im Posteingang (manuelle Prüfung nötig)."
        ),
        foreground="gray",
    ).pack(anchor="w", padx=8, pady=(8,0))
    app.tree_pruefen = make_tree(app.tab_pruefen, ["Grund", "Zu prüfende Dokumente", "Absender", "Betreff", "Rückblick/Ausblick/Feedback kopiert?"], bind_sort=lambda tree: bind_treeview_sort(tree))

    ttk.Label(
        app.tab_skip,
        text=("Keine MD-Anhänge gefunden: E-Mail wurde übersprungen."),
        foreground="gray",
    ).pack(anchor="w", padx=8, pady=(8,0))
    app.tree_skip = make_tree(app.tab_skip, ["Absender", "Betreff", "Grund"], bind_sort=lambda tree: bind_treeview_sort(tree))

    # Layout-Weights
    parent.grid_rowconfigure(3, weight=1)
    parent.grid_columnconfigure(5, weight=1)


