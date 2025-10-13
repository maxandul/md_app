from __future__ import annotations

"""
View-Modul für den Tab "Maileingang verwalten (Rücklauf)".

Minimal-invasive Extraktion des UI-Aufbaus. Verwendet weiterhin Callbacks und
Hilfsmethoden der App-Instanz (z. B. `on_scan_real`, `_make_tree`).
"""

from pathlib import Path
import tkinter as tk
from tkinter import ttk

from app.constants import MDConstants
from app.data_loader import load_config
from app.utils import create_info_button
from app.views.ui_utils import make_tree, bind_treeview_sort, autosize_tree_columns


def build_ruecklauf(parent: ttk.Frame, app) -> None:
    """Erstellt den Rücklauf-Tab-Inhalt in `parent` und bindet UI-Elemente an `app`."""
    # Toolbar
    toolbar = ttk.Frame(parent)
    toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

    create_info_button(
        parent=toolbar,
        title="Info • Maileingang verwalten (Rücklauf)",
        text=(
            "Eingehende MD-Dokumente aus dem Outlook-Posteingang verarbeiten\n\n"
            "Ablauf:\n"
            "1) 'Posteingang scannen' durchsucht das Gruppenpostfach 'VD-GS HR'.\n"
            "2) Anhänge werden analysiert und nach Typ klassifiziert:\n"
            "   • Rückblick (Word/PDF)\n"
            "   • Ausblick (Word/PDF)\n"
            "   • Feedback (PDF)\n"
            "   • Probezeit-Rückblick\n"
            "   • Sonstige Dateien\n\n"
            "3) Ergebnis-Kategorien:\n"
            "   ✓ Kopiert & verschoben: Nur MD-Anhänge gefunden → Dateien gespeichert, Mail verschoben nach '12 Mitarbeitenden-Dialog'.\n"
            "   ⚠ Prüfen erforderlich: Fremde/Probezeit-Anhänge gefunden → MD-Dateien gespeichert, Mail bleibt im Posteingang.\n"
            "   ○ Übersprungen: Keine MD-Anhänge → Mail unverändert gelassen.\n\n"
            "Zielordner: Anhänge werden nach 'ruecklauf/unverarbeitet' kopiert (anpassbar im Eingabefeld)."
        ),
        side="right",
    )

    from app.controllers.ruecklauf_controller import scan_real
    ttk.Button(toolbar, text="Posteingang scannen", command=lambda: scan_real(app)).pack(side="left", padx=(0, 8))
    ttk.Label(toolbar, text="Ziel für neue Anhänge:").pack(side="left", padx=(16,4))
    # Standard-Ziel aus Konfiguration (<root>/ruecklauf/unverarbeitet)
    CFG = load_config()
    # Korrektur: Von views/ aus 2 Ebenen hoch zur app/, Config-Pfade sind relativ zu app/
    app.inbox_target_var = tk.StringVar(value=str((Path(__file__).parent.parent / CFG["paths"]["ruecklauf"]["unverarbeitet"]).resolve()))
    entry_target = ttk.Entry(toolbar, textvariable=app.inbox_target_var, width=60)
    entry_target.pack(side="left", padx=(0,8))

    # Statuslabels
    app.inbox_status = ttk.Label(parent, text="Bereit.", foreground="gray")
    app.inbox_status.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 4))

    app.ruecklauf_status = ttk.Label(parent, text="Noch kein Scan durchgeführt.", foreground="gray")
    app.ruecklauf_status.grid(row=2, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

    # ========== Sektion 1: Kopiert & verschoben ==========
    frame_ok = ttk.LabelFrame(parent, text="✓ Kopiert & verschoben", padding=8)
    frame_ok.grid(row=3, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0, 4))
    
    ttk.Label(
        frame_ok,
        text="Nur MD-Anhänge gefunden: Dateien wurden gespeichert und die E-Mail in den Ordner '12 Mitarbeitenden-Dialog' verschoben.",
        foreground="gray",
        wraplength=800
    ).pack(anchor="w", padx=4, pady=(0, 4))
    
    app.tree_ok = make_tree(
        frame_ok, 
        ["Datei", "Zielordner", "Absender", "Betreff"], 
        bind_sort=lambda tree: bind_treeview_sort(tree),
        height=8
    )

    # ========== Sektion 2: Prüfen erforderlich ==========
    frame_pruefen = ttk.LabelFrame(parent, text="⚠ Prüfen erforderlich", padding=8)
    frame_pruefen.grid(row=4, column=0, columnspan=6, sticky="nsew", padx=8, pady=4)
    
    ttk.Label(
        frame_pruefen,
        text="Fremde Anhänge oder Sonderfälle gefunden: MD-Dateien wurden gespeichert; die E-Mail blieb im Posteingang (manuelle Prüfung nötig).",
        foreground="gray",
        wraplength=800
    ).pack(anchor="w", padx=4, pady=(0, 4))
    
    app.tree_pruefen = make_tree(
        frame_pruefen, 
        ["Grund", "Zu prüfende Dokumente", "Absender", "Betreff", "Rückblick/Ausblick/Feedback kopiert?"], 
        bind_sort=lambda tree: bind_treeview_sort(tree),
        height=8
    )

    # ========== Sektion 3: Übersprungen ==========
    frame_skip = ttk.LabelFrame(parent, text="○ Übersprungen", padding=8)
    frame_skip.grid(row=5, column=0, columnspan=6, sticky="nsew", padx=8, pady=(4, 8))
    
    ttk.Label(
        frame_skip,
        text="Keine MD-Anhänge gefunden: E-Mail wurde übersprungen.",
        foreground="gray",
        wraplength=800
    ).pack(anchor="w", padx=4, pady=(0, 4))
    
    app.tree_skip = make_tree(
        frame_skip, 
        ["Absender", "Betreff", "Grund"], 
        bind_sort=lambda tree: bind_treeview_sort(tree),
        height=8
    )

    # Layout-Weights: Jede Sektion bekommt gleichmäßig Platz
    parent.grid_rowconfigure(3, weight=1)  # Kopiert & verschoben
    parent.grid_rowconfigure(4, weight=1)  # Prüfen erforderlich
    parent.grid_rowconfigure(5, weight=1)  # Übersprungen
    parent.grid_columnconfigure(5, weight=1)


