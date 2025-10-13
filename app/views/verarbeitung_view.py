from __future__ import annotations

"""
View-Modul für den Tab "MD-Dokumente verarbeiten".

Minimal-invasive Extraktion des UI-Aufbaus. Verwendet weiterhin Callbacks und
Hilfsmethoden der App-Instanz (z. B. `on_run_full_processing`, `_bind_treeview_sort`).
"""

from datetime import date
import tkinter as tk
from tkinter import ttk

from app.constants import MDConstants
from app.utils import create_info_button
from app.data_loader import load_config


def build_verarbeitung(parent: ttk.Frame, app) -> None:
    """Erstellt den Verarbeitung-Tab-Inhalt in `parent` und bindet UI-Elemente an `app`."""
    CFG = load_config()

    # Toolbar
    bar = ttk.Frame(parent)
    bar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

    # Jahrwahl mit Default bis April = Vorjahr
    ttk.Label(bar, text="Durchlauf-Jahr:").pack(side="left", padx=(0,4))
    app.proc_year_var = tk.IntVar(value=app._default_proc_year())
    proc_year_box = ttk.Combobox(
        bar,
        textvariable=app.proc_year_var,
        values=[date.today().year-1, date.today().year, date.today().year+1],
        state="readonly",
        width=8,
    )
    proc_year_box.pack(side="left", padx=(0, 12))

    # RPA-Zielverzeichnis (übersteuerbar)
    ttk.Label(bar, text="RPA-Ziel:").pack(side="left", padx=(0,4))
    app.rpa_target_var = tk.StringVar(value=str((CFG["paths"]["rpa_input_dir"])) )
    ttk.Entry(bar, textvariable=app.rpa_target_var, width=40).pack(side="left", padx=(0,12))

    # Batchgröße
    ttk.Label(bar, text="Batch:").pack(side="left", padx=(0,4))
    app.batch_size_var = tk.IntVar(value=MDConstants.PROC_DEFAULT_BATCH)
    ttk.Entry(bar, textvariable=app.batch_size_var, width=6).pack(side="left", padx=(0,12))

    from app.controllers.verarbeitung_controller import run_full_processing
    ttk.Button(bar, text="Verarbeitung starten", command=lambda: run_full_processing(app)).pack(side="left")

    # Info-Button
    create_info_button(
        parent=bar,
        title="Info • MD-Dokumente verarbeiten",
        text=(
            "Eingegangene MD-Dokumente validieren und exportieren\n\n"
            "Dieser Tab verarbeitet Dokumente aus dem Ordner 'ruecklauf/unverarbeitet',\n"
            "die zuvor via 'Maileingang verwalten' dort gespeichert wurden.\n\n"
            "Drei-Schritte-Prozess:\n\n"
            "1) DOCX prüfen:\n"
            "   • Liest Word-Dokumente (Rückblick/Ausblick)\n"
            "   • Validiert Pflichtfelder (Name, PN, Gesamteindruck)\n"
            "   • Prüft gegen SAP-Stammdaten\n"
            "   • Aktualisiert Tracking-System\n"
            "   → Ergebnis: Status 'ok' oder 'manuell'/'prüfung_nötig'\n\n"
            "2) Export & Verschieben:\n"
            "   • Erstellt SAP-Massenupload (sap_massenupload/massenupload.xlsx)\n"
            "   • Erstellt DataScience-Export (tracking/ds_export/docx_extract.csv)\n"
            f"   • Verschiebt 'ok' → '{MDConstants.VERARBEITET_DIR}'\n"
            f"   • Verschiebt 'manuell' → '{MDConstants.UNVERARBEITET_DIR}/{MDConstants.MANUELL_DIR}'\n\n"
            "3) PDFs verarbeiten:\n"
            "   • Erkennt Dokumenttyp aus Dateinamen (Rückblick/Ausblick/Feedback)\n"
            "   • Extrahiert Personalnummer\n"
            "   • Feedback → ruecklauf/feedbacks/\n"
            "   • Rückblick/Ausblick → RPA-Zielordner (für Roboter-Upload)\n"
            "   • Aktualisiert Tracking-System\n\n"
            "Einstellungen:\n"
            "• Durchlauf-Jahr: Steuert RB/AB-Zuordnung (Standard: Aktuelles Jahr bis April, sonst Vorjahr)\n"
            "• Batchgröße: Begrenzt Anzahl zu verarbeitender Dateien pro Lauf\n"
            "• RPA-Zielordner: Wohin PDFs für SAP-Upload verschoben werden"
        ),
        side="right",
    )

    # Statuslabel
    app.proc_status = ttk.Label(parent, text="Noch kein Lauf.", foreground="gray")
    app.proc_status.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

    # Überschrift und Erklärung für DOCX-Verarbeitung
    ttk.Label(parent, text="Word-Dokumente:", font=("TkDefaultFont", 10, "bold")).grid(
        row=2, column=0, columnspan=6, sticky="w", padx=8, pady=(8, 4)
    )

    docx_info = ttk.Label(
        parent,
        text=(
            "Verarbeitete Word-Dokumente mit extrahierten Steuerelement-Inhalten. "
            "Status zeigt ob Dokument korrekt verarbeitet wurde oder manuelle Prüfung benötigt."
        ),
        foreground="gray",
        wraplength=800,
    )
    docx_info.grid(row=3, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

    # Ergebnis-Tabelle DOCX
    cols = ["Datei", "Typ", "PN", "Name", "Status", "Grund", "Ziel"]
    app.tree_proc = ttk.Treeview(parent, columns=cols, show="headings", height=12)
    for c in cols:
        app.tree_proc.heading(c, text=c)
        if c == "Ziel":
            app.tree_proc.column(c, width=200, anchor="w")
        elif c in ["PN", "Status"]:
            app.tree_proc.column(c, width=100, anchor="w")
        elif c == "Grund":
            app.tree_proc.column(c, width=260, anchor="w")
        else:
            app.tree_proc.column(c, width=180, anchor="w")
    app.tree_proc.grid(row=4, column=0, columnspan=6, sticky="nsew", padx=8, pady=8)
    app._bind_treeview_sort(app.tree_proc, numeric_like={"PN"})

    # Überschrift und Erklärung für PDF-Verarbeitung
    ttk.Label(parent, text="PDF-Dokumente:", font=("TkDefaultFont", 10, "bold")).grid(
        row=5, column=0, columnspan=6, sticky="w", padx=8, pady=(8, 4)
    )

    pdf_info = ttk.Label(
        parent,
        text=(
            "Verarbeitete PDF-Dokumente basierend auf Dateiname. "
            "Bei Mehrfachanstellungen ist manuelle Prüfung erforderlich. "
            "Ziel zeigt, wohin die Datei verschoben wurde."
        ),
        foreground="gray",
        wraplength=800,
    )
    pdf_info.grid(row=6, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

    # Ergebnis-Tabelle PDFs (gleiche Spalten wie DOCX für Konsistenz)
    pdf_cols = ["Datei", "Typ", "PN", "Name", "Status", "Grund", "Ziel"]
    app.tree_pdfs = ttk.Treeview(parent, columns=pdf_cols, show="headings", height=6)
    for c in pdf_cols:
        app.tree_pdfs.heading(c, text=c)
        if c == "Ziel":
            app.tree_pdfs.column(c, width=200, anchor="w")
        elif c in ["PN", "Status"]:
            app.tree_pdfs.column(c, width=100, anchor="w")
        else:
            app.tree_pdfs.column(c, width=140, anchor="w")
    app.tree_pdfs.grid(row=7, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0, 8))
    app._bind_treeview_sort(app.tree_pdfs, numeric_like={"PN"})

    # Grid-Konfiguration
    parent.grid_rowconfigure(4, weight=3)  # DOCX-Tabelle größer
    parent.grid_rowconfigure(7, weight=1)  # PDFs kleiner
    parent.grid_columnconfigure(5, weight=1)


