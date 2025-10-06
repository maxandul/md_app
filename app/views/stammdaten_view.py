from __future__ import annotations

"""
View-Modul für den Tab "SAP Stammdaten prüfen".

Diese minimal-invasive Extraktion kapselt ausschließlich den UI-Aufbau und
nutzt weiterhin Callbacks/Methoden der App-Instanz (z. B. `on_check_stammdaten`).
"""

from tkinter import ttk

from app.utils import create_info_button
from app.views.ui_utils import autosize_tree_columns


def build_stammdaten(parent: ttk.Frame, app) -> None:
    """Erstellt den Tab-Inhalt für die SAP Stammdaten-Validierung in `parent`.

    Verwendet vorhandene Logik der App (z. B. `on_check_stammdaten`).
    """
    bar = ttk.Frame(parent)
    bar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

    from app.controllers.stammdaten_controller import check_stammdaten
    ttk.Button(bar, text="Aktualisieren", command=lambda: check_stammdaten(app)).pack(side="left")

    create_info_button(
        parent=bar,
        text=(
            "Überblick:\n"
            "• 'Aktualisieren' lädt EXPORT.xlsx und prüft Pflichtspalten.\n"
            "• 'Prüfpunkte' zeigen Lade- und Spalten-Checks.\n"
            "• 'Auffällige Einträge' listet BsGrd=0, doppelte PN und ungültige VG-PN.\n\n"
            "Vorbereitung EXPORT.xlsx:\n"
            "1) Spalten (erste Zeile, exakt):\n"
            "   ID_NO_ZERO, Rufname, Nachname, OE Bez., OE Kurzb., Plans. Bez.,\n"
            "   lange ID/Nummer, Dir. Vorgesetzter (PN), BsGrd\n"
            "2) Keine zusammengeführten Zellen oder zusätzlichen Kopfzeilen.\n"
            "3) Formate: Text/Allgemein ok; Datumsfelder dürfen echte Datumswerte sein.\n"
            "4) Personalnummern: Führende Nullen erlaubt – Längenabgleich erfolgt automatisch.\n"
            "5) Dateiname & Ort: 'EXPORT.xlsx' im Ordner 'sap_stammdaten'.\n\n"
            "Bei Abweichungen zeigt der Tab fehlende Spalten und problematische Einträge."
        ),
        title="Info • SAP Stammdaten prüfen",
        side="right",
    )

    app.lbl_fileinfo = ttk.Label(parent, text="Noch nicht geprüft.", foreground="gray")
    app.lbl_fileinfo.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0,8))

    ttk.Label(
        parent,
        text=(
            "Prüfpunkte: Hier siehst du, ob Pflichtspalten vorhanden sind und ob das Laden der Datei geklappt hat."
        ),
        foreground="gray",
    ).grid(row=2, column=0, columnspan=6, sticky="w", padx=8, pady=(0,4))

    cols_missing = ["Prüfpunkte", "Ergebnis"]
    app.tree_checks = ttk.Treeview(parent, columns=cols_missing, show="headings", height=6)
    for c in cols_missing:
        app.tree_checks.heading(c, text=c)
        app.tree_checks.column(c, width=360 if c == "Prüfpunkte" else 220, anchor="w")
    app.tree_checks.grid(row=3, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))

    ttk.Label(
        parent,
        text=(
            "Auffällige Einträge: Zeigt Datensätze mit BsGrd=0, doppelter PersNr (ID_NO_ZERO) oder ungültiger VG-PN. "
            "Diese sind informativ – sie werden NICHT automatisch vom Versand ausgeschlossen."
        ),
        foreground="gray",
    ).grid(row=4, column=0, columnspan=6, sticky="w", padx=8, pady=(8,4))

    cols_findings = ["Kategorie", "PersNr", "Nachname", "Vorname", "Details"]
    app.tree_findings = ttk.Treeview(parent, columns=cols_findings, show="headings", height=12)
    for c in cols_findings:
        app.tree_findings.heading(c, text=c)
        app.tree_findings.column(c, width=160 if c not in ("Details",) else 320, anchor="w")
    app.tree_findings.grid(row=5, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))

    # Sortierbare Findings-Tabelle
    def _sort_tree_by(tree: ttk.Treeview, col: str, descending: bool) -> None:
        data = [(tree.set(k, col), k) for k in tree.get_children("")]
        def _to_key(v: str):  # einfache Variante ohne Typenexport
            try:
                return float(v)
            except Exception:
                return (v or "").lower()
        data.sort(key=lambda t: _to_key(t[0]), reverse=descending)
        for idx, (_, k) in enumerate(data):
            tree.move(k, "", idx)
        tree.heading(col, command=lambda _c=col: _sort_tree_by(tree, _c, not descending))

    for c in cols_findings:
        app.tree_findings.heading(c, command=lambda _c=c: _sort_tree_by(app.tree_findings, _c, False))

    parent.grid_rowconfigure(5, weight=1)
    parent.grid_columnconfigure(5, weight=1)


