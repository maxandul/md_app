from __future__ import annotations

"""
View-Modul für den Tab "SAP Stammdaten prüfen".

Diese minimal-invasive Extraktion kapselt ausschließlich den UI-Aufbau und
nutzt weiterhin Callbacks/Methoden der App-Instanz (z. B. `on_check_stammdaten`).
"""

from tkinter import ttk

from app.utils import create_info_button
from app.views.ui_utils import autosize_tree_columns
from app.theme import configure_treeview_for_alternating_rows


def build_stammdaten(parent: ttk.Frame, app) -> None:
    """Erstellt den Tab-Inhalt für die SAP Stammdaten-Validierung in `parent`.

    Verwendet vorhandene Logik der App (z. B. `on_check_stammdaten`).
    """
    bar = ttk.Frame(parent)
    bar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

    from app.controllers.stammdaten_controller import check_stammdaten
    ttk.Button(bar, text="Aktualisieren", command=lambda: check_stammdaten(app), style='Primary.TButton').pack(side="left")

    create_info_button(
        parent=bar,
        text=(
            "Überblick:\n"
            "• 'Aktualisieren' lädt EXPORT.xlsx und prüft Pflichtspalten.\n"
            "• 'Prüfpunkte' zeigen Lade- und Spalten-Checks.\n"
            "• 'Auffällige Einträge' listet BsGrd=0, doppelte PN und ungültige VG-PN.\n"
            "   Diese müssen ggf. manuell korrigiert/gelöscht werden.\n\n"
            "Vorbereitung EXPORT.xlsx:\n"
            "1) ad-hoc Query 'VD_MD' aus SAP exportieren\n"
            "2) Dateiname & Ort: 'EXPORT.xlsx' im Ordner 'sap_stammdaten'.\n"
            "3) Spalten (erste Zeile, exakt):\n"
            "   ID_NO_ZERO, Rufname, Nachname, OE Bez., OE Kurzb., Plans. Bez.,\n"
            "   lange ID/Nummer, Dir. Vorgesetzter (PN), BsGrd\n"
            "4) Alle GsGrd = 0 löschen.\n"
            "5) Mehrere Bewilligungen führen zu mehreren Zeilen. Alle Bewilligungen\n"
            "   in eine Zeile zusammenführen. Restliche Zeilen löschen.\n"
            "6) ID_NO_ZERO auf Duplikate prüfen und Duplikatelöschen.\n"
            "   Mehrfachanstellungen brauchen pro Anstellung eine Zeile.\n"
            "7) Zeilen mit Dir. Vorgesetzter (PN) = 0 sind erlaubt und können\n"
            "   später ergänzt werden."
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
        style='InfoText.TLabel',
    ).grid(row=2, column=0, columnspan=6, sticky="w", padx=8, pady=(0,4))

    cols_missing = ["Prüfpunkte", "Ergebnis"]
    app.tree_checks = ttk.Treeview(parent, columns=cols_missing, show="headings", height=6)
    for c in cols_missing:
        app.tree_checks.heading(c, text=c)
        app.tree_checks.column(c, width=360 if c == "Prüfpunkte" else 220, anchor="w")
    configure_treeview_for_alternating_rows(app.tree_checks)
    app.tree_checks.grid(row=3, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,0))
    
    scrollbar_y_checks = ttk.Scrollbar(parent, orient="vertical", command=app.tree_checks.yview)
    scrollbar_y_checks.grid(row=3, column=6, sticky="ns")
    app.tree_checks.configure(yscrollcommand=scrollbar_y_checks.set)
    
    scrollbar_x_checks = ttk.Scrollbar(parent, orient="horizontal", command=app.tree_checks.xview)
    scrollbar_x_checks.grid(row=4, column=0, columnspan=6, sticky="ew", padx=8, pady=(0,8))
    app.tree_checks.configure(xscrollcommand=scrollbar_x_checks.set)

    ttk.Label(
        parent,
        text=(
            "Auffällige Einträge: Zeigt Zeilen mit BsGrd=0, doppelter PersNr (ID_NO_ZERO) oder ungültiger VG-PN. "
            "Diese sind zu prüfen, können aber berechtigt sein (s. Infobox)."
        ),
        style='InfoText.TLabel',
    ).grid(row=5, column=0, columnspan=6, sticky="w", padx=8, pady=(8,4))

    cols_findings = ["Kategorie", "PersNr", "Nachname", "Vorname", "Details"]
    app.tree_findings = ttk.Treeview(parent, columns=cols_findings, show="headings", height=12)
    for c in cols_findings:
        app.tree_findings.heading(c, text=c)
        app.tree_findings.column(c, width=160 if c not in ("Details",) else 320, anchor="w")
    configure_treeview_for_alternating_rows(app.tree_findings)
    app.tree_findings.grid(row=6, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,0))
    
    scrollbar_y_findings = ttk.Scrollbar(parent, orient="vertical", command=app.tree_findings.yview)
    scrollbar_y_findings.grid(row=6, column=6, sticky="ns")
    app.tree_findings.configure(yscrollcommand=scrollbar_y_findings.set)
    
    scrollbar_x_findings = ttk.Scrollbar(parent, orient="horizontal", command=app.tree_findings.xview)
    scrollbar_x_findings.grid(row=7, column=0, columnspan=6, sticky="ew", padx=8, pady=(0,8))
    app.tree_findings.configure(xscrollcommand=scrollbar_x_findings.set)

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
            # Tags nach Sortierung neu setzen für alternierende Zeilen
            from app.views.ui_utils import _reapply_alternating_tags
            _reapply_alternating_tags(tree, idx, k)
        tree.heading(col, command=lambda _c=col: _sort_tree_by(tree, _c, not descending))

    for c in cols_findings:
        app.tree_findings.heading(c, command=lambda _c=c: _sort_tree_by(app.tree_findings, _c, False))

    parent.grid_rowconfigure(6, weight=1)
    parent.grid_columnconfigure(5, weight=1)


