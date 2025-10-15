from __future__ import annotations

"""
View-Modul für den Tab "MD-Dashboard".

Minimal-invasive Extraktion des UI-Aufbaus. Verwendet weiterhin Callbacks und
Hilfsmethoden der App-Instanz (z. B. `on_refresh_dashboard`).
"""

from tkinter import ttk

from app.constants import MDConstants, ProcStatus, DashTag
from app.utils import create_info_button
from app.theme import configure_treeview_for_alternating_rows


def build_dashboard(parent: ttk.Frame, app) -> None:
    """Erstellt den Dashboard-Tab-Inhalt in `parent` und bindet UI-Elemente an `app`."""
    parent.grid_columnconfigure(0, weight=1)
    parent.grid_rowconfigure(3, weight=1)

    # Toolbar
    toolbar = ttk.Frame(parent)
    toolbar.grid(row=0, column=0, sticky="ew", padx=8, pady=8)

    from app.controllers.dashboard_controller import refresh_dashboard, manual_adjustment
    from app.controllers.erinnerung_controller import send_reminders, save_reminders_as_draft, preview_reminders
    
    # Buttons in gewünschter Reihenfolge
    ttk.Button(toolbar, text="Aktualisieren", command=lambda: refresh_dashboard(app), style='Primary.TButton').pack(side="left", padx=(0, 8))
    ttk.Button(toolbar, text="Manuelle Anpassung", command=lambda: manual_adjustment(app), style='Primary.TButton').pack(side="left", padx=(0, 8))
    ttk.Button(toolbar, text="Erinnerung versenden", command=lambda: send_reminders(app), style='Primary.TButton').pack(side="left", padx=(0, 8))
    ttk.Button(toolbar, text="Erinnerung als Entwurf speichern", command=lambda: save_reminders_as_draft(app)).pack(side="left", padx=(0, 8))
    ttk.Button(toolbar, text="Vorschau generieren", command=lambda: preview_reminders(app)).pack(side="left")

    # Info-Button
    create_info_button(
        parent=toolbar,
        title="Info • MD-Dashboard",
        text=(
            "Übersicht aller versendeten Dokumente und deren Rücklauf-Status\n\n"
            "Funktionen:\n\n"
            "• Aktualisieren: Lädt aktuelle Tracking-Daten aus 'tracking/md_logging_{jahr}.csv'.\n"
            "  Das Jahr richtet sich nach dem oben ausgewählten MD-Durchlaufjahr.\n\n"
            "• Filter:\n"
            "  - Name: Suche nach Vorgesetzten- oder Mitarbeiter-Namen\n"
            "  - Status: Zeige nur Einträge mit bestimmtem Status\n"
            "    · ausstehend: Dokument noch nicht eingegangen\n"
            "    · erhalten: Dokuemnt eingegangen\n"
            "    · prüfung_nötig: Fehler bei Verarbeitung (siehe Grund)\n"
            "    · erübrigt: Manuell als nicht mehr relevant markiert\n"
            "  - OE: Suche nach Organisationseinheit (enthält-Logik)\n"
            "    Beispiel: 'Human Resources' findet alle HR-Mitarbeiter\n\n"
            "• Manuelle Anpassung:\n"
            "  Einzelnen Eintrag auswählen und Status/Grund ändern.\n"
            "  Nützlich bei Sonderfällen oder Korrekturen.\n\n"
            "• Erinnerungen senden:\n"
            "  1) Eine oder mehrere Zeilen auswählen (mit Strg/Shift)\n"
            "  2) 'E-Mail-Vorschau' zeigt geplante Erinnerungsmails\n"
            "  3) 'Erinnerungen versenden' sendet Mails an Vorgesetzte\n"
            "  Emails werden pro Vorgesetzten gruppiert und enthalten\n"
            "  nur die ausstehenden Dokumente der ausgewählten Einträge.\n\n"
            "Spalten-Erklärung:\n"
            "• vg_pn/vg_name: Vorgesetzte/r\n"
            "• ma_pn/ma_name: Mitarbeiter/in\n"
            "• doc_type: Dokumenttyp (Rückblick/Ausblick Word/PDF)\n"
            "• erwartet/erhalten: Anzahl Dokumente\n"
            "• versendet_am: Zeitpunkt des ursprünglichen Versands\n"
            "• zuletzt_erinnert_am: Zeitpunkt der letzten Erinnerung"
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

    ttk.Label(filter_frame, text="OE:").pack(side="left", padx=(0, 4))
    app.dash_oe_search = ttk.Entry(filter_frame, width=20)
    app.dash_oe_search.pack(side="left")
    app.dash_oe_search.bind("<KeyRelease>", lambda e: refresh_dashboard(app))

    # Status-Überschrift
    ttk.Label(parent, text="Status-Übersicht:", style='SectionHeading.TLabel').grid(row=2, column=0, sticky="w", padx=8, pady=(0, 4))

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
        "OE-Kette",
    ]
    app.tree_dashboard = ttk.Treeview(parent, columns=cols, show="headings", height=15)
    configure_treeview_for_alternating_rows(app.tree_dashboard)

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
        elif c == "OE-Kette":
            app.tree_dashboard.column(c, width=250, anchor="w")
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

    app.tree_dashboard.grid(row=3, column=0, sticky="nsew", padx=8, pady=(0, 0))

    # Scrollbars
    scrollbar_dash_y = ttk.Scrollbar(parent, orient="vertical", command=app.tree_dashboard.yview)
    scrollbar_dash_y.grid(row=3, column=1, sticky="ns")
    app.tree_dashboard.configure(yscrollcommand=scrollbar_dash_y.set)
    
    scrollbar_dash_x = ttk.Scrollbar(parent, orient="horizontal", command=app.tree_dashboard.xview)
    scrollbar_dash_x.grid(row=4, column=0, sticky="ew", padx=8, pady=(0, 8))
    app.tree_dashboard.configure(xscrollcommand=scrollbar_dash_x.set)

    # Initial load
    from app.controllers.dashboard_controller import refresh_dashboard
    refresh_dashboard(app)


