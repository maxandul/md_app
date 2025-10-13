from __future__ import annotations

"""
View-Modul für den Versand-Tab.

Dieses Modul kapselt den Aufbau des Versand-Bereichs und verwendet dabei
weiterhin die bestehenden Unterfunktionen der `App`-Klasse (`build_massenversand`,
`build_einzelversand`, `build_vg_ma_creation`).
"""

from datetime import date
import tkinter as tk
from tkinter import ttk

from app.data_loader import load_employees, build_manager_index
from app.utils import create_info_button
from app.constants import MDConstants


def build_versand(parent: ttk.Frame, app) -> None:
    """Erstellt den Versand-Tab in `parent` und bindet ihn an die gegebene `app`.

    Minimal-invasive Extraktion: Diese Funktion übernimmt die Orchestrierung
    (Notebook, Frames, Daten laden), delegiert die Detail-UI an bestehende
    Methoden der `App`-Instanz. So können wir Schritt für Schritt weitere Teile
    auslagern, ohne Verhalten zu ändern.
    """
    # Daten laden (werden auf der App-Instanz gehalten)
    app.df = load_employees()
    app.mgr_index = build_manager_index(app.df)

    # Notebook für 3 Unterbereiche
    app.versand_notebook = ttk.Notebook(parent)
    app.versand_notebook.pack(fill="both", expand=True, padx=8, pady=8)

    # Tab 1: Massenversand
    app.frame_massenversand = ttk.Frame(app.versand_notebook)
    app.versand_notebook.add(app.frame_massenversand, text="Massenversand")
    build_massenversand(app)

    # Tab 2: Einzelversand
    app.frame_einzelversand = ttk.Frame(app.versand_notebook)
    app.versand_notebook.add(app.frame_einzelversand, text="Einzelversand")
    build_einzelversand(app)

    # Tab 3: Neues VG-MA-Verhältnis
    app.frame_vg_ma = ttk.Frame(app.versand_notebook)
    app.versand_notebook.add(app.frame_vg_ma, text="Neues VG-MA-Verhältnis")
    build_vg_ma_creation(app)


def build_massenversand(app) -> None:
    """Erstellt den Massenversand Tab für den jährlichen MD-Durchlauf."""
    toolbar = ttk.Frame(app.frame_massenversand)
    toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

    app.rb_year_var = tk.IntVar(value=date.today().year)
    app.ab_year_var = tk.IntVar(value=app.rb_year_var.get() + 1)

    ttk.Label(toolbar, text="Jahr:").pack(side="left", padx=(0, 4))
    jahr_box = ttk.Combobox(
        toolbar,
        textvariable=app.rb_year_var,
        values=[date.today().year-1, date.today().year, date.today().year+1],
        state="readonly",
        width=8
    )
    jahr_box.pack(side="left", padx=(0, 8))

    app.year_label = ttk.Label(toolbar, text="")
    app.year_label.pack(side="left", padx=(0, 20))

    def _update_year_label(*_):
        rb = app.rb_year_var.get()
        app.ab_year_var.set(rb + 1)
        app.year_label.config(text=f"Rückblick: {rb} / Ausblick: {rb+1}")

    jahr_box.bind("<<ComboboxSelected>>", _update_year_label)
    _update_year_label()

    ttk.Label(toolbar, text="Suche:").pack(side="left", padx=(0, 4))
    app.filter_var = tk.StringVar()
    entry = ttk.Entry(toolbar, textvariable=app.filter_var, width=36)
    entry.pack(side="left", padx=(0, 8))
    from app.controllers.versand_controller import refresh_mgr_table
    entry.bind("<KeyRelease>", lambda e: refresh_mgr_table(app))

    create_info_button(
        parent=toolbar,
        title="Info • Massenversand",
        text=(
            "Massenversand für den jährlichen MD-Durchlauf\n\n"
            "Verwendung:\n"
            "1) Rückblick-Jahr und Ausblick-Jahr wählen (z.B. 2025/2026).\n"
            "2) Optional: Suchfeld nutzen um nach Namen, OE oder Personalnummer zu filtern.\n"
            "3) Vorgesetzte auswählen (Mehrfachauswahl mit Strg/Shift möglich).\n"
            "4) Vorschau: 'E-Mail-Vorschau' zeigt alle zu versendenden Mails.\n"
            "5) Versand: 'Generieren & Versenden' erstellt Dokumente und verschickt sofort.\n"
            "            'Generieren & Als Entwurf' erstellt Outlook-Entwürfe zur manuellen Prüfung.\n\n"
            "Automatische Dokumenten-Logik pro Mitarbeiter/in:\n"
            "• Austritt Okt-Jan: nur Rückblick\n"
            "• Probezeit-Ende Okt-Jan: Rückblick Probezeit + Ausblick\n"
            "• Probezeit-Ende Jun-Sep: nur Ausblick\n"
            "• Standard: Rückblick + Ausblick\n\n"
            "Pro Vorgesetzten wird zusätzlich eine Feedback-Vorlage erstellt.\n"
            "Alle Dokumente werden im Tracking-System erfasst (außer Probezeit-Rückblick)."
        ),
        side="right",
    )

    ttk.Label(app.frame_massenversand, text="Vorgesetzte:").grid(row=1, column=0, sticky="w", padx=8, pady=(0,4))
    cols = ["PN", "Nachname", "Vorname", "OE", "Anzahl MA"]
    app.tree = ttk.Treeview(app.frame_massenversand, columns=cols, show="headings", height=12)
    for c in cols:
        app.tree.heading(c, text=c)
        app.tree.column(c, width=120 if c in ("PN", "Anzahl MA") else 200, anchor="w")
    app.tree.grid(row=2, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))
    app._bind_treeview_sort(app.tree, numeric_like={"PN", "Anzahl MA"})

    scrollbar = ttk.Scrollbar(app.frame_massenversand, orient="vertical", command=app.tree.yview)
    scrollbar.grid(row=2, column=6, sticky="ns")
    app.tree.configure(yscrollcommand=scrollbar.set)

    btn_frame = ttk.Frame(app.frame_massenversand)
    btn_frame.grid(row=3, column=0, columnspan=6, sticky="ew", padx=8, pady=8)
    from app.controllers.versand_controller import send_managers, preview_managers
    ttk.Button(btn_frame, text="Generieren & Versenden", command=lambda: send_managers(app, mode="send")).pack(side="left")
    ttk.Button(btn_frame, text="Generieren & Als Entwurf speichern", command=lambda: send_managers(app, mode="display")).pack(side="left", padx=(8,0))
    ttk.Button(btn_frame, text="Vorschau generieren", command=lambda: preview_managers(app)).pack(side="left", padx=(8,0))

    prog_frame = ttk.Frame(app.frame_massenversand)
    prog_frame.grid(row=4, column=0, columnspan=6, sticky="ew", padx=8, pady=(0,8))
    app.ms_progress = ttk.Progressbar(prog_frame, orient="horizontal", length=300, mode="determinate")
    app.ms_progress.pack(side="left")
    app.ms_status = ttk.Label(prog_frame, text="Bereit.", foreground="gray")
    app.ms_status.pack(side="left", padx=(8,0))

    app.frame_massenversand.grid_columnconfigure(0, weight=1)
    app.frame_massenversand.grid_rowconfigure(2, weight=1)

    refresh_mgr_table(app)


def build_einzelversand(app) -> None:
    """Erstellt den Einzelversand Tab für unterjährige MD-Versendung."""
    toolbar = ttk.Frame(app.frame_einzelversand)
    toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

    app.rb_year_var_einzel = tk.IntVar(value=date.today().year)
    app.ab_year_var_einzel = tk.IntVar(value=app.rb_year_var_einzel.get() + 1)

    ttk.Label(toolbar, text="Jahr:").pack(side="left", padx=(0, 4))
    jahr_box = ttk.Combobox(
        toolbar,
        textvariable=app.rb_year_var_einzel,
        values=[date.today().year-1, date.today().year, date.today().year+1],
        state="readonly",
        width=8
    )
    jahr_box.pack(side="left", padx=(0, 8))

    app.year_label_einzel = ttk.Label(toolbar, text="")
    app.year_label_einzel.pack(side="left", padx=(0, 20))

    def _update_year_label_einzel(*_):
        rb = app.rb_year_var_einzel.get()
        app.ab_year_var_einzel.set(rb + 1)
        app.year_label_einzel.config(text=f"Rückblick: {rb} / Ausblick: {rb+1}")

    jahr_box.bind("<<ComboboxSelected>>", _update_year_label_einzel)
    _update_year_label_einzel()

    create_info_button(
        parent=toolbar,
        title="Info • Einzelversand",
        text=(
            "Einzelversand für unterjährige MD-Gespräche\n\n"
            "Verwendung:\n"
            "1) Rückblick-Jahr und Ausblick-Jahr wählen.\n"
            "2) EINEN Vorgesetzten aus der oberen Liste auswählen.\n"
            "3) In der unteren Liste erscheinen die zugeordneten Mitarbeitenden.\n"
            "4) Gewünschte Mitarbeitende auswählen (Mehrfachauswahl möglich).\n"
            "5) Dokumenttypen aktivieren:\n"
            "   [x] Rückblick: Vergangenes Jahr reflektieren\n"
            "   [x] Ausblick: Ziele für kommendes Jahr\n"
            "   [x] Rückblick Probezeit: Für MA mit Probezeit-Ende\n"
            "6) Vorschau: 'E-Mail-Vorschau' zeigt die zu versendende Mail.\n"
            "7) Versand: 'Generieren & Versenden' oder 'Als Entwurf speichern'.\n\n"
            "Hinweis: Im Gegensatz zum Massenversand können Sie hier manuell\n"
            "steuern, welche Dokumenttypen erstellt werden.\n"
            "Rückblick Probezeit wird NICHT im Tracking erfasst."
        ),
        side="right",
    )

    ttk.Label(app.frame_einzelversand, text="Vorgesetzte:").grid(row=1, column=0, sticky="w", padx=8, pady=(0,4))
    cols = ["PN", "Nachname", "Vorname", "OE", "Anzahl MA"]
    app.tree_einzel = ttk.Treeview(app.frame_einzelversand, columns=cols, show="headings", height=6)
    for c in cols:
        app.tree_einzel.heading(c, text=c)
        app.tree_einzel.column(c, width=120 if c in ("PN", "Anzahl MA") else 200, anchor="w")
    app.tree_einzel.grid(row=2, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))
    app._bind_treeview_sort(app.tree_einzel, numeric_like={"PN", "Anzahl MA"})

    ttk.Label(app.frame_einzelversand, text="Mitarbeitende:").grid(row=3, column=0, sticky="w", padx=8, pady=(0,4))
    app.subs_tree = ttk.Treeview(app.frame_einzelversand, columns=["PN", "Nachname", "Vorname", "OE"], show="headings", height=6)
    for c in ["PN", "Nachname", "Vorname", "OE"]:
        app.subs_tree.heading(c, text=c)
        app.subs_tree.column(c, width=80 if c == "PN" else 100, anchor="w")
    app.subs_tree.grid(row=4, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))
    app._bind_treeview_sort(app.subs_tree, numeric_like={"PN"})

    doc_frame = ttk.LabelFrame(app.frame_einzelversand, text="Dokumenttypen")
    doc_frame.grid(row=5, column=0, columnspan=6, sticky="ew", padx=8, pady=(0,8))

    app.var_rb = tk.BooleanVar()
    app.var_ab = tk.BooleanVar()
    app.var_pz = tk.BooleanVar()

    ttk.Checkbutton(doc_frame, text="Rückblick", variable=app.var_rb).pack(side="left", padx=8, pady=8)
    ttk.Checkbutton(doc_frame, text="Ausblick", variable=app.var_ab).pack(side="left", padx=8, pady=8)
    ttk.Checkbutton(doc_frame, text="Probezeit", variable=app.var_pz).pack(side="left", padx=8, pady=8)

    btn_frame = ttk.Frame(app.frame_einzelversand)
    btn_frame.grid(row=6, column=0, columnspan=6, sticky="ew", padx=8, pady=8)
    from app.controllers.versand_controller import send_selected_employees, preview_selected
    ttk.Button(btn_frame, text="Generieren & Versenden", command=lambda: send_selected_employees(app, mode="send")).pack(side="left")
    ttk.Button(btn_frame, text="Generieren & Als Entwurf speichern", command=lambda: send_selected_employees(app, mode="display")).pack(side="left", padx=(8,0))
    ttk.Button(btn_frame, text="Vorschau generieren", command=lambda: preview_selected(app)).pack(side="left", padx=(8,0))

    prog_frame_e = ttk.Frame(app.frame_einzelversand)
    prog_frame_e.grid(row=7, column=0, columnspan=6, sticky="ew", padx=8, pady=(0,8))
    app.es_progress = ttk.Progressbar(prog_frame_e, orient="horizontal", length=300, mode="indeterminate")
    app.es_progress.pack(side="left")
    app.es_status = ttk.Label(prog_frame_e, text="Bereit.", foreground="gray")
    app.es_status.pack(side="left", padx=(8,0))

    app.frame_einzelversand.grid_columnconfigure(0, weight=1)
    app.frame_einzelversand.grid_rowconfigure(2, weight=1)
    app.frame_einzelversand.grid_rowconfigure(4, weight=1)

    from app.controllers.versand_controller import refresh_mgr_table_einzel
    refresh_mgr_table_einzel(app)

    def _on_mgr_select_einzel(*_):
        sel = app.tree_einzel.selection()
        if not sel:
            return
        # Echte PN aus den Spaltenwerten der ausgewählten Zeile lesen
        try:
            values = app.tree_einzel.item(sel[0], "values")
            vg_pn = str(values[0]) if values else ""
        except Exception:
            vg_pn = ""
        pack = app.mgr_index.get(vg_pn)
        if not pack:
            return
        subs = pack["subs"]
        for item in app.subs_tree.get_children():
            app.subs_tree.delete(item)
        for i, (_, r) in enumerate(subs.iterrows()):
            unique_iid = f"{str(r.get('ID_NO_ZERO',''))}_{i}"
            app.subs_tree.insert("", "end", iid=unique_iid, values=[
                str(r.get("ID_NO_ZERO","")),
                str(r.get("Nachname","")),
                str(r.get("Rufname","")),
                str(r.get("OE Kurzb.","")),
            ])

    app.tree_einzel.bind("<<TreeviewSelect>>", _on_mgr_select_einzel)


def build_vg_ma_creation(app) -> None:
    """Erstellt den Tab für neue VG-MA-Verhältnisse."""
    toolbar = ttk.Frame(app.frame_vg_ma)
    toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

    create_info_button(
        parent=toolbar,
        title="Info • VG-MA-Verhältnis anlegen",
        text=(
            "Neue Vorgesetzten-Mitarbeiter-Beziehung erstellen\n\n"
            "Wann benötigt:\n"
            "• Bei Neueinstellungen: Neuer MA muss einem VG zugeordnet werden\n"
            "• Bei Wechseln: MA wechselt zu anderem VG\n"
            "• Bei Vertretungen: Temporäre Zuordnung erforderlich\n\n"
            "Vorgehen:\n"
            "1) Suche: Vorgesetzten links im Suchfeld finden und auswählen.\n"
            "2) Suche: Mitarbeiter/in rechts im Suchfeld finden und auswählen.\n"
            "3) 'Neue Beziehung erstellen' klicken.\n"
            "4) System fügt neue Zeile in EXPORT.xlsx hinzu.\n"
            "5) App neu starten, damit Änderungen wirksam werden.\n\n"
            "Hinweis: Diese Funktion erzeugt einen neuen Datensatz mit:\n"
            "• Kopie der MA-Stammdaten\n"
            "• Neue Zuordnung zum gewählten VG\n"
            "• BsGrd=0 (wird später in SAP korrigiert)\n\n"
            "Warnung: Änderungen sind sofort in EXPORT.xlsx gespeichert!"
        ),
        side="right",
    )

    vg_frame = ttk.LabelFrame(app.frame_vg_ma, text="Vorgesetzte")
    vg_frame.grid(row=1, column=0, sticky="nsew", padx=(8,4), pady=8)

    vg_search_frame = ttk.Frame(vg_frame)
    vg_search_frame.pack(fill="x", padx=8, pady=8)
    ttk.Label(vg_search_frame, text="Suche:").pack(side="left", padx=(0,4))
    app.vg_search_var = tk.StringVar()
    vg_search_entry = ttk.Entry(vg_search_frame, textvariable=app.vg_search_var, width=20)
    vg_search_entry.pack(side="left", padx=(0,8))
    from app.controllers.versand_controller import refresh_vg_list, refresh_ma_list
    vg_search_entry.bind("<KeyRelease>", lambda e: refresh_vg_list(app))
    ttk.Button(vg_search_frame, text="Aktualisieren", command=lambda: refresh_vg_list(app)).pack(side="left")

    app.vg_tree = ttk.Treeview(vg_frame, columns=["PN", "Nachname", "Vorname", "OE"], show="headings", height=8)
    for c in ["PN", "Nachname", "Vorname", "OE"]:
        app.vg_tree.heading(c, text=c)
        app.vg_tree.column(c, width=80 if c == "PN" else 120, anchor="w")
    app.vg_tree.pack(fill="both", expand=True, padx=8, pady=(0,8))
    app._bind_treeview_sort(app.vg_tree, numeric_like={"PN"})

    ma_frame = ttk.LabelFrame(app.frame_vg_ma, text="Mitarbeitende")
    ma_frame.grid(row=1, column=1, sticky="nsew", padx=(4,8), pady=8)

    ma_search_frame = ttk.Frame(ma_frame)
    ma_search_frame.pack(fill="x", padx=8, pady=8)
    ttk.Label(ma_search_frame, text="Suche:").pack(side="left", padx=(0,4))
    app.ma_search_var = tk.StringVar()
    ma_search_entry = ttk.Entry(ma_search_frame, textvariable=app.ma_search_var, width=20)
    ma_search_entry.pack(side="left", padx=(0,8))
    ma_search_entry.bind("<KeyRelease>", lambda e: refresh_ma_list(app))
    ttk.Button(ma_search_frame, text="Aktualisieren", command=lambda: refresh_ma_list(app)).pack(side="left")

    app.ma_tree = ttk.Treeview(ma_frame, columns=["PN", "Nachname", "Vorname", "OE"], show="headings", height=8)
    for c in ["PN", "Nachname", "Vorname", "OE"]:
        app.ma_tree.heading(c, text=c)
        app.ma_tree.column(c, width=80 if c == "PN" else 120, anchor="w")
    app.ma_tree.pack(fill="both", expand=True, padx=8, pady=(0,8))
    app._bind_treeview_sort(app.ma_tree, numeric_like={"PN"})

    app.selection_status = ttk.Label(app.frame_vg_ma, text="Keine Auswahl", foreground="gray")
    app.selection_status.grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(0,8))

    btn_frame = ttk.Frame(app.frame_vg_ma)
    btn_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
    from controllers.versand_controller import create_vg_ma_relationship, preview_managers, preview_selected
    ttk.Button(btn_frame, text="Neue Beziehung erstellen", command=lambda: create_vg_ma_relationship(app)).pack(side="left", padx=(0,8))

    app.frame_vg_ma.grid_columnconfigure(0, weight=1)
    app.frame_vg_ma.grid_columnconfigure(1, weight=1)
    app.frame_vg_ma.grid_rowconfigure(1, weight=1)

    refresh_vg_list(app)
    refresh_ma_list(app)

    from app.controllers.versand_controller import update_selection_status
    def _on_vg_select(*_):
        update_selection_status(app)
    def _on_ma_select(*_):
        update_selection_status(app)

    app.vg_tree.bind("<<TreeviewSelect>>", _on_vg_select)
    app.ma_tree.bind("<<TreeviewSelect>>", _on_ma_select)

