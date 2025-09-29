# app/main.py
"""
MD-Prozess-Tool - Hauptanwendung

Diese Anwendung verwaltet den Mitarbeitenden-Dialog (MD) Prozess:
- SAP Stammdaten pruefen und validieren
- MD-Dokumente generieren und versenden
- Ruecklauf verarbeiten und tracken
- Dashboard fuer Status-Uebersicht

Autor: HR-Team
Version: 1.0
"""
from pathlib import Path
import tempfile
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
import os
import win32com.client
import pandas as pd

from .data_loader import load_employees, load_config, build_manager_index
from .utils import create_info_button
from .dispatch import build_and_send_for_manager
from .doc_processing import process_docx_folder, export_sap_massenupload, export_ds_csv, move_after_processing, process_pdfs
from .simple_tracking import SimpleTrackingSystem
from .org_structure import build_org_structure, export_org_structure, analyze_org_structure

CFG = load_config()

class App(tk.Tk):
    """
    Hauptanwendung für das MD-Prozess-Tool.
    
    Bietet eine GUI mit 5 Hauptbereichen:
    1. SAP Stammdaten prüfen - Validierung der Mitarbeiterdaten
    2. MD-Versand - Generierung und Versand von MD-Dokumenten
    3. Maileingang verwalten - Verarbeitung eingehender MD-Dokumente
    4. MD-Dokumente verarbeiten - Verarbeitung und Export
    5. MD-Dashboard - Status-Übersicht und Tracking
    """
    def __init__(self):
        """Initialisiert die Hauptanwendung und erstellt die GUI-Struktur."""
        super().__init__()
        self.title("MD-Prozess-Tool")
        self.geometry("1200x800")

        # Jahr-Variable ZENTRAL anlegen (wichtig, sonst None im Callback)
        self.jahr_var = tk.IntVar(value=date.today().year)
        
        # Tracking-System initialisieren
        self.tracking = SimpleTrackingSystem()

        # Notebook mit Tabs für die verschiedenen Funktionsbereiche
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        # Frame-Container für jeden Tab
        self.frame_stammdaten = ttk.Frame(self.notebook)
        self.frame_versand = ttk.Frame(self.notebook)
        self.frame_ruecklauf = ttk.Frame(self.notebook)
        self.frame_verarbeitung = ttk.Frame(self.notebook)
        self.frame_dashboard = ttk.Frame(self.notebook)

        # Tabs hinzufügen
        self.notebook.add(self.frame_stammdaten, text="SAP Stammdaten prüfen")
        self.notebook.add(self.frame_versand, text="MD-Versand")
        self.notebook.add(self.frame_ruecklauf, text="Maileingang verwalten")
        self.notebook.add(self.frame_verarbeitung, text="MD-Dokumente verarbeiten")
        self.notebook.add(self.frame_dashboard, text="MD-Dashboard")

        # Alle Tabs aufbauen
        self.build_stammdaten()
        self.build_versand()
        self.build_ruecklauf()
        self.build_verarbeitung()
        self.build_dashboard()

    # ---------------------------
    # STAMMDATEN PRÜFEN
    # ---------------------------
    def build_stammdaten(self):
        """
        Erstellt den Tab für die SAP Stammdaten-Validierung.
        
        Funktionen:
        - Lädt und prüft EXPORT.xlsx auf Vollständigkeit
        - Validiert Pflichtspalten und Datenqualität
        - Zeigt auffällige Einträge (BsGrd=0, Duplikate, fehlende VG-PN)
        """
        bar = ttk.Frame(self.frame_stammdaten)
        bar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

        ttk.Button(bar, text="Aktualisieren", command=self.on_check_stammdaten).pack(side="left")
        
        # Einheitlicher Info-Button für den Tab (inkl. Vorbereitungshinweise)
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

        self.lbl_fileinfo = ttk.Label(self.frame_stammdaten, text="Noch nicht geprüft.", foreground="gray")
        self.lbl_fileinfo.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0,8))

        ttk.Label(self.frame_stammdaten,
                  text="Prüfpunkte: Hier siehst du, ob Pflichtspalten vorhanden sind und ob das Laden der Datei geklappt hat.",
                  foreground="gray").grid(row=2, column=0, columnspan=6, sticky="w", padx=8, pady=(0,4))

        cols_missing = ["Prüfpunkte", "Ergebnis"]
        self.tree_checks = ttk.Treeview(self.frame_stammdaten, columns=cols_missing, show="headings", height=6)
        for c in cols_missing:
            self.tree_checks.heading(c, text=c)
            self.tree_checks.column(c, width=360 if c=="Prüfpunkte" else 220, anchor="w")
        self.tree_checks.grid(row=3, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))

        ttk.Label(self.frame_stammdaten,
                  text="Auffällige Einträge: Zeigt Datensätze mit BsGrd=0, doppelter PersNr (ID_NO_ZERO) oder ungültiger VG-PN. Diese sind informativ – sie werden NICHT automatisch vom Versand ausgeschlossen.",
                  foreground="gray").grid(row=4, column=0, columnspan=6, sticky="w", padx=8, pady=(8,4))

        cols_findings = ["Kategorie", "PersNr", "Nachname", "Vorname", "Details"]
        self.tree_findings = ttk.Treeview(self.frame_stammdaten, columns=cols_findings, show="headings", height=12)
        for c in cols_findings:
            self.tree_findings.heading(c, text=c)
            self.tree_findings.column(c, width=160 if c not in ("Details",) else 320, anchor="w")
        self.tree_findings.grid(row=5, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))

        # Sortierbare Findings-Tabelle
        def _sort_tree_by(tree: ttk.Treeview, col: str, descending: bool):
            data = [(tree.set(k, col), k) for k in tree.get_children("")]
            def _to_key(v):
                try:
                    return float(v)
                except Exception:
                    return (v or "").lower()
            data.sort(key=lambda t: _to_key(t[0]), reverse=descending)
            for idx, (_, k) in enumerate(data):
                tree.move(k, "", idx)
            tree.heading(col, command=lambda _c=col: _sort_tree_by(tree, _c, not descending))

        for c in cols_findings:
            self.tree_findings.heading(c, command=lambda _c=c: _sort_tree_by(self.tree_findings, _c, False))

        self.frame_stammdaten.grid_rowconfigure(5, weight=1)
        self.frame_stammdaten.grid_columnconfigure(5, weight=1)

    def on_check_stammdaten(self):
        """
        Führt die SAP Stammdaten-Validierung durch.
        
        Prüfungen:
        1. Datei-Zugriff und Metadaten
        2. Pflichtspalten-Vollständigkeit
        3. Datenqualität (BsGrd=0, Duplikate, fehlende VG-PN)
        """
        from datetime import datetime
        from .data_loader import CFG
        import pandas as pd
        xlsx_path = Path(__file__).parent / CFG["paths"]["sap_stammdaten"]

        # 1. Dateiinfo und Metadaten prüfen
        try:
            mtime = datetime.fromtimestamp((Path(xlsx_path)).stat().st_mtime)
            self.lbl_fileinfo.config(text=f"Datei: {xlsx_path} • Letzte Änderung: {mtime:%Y-%m-%d %H:%M}", foreground="black")
        except Exception as e:
            self.lbl_fileinfo.config(text=f"Datei: {xlsx_path} • Fehler beim Lesen: {e}", foreground="red")

        # 2. Vorherige Ergebnisse löschen
        for t in (self.tree_checks, self.tree_findings):
            for i in t.get_children():
                t.delete(i)

        # 3. Excel-Datei laden
        try:
            df = load_employees()
        except Exception as e:
            self.tree_checks.insert("", "end", values=["Datei laden", f"Fehler: {e}"])
            return

        # 4. Pflichtspalten-Validierung
        required_cols = [
            "ID_NO_ZERO","Rufname","Nachname","OE Bez.","OE Kurzb.",
            "Plans. Bez.","lange ID/Nummer","Dir. Vorgesetzter (PN)","BsGrd"
        ]
        df_cols = set(df.columns)
        missing = [c for c in required_cols if c not in df_cols]
        if missing:
            self.tree_checks.insert("", "end", values=["Pflichtspalten vorhanden", f"FEHLT: {', '.join(missing)}"])
        else:
            self.tree_checks.insert("", "end", values=["Pflichtspalten vorhanden", "OK"])

        # 5. Beschäftigungsgrad = 0 prüfen
        if "BsGrd" in df.columns:
            bs0 = df[df["BsGrd"].astype(str).str.strip().isin(["0","0.0"])]
            for _, r in bs0.iterrows():
                self.tree_findings.insert("", "end", values=[
                    "BsGrd=0",
                    str(r.get("ID_NO_ZERO","")),
                    str(r.get("Nachname","")),
                    str(r.get("Rufname","")),
                    "Beschäftigungsgrad=0"
                ])

        # 6. Duplikat-Erkennung bei Personalnummern
        if "ID_NO_ZERO" in df.columns:
            pns = df["ID_NO_ZERO"].astype(str).str.strip()
            dup_mask = pns.duplicated(keep=False)
            dups = df[dup_mask].copy()
            dups = dups.assign(_pn=pns[dup_mask]).sort_values("_pn")
            for _, r in dups.iterrows():
                self.tree_findings.insert("", "end", values=[
                    "Duplikat PN",
                    str(r.get("ID_NO_ZERO","")),
                    str(r.get("Nachname","")),
                    str(r.get("Rufname","")),
                    "Mehrfach vorhandene Personalnummer"
                ])

        # 7. Vorgesetzten-PN Validierung (leer oder nur Nullen)
        if "Dir. Vorgesetzter (PN)" in df.columns:
            vg_col = df["Dir. Vorgesetzter (PN)"].astype(str).str.strip()
            bad_vg = vg_col.isin(["", "nan", "NaN", "None"]) | vg_col.str.fullmatch(r"0+")
            vg0 = df[bad_vg]
            for _, r in vg0.iterrows():
                self.tree_findings.insert("", "end", values=[
                    "VG_PN fehlend/0",
                    str(r.get("ID_NO_ZERO","")),
                    str(r.get("Nachname","")),
                    str(r.get("Rufname","")),
                    "Kein gültiger Wert in 'Dir. Vorgesetzter (PN)'"
                ])

    

    # ---------------------------
    # VERSAND
    # ---------------------------
    def build_versand(self):
        # Daten laden
        self.df = load_employees()
        self.mgr_index = build_manager_index(self.df)

        # Notebook für 3 Tabs
        self.versand_notebook = ttk.Notebook(self.frame_versand)
        self.versand_notebook.pack(fill="both", expand=True, padx=8, pady=8)

        # Tab 1: Massenversand
        self.frame_massenversand = ttk.Frame(self.versand_notebook)
        self.versand_notebook.add(self.frame_massenversand, text="Massenversand")
        self.build_massenversand()

        # Tab 2: Einzelversand
        self.frame_einzelversand = ttk.Frame(self.versand_notebook)
        self.versand_notebook.add(self.frame_einzelversand, text="Einzelversand")
        self.build_einzelversand()

        # Tab 3: Neues VG-MA-Verhältnis
        self.frame_vg_ma = ttk.Frame(self.versand_notebook)
        self.versand_notebook.add(self.frame_vg_ma, text="Neues VG-MA-Verhältnis")
        self.build_vg_ma_creation()

    def build_massenversand(self):
        """Massenversand Tab (Jahreslauf)"""
        # Toolbar
        toolbar = ttk.Frame(self.frame_massenversand)
        toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

        # Jahr wählen
        self.rb_year_var = tk.IntVar(value=date.today().year)
        self.ab_year_var = tk.IntVar(value=self.rb_year_var.get() + 1)

        ttk.Label(toolbar, text="Jahr:").pack(side="left", padx=(0, 4))
        jahr_box = ttk.Combobox(
            toolbar,
            textvariable=self.rb_year_var,
            values=[date.today().year-1, date.today().year, date.today().year+1],
            state="readonly",
            width=8
        )
        jahr_box.pack(side="left", padx=(0, 8))

        self.year_label = ttk.Label(toolbar, text="")
        self.year_label.pack(side="left", padx=(0, 20))

        def _update_year_label(*_):
            rb = self.rb_year_var.get()
            self.ab_year_var.set(rb + 1)
            self.year_label.config(text=f"Rückblick: {rb} / Ausblick: {rb+1}")

        jahr_box.bind("<<ComboboxSelected>>", _update_year_label)
        _update_year_label()

        # Suche
        ttk.Label(toolbar, text="Suche:").pack(side="left", padx=(0, 4))
        self.filter_var = tk.StringVar()
        entry = ttk.Entry(toolbar, textvariable=self.filter_var, width=36)
        entry.pack(side="left", padx=(0, 8))
        entry.bind("<KeyRelease>", lambda e: self._refresh_mgr_table())

        # Info-Button
        create_info_button(
            parent=toolbar,
            title="Info • Massenversand",
            text=(
                "Massenversand (Jahreslauf)\n"
                "1) Jahr wählen.\n"
                "2) Optional über 'Suche' filtern.\n"
                "3) Vorgesetzte markieren (Mehrfachauswahl möglich).\n"
                "4) Button wählen: 'Generieren & Versenden' oder 'Generieren & Als Entwurf speichern'.\n\n"
                "Dokumenten-Logik pro Mitarbeitende/n:\n"
                "- Austritt zwischen Okt (Y) und Jan (Y+1): nur Rückblick.\n"
                "- Ende Probezeit zwischen Okt (Y) und Jan (Y+1): Rückblick Probezeit + Ausblick.\n"
                "- Ende Probezeit zwischen Jun und Sep (Y): nur Ausblick.\n"
                "- Sonst: Rückblick + Ausblick.\n"
            ),
            side="right",
        )

        # Vorgesetzten-Tabelle
        ttk.Label(self.frame_massenversand, text="Vorgesetzte:").grid(row=1, column=0, sticky="w", padx=8, pady=(0,4))
        cols = ["PN", "Nachname", "Vorname", "OE", "Anzahl MA"]
        self.tree = ttk.Treeview(self.frame_massenversand, columns=cols, show="headings", height=12)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120 if c in ("PN", "Anzahl MA") else 200, anchor="w")
        self.tree.grid(row=2, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))

        # Scrollbar
        scrollbar = ttk.Scrollbar(self.frame_massenversand, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=2, column=6, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Button
        btn_frame = ttk.Frame(self.frame_massenversand)
        btn_frame.grid(row=3, column=0, columnspan=6, sticky="ew", padx=8, pady=8)
        ttk.Button(btn_frame, text="Generieren & Versenden", command=lambda: self.on_send_managers(mode="send")).pack(side="left")
        ttk.Button(btn_frame, text="Generieren & Als Entwurf speichern", command=lambda: self.on_send_managers(mode="display")).pack(side="left", padx=(8,0))
        ttk.Button(btn_frame, text="Vorschau generieren", command=self.on_preview_managers).pack(side="left", padx=(8,0))

        # Grid-Konfiguration
        self.frame_massenversand.grid_columnconfigure(0, weight=1)
        self.frame_massenversand.grid_rowconfigure(2, weight=1)

        # Initial load
        self._refresh_mgr_table()

    def build_einzelversand(self):
        """Einzelversand Tab (Unterjährig)"""
        # Toolbar
        toolbar = ttk.Frame(self.frame_einzelversand)
        toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

        # Jahr wählen
        self.rb_year_var_einzel = tk.IntVar(value=date.today().year)
        self.ab_year_var_einzel = tk.IntVar(value=self.rb_year_var_einzel.get() + 1)

        ttk.Label(toolbar, text="Jahr:").pack(side="left", padx=(0, 4))
        jahr_box = ttk.Combobox(
            toolbar,
            textvariable=self.rb_year_var_einzel,
            values=[date.today().year-1, date.today().year, date.today().year+1],
            state="readonly",
            width=8
        )
        jahr_box.pack(side="left", padx=(0, 8))

        self.year_label_einzel = ttk.Label(toolbar, text="")
        self.year_label_einzel.pack(side="left", padx=(0, 20))

        def _update_year_label_einzel(*_):
            rb = self.rb_year_var_einzel.get()
            self.ab_year_var_einzel.set(rb + 1)
            self.year_label_einzel.config(text=f"Rückblick: {rb} / Ausblick: {rb+1}")

        jahr_box.bind("<<ComboboxSelected>>", _update_year_label_einzel)
        _update_year_label_einzel()

        # Info-Button
        create_info_button(
            parent=toolbar,
            title="Info • Einzelversand",
            text=(
                "Einzelversand (unterjährig)\n"
                "1) Jahr wählen.\n"
                "2) Einen Vorgesetzten markieren.\n"
                "3) Mitarbeitende auswählen.\n"
                "4) Dokumenttypen ankreuzen.\n"
                "5) Button wählen: 'Generieren & Versenden' oder 'Generieren & Als Entwurf speichern'."
            ),
            side="right",
        )

        # Vorgesetzten-Tabelle
        ttk.Label(self.frame_einzelversand, text="Vorgesetzte:").grid(row=1, column=0, sticky="w", padx=8, pady=(0,4))
        cols = ["PN", "Nachname", "Vorname", "OE", "Anzahl MA"]
        self.tree_einzel = ttk.Treeview(self.frame_einzelversand, columns=cols, show="headings", height=6)
        for c in cols:
            self.tree_einzel.heading(c, text=c)
            self.tree_einzel.column(c, width=120 if c in ("PN", "Anzahl MA") else 200, anchor="w")
        self.tree_einzel.grid(row=2, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))

        # Mitarbeitenden-Tabelle
        ttk.Label(self.frame_einzelversand, text="Mitarbeitende:").grid(row=3, column=0, sticky="w", padx=8, pady=(0,4))
        self.subs_tree = ttk.Treeview(self.frame_einzelversand, columns=["PN", "Nachname", "Vorname", "OE"], show="headings", height=6)
        for c in ["PN", "Nachname", "Vorname", "OE"]:
            self.subs_tree.heading(c, text=c)
            self.subs_tree.column(c, width=80 if c == "PN" else 100, anchor="w")
        self.subs_tree.grid(row=4, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0,8))

        # Dokumenttypen
        doc_frame = ttk.LabelFrame(self.frame_einzelversand, text="Dokumenttypen")
        doc_frame.grid(row=5, column=0, columnspan=6, sticky="ew", padx=8, pady=(0,8))
        
        self.var_rb = tk.BooleanVar()
        self.var_ab = tk.BooleanVar()
        self.var_pz = tk.BooleanVar()
        
        ttk.Checkbutton(doc_frame, text="Rückblick", variable=self.var_rb).pack(side="left", padx=8, pady=8)
        ttk.Checkbutton(doc_frame, text="Ausblick", variable=self.var_ab).pack(side="left", padx=8, pady=8)
        ttk.Checkbutton(doc_frame, text="Probezeit", variable=self.var_pz).pack(side="left", padx=8, pady=8)

        # Button
        btn_frame = ttk.Frame(self.frame_einzelversand)
        btn_frame.grid(row=6, column=0, columnspan=6, sticky="ew", padx=8, pady=8)
        ttk.Button(btn_frame, text="Generieren & Versenden", command=lambda: self.on_send_selected_employees(mode="send")).pack(side="left")
        ttk.Button(btn_frame, text="Generieren & Als Entwurf speichern", command=lambda: self.on_send_selected_employees(mode="display")).pack(side="left", padx=(8,0))
        ttk.Button(btn_frame, text="Vorschau generieren", command=self.on_preview_selected).pack(side="left", padx=(8,0))

        # Grid-Konfiguration
        self.frame_einzelversand.grid_columnconfigure(0, weight=1)
        self.frame_einzelversand.grid_rowconfigure(2, weight=1)
        self.frame_einzelversand.grid_rowconfigure(4, weight=1)

        # Initial load
        self._refresh_mgr_table_einzel()

        # Callback für VG-Auswahl
        def _on_mgr_select_einzel(*_):
            sel = self.tree_einzel.selection()
            if not sel:
                return
            vg_pn = sel[0]
            pack = self.mgr_index.get(vg_pn)
            if not pack:
                return
            subs = pack["subs"]
            # Treeview leeren
            for item in self.subs_tree.get_children():
                self.subs_tree.delete(item)
            for _, r in subs.iterrows():
                self.subs_tree.insert("", "end", iid=str(r.get("ID_NO_ZERO","")), values=[
                    str(r.get("ID_NO_ZERO","")),
                    str(r.get("Nachname","")),
                    str(r.get("Rufname","")),
                    str(r.get("OE Kurzb.","")),
                ])

        self.tree_einzel.bind("<<TreeviewSelect>>", _on_mgr_select_einzel)

    def build_vg_ma_creation(self):
        """Neues VG-MA-Verhältnis Tab"""
        # Toolbar
        toolbar = ttk.Frame(self.frame_vg_ma)
        toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

        # Info-Button
        create_info_button(
            parent=toolbar,
            title="Info • VG-MA-Verhältnis",
            text=(
                "Neues VG-MA-Verhältnis anlegen\n"
                "1) VG links auswählen.\n"
                "2) MA rechts auswählen.\n"
                "3) 'Neue Beziehung erstellen' legt eine neue Zeile in EXPORT.xlsx an.\n"
                "4) App neu starten, um Änderungen zu laden."
            ),
            side="right",
        )

        # Vorgesetzten-Liste
        vg_frame = ttk.LabelFrame(self.frame_vg_ma, text="Vorgesetzte")
        vg_frame.grid(row=1, column=0, sticky="nsew", padx=(8,4), pady=8)
        
        # VG-Suche
        vg_search_frame = ttk.Frame(vg_frame)
        vg_search_frame.pack(fill="x", padx=8, pady=8)
        ttk.Label(vg_search_frame, text="Suche:").pack(side="left", padx=(0,4))
        self.vg_search_var = tk.StringVar()
        vg_search_entry = ttk.Entry(vg_search_frame, textvariable=self.vg_search_var, width=20)
        vg_search_entry.pack(side="left", padx=(0,8))
        vg_search_entry.bind("<KeyRelease>", lambda e: self._refresh_vg_list())
        ttk.Button(vg_search_frame, text="Aktualisieren", command=self._refresh_vg_list).pack(side="left")

        # VG-Treeview
        self.vg_tree = ttk.Treeview(vg_frame, columns=["PN", "Nachname", "Vorname", "OE"], show="headings", height=8)
        for c in ["PN", "Nachname", "Vorname", "OE"]:
            self.vg_tree.heading(c, text=c)
            self.vg_tree.column(c, width=80 if c == "PN" else 120, anchor="w")
        self.vg_tree.pack(fill="both", expand=True, padx=8, pady=(0,8))

        # Mitarbeitenden-Liste
        ma_frame = ttk.LabelFrame(self.frame_vg_ma, text="Mitarbeitende")
        ma_frame.grid(row=1, column=1, sticky="nsew", padx=(4,8), pady=8)
        
        # MA-Suche
        ma_search_frame = ttk.Frame(ma_frame)
        ma_search_frame.pack(fill="x", padx=8, pady=8)
        ttk.Label(ma_search_frame, text="Suche:").pack(side="left", padx=(0,4))
        self.ma_search_var = tk.StringVar()
        ma_search_entry = ttk.Entry(ma_search_frame, textvariable=self.ma_search_var, width=20)
        ma_search_entry.pack(side="left", padx=(0,8))
        ma_search_entry.bind("<KeyRelease>", lambda e: self._refresh_ma_list())
        ttk.Button(ma_search_frame, text="Aktualisieren", command=self._refresh_ma_list).pack(side="left")

        # MA-Treeview
        self.ma_tree = ttk.Treeview(ma_frame, columns=["PN", "Nachname", "Vorname", "OE"], show="headings", height=8)
        for c in ["PN", "Nachname", "Vorname", "OE"]:
            self.ma_tree.heading(c, text=c)
            self.ma_tree.column(c, width=80 if c == "PN" else 120, anchor="w")
        self.ma_tree.pack(fill="both", expand=True, padx=8, pady=(0,8))

        # Auswahl-Status
        self.selection_status = ttk.Label(self.frame_vg_ma, text="Keine Auswahl", foreground="gray")
        self.selection_status.grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(0,8))

        # Buttons
        btn_frame = ttk.Frame(self.frame_vg_ma)
        btn_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        ttk.Button(btn_frame, text="Neue Beziehung erstellen", command=self.on_create_vg_ma_relationship).pack(side="left", padx=(0,8))

        # Grid-Konfiguration
        self.frame_vg_ma.grid_columnconfigure(0, weight=1)
        self.frame_vg_ma.grid_columnconfigure(1, weight=1)
        self.frame_vg_ma.grid_rowconfigure(1, weight=1)

        # Initial load
        self._refresh_vg_list()
        self._refresh_ma_list()

        # Callbacks für Auswahl
        def _on_vg_select(*_):
            self._update_selection_status()
        def _on_ma_select(*_):
            self._update_selection_status()
        
        self.vg_tree.bind("<<TreeviewSelect>>", _on_vg_select)
        self.ma_tree.bind("<<TreeviewSelect>>", _on_ma_select)

    def on_send_managers(self, mode: str = None):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Hinweis", "Bitte mindestens eine/n Vorgesetzte/n auswählen.")
            return

        rb_year = self.rb_year_var.get()
        ab_year = self.ab_year_var.get()
        out_root = Path(__file__).parent.parent / "tracking" / "versand"
        out_root.mkdir(parents=True, exist_ok=True)

        errors = []
        for vg_pn in sel:
            pack = self.mgr_index.get(vg_pn)
            if not pack:
                errors.append(f"Kein Paket für VG_PN {vg_pn}")
                continue
            mgr = pack["manager"]
            subs = pack["subs"]
            if mgr is None:
                errors.append(f"Vorgesetzte/r mit PN {vg_pn} nicht in EXPORT.xlsx gefunden.")
                continue

            try:
                build_and_send_for_manager(
                    mgr_row=mgr,
                    subs_df=subs,
                    rb_year=rb_year,
                    ab_year=ab_year,
                    today=date.today(),
                    out_root=out_root,
                    managers_index=self.mgr_index,
                    include_feedback=True,
                    send_mode=mode,
                )
                
                # Tracking: Logge Versand für jeden Mitarbeiter
                mgr_name = f"{mgr.get('Rufname','')} {mgr.get('Nachname','')}"
                
                # Erst: Feedback einmal pro Vorgesetzten loggen
                self.tracking.log_feedback_for_manager(vg_pn, mgr_name, rb_year, self.mgr_index)
                
                # Dann: Dokumente pro Mitarbeiter loggen
                for _, emp in subs.iterrows():
                    emp_pn = str(emp.get("ID_NO_ZERO", "")).strip()
                    emp_name = f"{emp.get('Rufname','')} {emp.get('Nachname','')}"
                    
                    # Bestimme Dokumenttypen basierend auf Mitarbeiter-Status
                    doc_types = []
                    
                    # Rückblick für alle
                    doc_types.append("rueckblick")
                    
                    # Ausblick nur wenn nicht austretend
                    if pd.isna(emp.get("Austritt")) or emp.get("Austritt") == "":
                        doc_types.append("ausblick")
                    
                    # Probezeit-Rückblick wenn Probezeit Ende zwischen Okt-Jan
                    if not pd.isna(emp.get("Ende Probezeit")):
                        probezeit_ende = emp.get("Ende Probezeit")
                        if isinstance(probezeit_ende, pd.Timestamp):
                            month = probezeit_ende.month
                            if month in [10, 11, 12, 1]:
                                doc_types.append("rueckblick_probezeit")
                    
                    # Logge Versand (ohne Feedback, da bereits oben geloggt)
                    self.tracking.log_versand(
                        mgr_pn=vg_pn,
                        mgr_name=mgr_name,
                        emp_pn=emp_pn,
                        emp_name=emp_name,
                        doc_types=doc_types,
                        rb_year=rb_year,
                        ab_year=ab_year,
                        include_feedback=False
                    )
                
            except Exception as e:
                errors.append(f"{mgr.get('Rufname','')} {mgr.get('Nachname','')} ({vg_pn}): {e}")

        if errors:
            messagebox.showerror("Abschluss mit Fehlern", "\n".join(errors))
        else:
            messagebox.showinfo("Fertig", "Versand ausgefuehrt (siehe Outbox/gesendete Elemente).")

    def on_send_selected_employees(self, mode: str = None):
        sel_mgrs = self.tree_einzel.selection()
        if len(sel_mgrs) != 1:
            messagebox.showwarning("Hinweis", "Bitte genau eine/n Vorgesetzte/n auswählen.")
            return
        vg_pn = sel_mgrs[0]
        pack = self.mgr_index.get(vg_pn)
        if not pack:
            messagebox.showerror("Fehler", f"Kein Paket für VG_PN {vg_pn}")
            return

        subs = pack["subs"].copy()
        sel_subs = self.subs_tree.selection()
        if not sel_subs:
            messagebox.showwarning("Hinweis", "Bitte mindestens eine/n Mitarbeitende/n auswählen.")
            return

        # Filter auf ausgewählte PN
        sel_pns = set(sel_subs)
        if "ID_NO_ZERO" not in subs.columns:
            messagebox.showerror("Fehler", "Stammdaten enthalten keine Spalte ID_NO_ZERO.")
            return
        subs_filtered = subs[subs["ID_NO_ZERO"].astype(str).isin(sel_pns)]
        if subs_filtered.empty:
            messagebox.showerror("Fehler", "Keine passenden Mitarbeitenden gefunden.")
            return

        # Dokumenttypen aus Auswahl
        types = []
        if self.var_rb.get():
            types.append("Rückblick")
        if self.var_ab.get():
            types.append("Ausblick")
        if self.var_pz.get():
            types.append("Rückblick_Probezeit")
        if not types:
            messagebox.showwarning("Hinweis", "Bitte mindestens einen Dokumenttyp auswählen.")
            return

        rb_year = self.rb_year_var_einzel.get()
        ab_year = self.ab_year_var_einzel.get()
        out_root = Path(__file__).parent.parent / "tracking" / "versand"
        out_root.mkdir(parents=True, exist_ok=True)

        mgr = pack["manager"]
        try:
            build_and_send_for_manager(
                mgr_row=mgr,
                subs_df=subs_filtered,
                rb_year=rb_year,
                ab_year=ab_year,
                today=date.today(),
                out_root=out_root,
                managers_index=self.mgr_index,
                doc_types_override=types,
                include_feedback=False,
                send_mode=mode,
            )
            
            # Tracking: Logge Versand für jeden ausgewählten Mitarbeiter
            mgr_name = f"{mgr.get('Rufname','')} {mgr.get('Nachname','')}"
            
            # Erst: Feedback einmal pro Vorgesetzten loggen (nur wenn Feedback ausgewählt)
            if "Feedback" in types:
                self.tracking.log_feedback_for_manager(vg_pn, mgr_name, rb_year, self.mgr_index)
            
            # Dann: Dokumente pro Mitarbeiter loggen
            for _, emp in subs_filtered.iterrows():
                emp_pn = str(emp.get("ID_NO_ZERO", "")).strip()
                emp_name = f"{emp.get('Rufname','')} {emp.get('Nachname','')}"
                
                # Konvertiere UI-Typen zu internen Typen (ohne Feedback)
                doc_types = []
                for ui_type in types:
                    if ui_type == "Rückblick":
                        doc_types.append("rueckblick")
                    elif ui_type == "Ausblick":
                        doc_types.append("ausblick")
                    elif ui_type == "Rückblick_Probezeit":
                        doc_types.append("rueckblick_probezeit")
                    # Feedback wird separat oben geloggt
                
                # Logge Versand (ohne Feedback)
                self.tracking.log_versand(
                    mgr_pn=vg_pn,
                    mgr_name=mgr_name,
                    emp_pn=emp_pn,
                    emp_name=emp_name,
                    doc_types=doc_types,
                    rb_year=rb_year,
                    ab_year=ab_year,
                    include_feedback=False
                )
            
        except Exception as e:
            messagebox.showerror("Fehler", str(e))
            return

        messagebox.showinfo("Fertig", "Unterlagen erzeugt und E-Mail vorbereitet/gesendet.")

    # ---------------------------
    # RÜCKLAUF (Prototyp GUI)
    # ---------------------------
            
    def build_ruecklauf(self):
        # --- Toolbar Frame ---
        toolbar = ttk.Frame(self.frame_ruecklauf)
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

        # Scan-Button
        ttk.Button(toolbar, text="Posteingang scannen", command=self.on_scan_real).pack(side="left", padx=(0, 8))
        ttk.Label(toolbar, text="Ziel für neue Anhänge:").pack(side="left", padx=(16,4))
        self.inbox_target_var = tk.StringVar(value=str((Path(__file__).parent.parent / "ruecklauf" / "unverarbeitet").resolve()))
        entry_target = ttk.Entry(toolbar, textvariable=self.inbox_target_var, width=60)
        entry_target.pack(side="left", padx=(0,8))

        # Status-Zeile
        self.ruecklauf_status = ttk.Label(self.frame_ruecklauf, text="Noch kein Scan durchgeführt.", foreground="gray")
        self.ruecklauf_status.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

        # Notebook für Ergebnislisten
        self.ruecklauf_nb = ttk.Notebook(self.frame_ruecklauf)
        self.ruecklauf_nb.grid(row=2, column=0, columnspan=6, sticky="nsew", padx=8, pady=8)

        # Tabs
        self.tab_ok = ttk.Frame(self.ruecklauf_nb)
        self.tab_pruefen = ttk.Frame(self.ruecklauf_nb)
        self.tab_skip = ttk.Frame(self.ruecklauf_nb)

        self.ruecklauf_nb.add(self.tab_ok, text="Kopiert & verschoben")
        self.ruecklauf_nb.add(self.tab_pruefen, text="Prüfen erforderlich")
        self.ruecklauf_nb.add(self.tab_skip, text="Übersprungen")

        # Treeviews für jede Kategorie
        ttk.Label(self.tab_ok,
                  text="Nur MD-Anhänge gefunden: Dateien wurden gespeichert und die E-Mail in den Ordner '12 Mitarbeitenden-Dialog' verschoben.",
                  foreground="gray").pack(anchor="w", padx=8, pady=(8,0))
        self.tree_ok = self._make_tree(self.tab_ok, ["Datei", "Zielordner", "Absender", "Betreff"])
        ttk.Label(self.tab_pruefen,
                  text="Probezeit- und/oder fremde Anhänge gefunden: MD-Dateien wurden gespeichert, die E-Mail blieb im Posteingang (manuelle Prüfung nötig).",
                  foreground="gray").pack(anchor="w", padx=8, pady=(8,0))
        self.tree_pruefen = self._make_tree(self.tab_pruefen, ["Grund", "Zu prüfende Dokumente", "Absender", "Betreff", "Rückblick/Ausblick/Feedback kopiert?"])
        ttk.Label(self.tab_skip,
                  text="Keine MD-Anhänge gefunden: E-Mail wurde übersprungen.",
                  foreground="gray").pack(anchor="w", padx=8, pady=(8,0))
        self.tree_skip = self._make_tree(self.tab_skip, ["Absender", "Betreff", "Grund"])

        # Layout-Weights
        self.frame_ruecklauf.grid_rowconfigure(2, weight=1)
        self.frame_ruecklauf.grid_columnconfigure(5, weight=1)



    def _make_tree(self, parent, cols):
        """Helper: erstellt Treeview mit Scrollbar"""
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=180 if c != "Betreff" else 300, anchor="w")
        tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        return tree


    def on_scan_real(self):
        """Scan der Shared Mailbox 'VD-GS HR' und Verarbeitung nach Regeln"""
        for t in [self.tree_ok, self.tree_pruefen, self.tree_skip]:
            for i in t.get_children():
                t.delete(i)

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        mailbox = outlook.Folders["VD-GS HR"]
        inbox = mailbox.Folders["Posteingang"]
        target_folder = inbox.Folders["12 Mitarbeitenden-Dialog"]

        md_keywords = ["rückblick", "rueckblick", "ausblick", "feedback"]
        kw_probezeit = "probezeit"
        allowed_exts = [".docx", ".pdf"]

        base_path = Path(self.inbox_target_var.get())
        base_path.mkdir(parents=True, exist_ok=True)

        found, copied, moved, to_check, skipped = 0, 0, 0, 0, 0

        for mail in inbox.Items:
            found += 1
            sender = self._get_sender_address(mail)
            subject = str(mail.Subject or "")

            # Anhänge einsammeln
            files = []
            for att in mail.Attachments:
                try:
                    fname = str(att.FileName or "").strip()
                except Exception:
                    continue
                if not fname or "." not in fname:
                    continue
                files.append((fname, att))

            if not files:
                self.tree_skip.insert("", "end", values=[sender, subject, "Keine Anhänge"])
                skipped += 1
                continue

            # Klassifizieren
            md_files = [f for f, _ in files if any(k in f.lower() for k in md_keywords) and os.path.splitext(f)[1].lower() in allowed_exts]
            probezeit_files = [f for f, _ in files if kw_probezeit in f.lower() and os.path.splitext(f)[1].lower() in allowed_exts]
            other_files = [f for f, _ in files if f not in md_files and f not in probezeit_files]

            if md_files and not probezeit_files and not other_files:
                # Sauber → kopieren & verschieben
                for fname, att in files:
                    if fname in md_files:
                        save_path = base_path / fname
                        att.SaveAsFile(str(save_path))
                        self.tree_ok.insert("", "end", values=[fname, str(base_path), sender, subject])
                        copied += 1
                mail.Move(target_folder)
                moved += 1

            elif probezeit_files:
                # Probezeit → MD kopieren, Mail bleibt
                # Probezeit: nicht verschieben/kopieren, nur als Prüf-Fall listen
                grund = "Probezeit"
                self.tree_pruefen.insert("", "end", values=[grund, ", ".join(probezeit_files) or "Keine", sender, subject, "Keine"])
                to_check += 1

            elif md_files and other_files:
                # Gemischt → MD kopieren, Mail bleibt
                copied_names = []
                for fname, att in files:
                    if fname in md_files:
                        save_path = base_path / fname
                        att.SaveAsFile(str(save_path))
                        copied += 1
                        copied_names.append(fname)
                grund = "Fremde Anhänge"
                self.tree_pruefen.insert("", "end", values=[grund, ", ".join(other_files) or "Keine", sender, subject, ", ".join(copied_names) or "Keine"])
                to_check += 1

            elif probezeit_files and other_files:
                # Probezeit + Fremd
                # Probezeit + Fremde: nichts kopieren, nur listen
                grund = "Probezeit + Fremde Anhänge"
                zu_pruefen = probezeit_files + other_files
                self.tree_pruefen.insert("", "end", values=[grund, ", ".join(zu_pruefen) or "Keine", sender, subject, "Keine"])
                to_check += 1

            else:
                # Keine MD-Anhänge → skip
                self.tree_skip.insert("", "end", values=[sender, subject, "Keine MD-Anhänge"])
                skipped += 1

        self.ruecklauf_status.config(
            text=f"Scan abgeschlossen: {found} Mails • {copied} Anhänge kopiert • {moved} verschoben • {to_check} prüfen • {skipped} übersprungen",
            foreground="black"
        )

    def _get_sender_address(self, mail):
        """Versucht, eine saubere SMTP-Adresse für den Absender zurückzugeben."""
        try:
            sender = mail.Sender
            if sender and sender.AddressEntryUserType == 0:  # 0 = ExchangeUser
                ex_user = sender.GetExchangeUser()
                if ex_user:
                    return ex_user.PrimarySmtpAddress
        except Exception:
            pass
        try:
            return mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        except Exception:
            pass
        return str(mail.SenderEmailAddress or "")

    # ---------------------------
    # VERARBEITUNG
    # ---------------------------

    def build_verarbeitung(self):
        # Toolbar
        bar = ttk.Frame(self.frame_verarbeitung)
        bar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

        # Jahrwahl mit Default bis April = Vorjahr
        ttk.Label(bar, text="Durchlauf-Jahr:").pack(side="left", padx=(0,4))
        self.proc_year_var = tk.IntVar(value=self._default_proc_year())
        proc_year_box = ttk.Combobox(bar, textvariable=self.proc_year_var,
                                     values=[date.today().year-1, date.today().year, date.today().year+1],
                                     state="readonly", width=8)
        proc_year_box.pack(side="left", padx=(0, 12))

        # RPA-Zielverzeichnis (übersteuerbar)
        ttk.Label(bar, text="RPA-Ziel:").pack(side="left", padx=(0,4))
        self.rpa_target_var = tk.StringVar(value=str((Path(CFG["paths"]["rpa_input_dir"]).resolve())))
        ttk.Entry(bar, textvariable=self.rpa_target_var, width=40).pack(side="left", padx=(0,12))

        # Batchgröße
        ttk.Label(bar, text="Batch:").pack(side="left", padx=(0,4))
        self.batch_size_var = tk.IntVar(value=100)
        ttk.Entry(bar, textvariable=self.batch_size_var, width=6).pack(side="left", padx=(0,12))

        ttk.Button(bar, text="DOCX prüfen & extrahieren",
                command=self.on_process_docx).pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="Export (SAP+DS) & verschieben",
                command=self.on_export_and_move).pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="PDFs verarbeiten",
                command=self.on_process_pdfs).pack(side="left")

        # Info-Button
        create_info_button(
            parent=bar,
            title="Info • MD-Dokumente verarbeiten",
            text=(
                "MD-Dokumente verarbeiten\n"
                "DOCX prüfen & extrahieren: Liest DOCX im Ordner 'ruecklauf/unverarbeitet' und extrahiert Status.\n"
                "Export (SAP+DS) & verschieben: Schreibt Exporte und verschiebt Dateien nach 'archiv' bzw. 'manuell'.\n"
                "PDFs verarbeiten: Verteilt eingehende PDFs nach RPA-Zielordner.\n"
                "Batchgröße begrenzt die Anzahl je Lauf; Durchlauf-Jahr steuert RB/AB-Zuordnung."
            ),
            side="right",
        )

        # Statuslabel
        self.proc_status = ttk.Label(self.frame_verarbeitung, text="Noch kein Lauf.", foreground="gray")
        self.proc_status.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

        # Überschrift und Erklärung für DOCX-Verarbeitung
        ttk.Label(self.frame_verarbeitung, text="DOCX-Dokumente (Word-Vorlagen):", 
                 font=("TkDefaultFont", 10, "bold")).grid(row=2, column=0, columnspan=6, sticky="w", padx=8, pady=(8, 4))
        
        docx_info = ttk.Label(self.frame_verarbeitung, 
                             text="Verarbeitete Word-Dokumente mit extrahierten Steuerelement-Inhalten. Status zeigt ob Dokument korrekt verarbeitet wurde oder manuelle Prüfung benötigt.",
                             foreground="gray", wraplength=800)
        docx_info.grid(row=3, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

        # Ergebnis-Tabelle DOCX
        cols = ["Datei", "Typ", "PN", "Name", "Status", "Grund", "Gesamteindruck (RB)"]
        self.tree_proc = ttk.Treeview(self.frame_verarbeitung, columns=cols, show="headings", height=12)
        for c in cols:
            self.tree_proc.heading(c, text=c)
            self.tree_proc.column(c, width=180 if c not in ("Datei", "Grund") else 260, anchor="w")
        self.tree_proc.grid(row=4, column=0, columnspan=6, sticky="nsew", padx=8, pady=8)

        # Überschrift und Erklärung für PDF-Verarbeitung
        ttk.Label(self.frame_verarbeitung, text="PDF-Dokumente (eingescannte/konvertierte Dokumente):", 
                 font=("TkDefaultFont", 10, "bold")).grid(row=5, column=0, columnspan=6, sticky="w", padx=8, pady=(8, 4))
        
        pdf_info = ttk.Label(self.frame_verarbeitung, 
                            text="Verarbeitete PDF-Dokumente basierend auf Dateiname. Bei Mehrfachanstellungen ist manuelle Prüfung erforderlich. Ziel zeigt wohin Datei verschoben wurde.",
                            foreground="gray", wraplength=800)
        pdf_info.grid(row=6, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

        # Ergebnis-Tabelle PDFs (gleiche Spalten wie DOCX für Konsistenz)
        pdf_cols = ["Datei", "Typ", "PN", "Name", "Status", "Grund", "Ziel"]
        self.tree_pdfs = ttk.Treeview(self.frame_verarbeitung, columns=pdf_cols, show="headings", height=6)
        for c in pdf_cols:
            self.tree_pdfs.heading(c, text=c)
            if c == "Ziel":
                self.tree_pdfs.column(c, width=200, anchor="w")
            elif c in ["PN", "Status"]:
                self.tree_pdfs.column(c, width=100, anchor="w")
            else:
                self.tree_pdfs.column(c, width=140, anchor="w")
        self.tree_pdfs.grid(row=7, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0, 8))

        # Grid-Konfiguration
        self.frame_verarbeitung.grid_rowconfigure(4, weight=3)  # DOCX-Tabelle größer
        self.frame_verarbeitung.grid_rowconfigure(7, weight=1)  # PDFs kleiner
        self.frame_verarbeitung.grid_columnconfigure(5, weight=1)
    def _default_proc_year(self) -> int:
        today = date.today()
        return today.year - 1 if today.month <= 4 else today.year


    def on_process_docx(self):
        # Quelle: /ruecklauf/unverarbeitet
        input_dir = Path(__file__).parent.parent / "ruecklauf" / "unverarbeitet"

        # Tabelle leeren
        for i in self.tree_proc.get_children():
            self.tree_proc.delete(i)

        # Stammdaten laden
        sap_df = load_employees()

        # Prozess laufen lassen (Batch)
        try:
            max_files = int(self.batch_size_var.get()) if self.batch_size_var.get() else None
        except Exception:
            max_files = None
        durchlauf_jahr = self.proc_year_var.get()
        results = process_docx_folder(input_dir, sap_df, max_files=max_files, durchlauf_jahr=durchlauf_jahr)

        ok_count = 0
        man_count = 0

        for r in results:
            gi = r["extras"].get("rb_gesamteindruck", "") if isinstance(r.get("extras"), dict) else ""
            self.tree_proc.insert("", "end", values=[
                r.get("file",""), r.get("typ",""), r.get("pn",""), r.get("name",""),
                r.get("status",""), r.get("reason",""), gi
            ])
            if r.get("status") == "ok":
                ok_count += 1
            elif r.get("status") == "manuell":
                man_count += 1

        self.proc_status.config(
            text=f"DOCX geprüft: {len(results)} Dateien • OK: {ok_count} • Manuell: {man_count}",
            foreground="black"
        )

        self._last_docx_results = results


    def on_export_and_move(self):
        input_dir = Path(__file__).parent.parent / "ruecklauf" / "unverarbeitet"
        sap_out = Path(__file__).parent.parent / "sap_massenupload" / "massenupload.xlsx"
        ds_out  = Path(__file__).parent.parent / "tracking" / "ds_export" / "docx_extract.csv"

        # 1) Wenn keine letzte Prüfung im UI vorhanden ist, einmal neu prüfen
        if not hasattr(self, "_last_docx_results"):
            sap_df = load_employees()
            self._last_docx_results = process_docx_folder(input_dir, sap_df)

        results = self._last_docx_results

        # 2) Exporte schreiben
        try:
            sap_df = load_employees()
            export_sap_massenupload(results, sap_df, sap_out)
            export_ds_csv(results, ds_out, sap_df)
        except Exception as e:
            messagebox.showerror("Export-Fehler", f"Export fehlgeschlagen:\n{e}")
            return

        # 3) Verschieben
        try:
            moved_ok, moved_man = move_after_processing(input_dir, results)
        except Exception as e:
            messagebox.showerror("Verschiebefehler", f"Verschieben fehlgeschlagen:\n{e}")
            return

        messagebox.showinfo(
            "Fertig",
            f"Export geschrieben:\n- SAP: {sap_out}\n- DS:  {ds_out}\n\n"
            f"Verschoben:\n- OK → archiv: {moved_ok}\n- manuell → manuell: {moved_man}"
        )

    def on_export_ds(self):
        if not hasattr(self, "_last_docx_results"):
            messagebox.showwarning("Hinweis", "Bitte zuerst DOCX prüfen.")
            return
        out_csv = Path("export") / "ds_export.csv"
        sap_df = load_employees()
        export_ds_csv(self._last_docx_results, out_csv, sap_df)
        messagebox.showinfo("Erfolg", f"DS-Export geschrieben: {out_csv}")

    def on_process_pdfs(self):
        in_dir = Path(__file__).parent.parent / "ruecklauf" / "unverarbeitet"
        out_root = Path(self.rpa_target_var.get())  # RPA Ziel

        if hasattr(self, "sap_df"):
            sap_df = self.sap_df
        else:
            from app.data_loader import load_employees
            sap_df = load_employees()
            self.sap_df = sap_df

        try:
            max_files = int(self.batch_size_var.get()) if self.batch_size_var.get() else None
        except Exception:
            max_files = None

        # process_pdfs verarbeitet alles im Ordner; einfache Batch-Variante: begrenze vorab
        all_pdfs = sorted(in_dir.glob("*.pdf"))
        if max_files is not None:
            process_list = all_pdfs[:max_files]
        else:
            process_list = all_pdfs

        # Temporärer Unterlauf: verschiebe selektiv in temp und rufe process_pdfs darauf auf
        temp_dir = in_dir / "_batch_tmp"
        temp_dir.mkdir(parents=True, exist_ok=True)
        moved = []
        for p in process_list:
            tgt = temp_dir / p.name
            try:
                p.rename(tgt)
                moved.append(tgt)
            except Exception:
                pass

        durchlauf_jahr = self.proc_year_var.get()
        results = process_pdfs(temp_dir, out_root, sap_df, durchlauf_jahr=durchlauf_jahr)

        # Treeview leeren
        for item in self.tree_pdfs.get_children():
            self.tree_pdfs.delete(item)

        # Ergebnisse einfüllen (konsistent mit DOCX-Anzeige)
        for r in results:
            self.tree_pdfs.insert("", "end", values=(
                r.get("file", ""),
                r.get("typ", ""),
                r.get("pn", ""),
                r.get("name", ""),
                r.get("status", ""),
                r.get("reason", ""),
                r.get("target", "")
            ))

        from tkinter import messagebox
        messagebox.showinfo("Fertig", f"{len(results)} PDFs verarbeitet.")
        # Aufräumen: restliche Dateien im temp zurück nach unverarbeitet (sollten keine sein)
        for p in temp_dir.glob("*.pdf"):
            try:
                p.rename(in_dir / p.name)
            except Exception:
                pass
        try:
            temp_dir.rmdir()
        except Exception:
            pass

    # ---------------------------
    # MD-DASHBOARD
    # ---------------------------
    def build_dashboard(self):
        """MD-Dashboard: Status-Übersicht und manuelle Anpassungen"""
        self.frame_dashboard.grid_columnconfigure(0, weight=1)
        self.frame_dashboard.grid_rowconfigure(3, weight=1)  # Haupt-Treeview
        
        # Toolbar
        toolbar = ttk.Frame(self.frame_dashboard)
        toolbar.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        
        ttk.Button(toolbar, text="Aktualisieren", command=self.on_refresh_dashboard).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Manuelle Anpassung", command=self.on_manual_adjustment).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Export CSV", command=self.on_export_dashboard).pack(side="left")

        # Info-Button
        create_info_button(
            parent=toolbar,
            title="Info • MD-Dashboard",
            text=(
                "MD-Dashboard\n"
                "- 'Aktualisieren' lädt die aktuellen Tracking-Daten.\n"
                "- Filter nach Manager-PN, Status und Jahr möglich.\n"
                "- 'Manuelle Anpassung' erlaubt Statuskorrekturen je Eintrag.\n"
                "- 'Export CSV' speichert die aktuell gefilterte Ansicht."
            ),
            side="right",
        )
        
        # Filter
        filter_frame = ttk.Frame(self.frame_dashboard)
        filter_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        
        ttk.Label(filter_frame, text="Filter:").pack(side="left", padx=(0, 8))
        
        # Manager Filter
        ttk.Label(filter_frame, text="Manager PN:").pack(side="left", padx=(0, 4))
        self.dash_mgr_filter = ttk.Entry(filter_frame, width=10)
        self.dash_mgr_filter.pack(side="left", padx=(0, 16))
        
        # Status Filter
        ttk.Label(filter_frame, text="Status:").pack(side="left", padx=(0, 4))
        self.dash_status_filter = ttk.Combobox(filter_frame, width=12, values=["", "ausstehend", "erhalten", "prüfung_nötig", "erübrigt"])
        self.dash_status_filter.pack(side="left", padx=(0, 16))
        
        # Jahr Filter
        ttk.Label(filter_frame, text="Jahr:").pack(side="left", padx=(0, 4))
        self.dash_year_filter = ttk.Entry(filter_frame, width=6)
        self.dash_year_filter.pack(side="left", padx=(0, 16))
        
        ttk.Button(filter_frame, text="Anwenden", command=self.on_apply_dashboard_filters).pack(side="left")
        
        # Status-Übersicht
        ttk.Label(self.frame_dashboard, text="Status-Übersicht:").grid(row=2, column=0, sticky="w", padx=8, pady=(0, 4))
        
        # Haupt-Treeview
        cols = ["Log-ID", "VG PN", "VG Name", "MA PN", "MA Name", "Dokument-Typ", "Erwartet", "Erhalten", "Status", "Status Grund", "Versendet am", "Zuletzt erinnert am"]
        self.tree_dashboard = ttk.Treeview(self.frame_dashboard, columns=cols, show="headings", height=15)
        
        # Status-Farben definieren (leichte Farben)
        self.tree_dashboard.tag_configure("status_ausstehend", background="#ffebee")    # Sehr helles Rot
        self.tree_dashboard.tag_configure("status_erhalten", background="#e8f5e8")     # Sehr helles Grün
        self.tree_dashboard.tag_configure("status_eruobrigt", background="#e8f5e8")    # Sehr helles Grün (wie erhalten)
        self.tree_dashboard.tag_configure("status_pruefung_noetig", background="#fff3e0")  # Sehr helles Orange
        
        for c in cols:
            self.tree_dashboard.heading(c, text=c)
            if c in ["Log-ID", "VG PN", "MA PN"]:
                self.tree_dashboard.column(c, width=100, anchor="w")
            elif c in ["Status", "Dokument-Typ"]:
                self.tree_dashboard.column(c, width=120, anchor="w")
            elif c in ["Erwartet", "Erhalten"]:
                self.tree_dashboard.column(c, width=80, anchor="w")
            elif c in ["Zuletzt erinnert am", "Versendet am"]:
                self.tree_dashboard.column(c, width=140, anchor="w")
            else:
                self.tree_dashboard.column(c, width=150, anchor="w")
        
        self.tree_dashboard.grid(row=3, column=0, sticky="nsew", padx=8, pady=(0, 8))
        
        # Scrollbar
        scrollbar_dash = ttk.Scrollbar(self.frame_dashboard, orient="vertical", command=self.tree_dashboard.yview)
        scrollbar_dash.grid(row=3, column=1, sticky="ns")
        self.tree_dashboard.configure(yscrollcommand=scrollbar_dash.set)
        
        # Initial load
        self.on_refresh_dashboard()
    
    def _get_status_tag(self, status: str) -> str:
        """Bestimmt den Farb-Tag basierend auf dem Status."""
        status_lower = str(status).lower().strip()
        
        if status_lower == "ausstehend":
            return "status_ausstehend"
        elif status_lower == "erhalten":
            return "status_erhalten"
        elif status_lower == "erübrigt":
            return "status_eruobrigt"
        elif status_lower == "prüfung_nötig":
            return "status_pruefung_noetig"
        else:
            return ""  # Keine Farbe für unbekannte Status
    
    def on_refresh_dashboard(self):
        """Lädt Dashboard-Daten neu"""
        # Treeview leeren
        for item in self.tree_dashboard.get_children():
            self.tree_dashboard.delete(item)
        
        try:
            # Filter anwenden
            mgr_filter = self.dash_mgr_filter.get().strip()
            status_filter = self.dash_status_filter.get().strip()
            year_filter = self.dash_year_filter.get().strip()
            
            year_int = int(year_filter) if year_filter.isdigit() else None
            
            # Daten laden
            df = self.tracking.get_dashboard_data(
                filter_mgr=mgr_filter,
                filter_status=status_filter
            )
            
            if df.empty:
                return
            
            # Daten in Treeview einfügen
            for _, row in df.iterrows():
                # Hilfsfunktion um NaN-Werte zu leeren Strings zu konvertieren
                def safe_value(val):
                    if pd.isna(val) or val == "nan" or val == "NaN":
                        return ""
                    return str(val) if val is not None else ""
                
                # Status für Farb-Tag bestimmen
                status = safe_value(row.get("status", ""))
                status_tag = self._get_status_tag(status)
                
                # Eintrag mit Farb-Tag einfügen
                item_id = self.tree_dashboard.insert("", "end", values=(
                    safe_value(row.get("log_id", "")),
                    safe_value(row.get("vg_pn", "")),
                    safe_value(row.get("vg_name", "")),
                    safe_value(row.get("ma_pn", "")),
                    safe_value(row.get("ma_name", "")),
                    safe_value(row.get("doc_type", "")),
                    safe_value(row.get("erwartet", "")),
                    safe_value(row.get("erhalten", "")),
                    safe_value(row.get("status", "")),
                    safe_value(row.get("status_grund", "")),
                    safe_value(row.get("versendet_am", "")),
                    safe_value(row.get("zuletzt_erinnert_am", ""))
                ), tags=(status_tag,) if status_tag else ())
                
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Dashboard-Daten: {e}")
    
    def on_apply_dashboard_filters(self):
        """Wendet Filter an"""
        self.on_refresh_dashboard()
    
    def on_manual_adjustment(self):
        """Öffnet Dialog für manuelle Anpassung"""
        selection = self.tree_dashboard.selection()
        if not selection:
            messagebox.showwarning("Warnung", "Bitte wählen Sie einen Eintrag aus.")
            return
        
        item = self.tree_dashboard.item(selection[0])
        values = item["values"]
        
        if len(values) < 6:
            messagebox.showerror("Fehler", "Ungültiger Eintrag ausgewählt.")
            return
        
        log_id = values[0]
        vg_pn = values[1]
        ma_pn = values[3]
        doc_type = values[5]
        
        # Dialog für manuelle Anpassung
        dialog = tk.Toplevel(self)
        dialog.title("Manuelle Anpassung")
        dialog.geometry("500x300")
        dialog.transient(self)
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"Log-ID: {log_id}").pack(pady=8)
        ttk.Label(dialog, text=f"VG PN: {vg_pn}, MA PN: {ma_pn}").pack(pady=4)
        ttk.Label(dialog, text=f"Dokument-Typ: {doc_type}").pack(pady=4)
        
        # Status-Auswahl
        ttk.Label(dialog, text="Neuer Status:").pack(pady=(16, 4))
        status_var = tk.StringVar(value="erhalten")
        status_combo = ttk.Combobox(dialog, textvariable=status_var, 
                                  values=["ausstehend", "erhalten", "prüfung_nötig", "erübrigt"])
        status_combo.pack(pady=4)
        
        # Grund
        ttk.Label(dialog, text="Status Grund:").pack(pady=(16, 4))
        reason_var = tk.StringVar()
        reason_combo = ttk.Combobox(dialog, textvariable=reason_var, width=40,
                                   values=["", "Grund_Prüfung (aus Verarbeitung)", "Krankheit/Unfall", 
                                          "anderer VG", "Austritt", "sonstiges"])
        reason_combo.pack(pady=4)
        
        def apply_adjustment():
            new_status = status_var.get()
            reason = reason_var.get().strip()
            
            try:
                if self.tracking.manual_status_update(vg_pn, ma_pn, doc_type, new_status, reason):
                    messagebox.showinfo("Erfolg", "Anpassung gespeichert.")
                    dialog.destroy()
                    self.on_refresh_dashboard()
                else:
                    messagebox.showerror("Fehler", "Anpassung fehlgeschlagen - Eintrag nicht gefunden.")
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler beim Speichern: {e}")
        
        ttk.Button(dialog, text="Anwenden", command=apply_adjustment).pack(pady=16)
        ttk.Button(dialog, text="Abbrechen", command=dialog.destroy).pack()
    
    def on_export_dashboard(self):
        """Exportiert Dashboard-Daten als CSV"""
        try:
            from tkinter import filedialog
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            if not filename:
                return
            
            # Aktuelle Filter anwenden
            mgr_filter = self.dash_mgr_filter.get().strip()
            status_filter = self.dash_status_filter.get().strip()
            year_filter = self.dash_year_filter.get().strip()
            
            year_int = int(year_filter) if year_filter.isdigit() else None
            
            df = self.tracking.get_dashboard_data(
                filter_mgr=mgr_filter,
                filter_status=status_filter,
                filter_year=year_int
            )
            
            df.to_csv(filename, sep=";", index=False, encoding="utf-8-sig")
            messagebox.showinfo("Erfolg", f"Dashboard-Daten exportiert nach: {filename}")
            
        except Exception as e:
            messagebox.showerror("Fehler", f"Export fehlgeschlagen: {e}")

    # ---------------------------
    # NEUE VERSAND-METHODEN
    # ---------------------------
    def _refresh_mgr_table(self):
        """Aktualisiert die VG-Tabelle im Massenversand Tab"""
        # Treeview leeren
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Filter anwenden
        filter_text = self.filter_var.get().lower() if hasattr(self, 'filter_var') else ""
        
        for vg_pn, pack in self.mgr_index.items():
            mgr = pack["manager"]
            if mgr is None:
                continue
                
            # Filter prüfen
            if filter_text:
                mgr_text = f"{mgr.get('Nachname','')} {mgr.get('Rufname','')} {mgr.get('OE Kurzb.','')}".lower()
                if filter_text not in mgr_text:
                    continue
            
            # Anzahl MA
            subs_count = len(pack["subs"])
            
            self.tree.insert("", "end", iid=vg_pn, values=[
                vg_pn,
                mgr.get("Nachname", ""),
                mgr.get("Rufname", ""),
                mgr.get("OE Kurzb.", ""),
                subs_count
            ])

    def _refresh_mgr_table_einzel(self):
        """Aktualisiert die VG-Tabelle im Einzelversand Tab"""
        # Treeview leeren
        for item in self.tree_einzel.get_children():
            self.tree_einzel.delete(item)
        
        # Filter anwenden
        filter_text = self.filter_var.get().lower() if hasattr(self, 'filter_var') else ""
        
        for vg_pn, pack in self.mgr_index.items():
            mgr = pack["manager"]
            if mgr is None:
                continue
                
            # Filter prüfen
            if filter_text:
                mgr_text = f"{mgr.get('Nachname','')} {mgr.get('Rufname','')} {mgr.get('OE Kurzb.','')}".lower()
                if filter_text not in mgr_text:
                    continue
            
            # Anzahl MA
            subs_count = len(pack["subs"])
            
            self.tree_einzel.insert("", "end", iid=vg_pn, values=[
                vg_pn,
                mgr.get("Nachname", ""),
                mgr.get("Rufname", ""),
                mgr.get("OE Kurzb.", ""),
                subs_count
            ])

    def _refresh_vg_list(self):
        """Aktualisiert die VG-Liste im VG-MA Tab"""
        # Treeview leeren
        for item in self.vg_tree.get_children():
            self.vg_tree.delete(item)
        
        # Filter anwenden
        filter_text = self.vg_search_var.get().lower()
        
        for vg_pn, pack in self.mgr_index.items():
            mgr = pack["manager"]
            if mgr is None:
                continue
                
            # Filter prüfen
            if filter_text:
                mgr_text = f"{mgr.get('Nachname','')} {mgr.get('Rufname','')} {mgr.get('OE Kurzb.','')}".lower()
                if filter_text not in mgr_text:
                    continue
            
            self.vg_tree.insert("", "end", iid=vg_pn, values=[
                vg_pn,
                mgr.get("Nachname", ""),
                mgr.get("Rufname", ""),
                mgr.get("OE Kurzb.", "")
            ])

    def _refresh_ma_list(self):
        """Aktualisiert die MA-Liste im VG-MA Tab"""
        # Treeview leeren
        for item in self.ma_tree.get_children():
            self.ma_tree.delete(item)
        
        # Filter anwenden
        filter_text = self.ma_search_var.get().lower()
        
        # Set für bereits hinzugefügte IDs
        added_ids = set()
        
        for idx, emp in self.df.iterrows():
            # Filter prüfen
            if filter_text:
                emp_text = f"{emp.get('Nachname','')} {emp.get('Rufname','')} {emp.get('OE Kurzb.','')}".lower()
                if filter_text not in emp_text:
                    continue
            
            # Eindeutige ID erstellen (Index + PN für Duplikate)
            emp_id = f"emp_{idx}"  # Verwende Index als ID, nicht PN
            if emp_id in added_ids:
                emp_id = f"{emp_id}_{idx}"  # Eindeutige ID für Duplikate
            
            added_ids.add(emp_id)
            
            self.ma_tree.insert("", "end", iid=emp_id, values=[
                str(emp.get("ID_NO_ZERO", "")),
                emp.get("Nachname", ""),
                emp.get("Rufname", ""),
                emp.get("OE Kurzb.", "")
            ])

    def _update_selection_status(self):
        """Aktualisiert den Auswahl-Status im VG-MA Tab"""
        vg_sel = self.vg_tree.selection()
        ma_sel = self.ma_tree.selection()
        
        if vg_sel and ma_sel:
            vg_item = self.vg_tree.item(vg_sel[0])
            ma_item = self.ma_tree.item(ma_sel[0])
            vg_name = f"{vg_item['values'][2]} {vg_item['values'][1]}"
            ma_name = f"{ma_item['values'][2]} {ma_item['values'][1]}"
            self.selection_status.config(
                text=f"Ausgewählt: VG {vg_name} ({vg_item['values'][0]}) ↔ MA {ma_name} ({ma_item['values'][0]})",
                foreground="black"
            )
        else:
            self.selection_status.config(text="Keine Auswahl", foreground="gray")

    def on_create_vg_ma_relationship(self):
        """Erstellt neue VG-MA-Beziehung in EXPORT.xlsx"""
        vg_sel = self.vg_tree.selection()
        ma_sel = self.ma_tree.selection()
        
        if not vg_sel or not ma_sel:
            messagebox.showwarning("Warnung", "Bitte wählen Sie sowohl einen Vorgesetzten als auch einen Mitarbeitenden aus.")
            return
        
        vg_item = self.vg_tree.item(vg_sel[0])
        ma_item = self.ma_tree.item(ma_sel[0])
        vg_pn = str(vg_item['values'][0])  # Explizit als String konvertieren
        ma_pn = str(ma_item['values'][0])   # Explizit als String konvertieren
        
        try:
            # Lade aktuelle EXPORT.xlsx mit gleicher Normalisierung wie load_employees()
            xlsx_path = Path(__file__).parent.parent / "sap_stammdaten" / "EXPORT.xlsx"
            df = pd.read_excel(xlsx_path)
            
            # Gleiche Datentyp-Konvertierung wie in load_employees()
            for col, t in {"ID_NO_ZERO": str, "Dir. Vorgesetzter (PN)": str, "Ans.": str}.items():
                if col in df.columns:
                    df[col] = df[col].astype(t)
            df.columns = [c.strip() for c in df.columns]
            
            # Personalnummern normalisieren (gleiche Logik wie in load_employees)
            if "ID_NO_ZERO" in df.columns:
                df["ID_NO_ZERO"] = df["ID_NO_ZERO"].astype(str).str.strip()
            if "Dir. Vorgesetzter (PN)" in df.columns:
                df["Dir. Vorgesetzter (PN)"] = df["Dir. Vorgesetzter (PN)"].astype(str).str.strip()
            
            # IDs mit führenden Nullen angleichen
            if "ID_NO_ZERO" in df.columns and "Dir. Vorgesetzter (PN)" in df.columns:
                max_len = max(
                    df["ID_NO_ZERO"].str.len().max(),
                    df["Dir. Vorgesetzter (PN)"].str.len().max()
                )
                df["ID_NO_ZERO"] = df["ID_NO_ZERO"].str.zfill(max_len)
                df["Dir. Vorgesetzter (PN)"] = df["Dir. Vorgesetzter (PN)"].str.zfill(max_len)
            
            # Finde MA-Zeile mit normalisierter PN
            ma_row = df[df["ID_NO_ZERO"].astype(str) == ma_pn]
            if ma_row.empty:
                # Debug: Zeige verfügbare PNs und vergleiche mit self.df
                available_pns = df["ID_NO_ZERO"].astype(str).tolist()[:10]  # Erste 10 PNs aus Excel
                loaded_pns = self.df["ID_NO_ZERO"].astype(str).tolist()[:10]  # Erste 10 PNs aus geladenem DataFrame
                
                # Prüfe ob PN in geladenem DataFrame existiert
                ma_in_loaded = self.df[self.df["ID_NO_ZERO"].astype(str) == ma_pn]
                
                debug_msg = f"Mitarbeiter mit PN {ma_pn} nicht in EXPORT.xlsx gefunden.\n\n"
                debug_msg += f"Verfügbare PNs in Excel (erste 10): {available_pns}\n"
                debug_msg += f"Verfügbare PNs in geladenem DF (erste 10): {loaded_pns}\n"
                debug_msg += f"Gesuchte PN: '{ma_pn}' (Typ: {type(ma_pn)})\n"
                debug_msg += f"PN in geladenem DF gefunden: {not ma_in_loaded.empty}\n"
                debug_msg += f"Excel-Datei Zeilen: {len(df)}, Geladener DF Zeilen: {len(self.df)}"
                
                messagebox.showerror("Fehler", debug_msg)
                return
            
            # Erstelle Kopie der MA-Zeile
            new_row = ma_row.iloc[0].copy()
            new_row["Dir. Vorgesetzter (PN)"] = vg_pn
            
            # Füge neue Zeile hinzu
            df = pd.concat([df, new_row.to_frame().T], ignore_index=True)
            
            # Speichere zurück
            df.to_excel(xlsx_path, index=False)
            
            messagebox.showinfo("Erfolg", 
                f"Neue VG-MA-Beziehung erstellt:\n"
                f"VG: {vg_item['values'][2]} {vg_item['values'][1]} ({vg_pn})\n"
                f"MA: {ma_item['values'][2]} {ma_item['values'][1]} ({ma_pn})\n\n"
                f"Bitte starten Sie die App neu, um die Änderungen zu laden."
            )
            
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Erstellen der Beziehung: {e}")


    # ---------------------------
    # VORSCHAU (NEU)
    # ---------------------------
    def _render_mail_preview(self, subject_tpl: str, body_tpl: str, anrede: str, rb_year: int, ab_year: int):
        try:
            subject = (subject_tpl or "").format(rb_year=rb_year, ab_year=ab_year)
        except Exception:
            subject = subject_tpl or ""
        try:
            body = (body_tpl or "").format(anrede=anrede, rb_year=rb_year, ab_year=ab_year)
        except Exception:
            body = body_tpl or ""

        html = (
            "<html><head><meta charset=\"utf-8\">"
            "<style>"
            "body{font-family:Arial, Helvetica, sans-serif;font-size:14px;line-height:1.5;color:#222;margin:16px;}"
            "h3{margin:0 0 8px 0;font-weight:600;}"
            "hr{border:none;border-top:1px solid #ccc;margin:12px 0;}"
            "</style></head><body>"
            f"<h3>Betreff: {subject}</h3><hr/>"
            f"{body}"
            "</body></html>"
        )

        with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as f:
            f.write(html)
            temp_path = f.name
        webbrowser.open(Path(temp_path).as_uri())

    def on_preview_managers(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Hinweis", "Bitte mindestens eine/n Vorgesetzte/n auswählen.")
            return
        vg_pn = sel[0]
        pack = self.mgr_index.get(vg_pn)
        if not pack or pack.get("manager") is None:
            messagebox.showerror("Fehler", "Kein gültiger Vorgesetzten-Datensatz gefunden.")
            return

        mgr = pack["manager"]
        vg_vorname = str(mgr.get("Rufname", "")).strip()
        anrede = f"Hallo {vg_vorname}" if vg_vorname else "Hallo"

        rb_year = self.rb_year_var.get()
        ab_year = self.ab_year_var.get()

        subject_tpl = CFG.get("mail", {}).get("subject_template", "MD-Unterlagen Durchlauf {rb_year}/{ab_year}")
        body_tpl = CFG.get("mail", {}).get("body_html_template", "")

        self._render_mail_preview(subject_tpl, body_tpl, anrede, rb_year, ab_year)

    def on_preview_selected(self):
        sel_mgrs = self.tree_einzel.selection()
        if len(sel_mgrs) != 1:
            messagebox.showwarning("Hinweis", "Bitte genau eine/n Vorgesetzte/n auswählen.")
            return
        vg_pn = sel_mgrs[0]
        pack = self.mgr_index.get(vg_pn)
        if not pack or pack.get("manager") is None:
            messagebox.showerror("Fehler", "Kein gültiger Vorgesetzten-Datensatz gefunden.")
            return

        mgr = pack["manager"]
        vg_vorname = str(mgr.get("Rufname", "")).strip()
        anrede = f"Hallo {vg_vorname}" if vg_vorname else "Hallo"

        rb_year = self.rb_year_var_einzel.get()
        ab_year = self.ab_year_var_einzel.get()

        subject_tpl = CFG.get("mail_underjaehrig", {}).get("subject_template", "MD-Unterlagen")
        body_tpl = CFG.get("mail_underjaehrig", {}).get("body_html_template", "")

        self._render_mail_preview(subject_tpl, body_tpl, anrede, rb_year, ab_year)

if __name__ == "__main__":
    App().mainloop()