# app/main.py
"""
MD-Prozess-Tool - Hauptanwendung

Diese Anwendung verwaltet den Mitarbeitenden-Dialog (MD) Prozess:
- SAP Stammdaten pruefen und validieren
- MD-Dokumente generieren und versenden
- Ruecklauf verarbeiten und tracken
- Dashboard fuer Status-Uebersicht

Autor: VD GS HR
Version: 1.0
"""
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
import pandas as pd
from typing import Dict, Any, List, Optional, Tuple

# Imports - kompatibel mit direktem Aufruf und Modul-Import
import sys
import os

# Füge das Projektverzeichnis zum Python-Pfad hinzu (nur bei direktem Aufruf)
if __name__ == "__main__":
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, project_root)
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Versuche zuerst absolute Imports, dann relative
try:
    from data_loader import load_employees, load_config, build_manager_index, validate_config
    from constants import ProcStatus, DashTag, MDConstants
    from logging_config import setup_logging
    from utils import create_info_button
    from services.tracking_service import SimpleTrackingSystem
except ImportError:
    # Fallback für Modul-Import
    from .data_loader import load_employees, load_config, build_manager_index, validate_config
    from .constants import ProcStatus, DashTag, MDConstants
    from .logging_config import setup_logging
    from .utils import create_info_button
    from .services.tracking_service import SimpleTrackingSystem

CFG = load_config()
# Logging initialisieren: Konsole + Logdatei unter tracking/app.log
try:
    from pathlib import Path as _Path
    tracking_dir_rel = (CFG.get("paths", {}) or {}).get("tracking_dir", "../tracking")
    log_dir = (_Path(__file__).parent / tracking_dir_rel).resolve()
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = str((log_dir / "app.log").resolve())
except Exception:
    log_file = None
setup_logging(log_level="INFO", log_file=log_file)
try:
    validate_config(CFG)
except Exception as e:
    # Frühzeitige Nutzerinfo; App startet trotzdem, UI zeigt Fehler bei Nutzung
    try:
        from tkinter import messagebox
        messagebox.showwarning("Konfiguration", f"Hinweis: Konfiguration unvollständig/Prüfung: {e}")
    except Exception:
        pass

# Zentrale Konstanten zur Eliminierung von Magic Numbers/Strings
# Konstanten werden aus .constants importiert (MDConstants)

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
    
    # Type Hints für wichtige Attribute
    jahr_var: tk.IntVar
    tracking: SimpleTrackingSystem
    df: pd.DataFrame
    mgr_index: Dict[str, Dict[str, Any]]
    
    # GUI-Komponenten
    notebook: ttk.Notebook
    frame_stammdaten: ttk.Frame
    frame_versand: ttk.Frame
    frame_ruecklauf: ttk.Frame
    frame_verarbeitung: ttk.Frame
    frame_dashboard: ttk.Frame
    
    # Versand-spezifische Komponenten
    versand_notebook: ttk.Notebook
    frame_massenversand: ttk.Frame
    frame_einzelversand: ttk.Frame
    frame_vg_ma: ttk.Frame
    
    # Treeview-Komponenten
    tree: ttk.Treeview
    tree_einzel: ttk.Treeview
    tree_checks: ttk.Treeview
    tree_findings: ttk.Treeview
    tree_ok: ttk.Treeview
    tree_pruefen: ttk.Treeview
    tree_skip: ttk.Treeview
    tree_proc: ttk.Treeview
    tree_pdfs: ttk.Treeview
    tree_dashboard: ttk.Treeview
    vg_tree: ttk.Treeview
    ma_tree: ttk.Treeview
    subs_tree: ttk.Treeview
    
    # Status-Labels
    lbl_fileinfo: ttk.Label
    inbox_status: ttk.Label
    ruecklauf_status: ttk.Label
    proc_status: ttk.Label
    ms_status: ttk.Label
    es_status: ttk.Label
    selection_status: ttk.Label
    
    # Progress Bars
    ms_progress: ttk.Progressbar
    es_progress: ttk.Progressbar
    
    # String Variables
    filter_var: tk.StringVar
    inbox_target_var: tk.StringVar
    rb_year_var: tk.IntVar
    ab_year_var: tk.IntVar
    rb_year_var_einzel: tk.IntVar
    ab_year_var_einzel: tk.IntVar
    proc_year_var: tk.IntVar
    rpa_target_var: tk.StringVar
    batch_size_var: tk.IntVar
    vg_search_var: tk.StringVar
    ma_search_var: tk.StringVar
    dash_name_search: ttk.Entry
    dash_status_filter: ttk.Combobox
    
    # Checkbox Variables
    var_rb: tk.BooleanVar
    var_ab: tk.BooleanVar
    var_pz: tk.BooleanVar
    
    # Interne Daten
    _last_docx_results: List[Dict[str, Any]]
    sap_df: pd.DataFrame
    
    def __init__(self) -> None:
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
        self._build_tabs()

    def _bind_treeview_sort(self, tree, numeric_like=None):
        """Bindet klickbare Sortierung an alle Spalten eines Treeviews.
        
        Diese Methode wird von den Views verwendet und delegiert an ui_utils.
        """
        try:
            from views.ui_utils import bind_treeview_sort
        except ImportError:
            from .views.ui_utils import bind_treeview_sort
        bind_treeview_sort(tree, numeric_like)

    def _build_tabs(self) -> None:
        """Baut alle Tabs mit Fehlerbehandlung auf."""
        try:
            try:
                from views.stammdaten_view import build_stammdaten
            except ImportError:
                from .views.stammdaten_view import build_stammdaten
            build_stammdaten(self.frame_stammdaten, self)
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror(MDConstants.MSG_ERROR, f"Stammdaten-Tab konnte nicht geladen werden: {e}")
        
        try:
            try:
                from views.versand_view import build_versand
            except ImportError:
                from .views.versand_view import build_versand
            build_versand(self.frame_versand, self)
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror(MDConstants.MSG_ERROR, f"Versand-Tab konnte nicht geladen werden: {e}")
        
        try:
            try:
                from views.ruecklauf_view import build_ruecklauf
            except ImportError:
                from .views.ruecklauf_view import build_ruecklauf
            build_ruecklauf(self.frame_ruecklauf, self)
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror(MDConstants.MSG_ERROR, f"Rücklauf-Tab konnte nicht geladen werden: {e}")
        
        try:
            try:
                from views.verarbeitung_view import build_verarbeitung
            except ImportError:
                from .views.verarbeitung_view import build_verarbeitung
            build_verarbeitung(self.frame_verarbeitung, self)
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror(MDConstants.MSG_ERROR, f"Verarbeitung-Tab konnte nicht geladen werden: {e}")
        
        try:
            try:
                from views.dashboard_view import build_dashboard
            except ImportError:
                from .views.dashboard_view import build_dashboard
            build_dashboard(self.frame_dashboard, self)
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror(MDConstants.MSG_ERROR, f"Dashboard-Tab konnte nicht geladen werden: {e}")

    def _default_proc_year(self):
        """Bestimmt das Standard-Verarbeitungsjahr basierend auf dem aktuellen Monat.
        
        Bis inkl. April -> Vorjahr, ab Mai -> aktuelles Jahr
        """
        from datetime import date
        current_month = date.today().month
        current_year = date.today().year
        
        if current_month <= MDConstants.PROC_MONTH_THRESHOLD:
            return current_year - 1  # Vorjahr
        else:
            return current_year  # Aktuelles Jahr


if __name__ == "__main__":
    App().mainloop()