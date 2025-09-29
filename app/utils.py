# app/utils.py
"""
Hilfsfunktionen für das MD-Prozess-Tool

Dieses Modul enthält:
- Datums- und Zeitfunktionen
- Dateinamen-Generierung
- GUI-Hilfsfunktionen (Info-Dialoge)
- Validierungsfunktionen

Autor: HR-Team
Version: 1.0
"""
from datetime import date
from calendar import monthrange
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

def date_in_range(d, start: date, end: date) -> bool:
    """
    Prüft ob ein Datum innerhalb eines Bereichs liegt.
    
    Args:
        d: Zu prüfendes Datum (kann pandas-Timestamp oder None sein)
        start: Startdatum des Bereichs
        end: Enddatum des Bereichs
        
    Returns:
        True wenn Datum im Bereich liegt, False sonst
    """
    if pd.isna(d):
        return False
    ts = pd.to_datetime(d, errors="coerce")
    if pd.isna(ts):
        return False
    dd = ts.date()
    return start <= dd <= end

def last_day_of_month(y: int, m: int) -> int:
    """
    Gibt den letzten Tag eines Monats zurück.
    
    Args:
        y: Jahr
        m: Monat (1-12)
        
    Returns:
        Anzahl Tage im Monat
    """
    return monthrange(y, m)[1]

def fixed_filename(typ: str, jahr: int | None, nachname_ma: str, vorname_ma: str, pn_ma: str,
                   nachname_vg: str | None = None, vorname_vg: str | None = None, pn_vg: str | None = None) -> str:
    """
    Erstellt standardisierte Dateinamen für MD-Dokumente.
    
    Format: Typ_Jahr_Nachname_Vorname_PN.docx
    Umlaute bleiben erhalten, Leerzeichen werden zu Unterstrichen.
    
    Args:
        typ: Dokumenttyp (Ausblick, Rückblick, Rückblick_Probezeit, Feedback)
        jahr: Jahr (nur für Ausblick/Rückblick)
        nachname_ma: Nachname des Mitarbeiters
        vorname_ma: Vorname des Mitarbeiters
        pn_ma: Personalnummer des Mitarbeiters
        nachname_vg: Nachname des Vorgesetzten (nur für Feedback)
        vorname_vg: Vorname des Vorgesetzten (nur für Feedback)
        pn_vg: Personalnummer des Vorgesetzten (nur für Feedback)
        
    Returns:
        Standardisierter Dateiname ohne Erweiterung
        
    Raises:
        ValueError: Bei unbekanntem Dokumenttyp
    """
    def clean(s: str) -> str:
        """Bereinigt String für Dateinamen (Leerzeichen -> _)."""
        return (s or "").strip().replace(" ", "_")

    if typ == "Ausblick":
        return f"Ausblick_{jahr}_{clean(nachname_ma)}_{clean(vorname_ma)}_{pn_ma}"
    if typ == "Rückblick":
        return f"Rückblick_{jahr}_{clean(nachname_ma)}_{clean(vorname_ma)}_{pn_ma}"
    if typ == "Rückblick_Probezeit":
        return f"Rückblick_Probezeit_{clean(nachname_ma)}_{clean(vorname_ma)}_{pn_ma}"
    if typ == "Feedback":
        return f"Feedback_{clean(nachname_vg)}_{clean(vorname_vg)}_{pn_vg}"
    raise ValueError(f"Unbekannter Typ: {typ}")


def show_info_dialog(parent: tk.Widget, title: str, text: str, width: int = 520, height: int = 360) -> None:
    """Zeigt einen scrollbaren Info-Dialog mit sauberem Zeilenumbruch.

    Verwendet ein Toplevel-Fenster mit Text + Scrollbar. Text wird im Word-Wrap dargestellt.
    """
    dlg = tk.Toplevel(parent)
    dlg.title(title or "Info")
    dlg.transient(parent)
    dlg.grab_set()
    try:
        # Position in der Nähe des Parent-Fensters
        x = parent.winfo_rootx() + 80
        y = parent.winfo_rooty() + 80
        dlg.geometry(f"{width}x{height}+{x}+{y}")
    except Exception:
        dlg.geometry(f"{width}x{height}")

    dlg.columnconfigure(0, weight=1)
    dlg.rowconfigure(0, weight=1)

    frame = ttk.Frame(dlg)
    frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    txt = tk.Text(frame, wrap="word", relief="flat")
    txt.grid(row=0, column=0, sticky="nsew")
    yscroll = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
    yscroll.grid(row=0, column=1, sticky="ns")
    txt.configure(yscrollcommand=yscroll.set)

    # Einheitliche Schrift (ähnlich Standard-UI)
    try:
        txt.configure(font=("Segoe UI", 10))
    except Exception:
        pass

    # Inhalt setzen
    txt.insert("1.0", text or "")
    txt.configure(state="disabled")

    btn_bar = ttk.Frame(dlg)
    btn_bar.grid(row=1, column=0, sticky="ew", padx=10, pady=(0,10))
    btn_bar.columnconfigure(0, weight=1)
    ok_btn = ttk.Button(btn_bar, text="OK", command=dlg.destroy)
    ok_btn.grid(row=0, column=0, sticky="e")


def create_info_button(parent: tk.Widget, text: str, title: str = "Hinweis", side: str = "left") -> ttk.Button:
    """Erstellt einen einheitlichen Info-Button, der eine messagebox mit Text zeigt.

    - Platzierung: pack(side=<side>)
    - Beschriftung: "ℹ Info"
    - Einheitliche Nutzung in allen Tabs
    """
    def _show():
        show_info_dialog(parent, title=title, text=text)
    btn = ttk.Button(parent, text="ℹ Info", command=_show)
    btn.pack(side=side)
    return btn
