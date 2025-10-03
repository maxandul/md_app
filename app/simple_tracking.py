# app/simple_tracking.py
"""
Vereinfachtes Tracking-System für MD-Prozess:
- Eine einzige CSV-Datei mit VG-zentrierter Sicht
- Pro VG: Zeilen für jeden MA (Rückblick/Ausblick Word+PDF) + eine Feedback-Zeile
- Status-Tracking: ausstehend → erhalten → prüfung_nötig/erübrigt
"""
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional
import csv

try:
    from .data_loader import load_config
except ImportError:
    from data_loader import load_config

CFG = load_config()

class SimpleTrackingSystem:
    def __init__(self):
        self.log_path = Path(__file__).parent.parent / "tracking" / "md_logging.csv"
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
    
    def log_versand(self, mgr_pn: str, mgr_name: str, emp_pn: str, emp_name: str, 
                   doc_types: List[str], rb_year: int, ab_year: int, include_feedback: bool = False) -> None:
        """
        Loggt einen Versand für einen VG und seine MA.
        Erstellt Zeilen für:
        - Pro MA: Rückblick Word, Rückblick PDF, Ausblick Word, Ausblick PDF (falls in doc_types)
        - Pro VG: Feedback PDF (nur wenn include_feedback=True)
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        entries = []
        
        # Pro MA: Rückblick und Ausblick (Word + PDF)
        for doc_type in doc_types:
            if doc_type == "rueckblick":
                entries.extend([
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, 
                                     "Rückblick Word", 1, 0, "ausstehend", "", timestamp),
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, 
                                     "Rückblick PDF", 1, 0, "ausstehend", "", timestamp)
                ])
            elif doc_type == "ausblick":
                entries.extend([
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, 
                                     "Ausblick Word", 1, 0, "ausstehend", "", timestamp),
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, 
                                     "Ausblick PDF", 1, 0, "ausstehend", "", timestamp)
                ])
        
        # Pro VG: Feedback PDF (nur wenn explizit angefordert)
        if include_feedback:
            # Prüfe ob Feedback für diesen VG bereits erwartet wird
            feedback_exists = self._has_feedback_for_manager(mgr_pn, rb_year)
            if not feedback_exists:
                # Zähle direkt unterstellte MA für erwartete Anzahl
                expected_feedback = self._count_direct_reports(mgr_pn)
                entries.append(
                    self._create_entry(mgr_pn, mgr_name, "", "", 
                                     "Feedback PDF", expected_feedback, 0, "ausstehend", "", timestamp)
                )
        
        # Schreibe alle Einträge
        for entry in entries:
            self._append_to_csv(entry)
    
    def log_feedback_for_manager(self, mgr_pn: str, mgr_name: str, rb_year: int, managers_index: dict = None) -> None:
        """
        Erstellt einen Feedback-Eintrag für einen Vorgesetzten.
        Sollte einmal pro Vorgesetzten aufgerufen werden, nicht pro Mitarbeiter.
        """
        # Prüfe ob Feedback für diesen VG bereits erwartet wird
        feedback_exists = self._has_feedback_for_manager(mgr_pn, rb_year)
        if feedback_exists:
            return  # Bereits vorhanden
        
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Zähle direkt unterstellte MA für erwartete Anzahl
        expected_feedback = self._count_direct_reports(mgr_pn, managers_index)
        
        entry = self._create_entry(mgr_pn, mgr_name, "", "", 
                                 "Feedback PDF", expected_feedback, 0, "ausstehend", "", timestamp)
        
        self._append_to_csv(entry)
    
    def mark_received(self, mgr_pn: str, emp_pn: str, doc_type: str) -> bool:
        """
        Markiert ein Dokument als empfangen.
        Für Feedback: nur mgr_pn relevant, emp_pn kann leer sein.
        """
        df = self._load_log()
        if df.empty:
            return False
        
        # Finde passende Zeile - konvertiere alle zu String für Vergleich
        if doc_type == "Feedback PDF":
            mask = (df["vg_pn"].astype(str) == str(mgr_pn)) & (df["ma_pn"].isna()) & (df["doc_type"] == "Feedback PDF")
        else:
            # Konvertiere ma_pn zu String und entferne .0, vergleiche mit emp_pn
            ma_pn_str = df["ma_pn"].astype(str).str.replace('.0', '', regex=False)
            mask = (df["vg_pn"].astype(str) == str(mgr_pn)) & (ma_pn_str == str(emp_pn)) & (df["doc_type"] == doc_type)
        
        if not mask.any():
            return False
        
        # Aktualisiere Zähler/Status
        df.loc[mask, "erhalten"] = df.loc[mask, "erhalten"].fillna(0).astype(int) + 1
        if doc_type == "Feedback PDF":
            # Erst als "erhalten" markieren, wenn erhalten >= erwartet
            erwartet = df.loc[mask, "erwartet"].fillna(0).astype(int)
            erhalten = df.loc[mask, "erhalten"].astype(int)
            fully_received = erhalten >= erwartet
            # Setze Status nur für die betroffene(n) Zeile(n)
            df.loc[mask & fully_received, "status"] = "erhalten"
            # Teilweise eingegangen -> Status bleibt (typisch: ausstehend)
        else:
            df.loc[mask, "status"] = "erhalten"
        
        # Speichere zurück
        self._save_log(df)
        return True
    
    def mark_error(self, mgr_pn: str, emp_pn: str, doc_type: str, error_reason: str) -> bool:
        """
        Markiert ein Dokument als fehlerhaft (prüfung_nötig).
        """
        df = self._load_log()
        if df.empty:
            return False
        
        # Finde passende Zeile - konvertiere alle zu String für Vergleich
        if doc_type == "Feedback PDF":
            mask = (df["vg_pn"].astype(str) == str(mgr_pn)) & (df["ma_pn"].isna()) & (df["doc_type"] == "Feedback PDF")
        else:
            # Konvertiere ma_pn zu String und entferne .0, vergleiche mit emp_pn
            ma_pn_str = df["ma_pn"].astype(str).str.replace('.0', '', regex=False)
            mask = (df["vg_pn"].astype(str) == str(mgr_pn)) & (ma_pn_str == str(emp_pn)) & (df["doc_type"] == doc_type)
        
        if not mask.any():
            return False
        
        # Aktualisiere Status
        df.loc[mask, "status"] = "prüfung_nötig"
        df.loc[mask, "status_grund"] = f"Grund_Prüfung (aus Verarbeitung): {error_reason}"
        
        # Speichere zurück
        self._save_log(df)
        return True
    
    def manual_status_update(self, mgr_pn: str, emp_pn: str, doc_type: str, 
                           new_status: str, reason: str = "") -> bool:
        """
        Manuelle Status-Änderung (z.B. erübrigt).
        """
        df = self._load_log()
        if df.empty:
            return False
        
        # Finde passende Zeile - konvertiere alle zu String für Vergleich
        if doc_type == "Feedback PDF":
            mask = (df["vg_pn"].astype(str) == str(mgr_pn)) & (df["ma_pn"].isna()) & (df["doc_type"] == "Feedback PDF")
        else:
            # Konvertiere ma_pn zu String und entferne .0, vergleiche mit emp_pn
            ma_pn_str = df["ma_pn"].astype(str).str.replace('.0', '', regex=False)
            mask = (df["vg_pn"].astype(str) == str(mgr_pn)) & (ma_pn_str == str(emp_pn)) & (df["doc_type"] == doc_type)
        
        if not mask.any():
            return False
        
        # Aktualisiere Status
        df.loc[mask, "status"] = new_status
        if reason:
            df.loc[mask, "status_grund"] = reason
        
        # Speichere zurück
        self._save_log(df)
        return True
    
    def update_entry(self, log_id: str, updates: dict) -> bool:
        """
        Aktualisiert beliebige Spalten eines Eintrags über die Log-ID.
        
        Args:
            log_id: Eindeutige Log-ID des Eintrags
            updates: Dictionary mit Spaltenname -> neuer Wert
            
        Returns:
            True wenn erfolgreich aktualisiert
        """
        df = self._load_log()
        if df.empty:
            return False
        
        # Finde Eintrag über Log-ID
        mask = df["log_id"] == log_id
        if not mask.any():
            return False
        
        # Aktualisiere alle angegebenen Spalten
        for column, value in updates.items():
            if column in df.columns:
                # Explizite dtype-Konvertierung um Future Warning zu vermeiden
                if column in ["erwartet", "erhalten"]:
                    # Numerische Spalten
                    df.loc[mask, column] = int(value) if str(value).isdigit() else value
                else:
                    # String-Spalten
                    df.loc[mask, column] = str(value)
        
        # Speichere zurück
        self._save_log(df)
        return True
    
    def get_dashboard_data(self, filter_mgr: str = "", filter_status: str = "", filter_year: int = None) -> pd.DataFrame:
        """
        Lädt Dashboard-Daten mit optionalen Filtern.
        """
        df = self._load_log()
        
        if filter_mgr:
            df = df[df["vg_pn"].str.contains(filter_mgr, na=False)]
        
        if filter_status:
            df = df[df["status"] == filter_status]
        
        if filter_year:
            df = df[df["versendet_am"].str.contains(str(filter_year), na=False)]
        
        # NaN-Werte durch leere Strings ersetzen für bessere Übersichtlichkeit
        df = df.fillna("")
        
        return df
    
    def _create_entry(self, mgr_pn: str, mgr_name: str, emp_pn: str, emp_name: str,
                     doc_type: str, erwartet: int, erhalten: int, status: str, 
                     status_grund: str, timestamp: str) -> Dict:
        """Erstellt einen Log-Eintrag."""
        return {
            "log_id": f"L{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{mgr_pn}_{emp_pn}_{doc_type.replace(' ', '_')}",
            "vg_pn": str(mgr_pn),
            "vg_name": str(mgr_name),
            "ma_pn": str(emp_pn),
            "ma_name": str(emp_name),
            "doc_type": str(doc_type),
            "erwartet": int(erwartet),
            "erhalten": int(erhalten),
            "status": str(status),
            "status_grund": str(status_grund),
            "versendet_am": str(timestamp),
            "zuletzt_erinnert_am": ""
        }
    
    def _has_feedback_for_manager(self, mgr_pn: str, rb_year: int) -> bool:
        """Prüft ob bereits Feedback für diesen VG erwartet wird."""
        df = self._load_log()
        if df.empty:
            return False
        
        return len(df[(df["vg_pn"] == mgr_pn) & 
                     (df["doc_type"] == "Feedback PDF") & 
                     (df["versendet_am"].str.contains(str(rb_year), na=False))]) > 0
    
    def _count_direct_reports(self, mgr_pn: str, managers_index: dict = None) -> int:
        """Zählt direkt unterstellte MA für Feedback-Erwartung."""
        if managers_index and mgr_pn in managers_index:
            subs_df = managers_index[mgr_pn]["subs"]
            return len(subs_df)
        # Fallback: Dummy-Wert wenn kein Index verfügbar
        return 3
    
    def _load_log(self) -> pd.DataFrame:
        """Lädt die Log-CSV."""
        if not self.log_path.exists():
            return pd.DataFrame()
        
        df = pd.read_csv(self.log_path, sep=";")
        
        # Explizite dtype-Konvertierung um Future Warnings zu vermeiden
        if not df.empty:
            # Numerische Spalten als int konvertieren
            for col in ["erwartet", "erhalten"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            
            # String-Spalten explizit als string
            for col in ["vg_pn", "vg_name", "ma_pn", "ma_name", "doc_type", "status", "status_grund", "versendet_am", "zuletzt_erinnert_am"]:
                if col in df.columns:
                    df[col] = df[col].astype(str)
        
        return df
    
    def _save_log(self, df: pd.DataFrame) -> None:
        """Speichert die Log-CSV."""
        df.to_csv(self.log_path, sep=";", index=False, encoding="utf-8-sig")
    
    def check_duplicate(self, filename: str, pn: str, doc_type: str = None, vg_pn: str = None) -> tuple[bool, str]:
        """
        Prüft ob ein Dokument bereits als "erhalten" verarbeitet wurde (Duplikat-Check).
        Eindeutige Identifikation über: VG-PN + MA-PN + Dokumenttyp
        Nur Dokumente mit Status "erhalten" sind Duplikate.
        Status "ausstehend" bedeutet: erwartetes Dokument, das verarbeitet werden darf.
        
        Args:
            filename: Name der Datei
            pn: Personalnummer des Mitarbeiters
            doc_type: Dokumenttyp ("Rückblick Word", "Rückblick PDF", "Ausblick Word", "Ausblick PDF")
            vg_pn: Personalnummer des Vorgesetzten (wichtig bei mehreren Anstellungen)
            
        Returns:
            (is_duplicate: bool, warning: str)
        """
        df = self._load_log()
        if df.empty:
            return False, ""
        
        # Konvertiere PNs zu String und entferne .0
        ma_pn_str = df["ma_pn"].astype(str).str.replace('.0', '', regex=False)
        vg_pn_str = df["vg_pn"].astype(str).str.replace('.0', '', regex=False)
        
        # Prüfe spezifischen Dokumenttyp falls angegeben
        if doc_type and vg_pn:
            # Eindeutige Identifikation über VG-PN + MA-PN + Dokumenttyp + Status
            duplicate_mask = (vg_pn_str == str(vg_pn)) & (ma_pn_str == str(pn)) & (df["doc_type"] == doc_type) & (df["status"] == "erhalten")
            if duplicate_mask.any():
                return True, f"VG {vg_pn} + MA {pn} + {doc_type} bereits als 'erhalten' verarbeitet"
        elif doc_type:
            # Fallback ohne VG-PN: nur MA-PN + Dokumenttyp
            duplicate_mask = (ma_pn_str == str(pn)) & (df["doc_type"] == doc_type) & (df["status"] == "erhalten")
            if duplicate_mask.any():
                return True, f"MA {pn} + {doc_type} bereits als 'erhalten' verarbeitet"
        else:
            # Fallback: Prüfe ob überhaupt schon ein "erhalten" Eintrag für diese PN existiert
            duplicate_mask = (ma_pn_str == str(pn)) & (df["status"] == "erhalten")
            if duplicate_mask.any():
                duplicate_entries = df[duplicate_mask]
                doc_types = duplicate_entries["doc_type"].unique().tolist()
                return True, f"MA {pn} bereits als 'erhalten' verarbeitet: {', '.join(doc_types)}"
        
        return False, ""

    def _append_to_csv(self, entry: Dict) -> None:
        """Fügt einen Eintrag zur CSV hinzu."""
        file_exists = self.log_path.exists()
        
        with open(self.log_path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=entry.keys(), delimiter=";")
            if not file_exists:
                writer.writeheader()
            writer.writerow(entry)
