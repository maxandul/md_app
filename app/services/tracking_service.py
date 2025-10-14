from __future__ import annotations

import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List
import csv

from app.data_loader import load_config

CFG = load_config()


class SimpleTrackingSystem:
    def __init__(self, jahr: int = None):
        """Initialisiert das Tracking-System für ein spezifisches MD-Durchlaufjahr.
        
        Args:
            jahr: MD-Durchlaufjahr (Rückblick-Jahr). Wenn None, wird es intelligent ermittelt:
                  - Oktober-Dezember: aktuelles Jahr
                  - Januar-April: Vorjahr (Nachläufer-Phase)
                  - Mai-September: aktuelles Jahr
        """
        tracking_dir = (CFG.get("paths", {}) or {}).get("tracking_dir", "../tracking")
        # Korrektur: Von services/ aus 2 Ebenen hoch zur Root, dann tracking_dir relativ auflösen
        base_dir = Path(__file__).parent.parent.parent  # Von services/ -> app/ -> md_app/
        tracking_path = (base_dir / tracking_dir.lstrip("../")).resolve()
        
        # Jahr ermitteln wenn nicht angegeben
        if jahr is None:
            jahr = self._detect_jahr()
        
        self.jahr = jahr
        
        # Jahr-spezifische Tracking-Datei
        self.log_path = tracking_path / f"md_logging_{jahr}.csv"
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
    
    def _detect_jahr(self) -> int:
        """Ermittelt das aktive MD-Durchlaufjahr basierend auf dem aktuellen Datum.
        
        Logik:
        - Oktober-Dezember: aktuelles Jahr (MD-Start)
        - Januar-April: Vorjahr (Nachläufer-Phase)
        - Mai-September: aktuelles Jahr (Vorbereitung/unterjährig)
        """
        heute = datetime.now()
        monat = heute.month
        jahr = heute.year
        
        if 10 <= monat <= 12:  # Okt-Dez: Aktuelles Jahr
            return jahr
        elif 1 <= monat <= 4:  # Jan-Apr: Vorjahr (Nachläufer)
            return jahr - 1
        else:  # Mai-Sep: Aktuelles Jahr
            return jahr

    # ... Implementation identisch zur bisherigen simple_tracking.SimpleTrackingSystem ...
    # Aus Platzgründen unverändert übernommen
    def log_versand(self, mgr_pn: str, mgr_name: str, emp_pn: str, emp_name: str,
                    doc_types: List[str], rb_year: int, ab_year: int, include_feedback: bool = False,
                    oe_bez_kette: str = "") -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entries = []
        for doc_type in doc_types:
            if doc_type == "rueckblick":
                entries.extend([
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, "Rückblick Word", 1, 0, "ausstehend", "", timestamp, oe_bez_kette),
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, "Rückblick PDF", 1, 0, "ausstehend", "", timestamp, oe_bez_kette),
                ])
            elif doc_type == "ausblick":
                entries.extend([
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, "Ausblick Word", 1, 0, "ausstehend", "", timestamp, oe_bez_kette),
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, "Ausblick PDF", 1, 0, "ausstehend", "", timestamp, oe_bez_kette),
                ])
        if include_feedback:
            feedback_exists = self._has_feedback_for_manager(mgr_pn, rb_year)
            if not feedback_exists:
                expected_feedback = self._count_direct_reports(mgr_pn)
                entries.append(self._create_entry(mgr_pn, mgr_name, "", "", "Feedback PDF", expected_feedback, 0, "ausstehend", "", timestamp, oe_bez_kette))
        for entry in entries:
            self._append_to_csv(entry)

    def log_feedback_for_manager(self, mgr_pn: str, mgr_name: str, rb_year: int, managers_index: dict = None, oe_bez_kette: str = "") -> None:
        feedback_exists = self._has_feedback_for_manager(mgr_pn, rb_year)
        if feedback_exists:
            return
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        expected_feedback = self._count_direct_reports(mgr_pn, managers_index)
        entry = self._create_entry(mgr_pn, mgr_name, "", "", "Feedback PDF", expected_feedback, 0, "ausstehend", "", timestamp, oe_bez_kette)
        self._append_to_csv(entry)

    # ... restliche Methoden unverändert übernehmen ...
    def _create_entry(self, mgr_pn: str, mgr_name: str, emp_pn: str, emp_name: str,
                      doc_type: str, erwartet: int, erhalten: int, status: str,
                      status_grund: str, timestamp: str, oe_bez_kette: str = "") -> Dict:
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
            "zuletzt_erinnert_am": "",
            "oe_bez_kette": str(oe_bez_kette),
        }

    def _append_to_csv(self, entry: Dict) -> None:
        file_exists = self.log_path.exists()
        with open(self.log_path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=entry.keys(), delimiter=";")
            if not file_exists:
                writer.writeheader()
            writer.writerow(entry)


    # Dashboard-APIs
    def get_dashboard_data(self, filter_status: str = "") -> pd.DataFrame:
        """Lädt die Dashboard-Daten aus der Tracking-CSV.

        filter_status: Optionaler Status-Filter (entspricht Werten in 'status').
        """
        if not self.log_path.exists():
            return pd.DataFrame()
        try:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            # Fallback-Encoding
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8")

        # Sicherstellen, dass erwartete Spalten existieren
        for col in [
            "log_id","vg_pn","vg_name","ma_pn","ma_name","doc_type",
            "erwartet","erhalten","status","status_grund","versendet_am","zuletzt_erinnert_am","oe_bez_kette"
        ]:
            if col not in df.columns:
                df[col] = ""

        if filter_status:
            df = df[df["status"].astype(str).str.strip().str.lower() == str(filter_status).strip().lower()]
        return df

    def update_entry(self, log_id: str, updates: Dict[str, str]) -> bool:
        """Aktualisiert Felder eines Eintrags anhand der log_id in der CSV.

        Returns True bei Erfolg, sonst False.
        """
        if not self.log_path.exists():
            return False
        try:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8")

        if "log_id" not in df.columns:
            return False
        mask = df["log_id"].astype(str) == str(log_id)
        if not mask.any():
            return False

        for key, value in (updates or {}).items():
            if key not in df.columns:
                # füge unbekannte Spalten dynamisch hinzu
                df[key] = ""
            df.loc[mask, key] = str(value)

        # Speichern mit unverändertem Format
        df.to_csv(self.log_path, sep=";", index=False, encoding="utf-8-sig")
        return True

    def mark_received_word(self, vg_pn: str, ma_pn: str, doc_type: str) -> bool:
        """Markiert ein Word-Dokument als empfangen (eindeutige Zuordnung über VG-PN + MA-PN).
        
        Args:
            vg_pn: Vorgesetzten-Personalnummer
            ma_pn: Mitarbeiter-Personalnummer
            doc_type: Dokumenttyp (z.B. "Rückblick Word", "Ausblick Word")
            
        Returns:
            True wenn Update erfolgreich, False sonst
        """
        if not self.log_path.exists():
            return False
            
        try:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8")
        
        # Eindeutige Zuordnung über vg_pn + ma_pn + doc_type
        mask = (
            (df["vg_pn"].astype(str).str.strip() == str(vg_pn).strip()) &
            (df["ma_pn"].astype(str).str.strip() == str(ma_pn).strip()) &
            (df["doc_type"].astype(str).str.strip() == str(doc_type).strip())
        )
        
        if not mask.any():
            # Kein passender Eintrag gefunden - Dokument war nicht erwartet
            return False
        
        # Erhöhe erhalten-Zähler und aktualisiere Status
        df.loc[mask, "erhalten"] = (pd.to_numeric(df.loc[mask, "erhalten"], errors="coerce").fillna(0) + 1).astype(int).astype(str)
        
        # Status auf "erhalten" setzen wenn erhalten >= erwartet
        erwartet = pd.to_numeric(df.loc[mask, "erwartet"], errors="coerce").fillna(0)
        erhalten = pd.to_numeric(df.loc[mask, "erhalten"], errors="coerce").fillna(0)
        df.loc[mask & (erhalten >= erwartet), "status"] = "erhalten"
        
        # Speichern
        df.to_csv(self.log_path, sep=";", index=False, encoding="utf-8-sig")
        return True

    def mark_received_pdf(self, ma_pn: str, doc_type: str) -> Dict[str, any]:
        """Markiert ein PDF-Dokument als empfangen (nur MA-PN bekannt).
        
        Bei Mehrfachanstellungen kann keine eindeutige Zuordnung erfolgen.
        
        Args:
            ma_pn: Mitarbeiter-Personalnummer
            doc_type: Dokumenttyp (z.B. "Rückblick PDF", "Ausblick PDF")
            
        Returns:
            Dictionary mit:
            - success: True wenn eindeutig zugeordnet
            - matched_count: Anzahl gefundener Einträge (0, 1, oder 2+)
            - message: Beschreibung des Ergebnisses
            - matched_entries: Liste der gefundenen Einträge (bei Mehrfachanstellung)
        """
        if not self.log_path.exists():
            return {
                "success": False,
                "matched_count": 0,
                "message": "Tracking-Datei nicht gefunden",
                "matched_entries": []
            }
            
        try:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8")
        
        # Suche alle Einträge mit dieser MA-PN + doc_type
        mask = (
            (df["ma_pn"].astype(str).str.strip() == str(ma_pn).strip()) &
            (df["doc_type"].astype(str).str.strip() == str(doc_type).strip())
        )
        
        matched_count = mask.sum()
        
        if matched_count == 0:
            return {
                "success": False,
                "matched_count": 0,
                "message": "Kein Tracking-Eintrag gefunden - Dokument war nicht erwartet",
                "matched_entries": []
            }
        
        elif matched_count == 1:
            # Eindeutige Zuordnung - Update durchführen
            df.loc[mask, "erhalten"] = (pd.to_numeric(df.loc[mask, "erhalten"], errors="coerce").fillna(0) + 1).astype(int).astype(str)
            
            # Status auf "erhalten" setzen wenn erhalten >= erwartet
            erwartet = pd.to_numeric(df.loc[mask, "erwartet"], errors="coerce").fillna(0)
            erhalten = pd.to_numeric(df.loc[mask, "erhalten"], errors="coerce").fillna(0)
            df.loc[mask & (erhalten >= erwartet), "status"] = "erhalten"
            
            # Speichern
            df.to_csv(self.log_path, sep=";", index=False, encoding="utf-8-sig")
            
            vg_info = df[mask].iloc[0]
            return {
                "success": True,
                "matched_count": 1,
                "message": f"Zugeordnet zu VG {vg_info.get('vg_name', '')} ({vg_info.get('vg_pn', '')})",
                "matched_entries": []
            }
        
        else:  # matched_count > 1
            # Mehrfachanstellung - keine automatische Zuordnung
            matched_entries = df[mask][["vg_pn", "vg_name", "log_id"]].to_dict('records')
            vg_list = ", ".join([f"{e['vg_name']} ({e['vg_pn']})" for e in matched_entries])
            
            return {
                "success": False,
                "matched_count": matched_count,
                "message": f"Mehrfachanstellung: {matched_count} Vorgesetzte - {vg_list}",
                "matched_entries": matched_entries
            }

    def mark_error(self, filename: str, ma_pn: str, doc_type: str, error_message: str, vg_pn: str = None) -> bool:
        """Markiert einen Tracking-Eintrag als fehlerhaft.
        
        Args:
            filename: Dateiname
            ma_pn: Mitarbeiter-Personalnummer
            doc_type: Dokumenttyp
            error_message: Fehlerbeschreibung
            vg_pn: Optional - Vorgesetzten-PN für eindeutige Zuordnung
            
        Returns:
            True wenn Update erfolgreich
        """
        if not self.log_path.exists():
            return False
            
        try:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8")
        
        # Zuordnung je nachdem ob VG-PN bekannt ist
        if vg_pn:
            # Eindeutige Zuordnung mit VG-PN
            mask = (
                (df["vg_pn"].astype(str).str.strip() == str(vg_pn).strip()) &
                (df["ma_pn"].astype(str).str.strip() == str(ma_pn).strip()) &
                (df["doc_type"].astype(str).str.strip() == str(doc_type).strip())
            )
        else:
            # Ohne VG-PN - nur mit MA-PN (kann mehrere treffen bei Mehrfachanstellung)
            mask = (
                (df["ma_pn"].astype(str).str.strip() == str(ma_pn).strip()) &
                (df["doc_type"].astype(str).str.strip() == str(doc_type).strip())
            )
        
        if not mask.any():
            return False
        
        # Setze Status und Fehlergrund
        df.loc[mask, "status"] = "prüfung_nötig"
        df.loc[mask, "status_grund"] = str(error_message)
        
        # Speichern
        df.to_csv(self.log_path, sep=";", index=False, encoding="utf-8-sig")
        return True

    def check_duplicate(self, filename: str, ma_pn: str, doc_type: str, vg_pn: str = None) -> tuple:
        """Prüft ob bereits ein Dokument dieses Typs für diese Person empfangen wurde.
        
        Args:
            filename: Dateiname
            ma_pn: Mitarbeiter-Personalnummer
            doc_type: Dokumenttyp
            vg_pn: Optional - Vorgesetzten-PN für eindeutige Zuordnung
            
        Returns:
            Tuple (is_duplicate: bool, warning_message: str)
        """
        if not self.log_path.exists():
            return (False, "")
            
        try:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8")
        
        # Zuordnung je nachdem ob VG-PN bekannt ist
        if vg_pn:
            mask = (
                (df["vg_pn"].astype(str).str.strip() == str(vg_pn).strip()) &
                (df["ma_pn"].astype(str).str.strip() == str(ma_pn).strip()) &
                (df["doc_type"].astype(str).str.strip() == str(doc_type).strip())
            )
        else:
            mask = (
                (df["ma_pn"].astype(str).str.strip() == str(ma_pn).strip()) &
                (df["doc_type"].astype(str).str.strip() == str(doc_type).strip())
            )
        
        if not mask.any():
            return (False, "")
        
        # Prüfe ob bereits erhalten
        erhalten = pd.to_numeric(df.loc[mask, "erhalten"], errors="coerce").fillna(0)
        
        if (erhalten > 0).any():
            return (True, f"Bereits {int(erhalten.max())} Dokument(e) für {doc_type} empfangen")
        
        return (False, "")

    def _count_direct_reports(self, mgr_pn: str, managers_index: dict = None) -> int:
        """Zählt Anzahl direkter Unterstellter eines Vorgesetzten."""
        if managers_index and mgr_pn in managers_index:
            subs = managers_index[mgr_pn].get("subs")
            if subs is not None:
                return len(subs)
        return 0

    def _has_feedback_for_manager(self, mgr_pn: str, rb_year: int) -> bool:
        """Prüft ob bereits ein Feedback-Eintrag für diesen Manager existiert."""
        if not self.log_path.exists():
            return False
        try:
            df = pd.read_csv(self.log_path, sep=";", dtype=str, encoding="utf-8-sig")
        except Exception:
            return False
        
        mask = (
            (df["vg_pn"].astype(str).str.strip() == str(mgr_pn).strip()) &
            (df["doc_type"].astype(str).str.strip() == "Feedback PDF")
        )
        return mask.any()

