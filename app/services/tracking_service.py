from __future__ import annotations

import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List
import csv

from app.data_loader import load_config

CFG = load_config()


class SimpleTrackingSystem:
    def __init__(self):
        tracking_dir = (CFG.get("paths", {}) or {}).get("tracking_dir", "../tracking")
        self.log_path = (Path(__file__).parent / tracking_dir).resolve() / "md_logging.csv"
        self.log_path.parent.mkdir(parents=True, exist_ok=True)

    # ... Implementation identisch zur bisherigen simple_tracking.SimpleTrackingSystem ...
    # Aus Platzgründen unverändert übernommen
    def log_versand(self, mgr_pn: str, mgr_name: str, emp_pn: str, emp_name: str,
                    doc_types: List[str], rb_year: int, ab_year: int, include_feedback: bool = False) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entries = []
        for doc_type in doc_types:
            if doc_type == "rueckblick":
                entries.extend([
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, "Rückblick Word", 1, 0, "ausstehend", "", timestamp),
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, "Rückblick PDF", 1, 0, "ausstehend", "", timestamp),
                ])
            elif doc_type == "ausblick":
                entries.extend([
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, "Ausblick Word", 1, 0, "ausstehend", "", timestamp),
                    self._create_entry(mgr_pn, mgr_name, emp_pn, emp_name, "Ausblick PDF", 1, 0, "ausstehend", "", timestamp),
                ])
        if include_feedback:
            feedback_exists = self._has_feedback_for_manager(mgr_pn, rb_year)
            if not feedback_exists:
                expected_feedback = self._count_direct_reports(mgr_pn)
                entries.append(self._create_entry(mgr_pn, mgr_name, "", "", "Feedback PDF", expected_feedback, 0, "ausstehend", "", timestamp))
        for entry in entries:
            self._append_to_csv(entry)

    def log_feedback_for_manager(self, mgr_pn: str, mgr_name: str, rb_year: int, managers_index: dict = None) -> None:
        feedback_exists = self._has_feedback_for_manager(mgr_pn, rb_year)
        if feedback_exists:
            return
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        expected_feedback = self._count_direct_reports(mgr_pn, managers_index)
        entry = self._create_entry(mgr_pn, mgr_name, "", "", "Feedback PDF", expected_feedback, 0, "ausstehend", "", timestamp)
        self._append_to_csv(entry)

    # ... restliche Methoden unverändert übernehmen ...
    def _create_entry(self, mgr_pn: str, mgr_name: str, emp_pn: str, emp_name: str,
                      doc_type: str, erwartet: int, erhalten: int, status: str,
                      status_grund: str, timestamp: str) -> Dict:
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
            "erwartet","erhalten","status","status_grund","versendet_am","zuletzt_erinnert_am"
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

