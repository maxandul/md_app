"""
Erinnerungs-Service für das MD-Prozess-Tool.

Dieser Service kapselt alle Funktionen zur Erinnerungsmail-Generierung
und -Versendung basierend auf ausgewählten Dashboard-Einträgen.
"""

from __future__ import annotations
from datetime import datetime, timedelta
from typing import Dict, List, Any
import pandas as pd

from data_loader import load_config
from mail_send import send_mail
from constants import MDConstants


class ErinnerungService:
    def __init__(self, app):
        self.app = app
        self.config = load_config()
        self.mail_config = self.config.get("mail_erinnerung", {})
        
    def send_reminders(self, selection: List[str], mode: str = "send") -> Dict[str, Any]:
        """
        Versendet oder speichert Erinnerungsmails für ausgewählte Dashboard-Einträge.
        
        Args:
            selection: Liste der ausgewählten Treeview-Item-IDs
            mode: "send" für Versenden, "display" für Entwurf
            
        Returns:
            Dictionary mit Ergebnis-Informationen
        """
        try:
            # Ausgewählte Einträge in DataFrame konvertieren
            selected_data = self._get_selected_data(selection)
            if selected_data.empty:
                return {"success": False, "error": "Keine gültigen Daten ausgewählt"}
            
            # Nach Vorgesetzten gruppieren
            grouped_data = self._group_by_manager(selected_data)
            if not grouped_data:
                return {"success": False, "error": "Keine Vorgesetzten-Daten gefunden"}
            
            emails_sent = 0
            errors = 0
            
            # Pro Vorgesetzten eine E-Mail erstellen
            for mgr_pn, mgr_data in grouped_data.items():
                try:
                    result = self._create_and_send_reminder(mgr_pn, mgr_data, mode)
                    if result["success"]:
                        emails_sent += 1
                        # Erinnerungsdatum im Tracking aktualisieren
                        self._update_reminder_date(mgr_pn, mgr_data["log_ids"])
                    else:
                        errors += 1
                        print(f"Fehler bei VG {mgr_pn}: {result['error']}")
                        
                except Exception as e:
                    errors += 1
                    print(f"Unerwarteter Fehler bei VG {mgr_pn}: {e}")
            
            return {
                "success": True,
                "emails_sent": emails_sent,
                "errors": errors
            }
            
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def preview_reminders(self, selection: List[str]) -> Dict[str, Any]:
        """
        Generiert Vorschau der Erinnerungsmails für ausgewählte Einträge.
        
        Args:
            selection: Liste der ausgewählten Treeview-Item-IDs
            
        Returns:
            Dictionary mit Vorschau-Daten
        """
        try:
            # Ausgewählte Einträge in DataFrame konvertieren
            selected_data = self._get_selected_data(selection)
            if selected_data.empty:
                return {"success": False, "error": "Keine gültigen Daten ausgewählt"}
            
            # Nach Vorgesetzten gruppieren
            grouped_data = self._group_by_manager(selected_data)
            if not grouped_data:
                return {"success": False, "error": "Keine Vorgesetzten-Daten gefunden"}
            
            previews = []
            
            # Pro Vorgesetzten eine Vorschau erstellen
            for mgr_pn, mgr_data in grouped_data.items():
                try:
                    preview = self._create_reminder_preview(mgr_pn, mgr_data)
                    if preview:
                        previews.append(preview)
                        
                except Exception as e:
                    print(f"Fehler bei Vorschau für VG {mgr_pn}: {e}")
            
            return {
                "success": True,
                "previews": previews
            }
            
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def _get_selected_data(self, selection: List[str]) -> pd.DataFrame:
        """Konvertiert ausgewählte Treeview-Items in DataFrame."""
        data = []
        
        for item_id in selection:
            item = self.app.tree_dashboard.item(item_id)
            values = item.get("values", [])
            if len(values) >= 12:
                data.append({
                    "log_id": values[0],
                    "vg_pn": values[1],
                    "vg_name": values[2],
                    "ma_pn": values[3],
                    "ma_name": values[4],
                    "doc_type": values[5],
                    "erwartet": values[6],
                    "erhalten": values[7],
                    "status": values[8],
                    "status_grund": values[9],
                    "versendet_am": values[10],
                    "zuletzt_erinnert_am": values[11]
                })
        
        return pd.DataFrame(data)
    
    def _group_by_manager(self, df: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
        """Gruppiert ausgewählte Einträge nach Vorgesetzten."""
        grouped = {}
        
        for _, row in df.iterrows():
            mgr_pn = str(row["vg_pn"]).strip()
            if mgr_pn not in grouped:
                grouped[mgr_pn] = {
                    "vg_name": row["vg_name"],
                    "documents": [],
                    "log_ids": []
                }
            
            grouped[mgr_pn]["documents"].append({
                "doc_type": row["doc_type"],
                "ma_name": row["ma_name"],
                "ma_pn": row["ma_pn"]
            })
            grouped[mgr_pn]["log_ids"].append(row["log_id"])
        
        return grouped
    
    def _create_and_send_reminder(self, mgr_pn: str, mgr_data: Dict[str, Any], mode: str) -> Dict[str, Any]:
        """Erstellt und versendet eine Erinnerungsmail für einen Vorgesetzten."""
        try:
            # Vorgesetzten-Daten abrufen
            vg_info = self._get_manager_info(mgr_pn)
            if not vg_info:
                return {"success": False, "error": f"Keine Vorgesetzten-Informationen für PN {mgr_pn}"}
            
            # E-Mail-Inhalt generieren
            subject, body = self._generate_email_content(vg_info, mgr_data["documents"])
            
            # E-Mail versenden oder als Entwurf speichern
            send_mail(
                to=vg_info["email"],
                subject=subject,
                html_body=body,
                mode_override=mode
            )
            
            return {"success": True}
            
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def _create_reminder_preview(self, mgr_pn: str, mgr_data: Dict[str, Any]) -> Dict[str, str]:
        """Erstellt eine Vorschau der Erinnerungsmail für einen Vorgesetzten."""
        try:
            # Vorgesetzten-Daten abrufen
            vg_info = self._get_manager_info(mgr_pn)
            if not vg_info:
                return None
            
            # E-Mail-Inhalt generieren
            subject, body = self._generate_email_content(vg_info, mgr_data["documents"])
            
            return {
                "to": vg_info["email"],
                "subject": subject,
                "body": body
            }
            
        except Exception as e:
            print(f"Fehler bei Vorschau-Generierung: {e}")
            return None
    
    def _get_manager_info(self, mgr_pn: str) -> Dict[str, str]:
        """Ruft Vorgesetzten-Informationen ab."""
        try:
            # Aus dem managers_index abrufen
            if hasattr(self.app, 'mgr_index') and self.app.mgr_index:
                if mgr_pn in self.app.mgr_index:
                    mgr_row = self.app.mgr_index[mgr_pn]["manager"]
                    if mgr_row is not None:
                        rufname = str(mgr_row.get('Rufname', '')).strip()
                        nachname = str(mgr_row.get('Nachname', '')).strip()
                        email = str(mgr_row.get("lange ID/Nummer", "")).strip()
                        
                        return {
                            "rufname": rufname,
                            "nachname": nachname,
                            "email": email
                        }
            
            # Fallback: Aus Dashboard-Daten
            df = self.app.tracking.get_dashboard_data()
            vg_rows = df[df["vg_pn"] == mgr_pn]
            if not vg_rows.empty:
                vg_name = str(vg_rows.iloc[0]["vg_name"]).strip()
                # E-Mail-Adresse ist nicht in Dashboard-Daten verfügbar
                return {
                    "rufname": vg_name.split()[0] if vg_name.split() else "",
                    "nachname": " ".join(vg_name.split()[1:]) if len(vg_name.split()) > 1 else "",
                    "email": ""  # Wird in der Validierung abgefangen
                }
            
            return None
            
        except Exception as e:
            print(f"Fehler beim Abrufen der VG-Informationen: {e}")
            return None
    
    def _generate_email_content(self, vg_info: Dict[str, str], documents: List[Dict[str, str]]) -> tuple[str, str]:
        """Generiert Betreff und Inhalt der Erinnerungsmail."""
        try:
            # Betreff generieren
            subject_template = self.mail_config.get("subject_template", "Erinnerung: Fehlende MD-Unterlagen")
            subject = subject_template
            
            # Dokumentliste generieren
            dokument_liste = ""
            for doc in documents:
                ma_name = str(doc["ma_name"]).strip()
                doc_type = str(doc["doc_type"]).strip()
                dokument_liste += f"      <li>{doc_type} für {ma_name}</li>\n"
            
            # Frist berechnen (heute + 10 Tage)
            frist_datum = (datetime.now() + timedelta(days=10)).strftime("%d.%m.%Y")
            
            # E-Mail-Inhalt generieren
            body_template = self.mail_config.get("body_html_template", "")
            body = body_template.format(
                rufname=vg_info["rufname"],
                dokument_liste=dokument_liste.strip(),
                frist_datum=frist_datum
            )
            
            return subject, body
            
        except Exception as e:
            raise Exception(f"Fehler bei E-Mail-Generierung: {e}")
    
    def _update_reminder_date(self, mgr_pn: str, log_ids: List[str]) -> None:
        """Aktualisiert das Erinnerungsdatum im Tracking-System."""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            for log_id in log_ids:
                self.app.tracking.update_entry(log_id, {"zuletzt_erinnert_am": timestamp})
                
        except Exception as e:
            print(f"Fehler beim Aktualisieren des Erinnerungsdatums: {e}")
