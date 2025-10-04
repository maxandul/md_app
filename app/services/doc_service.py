"""
Dokumentenverarbeitungs-Service für das MD-Prozess-Tool.

Dieser Service kapselt alle Funktionen zur Verarbeitung von DOCX- und PDF-Dokumenten,
einschließlich Validierung, Export und Verschiebung.
"""

from __future__ import annotations
from pathlib import Path
import re
import shutil
import unicodedata
import pandas as pd
from typing import Dict, Any
from tkinter import messagebox

from docx_tools import read_content_controls, detect_doc_type, map_rb_gesamteindruck
from simple_tracking import SimpleTrackingSystem
from constants import MDConstants, DocType, ProcStatus
from org_structure import build_org_structure
from data_loader import load_employees
from views.ui_utils import autosize_tree_columns


def build_sap_index(df: pd.DataFrame) -> dict[str, pd.Series]:
    """
    Erstellt Index für schnelle PN-Suche in SAP-Daten.
    
    Args:
        df: SAP-Stammdaten DataFrame
        
    Returns:
        Dictionary mit PN -> SAP-Datensatz
    """
    idx = {}
    for _, r in df.iterrows():
        pn = str(r.get("ID_NO_ZERO", "")).strip()
        if pn:
            idx[pn] = r
    return idx


def _strip_accents(s: str) -> str:
    """
    Normalisiert Text für Namensvergleich.
    
    Entfernt Akzente, konvertiert zu Kleinbuchstaben,
    normalisiert Leerzeichen und Bindestriche.
    """
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("-", " ")
    return s


def _name_matches(full_name_from_doc: str, sap_row: pd.Series) -> bool:
    """
    Vergleicht Namen aus Dokument mit SAP-Stammdaten.
    
    Args:
        full_name_from_doc: Vollständiger Name aus DOCX-Tag
        sap_row: SAP-Datensatz mit Rufname/Nachname
        
    Returns:
        True wenn Namen übereinstimmen (normalisiert)
    """
    doc = _strip_accents(full_name_from_doc)
    sap = _strip_accents(f"{sap_row.get('Rufname','')} {sap_row.get('Nachname','')}")
    return doc == sap


def _pn_in_sap(pn: str, sap_index: dict[str, pd.Series]) -> pd.Series | None:
    """
    Sucht Personalnummer in SAP-Index.
    
    Args:
        pn: Personalnummer zum Suchen
        sap_index: Dictionary mit PN -> SAP-Datensatz
        
    Returns:
        SAP-Datensatz oder None wenn nicht gefunden
    """
    if not pn:
        return None
    return sap_index.get(str(pn))


def _append_processing_log(results: list[dict], durchlauf_jahr: int):
    """Append-Protokollierung nach ruecklauf/logs/processing_log.csv"""
    from datetime import datetime
    import csv
    from pathlib import Path

    log_dir = Path(__file__).parent.parent / MDConstants.RUECKLAUF_DIR / MDConstants.LOGS_DIR
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "processing_log.csv"

    # CSV-Header falls neu
    header = [
        "timestamp", "durchlauf_jahr", "filename", "ext", "typ", "pn", "vg_pn",
        "status", "action", "reason", "target_path", "source_folder"
    ]

    # Append-Modus
    file_exists = log_file.exists()
    with open(log_file, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        if not file_exists:
            writer.writerow(header)

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in results:
            # Extrahiere PN/VG-PN aus Tags falls vorhanden
            pn = ""
            vg_pn = ""
            if isinstance(r.get("extras"), dict):
                tags = r.get("extras", {}).get("all_tags", {})
                pn = str(tags.get("rb_pn", "") or tags.get("ab_pn", "") or "")
                vg_pn = str(tags.get("rb_pn_vg", "") or tags.get("ab_pn_vg", "") or "")

            # Bestimme action
            action = "validated"
            if r.get("status") == ProcStatus.OK.value:
                action = "exported"
            elif r.get("status") == ProcStatus.MANUELL.value:
                action = "flagged_manual"

            writer.writerow([
                now,
                durchlauf_jahr,
                r.get("file", ""),
                Path(r.get("file", "")).suffix.lower(),
                r.get("typ", ""),
                pn,
                vg_pn,
                r.get("status", ""),
                action,
                r.get("reason", ""),
                r.get("target", ""),
                "unverarbeitet"
            ])


def process_docx_folder(input_dir: Path, sap_df: pd.DataFrame, max_files: int | None = None, durchlauf_jahr: int | None = None) -> list[dict]:
    """
    Scannt NUR *.docx direkt in input_dir (keine Unterordner),
    liest Tags mit python-docx, validiert gg. SAP und liefert eine Liste von Dicts:
      - file, typ, pn, name, status ('ok'|'manuell'|'unbekannt'), reason, extras (dict)
    """
    results: list[Dict[str, Any]] = []
    sap_idx = build_sap_index(sap_df)
    tracking = SimpleTrackingSystem()

    docx_files = sorted(input_dir.glob(f"*{MDConstants.ALLOWED_EXTENSIONS[0]}"))
    if max_files is not None:
        docx_files = docx_files[:max_files]
    for p in docx_files:  # nur Top-Level
        try:
            tags = read_content_controls(p)
        except Exception as e:
            results.append({
                "file": p.name, "typ": "Unbekannt", "pn": "", "name": "",
                "status": ProcStatus.MANUELL.value, "reason": f"DOCX beschädigt/lesefehler: {e}", "extras": {}
            })
            continue

        typ = detect_doc_type(tags)

        if typ == DocType.RUECKBLICK.value:
            name = (tags.get("rb_name") or "").strip()
            pn   = (tags.get("rb_pn") or "").strip()
            if not name or not pn:
                # VG-PN aus Tags extrahieren für Tracking
                vg_pn = (tags.get("rb_pn_vg") or "").strip()
                
                # Tracking: Markiere als empfangen aber fehlerhaft
                if vg_pn and pn:
                    tracking.mark_received(vg_pn, pn, "Rückblick Word")
                    tracking.mark_error(p.name, pn, "Rückblick Word", "Pflichtfelder rb_name/rb_pn fehlen")
                
                results.append({
                    "file": p.name, "typ": typ, "pn": pn, "name": name,
                    "status": ProcStatus.MANUELL.value, "reason": "Pflichtfelder rb_name/rb_pn fehlen", "extras": {"all_tags": tags}
                })
                continue

            sap_row = _pn_in_sap(pn, sap_idx)
            if sap_row is None:
                # VG-PN aus Tags extrahieren für Tracking
                vg_pn = (tags.get("rb_pn_vg") or "").strip()
                
                # Tracking: Markiere als empfangen aber fehlerhaft
                if vg_pn and pn:
                    tracking.mark_received(vg_pn, pn, "Rückblick Word")
                    tracking.mark_error(p.name, pn, "Rückblick Word", "PN nicht in SAP gefunden")
                
                results.append({
                    "file": p.name, "typ": typ, "pn": pn, "name": name,
                    "status": ProcStatus.MANUELL.value, "reason": "PN nicht in SAP gefunden", "extras": {"all_tags": tags}
                })
                continue

            if not _name_matches(name, sap_row):
                # VG-PN aus Tags extrahieren für Tracking
                vg_pn = (tags.get("rb_pn_vg") or "").strip()
                
                # Tracking: Markiere als empfangen aber fehlerhaft
                if vg_pn and pn:
                    tracking.mark_received(vg_pn, pn, "Rückblick Word")
                    tracking.mark_error(p.name, pn, "Rückblick Word", "Name passt nicht zu SAP")
                
                results.append({
                    "file": p.name, "typ": typ, "pn": pn, "name": name,
                    "status": ProcStatus.MANUELL.value, "reason": "Name passt nicht zu SAP", "extras": {"all_tags": tags}
                })
                continue

            gi_disp = (tags.get("rb_gesamteindruck") or "").strip()
            gi_code = map_rb_gesamteindruck(gi_disp) if gi_disp else ""
            if not gi_code:
                # VG-PN aus Tags extrahieren für Tracking
                vg_pn = (tags.get("rb_pn_vg") or "").strip()
                
                # Tracking: Markiere als empfangen aber fehlerhaft
                if vg_pn and pn:
                    tracking.mark_received(vg_pn, pn, "Rückblick Word")
                    tracking.mark_error(p.name, pn, "Rückblick Word", "rb_gesamteindruck fehlt/ungültig")
                
                results.append({
                    "file": p.name, "typ": typ, "pn": pn, "name": name,
                    "status": ProcStatus.PRUEFUNG_NOETIG.value, "reason": "rb_gesamteindruck fehlt/ungültig", "extras": {"all_tags": tags}
                })
                continue

            # VG-PN aus Tags extrahieren
            vg_pn = (tags.get("rb_pn_vg") or "").strip()
            
            # Duplikat-Erkennung - spezifisch für Rückblick Word mit VG-PN
            doc_type = "Rückblick Word"
            is_duplicate, warning = tracking.check_duplicate(p.name, pn, doc_type, vg_pn)
            if is_duplicate:
                # Tracking: Markiere als empfangen aber fehlerhaft
                if vg_pn and pn:
                    tracking.mark_received(vg_pn, pn, "Rückblick Word")
                    tracking.mark_error(p.name, pn, "Rückblick Word", f"Duplikat: {warning}")
                
                results.append({
                    "file": p.name, "typ": typ, "pn": pn, "name": name,
                    "status": ProcStatus.MANUELL.value, "reason": f"Duplikat: {warning}", "extras": {"all_tags": tags}
                })
                continue
            
            # Status im Tracking-System auf "erhalten" setzen
            if vg_pn:
                tracking.mark_received(vg_pn, pn, "Rückblick Word")
            
            results.append({
                "file": p.name, "typ": typ, "pn": pn, "name": name,
                "status": ProcStatus.OK.value, "reason": "",
                "extras": {"rb_gesamteindruck": gi_code, "all_tags": tags}
            })

        elif typ == DocType.AUSBLiCK.value:
            name = (tags.get("ab_name") or "").strip()
            pn   = (tags.get("ab_pn") or "").strip()
            if not name or not pn:
                # VG-PN aus Tags extrahieren für Tracking
                vg_pn = (tags.get("ab_pn_vg") or "").strip()
                
                # Tracking: Markiere als empfangen aber fehlerhaft
                if vg_pn and pn:
                    tracking.mark_received(vg_pn, pn, "Ausblick Word")
                    tracking.mark_error(p.name, pn, "Ausblick Word", "Pflichtfelder ab_name/ab_pn fehlen")
                
                results.append({
                    "file": p.name, "typ": typ, "pn": pn, "name": name,
                    "status": ProcStatus.MANUELL.value, "reason": "Pflichtfelder ab_name/ab_pn fehlen", "extras": {"all_tags": tags}
                })
                continue

            sap_row = _pn_in_sap(pn, sap_idx)
            if sap_row is None:
                # VG-PN aus Tags extrahieren für Tracking
                vg_pn = (tags.get("ab_pn_vg") or "").strip()
                
                # Tracking: Markiere als empfangen aber fehlerhaft
                if vg_pn and pn:
                    tracking.mark_received(vg_pn, pn, "Ausblick Word")
                    tracking.mark_error(p.name, pn, "Ausblick Word", "PN nicht in SAP gefunden")
                
                results.append({
                    "file": p.name, "typ": typ, "pn": pn, "name": name,
                    "status": ProcStatus.MANUELL.value, "reason": "PN nicht in SAP gefunden", "extras": {"all_tags": tags}
                })
                continue

            if not _name_matches(name, sap_row):
                # VG-PN aus Tags extrahieren für Tracking
                vg_pn = (tags.get("ab_pn_vg") or "").strip()
                
                # Tracking: Markiere als empfangen aber fehlerhaft
                if vg_pn and pn:
                    tracking.mark_received(vg_pn, pn, "Ausblick Word")
                    tracking.mark_error(p.name, pn, "Ausblick Word", "Name passt nicht zu SAP")
                
                results.append({
                    "file": p.name, "typ": typ, "pn": pn, "name": name,
                    "status": ProcStatus.MANUELL.value, "reason": "Name passt nicht zu SAP", "extras": {"all_tags": tags}
                })
                continue

            # VG-PN aus Tags extrahieren
            vg_pn = (tags.get("ab_pn_vg") or "").strip()
            
            # Duplikat-Erkennung - spezifisch für Ausblick Word mit VG-PN
            doc_type = "Ausblick Word"
            is_duplicate, warning = tracking.check_duplicate(p.name, pn, doc_type, vg_pn)
            if is_duplicate:
                # Tracking: Markiere als empfangen aber fehlerhaft
                if vg_pn and pn:
                    tracking.mark_received(vg_pn, pn, "Ausblick Word")
                    tracking.mark_error(p.name, pn, "Ausblick Word", f"Duplikat: {warning}")
                
                results.append({
                    "file": p.name, "typ": typ, "pn": pn, "name": name,
                    "status": ProcStatus.MANUELL.value, "reason": f"Duplikat: {warning}", "extras": {"all_tags": tags}
                })
                continue
            
            # Status im Tracking-System auf "erhalten" setzen
            if vg_pn:
                tracking.mark_received(vg_pn, pn, "Ausblick Word")
            
            results.append({
                "file": p.name, "typ": typ, "pn": pn, "name": name,
                "status": ProcStatus.OK.value, "reason": "", "extras": {"all_tags": tags}
            })

        else:
            results.append({
                "file": p.name, "typ": "Unbekannt", "pn": "", "name": "",
                "status": ProcStatus.MANUELL.value, "reason": "Dokumenttyp nicht erkannt (keine rb_/ab_-Tags)", "extras": {"all_tags": tags}
            })

    # Logging nach ruecklauf/logs/processing_log.csv
    if durchlauf_jahr is not None:
        _append_processing_log(results, durchlauf_jahr)

    return results


# Hilfsfunktionen
def _strip_accents(s: str) -> str:
    """
    Normalisiert Text für Namensvergleich.
    
    Entfernt Akzente, konvertiert zu Kleinbuchstaben,
    normalisiert Leerzeichen und Bindestriche.
    """
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("-", " ")
    return s


def _name_matches(full_name_from_doc: str, sap_row: pd.Series) -> bool:
    """
    Vergleicht Namen aus Dokument mit SAP-Stammdaten.
    
    Args:
        full_name_from_doc: Vollständiger Name aus DOCX-Tag
        sap_row: SAP-Datensatz mit Rufname/Nachname
        
    Returns:
        True wenn Namen übereinstimmen (normalisiert)
    """
    doc = _strip_accents(full_name_from_doc)
    sap = _strip_accents(f"{sap_row.get('Rufname','')} {sap_row.get('Nachname','')}")
    return doc == sap


def _pn_in_sap(pn: str, sap_index: dict[str, pd.Series]) -> pd.Series | None:
    """
    Sucht Personalnummer in SAP-Index.
    
    Args:
        pn: Personalnummer zum Suchen
        sap_index: Dictionary mit PN -> SAP-Datensatz
        
    Returns:
        SAP-Datensatz oder None wenn nicht gefunden
    """
    if not pn:
        return None
    return sap_index.get(str(pn))


def build_sap_index(df: pd.DataFrame) -> dict[str, pd.Series]:
    """
    Erstellt Index für schnelle PN-Suche in SAP-Daten.
    
    Args:
        df: SAP-Stammdaten DataFrame
        
    Returns:
        Dictionary mit PN -> SAP-Datensatz
    """
    idx = {}
    for _, r in df.iterrows():
        pn = str(r.get("ID_NO_ZERO", "")).strip()
        if pn:
            idx[pn] = r
    return idx


def _append_processing_log(results: list[dict], durchlauf_jahr: int):
    """Append-Protokollierung nach ruecklauf/logs/processing_log.csv"""
    from datetime import datetime
    import csv
    from pathlib import Path

    log_dir = Path(__file__).parent.parent.parent / MDConstants.RUECKLAUF_DIR / MDConstants.LOGS_DIR
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "processing_log.csv"

    # CSV-Header falls neu
    header = [
        "timestamp", "durchlauf_jahr", "filename", "ext", "typ", "pn", "vg_pn",
        "status", "action", "reason", "target_path", "source_folder"
    ]

    # Append-Modus
    file_exists = log_file.exists()
    with open(log_file, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        if not file_exists:
            writer.writerow(header)

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in results:
            # Extrahiere PN/VG-PN aus Tags falls vorhanden
            pn = ""
            vg_pn = ""
            if isinstance(r.get("extras"), dict):
                tags = r.get("extras", {}).get("all_tags", {})
                pn = str(tags.get("rb_pn", "") or tags.get("ab_pn", "") or "")
                vg_pn = str(tags.get("rb_pn_vg", "") or tags.get("ab_pn_vg", "") or "")

            # Bestimme action
            action = "validated"
            if r.get("status") == ProcStatus.OK.value:
                action = "exported"
            elif r.get("status") == ProcStatus.MANUELL.value:
                action = "flagged_manual"

            writer.writerow([
                now,
                durchlauf_jahr,
                r.get("file", ""),
                Path(r.get("file", "")).suffix.lower(),
                r.get("typ", ""),
                pn,
                vg_pn,
                r.get("status", ""),
                action,
                r.get("reason", ""),
                r.get("target", ""),
                "unverarbeitet"
            ])


# SAP Export Funktionen
SAP_COLS = [
    "PersNr",
    "Beurteilungsart",
    "Beginndatum IT9075",
    "Endedatum IT9075",
    "Ans.",
    "Datum MAB",
    "Beurteilungszeitraum von",
    "Beurteilungszeitraum bis",
    "Gesamtbeurteilung",
    "Zielerreichung",
    "Fachliche Kompetenz",
    "Sozialkompetenz (Verhalten)",
    "Führungskompetenz",
    "Nächster Termin",
]


def _parse_date(val: str):
    if not val:
        return pd.NaT
    try:
        return pd.to_datetime(val, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


def _safe_get(tags: dict, key: str) -> str:
    return str(tags.get(key, "") or "").strip()


def _map_beurteilungsart(typ: str | DocType) -> str:
    """Gibt die Beurteilungsart konsistent als String zurück.

    Unterstützt sowohl DocType-Enum als auch freie Strings.
    """
    if isinstance(typ, DocType):
        return typ.value
    if typ in (DocType.RUECKBLICK.value, DocType.AUSBLiCK.value):
        return typ
    return str(typ)


def export_sap_massenupload(results: list[dict], sap_df: pd.DataFrame, out_xlsx: Path):
    """
    Baut aus den 'ok'-Rückblick-DOCX-Ergebnissen die SAP-Massenupload-Tabelle
    gemäss Vorgabe (Beurteilungsart = '1', Datumslogik mit Eintritt/Austritt,
    Zeitraum-Felder an IT9075 geklemmt).
    """
    # PN -> Liste SAP-Zeilen Index (mehrere Anstellungen möglich)
    def _norm_pn(s: str) -> str:
        s = str(s or "").strip()
        s = s.lstrip("0")
        return s or "0"

    sap_idx = {}
    for _, r in sap_df.iterrows():
        pn_key = _norm_pn(r.get("ID_NO_ZERO", ""))
        sap_idx.setdefault(pn_key, []).append(r)

    def _d(s: str):
        if not s:
            return pd.NaT
        return pd.to_datetime(s, dayfirst=True, errors="coerce")

    rows = []

    for r in results:
        if r.get("status") != ProcStatus.OK.value or r.get("typ") != DocType.RUECKBLICK.value:
            continue

        # Tags holen (erst all_tags, sonst tags)
        tags = (r.get("extras", {}) or {}).get("all_tags", {}) or r.get("tags", {})
        pn = str(r.get("pn", "")).strip()
        if not pn:
            continue

        # SAP-Zeile wählen: erst PN + Vorgesetzter (PN) matchen, sonst Fallback auf PN
        pn_key = _norm_pn(pn)
        candidates = sap_idx.get(pn_key, [])

        # VG-PN aus Dokument-Tags (nur Rückblick hier)
        tags = (r.get("extras", {}) or {}).get("all_tags", {}) or r.get("tags", {})
        vg_tag = str(tags.get("rb_pn_vg", "") or "").strip()
        vg_key = _norm_pn(vg_tag)

        chosen = None
        if candidates and vg_key:
            for row in candidates:
                row_vg = _norm_pn(row.get("Dir. Vorgesetzter (PN)", ""))
                if row_vg == vg_key:
                    chosen = row
                    break
        if chosen is None and candidates:
            # Fallback: nimm die erste Zeile zur PN, wenn kein VG-Match
            chosen = candidates[0]

        # Eintritt/Austritt/Ans. aus gewählter Zeile
        eintritt = pd.to_datetime(chosen.get("Eintritt"), errors="coerce") if chosen is not None else pd.NaT
        austritt = pd.to_datetime(chosen.get("Austritt"), errors="coerce") if chosen is not None else pd.NaT
        ans = str(chosen.get("Ans.", "") or "").strip() if chosen is not None else ""

        # Dokument-Daten (Rückblick)
        doc_von = _d(tags.get("rb_datum_von", ""))
        doc_bis = _d(tags.get("rb_datum_bis", ""))
        datum_mab = _d(tags.get("rb_datum_gespraech", ""))

        # Rückblick-Jahr bestimmen
        if not pd.isna(doc_von):
            rb_year = doc_von.year
        elif not pd.isna(doc_bis):
            rb_year = doc_bis.year
        else:
            rb_year = pd.Timestamp.today().year

        # IT9075-Grenzen (RB-Jahr ∩ Beschäftigung)
        year_start = pd.Timestamp(year=rb_year, month=1, day=1)
        year_end   = pd.Timestamp(year=rb_year, month=12, day=31)

        begin_it9075 = year_start
        if not pd.isna(eintritt) and eintritt > begin_it9075:
            begin_it9075 = eintritt

        end_it9075 = year_end
        if not pd.isna(austritt) and austritt < end_it9075:
            end_it9075 = austritt

        # Beurteilungszeitraum: an IT9075 klemmen + Fallback
        period_von = doc_von if not pd.isna(doc_von) else begin_it9075
        if not pd.isna(period_von) and not pd.isna(begin_it9075):
            period_von = max(period_von, begin_it9075)

        period_bis = doc_bis if not pd.isna(doc_bis) else end_it9075
        if not pd.isna(period_bis) and not pd.isna(end_it9075):
            period_bis = min(period_bis, end_it9075)

        # Sanity: Beschäftigung gar nicht im RB-Jahr (Begin > End)
        if (not pd.isna(begin_it9075)) and (not pd.isna(end_it9075)) and (begin_it9075 > end_it9075):
            # Kein gültiges Intervall -> neutralisieren (optional: loggen)
            begin_it9075 = pd.NaT
            end_it9075 = pd.NaT
            # Fallbacks lassen wir stehen; wenn doc_von/bis leer, bleiben sie bereits auf NaT bzw. IT9075

        # Gesamteindruck (A–E) aus Plausi
        gesamt = (r.get("extras", {}) or {}).get("rb_gesamteindruck", "")

        rows.append({
            "PersNr": pn,
            "Beurteilungsart": "1",
            "Beginndatum IT9075": begin_it9075,
            "Endedatum IT9075": end_it9075,
            "Ans.": ans,
            "Datum MAB": datum_mab,
            "Beurteilungszeitraum von": period_von,
            "Beurteilungszeitraum bis": period_bis,
            "Gesamtbeurteilung": gesamt,
            "Zielerreichung": "",
            "Fachliche Kompetenz": "",
            "Sozialkompetenz (Verhalten)": "",
            "Führungskompetenz": "",
            "Nächster Termin": pd.NaT,
        })

    df_out = pd.DataFrame(rows, columns=SAP_COLS)
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    # Excel mit lokalem dd.mm.yyyy Zellformat
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as wr:
        df_out.to_excel(wr, index=False, sheet_name="Massenupload")
        ws = wr.sheets["Massenupload"]
        date_cols = ["Beginndatum IT9075","Endedatum IT9075","Datum MAB","Beurteilungszeitraum von","Beurteilungszeitraum bis","Nächster Termin"]
        from openpyxl.styles import numbers
        fmt = numbers.BUILTIN_FORMATS[14]  # 'm/d/yy' Basis, wir überschreiben auf deutsches
        for col_name in date_cols:
            if col_name in SAP_COLS:
                col_idx = SAP_COLS.index(col_name) + 1
                for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2):
                    for c in cell:
                        c.number_format = "DD.MM.YYYY"


def export_ds_csv(results: list[dict], out_csv: Path, sap_df: pd.DataFrame = None):
    """
    Breiter DS-Export: eine Zeile je Dokument, alle Tags (soweit vorhanden).
    Erweitert um Hierarchie-Informationen aus SAP Stammdaten.
    """
    rows = []
    
    # Hierarchie-Informationen vorbereiten falls SAP-Daten verfügbar
    hierarchy_data = {}
    if sap_df is not None:
        try:
            from org_structure import build_org_structure
            org_df = build_org_structure(sap_df)
            # Index für schnelle Suche nach PN
            hierarchy_data = org_df.set_index("Personalnummer").to_dict("index")
        except Exception as e:
            print(f"Warnung: Hierarchie-Daten konnten nicht geladen werden: {e}")
    
    for r in results:
        base = {
            "file": r.get("file", ""),
            "typ": r.get("typ", ""),
            "pn": r.get("pn", ""),
            "name": r.get("name", ""),
            "status": r.get("status", ""),
        }
        # nimm bevorzugt all_tags, sonst tags
        tags = (r.get("extras", {}) or {}).get("all_tags", {}) or r.get("tags", {})
        # flach mergen
        row = {**base, **tags}
        
        # Extras, die interessant sind
        if r.get("typ") == DocType.RUECKBLICK.value:
            row["rb_gesamteindruck_code"] = (r.get("extras", {}) or {}).get("rb_gesamteindruck", "")
        
        # Hierarchie-Informationen hinzufügen
        pn = r.get("pn", "")
        if pn and pn in hierarchy_data:
            h_data = hierarchy_data[pn]
            
            # Basis-Hierarchie-Informationen
            base_hierarchy = {
                "hierarchie_ebene": h_data.get("Hierarchie_Ebene", ""),
                "oe_kurz": h_data.get("OE_Kurz", ""),
                "oe_bez": h_data.get("OE_Bez", ""),
                "position": h_data.get("Position", ""),
                "vg_pn": h_data.get("Vorgesetzter_PN", ""),
                "vg_name": h_data.get("Vorgesetzter_Name", ""),
                "vg_oe": h_data.get("Vorgesetzter_OE", ""),
                "org_pfad": h_data.get("Organisations_Pfad", ""),
                "oe_hierarchie": h_data.get("OE_Hierarchie", ""),
                "oe_kette": h_data.get("OE_Kette", ""),
                "oe_bez_kette": h_data.get("OE_Bez_Kette", ""),
                "status_ma": h_data.get("Status", "")
            }
            
            # Nur Basis-Hierarchie und Ketten (keine org_level_* Spalten)
            row.update(base_hierarchy)
        else:
            # Leere Werte für fehlende Hierarchie-Daten
            empty_values = {
                "hierarchie_ebene": "",
                "oe_kurz": "",
                "oe_bez": "",
                "position": "",
                "vg_pn": "",
                "vg_name": "",
                "vg_oe": "",
                "org_pfad": "",
                "oe_hierarchie": "",
                "oe_kette": "",
                "oe_bez_kette": "",
                "status_ma": ""
            }
            
            row.update(empty_values)
        
        rows.append(row)

    df = pd.DataFrame(rows)
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out_csv, index=False, encoding="utf-8-sig")


def move_after_processing(input_dir: Path, results: list[dict]):
    """
    Verschiebt Dateien je nach Status:
      - OK  -> /ruecklauf/verarbeitet
      - manuell -> /ruecklauf/unverarbeitet/manuell
    Achtung: nur DOCX (hier); PDFs kommen im separaten Schritt.
    """
    import shutil
    
    verarbeitet_dir = input_dir / MDConstants.VERARBEITET_DIR
    manuell_dir = input_dir / MDConstants.MANUELL_DIR
    verarbeitet_dir.mkdir(parents=True, exist_ok=True)
    manuell_dir.mkdir(parents=True, exist_ok=True)
    tracking = SimpleTrackingSystem()

    def _unique_path(base_dir: Path, name: str) -> Path:
        target = base_dir / name
        if not target.exists():
            return target
        stem, suffix = Path(name).stem, Path(name).suffix
        i = 1
        while True:
            cand = base_dir / f"{stem} ({i}){suffix}"
            if not cand.exists():
                return cand
            i += 1

    moved_ok = moved_man = 0
    for r in results:
        fname = r.get("file")
        if not fname:
            continue
        src = input_dir / fname
        if not src.exists() or src.suffix.lower() != MDConstants.ALLOWED_EXTENSIONS[0]:
            continue

        if r.get("status") == ProcStatus.OK.value:
            dst = _unique_path(verarbeitet_dir, fname)
            shutil.move(str(src), str(dst))
            moved_ok += 1
            
            # Tracking: Markiere als erhalten
            pn = r.get("pn", "")
            if pn:
                tracking.mark_received(fname, pn, "word")
        elif r.get("status") in (ProcStatus.MANUELL.value, ProcStatus.PRUEFUNG_NOETIG.value):
            dst = _unique_path(manuell_dir, fname)
            shutil.move(str(src), str(dst))
            moved_man += 1
            
            # Tracking: Markiere als empfangen aber fehlerhaft
            pn = r.get("pn", "")
            if pn:
                tracking.mark_received(fname, pn, "word")
                tracking.mark_error(fname, pn, "word", r.get("reason", "Unbekannter Fehler"))

    return moved_ok, moved_man


# PDF Verarbeitung - Teil 1: Hilfsfunktionen und Hauptlogik
def _normalize(s: str) -> str:
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", s)
    return s.encode("ascii", "ignore").decode("ascii").lower()


def _ensure_pn_suffix(filename: str, pn: str) -> str:
    """Hängt _<pn> vor die .pdf-Endung, falls noch nicht vorhanden.
    Verändert den Rest des Namens nicht.
    """
    if not pn:
        return filename
    stem, ext = Path(filename).stem, Path(filename).suffix
    if ext.lower() != MDConstants.ALLOWED_EXTENSIONS[1]:
        ext = Path(filename).suffix  # unberührt
    # Bereits korrekt am Ende?
    if re.search(rf"_(?:{re.escape(pn)})$", stem):
        return filename
    # Falls Stem bereits mit PN endet (ohne Unterstrich), trotzdem standardisieren
    if re.search(rf"(?:^|[^0-9]){re.escape(pn)}$", stem):
        # entferne evtl. vorhandenes PN am Ende ohne Unterstrich
        stem = re.sub(rf"{re.escape(pn)}$", "", stem).rstrip("_")
    return f"{stem}_{pn}{ext}"


def process_pdfs(in_dir: Path, out_root: Path, sap_df: pd.DataFrame, durchlauf_jahr: int | None = None) -> list[dict]:
    """
    Bearbeitet PDFs im Eingangsordner:
    - Typ erkennen (Rückblick/Ausblick/Feedback/Probezeit)
    - PN & Namen aus Dateiname oder SAP ableiten
    - Zielname bilden
    - Datei verschieben
    Gibt Liste mit Log-Einträgen zurück.
    """
    results = []
    in_dir = Path(in_dir)
    out_root = Path(out_root)
    tracking = SimpleTrackingSystem()

    def _ensure_pn_suffix(filename: str, pn: str) -> str:
        """Hängt _<pn> vor die .pdf-Endung, falls noch nicht vorhanden.
        Verändert den Rest des Namens nicht.
        """
        if not pn:
            return filename
        stem, ext = Path(filename).stem, Path(filename).suffix
        if ext.lower() != MDConstants.ALLOWED_EXTENSIONS[1]:
            ext = Path(filename).suffix  # unberührt
        # Bereits korrekt am Ende?
        if re.search(rf"_(?:{re.escape(pn)})$", stem):
            return filename
        # Falls Stem bereits mit PN endet (ohne Unterstrich), trotzdem standardisieren
        if re.search(rf"(?:^|[^0-9]){re.escape(pn)}$", stem):
            # entferne evtl. vorhandenes PN am Ende ohne Unterstrich
            stem = re.sub(rf"{re.escape(pn)}$", "", stem).rstrip("_")
        return f"{stem}_{pn}{ext}"

    for pdf_path in in_dir.glob(f"*{MDConstants.ALLOWED_EXTENSIONS[1]}"):
        fname = pdf_path.name
        norm = _normalize(fname)
        target_dir = out_root
        target_name = None
        status = "ok"
        typ: DocType | None = None

        try:
            # Typ bestimmen
            if MDConstants.MD_KEYWORDS[0] in norm or MDConstants.MD_KEYWORDS[1] in norm:
                typ = DocType.RUECKBLICK
            elif MDConstants.MD_KEYWORDS[2] in norm:
                typ = DocType.AUSBLiCK
            elif MDConstants.MD_KEYWORDS[3] in norm:
                typ = DocType.FEEDBACK
            elif MDConstants.PROBEZEIT_KEYWORD in norm:
                typ = DocType.PROBEZEIT
                # Kein spezieller Probezeit-Ordner: bleibt im Standardziel

            # PN aus Dateiname ziehen (6-stellige Zahl, meist am Ende)
            # Sucht nach 6-stelligen Zahlen, isoliert oder mit _ davor/danach
            pn_match = re.search(r'(?:^|_)(\d{6})(?:_|\.|$)', fname)
            pn = pn_match.group(1) if pn_match else ""

            if not (pn and pn in sap_df["ID_NO_ZERO"].astype(str).values):
                status = ProcStatus.PRUEFUNG_NOETIG.value
                target_dir = Path(__file__).parent.parent / MDConstants.RUECKLAUF_DIR / MDConstants.UNVERARBEITET_DIR / MDConstants.MANUELL_DIR
                
                # Tracking: Markiere als empfangen aber fehlerhaft (nur wenn PN vorhanden)
                if pn and typ in (DocType.RUECKBLICK, DocType.AUSBLiCK):
                    # Versuche VG-PN zu finden, auch wenn PN nicht in SAP ist
                    vg_pn = ""
                    try:
                        # Suche nach ähnlichen PN in SAP (falls PN leicht abweicht)
                        similar_pns = sap_df[sap_df["ID_NO_ZERO"].astype(str).str.contains(pn, na=False)]
                        if not similar_pns.empty:
                            vg_pn = str(similar_pns.iloc[0].get("Dir. Vorgesetzter (PN)", "")).strip()
                    except:
                        pass
                    
                    if vg_pn:
                        doc_type = f"{typ.value} PDF"
                        tracking.mark_received(vg_pn, pn, doc_type)
                        tracking.mark_error(fname, pn, doc_type, "PN nicht in SAP-Daten gefunden")

            # Zielpfad festlegen
            target_dir.mkdir(parents=True, exist_ok=True)
            # Grundsätzlich Beschriftung nicht ändern, nur PN am Ende sicherstellen
            ensured_name = _ensure_pn_suffix(fname, pn)
            dest = target_dir / ensured_name

            # Duplikat-Erkennung (nur für Rückblick/Ausblick)
            if typ in (DocType.RUECKBLICK, DocType.AUSBLiCK) and pn:
                # Prüfe auf Mehrfachanstellungen
                matching_rows = sap_df[sap_df["ID_NO_ZERO"].astype(str) == pn]
                anzahl_anstellungen = len(matching_rows)
                
                if anzahl_anstellungen > 1:
                    # Mehrfachanstellung: Manuelle Prüfung erforderlich
                    status = ProcStatus.PRUEFUNG_NOETIG.value
                    target_dir = Path(__file__).parent.parent / MDConstants.RUECKLAUF_DIR / MDConstants.UNVERARBEITET_DIR / MDConstants.MANUELL_DIR
                    dest = target_dir / fname
                    target_dir.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(pdf_path), dest)
                    
                    # Tracking: Markiere als empfangen aber fehlerhaft
                    vg_pn = str(matching_rows.iloc[0].get("Dir. Vorgesetzter (PN)", "")).strip()
                    if vg_pn:
                        doc_type = f"{typ.value} PDF"
                        tracking.mark_received(vg_pn, pn, doc_type)
                        tracking.mark_error(fname, pn, doc_type, f"Mehrfachanstellung: {anzahl_anstellungen} Anstellungen gefunden")
                    
                    results.append({
                        "file": fname,
                        "typ": typ,
                        "pn": pn,
                        "name": f"{matching_rows.iloc[0].get('Rufname','')} {matching_rows.iloc[0].get('Nachname','')}".strip(),
                        "status": ProcStatus.MANUELL.value,
                        "reason": f"Mehrfachanstellung: {anzahl_anstellungen} Anstellungen gefunden - manuelle Zuordnung erforderlich",
                        "target": str(dest),
                        "extras": {"all_tags": {}, "anzahl_anstellungen": anzahl_anstellungen}
                    })
                    continue
                elif anzahl_anstellungen == 1:
                    # Einzelanstellung: Normale Duplikat-Prüfung
                    row = matching_rows.iloc[0]
                    vg_pn = str(row.get("Dir. Vorgesetzter (PN)", "")).strip()
                    
                    doc_type = f"{typ.value} PDF"
                    is_duplicate, warning = tracking.check_duplicate(fname, pn, doc_type, vg_pn)
                    if is_duplicate:
                        status = ProcStatus.PRUEFUNG_NOETIG.value
                        target_dir = Path(__file__).parent.parent / "ruecklauf" / "unverarbeitet" / "manuell"  # Korrekte Pfad für manuelle Prüfung
                        dest = target_dir / fname
                        target_dir.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(pdf_path), dest)
                        
                        # Tracking: Markiere als empfangen aber fehlerhaft
                        if vg_pn:
                            tracking.mark_received(vg_pn, pn, doc_type)
                            tracking.mark_error(fname, pn, doc_type, f"Duplikat: {warning}")
                        
                        results.append({
                            "file": fname,
                            "typ": typ,
                            "pn": pn,
                            "name": f"{row.get('Rufname','')} {row.get('Nachname','')}".strip(),
                            "status": ProcStatus.MANUELL.value,
                            "reason": f"Duplikat: {warning}",
                            "target": str(dest),
                            "extras": {"all_tags": {}}
                        })
                        continue
                else:
                    # Keine SAP-Daten gefunden
                    status = ProcStatus.PRUEFUNG_NOETIG.value
                    target_dir = Path(__file__).parent.parent / MDConstants.RUECKLAUF_DIR / MDConstants.UNVERARBEITET_DIR / MDConstants.MANUELL_DIR
                    dest = target_dir / fname
                    target_dir.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(pdf_path), dest)
                    
                    # Tracking: Markiere als empfangen aber fehlerhaft (nur wenn PN vorhanden)
                    if pn and typ in (DocType.RUECKBLICK, DocType.AUSBLiCK):
                        # Versuche VG-PN zu finden, auch wenn PN nicht in SAP ist
                        vg_pn = ""
                        try:
                            # Suche nach ähnlichen PN in SAP (falls PN leicht abweicht)
                            similar_pns = sap_df[sap_df["ID_NO_ZERO"].astype(str).str.contains(pn, na=False)]
                            if not similar_pns.empty:
                                vg_pn = str(similar_pns.iloc[0].get("Dir. Vorgesetzter (PN)", "")).strip()
                        except:
                            pass
                        
                        if vg_pn:
                            doc_type = f"{typ.value} PDF"
                            tracking.mark_received(vg_pn, pn, doc_type)
                            tracking.mark_error(fname, pn, doc_type, "PN nicht in SAP-Daten gefunden")
                    
                    results.append({
                        "file": fname,
                        "typ": typ.value if isinstance(typ, DocType) else str(typ),
                        "pn": pn,
                        "name": "",
                        "status": ProcStatus.MANUELL.value,
                        "reason": "PN nicht in SAP-Daten gefunden",
                        "target": str(dest),
                        "extras": {"all_tags": {}}
                    })
                    continue

            # Duplikate behandeln
            counter = 1
            while dest.exists():
                stem, ext = dest.stem, dest.suffix
                dest = target_dir / f"{stem}_{counter}{ext}"
                counter += 1

            shutil.move(str(pdf_path), dest)

            # Name aus SAP-Daten ermitteln
            name = ""
            if pn and pn in sap_df["ID_NO_ZERO"].astype(str).values:
                row = sap_df[sap_df["ID_NO_ZERO"].astype(str) == pn].iloc[0]
                name = f"{row.get('Rufname','')} {row.get('Nachname','')}".strip()

            # Tracking-System aktualisieren für erfolgreiche Verarbeitung
            if status == ProcStatus.OK.value and pn:
                if typ in (DocType.RUECKBLICK, DocType.AUSBLiCK):
                    # VG-PN ermitteln
                    vg_pn = ""
                    if pn in sap_df["ID_NO_ZERO"].astype(str).values:
                        row = sap_df[sap_df["ID_NO_ZERO"].astype(str) == pn].iloc[0]
                        vg_pn = str(row.get("Dir. Vorgesetzter (PN)", "")).strip()
                    doc_type = f"{typ.value} PDF"
                    tracking.mark_received(vg_pn, pn, doc_type)
                elif typ == DocType.FEEDBACK:
                    # Für Feedback ist PN die VG-PN
                    tracking.mark_received(pn, "", "Feedback PDF")

            results.append({
                "file": fname,
                "typ": typ,
                "pn": pn,
                "name": name,
                "status": status,
                "reason": "" if status == ProcStatus.OK.value else ("PN nicht in SAP-Daten gefunden" if not pn else "Prüfung nötig"),
                "target": str(dest),
                "extras": {"all_tags": {}}
            })

        except Exception as e:
            results.append({
                "file": fname,
                "typ": typ or "Unbekannt",
                "pn": "",
                "name": "",
                "status": ProcStatus.PRUEFUNG_NOETIG.value,
                "reason": f"Fehler bei Verarbeitung: {e}",
                "target": "",
                "extras": {"all_tags": {}}
            })
            manuell_dir = Path(__file__).parent.parent / MDConstants.RUECKLAUF_DIR / MDConstants.UNVERARBEITET_DIR / MDConstants.MANUELL_DIR
            manuell_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(pdf_path), (manuell_dir / fname))

    # Logging nach ruecklauf/logs/processing_log.csv
    if durchlauf_jahr is not None:
        _append_processing_log(results, durchlauf_jahr)

    return results


def run_full_processing(app) -> None:
    """Controller: Führt den dreistufigen Verarbeitungslauf aus."""
    try:
        if hasattr(app, "proc_status"):
            app.proc_status.config(text="Schritt 1/3: DOCX prüfen...", foreground="black")
            app.update_idletasks()
        process_docx(app)

        if hasattr(app, "proc_status"):
            app.proc_status.config(text="Schritt 2/3: Export (SAP+DS) & verschieben...", foreground="black")
            app.update_idletasks()
        export_and_move(app)

        if hasattr(app, "proc_status"):
            app.proc_status.config(text="Schritt 3/3: PDFs verarbeiten...", foreground="black")
            app.update_idletasks()
        process_pdfs_run(app)

        if hasattr(app, "proc_status"):
            app.proc_status.config(text="Gesamtverarbeitung abgeschlossen.", foreground="black")
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Verarbeitung abgebrochen: {e}")


def process_docx(app) -> None:
    """Controller: Prüft DOCX im Rücklauf und befüllt die DOCX-Tabelle im UI."""
    input_dir = Path(__file__).parent.parent / MDConstants.RUECKLAUF_DIR / MDConstants.UNVERARBEITET_DIR

    # Tabelle leeren
    for i in app.tree_proc.get_children():
        app.tree_proc.delete(i)

    # Stammdaten laden
    sap_df = load_employees()

    # Prozess laufen lassen (Batch)
    try:
        max_files = int(app.batch_size_var.get()) if app.batch_size_var.get() else None
    except Exception:
        max_files = None
    durchlauf_jahr = app.proc_year_var.get()
    results = process_docx_folder(input_dir, sap_df, max_files=max_files, durchlauf_jahr=durchlauf_jahr)

    ok_count = 0
    man_count = 0

    for r in results:
        status = r.get("status", "")
        if status == ProcStatus.OK.value:
            target = "verarbeitet"
        elif status in (ProcStatus.MANUELL.value, ProcStatus.PRUEFUNG_NOETIG.value):
            target = "manuell"
        else:
            target = ""

        app.tree_proc.insert(
            "",
            "end",
            values=[
                r.get("file",""),
                r.get("typ",""),
                r.get("pn",""),
                r.get("name",""),
                r.get("status",""),
                r.get("reason",""),
                target,
            ],
        )
        if status == ProcStatus.OK.value:
            ok_count += 1
        elif status == ProcStatus.MANUELL.value:
            man_count += 1

    # Spaltenbreiten an Inhalt anpassen
    try:
        autosize_tree_columns(app.tree_proc)
    except Exception:
        pass

    app.proc_status.config(
        text=f"DOCX geprüft: {len(results)} Dateien • OK: {ok_count} • Manuell: {man_count}",
        foreground="black",
    )

    app._last_docx_results = results


def export_and_move(app) -> None:
    """Controller: Schreibt Exporte (SAP, DS) und verschiebt DOCX gemäß Status."""
    input_dir = Path(__file__).parent.parent / MDConstants.RUECKLAUF_DIR / MDConstants.UNVERARBEITET_DIR
    sap_out = Path(__file__).parent.parent / "sap_massenupload" / "massenupload.xlsx"
    ds_out = Path(__file__).parent.parent / "tracking" / "ds_export" / "docx_extract.csv"

    # 1) Falls keine Ergebnisse, DOCX scan ausführen
    if not hasattr(app, "_last_docx_results"):
        sap_df = load_employees()
        app._last_docx_results = process_docx_folder(input_dir, sap_df)

    results = app._last_docx_results

    # 2) Exporte schreiben
    try:
        sap_df = load_employees()
        export_sap_massenupload(results, sap_df, sap_out)
        export_ds_csv(results, ds_out, sap_df)
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Export fehlgeschlagen:\n{e}")
        return

    # 3) Verschieben
    try:
        moved_ok, moved_man = move_after_processing(input_dir, results)
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Verschieben fehlgeschlagen:\n{e}")
        return

    messagebox.showinfo(
        MDConstants.MSG_FINISHED,
        f"Export geschrieben:\n- SAP: {sap_out}\n- DS:  {ds_out}\n\n"
        f"Verschoben:\n- OK → {MDConstants.VERARBEITET_DIR}: {moved_ok}\n- manuell → {MDConstants.MANUELL_DIR}: {moved_man}"
    )


def process_pdfs_run(app) -> None:
    """Controller: Verarbeitet PDFs im Rücklauf und aktualisiert die PDF-Tabelle im UI."""
    in_dir = Path(__file__).parent.parent / MDConstants.RUECKLAUF_DIR / MDConstants.UNVERARBEITET_DIR
    out_root = Path(app.rpa_target_var.get())

    if hasattr(app, "sap_df"):
        sap_df = app.sap_df
    else:
        sap_df = load_employees()
        app.sap_df = sap_df

    try:
        max_files = int(app.batch_size_var.get()) if app.batch_size_var.get() else None
    except Exception:
        max_files = None

    all_pdfs = sorted(in_dir.glob(f"*{MDConstants.ALLOWED_EXTENSIONS[1]}"))
    process_list = all_pdfs[:max_files] if max_files is not None else all_pdfs

    temp_dir = in_dir / "_batch_tmp"
    temp_dir.mkdir(parents=True, exist_ok=True)
    moved: list[Path] = []
    for p in process_list:
        tgt = temp_dir / p.name
        try:
            p.rename(tgt)
            moved.append(tgt)
        except Exception:
            pass

    durchlauf_jahr = app.proc_year_var.get()
    results = process_pdfs(temp_dir, out_root, sap_df, durchlauf_jahr=durchlauf_jahr)

    # Tabelle leeren
    for item in app.tree_pdfs.get_children():
        app.tree_pdfs.delete(item)

    for r in results:
        app.tree_pdfs.insert(
            "",
            "end",
            values=(
                r.get("file", ""),
                r.get("typ", ""),
                r.get("pn", ""),
                r.get("name", ""),
                r.get("status", ""),
                r.get("reason", ""),
                r.get("target", ""),
            ),
        )

    try:
        autosize_tree_columns(app.tree_pdfs)
    except Exception:
        pass

    messagebox.showinfo(MDConstants.MSG_FINISHED, f"{len(results)} PDFs verarbeitet.")

    # Aufräumen: restliche Dateien zurückschieben, falls vorhanden
    for p in temp_dir.glob(f"*{MDConstants.ALLOWED_EXTENSIONS[1]}"):
        try:
            p.rename(in_dir / p.name)
        except Exception:
            pass


