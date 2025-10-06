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
from logging_config import get_logger

from adapters.docx_reader import read_content_controls, detect_doc_type, map_rb_gesamteindruck
from services.tracking_service import SimpleTrackingSystem
from constants import MDConstants, DocType, ProcStatus
from services.org_structure_service import build_org_structure
from data_loader import load_employees, load_config
from views.ui_utils import autosize_tree_columns
from services.export_service import export_sap_massenupload, export_ds_csv

# Konfiguration einmalig laden für zentrale Pfade
CFG = load_config()
logger = get_logger()
from services.file_service import move_after_processing


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

    # Protokollverzeichnis aus Konfiguration beziehen
    log_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["logs_dir"]
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
            logger.warning("DOCX Lesefehler", extra={"file": p.name, "error": str(e)})
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

    log_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["logs_dir"]
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
                target_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["manuell"]
                
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

            # Zielpfad festlegen (vorläufig)
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
                    target_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["manuell"]
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
                    target_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["manuell"]
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

            # Business-Logik für Zielordner bei erfolgreichen PDFs:
            # - Rückblick/Ausblick: in RPA-Ziel (out_root)
            # - Feedback: in projektweiten Ordner ruecklauf/feedbacks
            if status == ProcStatus.OK.value and typ == DocType.FEEDBACK:
                # Feedback-PDFs projektweit unter <root>/ruecklauf/feedbacks ablegen
                fb_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["feedbacks"]
                fb_dir.mkdir(parents=True, exist_ok=True)
                # PN ggf. am Ende sicherstellen
                ensured_name = _ensure_pn_suffix(fname, pn)
                dest = fb_dir / ensured_name

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

                logger.info("PDF erfolgreich verschoben", extra={"file": fname, "target": str(dest)})
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
            logger.error("PDF-Verarbeitungsfehler", extra={"file": fname, "error": str(e)})
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
            # Manuelle PDFs projektweit unter <root>/ruecklauf/unverarbeitet/manuell ablegen
            manuell_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["manuell"]
            manuell_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(pdf_path), (manuell_dir / fname))

    # Logging nach ruecklauf/logs/processing_log.csv
    if durchlauf_jahr is not None:
        _append_processing_log(results, durchlauf_jahr)

    return results


def run_full_processing(app) -> None:
    """Controller: Führt den dreistufigen Verarbeitungslauf aus."""
    # Controller ist verantwortlich für UI/Fehleranzeigen; hier nur Exceptions werfen
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


def process_docx(app) -> None:
    """Controller: Prüft DOCX im Rücklauf und befüllt die DOCX-Tabelle im UI."""
    # DOCX-Eingang ist projektweit <root>/ruecklauf/unverarbeitet
    input_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["unverarbeitet"]

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
    input_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["unverarbeitet"]
    sap_out = Path(__file__).parent.parent / "sap_massenupload" / "massenupload.xlsx"
    ds_out = Path(__file__).parent.parent / "tracking" / "ds_export" / "docx_extract.csv"

    # 1) Falls keine Ergebnisse, DOCX scan ausführen
    if not hasattr(app, "_last_docx_results"):
        sap_df = load_employees()
        app._last_docx_results = process_docx_folder(input_dir, sap_df)

    results = app._last_docx_results

    # 2) Exporte schreiben
    sap_df = load_employees()
    export_sap_massenupload(results, sap_df, sap_out)
    export_ds_csv(results, ds_out, sap_df)

    # 3) Verschieben
    moved_ok, moved_man = move_after_processing(input_dir, results)

    # Erfolgsmeldung wird im Controller/GUI angezeigt


def process_pdfs_run(app) -> None:
    """Controller: Verarbeitet PDFs im Rücklauf und aktualisiert die PDF-Tabelle im UI."""
    in_dir = Path(__file__).parent / CFG["paths"]["ruecklauf"]["unverarbeitet"]
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

    # Erfolgsmeldung im Controller anzeigen

    # Aufräumen: restliche Dateien zurückschieben, falls vorhanden
    for p in temp_dir.glob(f"*{MDConstants.ALLOWED_EXTENSIONS[1]}"):
        try:
            p.rename(in_dir / p.name)
        except Exception:
            pass


