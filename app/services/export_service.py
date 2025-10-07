"""
Export-Service für das MD-Prozess-Tool.

Dieser Service kapselt alle Funktionen zur Datenexport,
einschließlich SAP-Massenupload und DS-Export.
"""

from __future__ import annotations
from pathlib import Path
import pandas as pd
from typing import Dict, Any

from app.constants import MDConstants, DocType, ProcStatus

# SAP-Spalten für Export
SAP_COLS = [
    "ID_NO_ZERO", "Rufname", "Nachname", "OE Bez.", "OE Kurzb.", 
    "Plans. Bez.", "lange ID/Nummer", "Dir. Vorgesetzter (PN)", "BsGrd"
]


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
    
    Exportiert nur:
    - DOCX-Dokumente (keine PDFs)
    - Erfolgreich verarbeitete Dokumente (Status = OK)
    - Daten werden angehängt (append mode), nicht überschrieben
    """
    rows = []
    
    # Hierarchie-Informationen vorbereiten falls SAP-Daten verfügbar
    hierarchy_data = {}
    if sap_df is not None:
        try:
            from app.services.org_structure_service import build_org_structure
            org_df = build_org_structure(sap_df)
            # Index für schnelle Suche nach PN
            hierarchy_data = org_df.set_index("Personalnummer").to_dict("index")
        except Exception as e:
            print(f"Warnung: Hierarchie-Daten konnten nicht geladen werden: {e}")
    
    for r in results:
        # Nur DOCX-Dokumente (keine PDFs) und nur erfolgreich verarbeitete (Status = OK)
        fname = r.get("file", "").lower()
        status = r.get("status", "")
        
        # Filtere PDFs und nicht-OK Status aus
        if fname.endswith(".pdf") or status != ProcStatus.OK.value:
            continue
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

    # Nur schreiben wenn Daten vorhanden
    if not rows:
        return
    
    df_new = pd.DataFrame(rows)
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    
    # Prüfen ob Datei existiert -> Append-Modus mit Spalten-Vereinigung
    if out_csv.exists():
        # Alte Daten einlesen
        try:
            df_old = pd.read_csv(out_csv, encoding="utf-8-sig")
            
            # Neue und alte Daten kombinieren
            # pd.concat vereinigt automatisch alle Spalten (union)
            # Fehlende Werte werden mit NaN gefüllt
            df_combined = pd.concat([df_old, df_new], ignore_index=True)
            
            # Komplett neu schreiben mit allen Spalten
            df_combined.to_csv(out_csv, mode='w', header=True, index=False, encoding="utf-8-sig")
        except Exception as e:
            # Fallback: Bei Fehler beim Einlesen nur neue Daten anhängen
            print(f"Warnung: Fehler beim Einlesen der bestehenden CSV: {e}")
            df_new.to_csv(out_csv, mode='a', header=False, index=False, encoding="utf-8-sig")
    else:
        # Neu erstellen mit Header
        df_new.to_csv(out_csv, mode='w', header=True, index=False, encoding="utf-8-sig")


def generate_export_data(results: list[dict], sap_df: pd.DataFrame) -> Dict[str, Any]:
    """
    Generiert Export-Daten für verschiedene Formate.
    """
    # Optional: Kann später implementiert werden
    pass
