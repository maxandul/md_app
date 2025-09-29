# ---------------------------
# Word lesen & Mappings
# ---------------------------

from __future__ import annotations
from pathlib import Path
from typing import Dict
from docx import Document

# Namespace für DOCX-XML
NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Dropdown-Anzeigenamen -> Wert
RB_GESAMTEINDRUCK_MAP = {
    "vorzüglich": "A",
    "sehr gut": "B",
    "gut": "C",
    "genügend": "D",
    "ungenügend": "E",
    # falls jemand direkt A-E einträgt:
    "a": "A", "b": "B", "c": "C", "d": "D", "e": "E",
}

def _norm(s: str) -> str:
    return (s or "").strip().lower()

def _is_empty_or_stop_content(text: str) -> bool:
    """
    Prüft ob Text nur aus leeren Zeichen oder Stoppzeichen besteht.
    Überspringt Inhalte mit nur:
    - Geschützten Leerzeichen (U+00A0)
    - Bindestrichen (-)
    - Leerzeichen
    - Tabs
    - Zeilenumbrüchen
    - Spezifische Platzhalter-Texte
    """
    if not text:
        return True
    
    # Spezifische Platzhalter-Texte ignorieren
    stop_phrases = [
        "Bitte wählen.",
        "Bitte wählen",
        "Datum",
        "dd.mm.yyyy",
        "MM/DD/YYYY",
        "TT.MM.JJJJ",
        "Bitte auswählen",
        "Bitte auswählen.",
        "Auswählen",
        "Wählen"
    ]
    
    text_normalized = text.strip()
    for phrase in stop_phrases:
        if phrase.lower() in text_normalized.lower():
            return True
    
    # Entferne alle Stoppzeichen und prüfe ob etwas übrig bleibt
    cleaned = text.replace('\u00A0', '')  # Geschütztes Leerzeichen
    cleaned = cleaned.replace('-', '')     # Bindestrich
    cleaned = cleaned.replace(' ', '')    # Normale Leerzeichen
    cleaned = cleaned.replace('\t', '')   # Tabs
    cleaned = cleaned.replace('\n', '')   # Zeilenumbrüche
    cleaned = cleaned.replace('\r', '')   # Carriage Return
    
    return len(cleaned.strip()) == 0

def map_rb_gesamteindruck(value_text: str) -> str:
    """Nimmt den Anzeigenamen ODER bereits A-E und gibt A-E zurück, sonst ''."""
    return RB_GESAMTEINDRUCK_MAP.get(_norm(value_text), "")

def read_content_controls(docx_path: str | Path) -> Dict[str, str]:
    """
    Liest alle Content Controls (Inhaltssteuerelemente) aus einer DOCX-Datei.
    - Verwendet das Tag-Attribut als Schlüssel.
    - Funktioniert ohne Word/COM, direkt über python-docx.
    - Leere Felder liefern ''.
    """
    values: Dict[str, str] = {}
    try:
        doc = Document(docx_path)
        for element in doc.element.body.iter():
            if element.tag.endswith("}sdt"):  # <w:sdt>
                sdt_pr = element.find(f".//{NS}sdtPr")
                sdt_content = element.find(f".//{NS}sdtContent")

                if sdt_pr is None or sdt_content is None:
                    continue

                # Tag oder Alias als Schlüssel nehmen
                tag_el = sdt_pr.find(f".//{NS}tag")
                alias_el = sdt_pr.find(f".//{NS}alias")

                key = None
                if tag_el is not None and tag_el.get(f"{NS}val"):
                    key = tag_el.get(f"{NS}val")
                elif alias_el is not None and alias_el.get(f"{NS}val"):
                    key = f"alias_{alias_el.get(f'{NS}val')}"

                if key:
                    # Text-Inhalt zusammensetzen
                    text_nodes = sdt_content.findall(f".//{NS}t")
                    text = " ".join(
                        t.text for t in text_nodes if t is not None and t.text
                    )
                    text = text.strip()
                    
                    # Überspringe Steuerelemente mit nur Stoppzeichen
                    if not _is_empty_or_stop_content(text):
                        values[key] = text
    except Exception as e:
        raise RuntimeError(f"DOCX beschädigt/Lesefehler: {e}")

    return values

def detect_doc_type(tags: dict) -> str:
    """
    Ermittelt Dokumenttyp grob anhand vorhandener Tags.
    'Rückblick' wenn rb_* vorkommt, 'Ausblick' wenn ab_*,
    sonst 'Unbekannt'.
    """
    has_rb = any(k.startswith("rb_") for k in tags.keys())
    has_ab = any(k.startswith("ab_") for k in tags.keys())
    if has_rb and not has_ab:
        return "Rückblick"
    if has_ab and not has_rb:
        return "Ausblick"
    # falls beides oder keines gefunden:
    return "Unbekannt"
