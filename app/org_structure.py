# app/org_structure.py
"""
Organisationsstruktur-Tool

Baut aus SAP Stammdaten eine Organisationsstruktur-Tabelle auf.
Verwendet die Hierarchie-Daten aus EXPORT.xlsx um:
- Organisationsbaum zu erstellen
- Hierarchie-Ebenen zu berechnen  
- OE-Struktur zu analysieren
- Management-Pfade zu visualisieren

Autor: HR-Team
Version: 1.0
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from collections import defaultdict, deque


def build_org_structure(sap_df: pd.DataFrame) -> pd.DataFrame:
    """
    Baut Organisationsstruktur-Tabelle aus SAP Stammdaten auf.
    
    Args:
        sap_df: DataFrame mit SAP Stammdaten (EXPORT.xlsx)
        
    Returns:
        DataFrame mit Organisationsstruktur-Informationen
    """
    if "ID_NO_ZERO" not in sap_df.columns or "Dir. Vorgesetzter (PN)" not in sap_df.columns:
        raise ValueError("Erforderliche Spalten 'ID_NO_ZERO' oder 'Dir. Vorgesetzter (PN)' fehlen")
    
    # Normalisiere Personalnummern
    df = sap_df.copy()
    df["ID_NO_ZERO"] = df["ID_NO_ZERO"].astype(str).str.strip()
    df["Dir. Vorgesetzter (PN)"] = df["Dir. Vorgesetzter (PN)"].astype(str).str.strip()
    
    # Erstelle Mitarbeiter-Index
    emp_index = {}
    for _, row in df.iterrows():
        pn = row["ID_NO_ZERO"]
        emp_index[pn] = {
            "name": f"{row.get('Rufname', '')} {row.get('Nachname', '')}".strip(),
            "oe_kurz": row.get("OE Kurzb.", ""),
            "oe_bez": row.get("OE Bez.", ""),
            "position": row.get("Plans. Bez.", ""),
            "vg_pn": row["Dir. Vorgesetzter (PN)"],
            "eintritt": row.get("Eintritt"),
            "austritt": row.get("Austritt")
        }
    
    # Berechne Hierarchie-Ebenen
    hierarchy_levels = _calculate_hierarchy_levels(emp_index)
    
    # Baue Organisationsstruktur-Tabelle
    org_rows = []
    for pn, emp_data in emp_index.items():
        vg_pn = emp_data["vg_pn"]
        vg_data = emp_index.get(vg_pn, {})
        
        # Hierarchie-Ebene
        level = hierarchy_levels.get(pn, -1)
        
        # Organisations-Pfad
        org_path = _build_org_path(pn, emp_index, hierarchy_levels)
        
        # OE-Hierarchie
        oe_hierarchy = _build_oe_hierarchy(emp_data["oe_kurz"], emp_index)

        # OE-Ketten (Top->Bottom) als Strings
        oe_chain_kurz, oe_chain_bez = _build_oe_chain(pn, emp_index)
        
        # OE-Ebenen aufbauen
        oe_levels = _build_oe_hierarchy_levels(pn, emp_index)
        
        # Basis-Daten
        row_data = {
            "Personalnummer": pn,
            "Name": emp_data["name"],
            "OE_Kurz": emp_data["oe_kurz"],
            "OE_Bez": emp_data["oe_bez"],
            "Position": emp_data["position"],
            "Hierarchie_Ebene": level,
            "Vorgesetzter_PN": vg_pn,
            "Vorgesetzter_Name": vg_data.get("name", ""),
            "Vorgesetzter_OE": vg_data.get("oe_kurz", ""),
            "Organisations_Pfad": org_path,
            "OE_Hierarchie": oe_hierarchy,
            "OE_Kette": oe_chain_kurz,
            "OE_Bez_Kette": oe_chain_bez,
            "Eintritt": emp_data["eintritt"],
            "Austritt": emp_data["austritt"],
            "Status": "Aktiv" if pd.isna(emp_data["austritt"]) else "Ausgeschieden"
        }
        
        # OE-Ebenen hinzufügen
        row_data.update(oe_levels)
        
        org_rows.append(row_data)
    
    return pd.DataFrame(org_rows)


def _calculate_hierarchy_levels(emp_index: Dict) -> Dict[str, int]:
    """
    Berechnet Hierarchie-Ebenen für alle Mitarbeiter.
    Level 0 = Top-Management (kein Vorgesetzter oder VG nicht in Daten)
    """
    levels = {}
    visited = set()
    
    def calculate_level(pn: str, path: set) -> int:
        if pn in path:  # Zirkuläre Referenz
            return 0
        if pn in levels:
            return levels[pn]
        
        emp_data = emp_index.get(pn)
        if not emp_data or not emp_data["vg_pn"] or emp_data["vg_pn"] not in emp_index:
            levels[pn] = 0  # Top-Level
            return 0
        
        vg_level = calculate_level(emp_data["vg_pn"], path | {pn})
        levels[pn] = vg_level + 1
        return levels[pn]
    
    for pn in emp_index:
        if pn not in visited:
            calculate_level(pn, set())
            visited.add(pn)
    
    return levels


def _build_org_path(pn: str, emp_index: Dict, levels: Dict[str, int]) -> str:
    """
    Baut Organisations-Pfad von Top-Level bis zum Mitarbeiter.
    """
    path = []
    current_pn = pn
    visited = set()
    
    while current_pn and current_pn not in visited:
        visited.add(current_pn)
        emp_data = emp_index.get(current_pn)
        if not emp_data:
            break
        
        name = emp_data["name"] or f"PN_{current_pn}"
        path.append(name)
        
        vg_pn = emp_data["vg_pn"]
        if not vg_pn or vg_pn not in emp_index:
            break
        current_pn = vg_pn
    
    return " > ".join(reversed(path))


def _build_oe_hierarchy_levels(emp_pn: str, emp_index: Dict, max_levels: int = 5) -> Dict[str, str]:
    """
    Baut OE-Hierarchie-Ebenen für einen Mitarbeiter basierend auf der Organisationsstruktur.
    
    Beispiel:
    - CEO: GL (ebene 0)
    - Abteilungsleiter A1: GL > A1 (ebene 1) 
    - Teamleiter1 A1: GL > A1 > T1 (ebene 2)
    - Mitarbeiter1 A1/T1: GL > A1 > T1 (ebene 3)
    
    Gibt Dict mit org_level_0, org_level_1, etc. zurück.
    """
    oe_levels = {}
    current_pn = emp_pn
    visited = set()
    
    # Sammle alle OE-Ebenen von Mitarbeiter bis zur Spitze
    oe_chain = []
    
    for level in range(max_levels):
        if not current_pn or current_pn in visited:
            break
            
        visited.add(current_pn)
        emp_data = emp_index.get(current_pn, {})
        oe_kurz = emp_data.get("oe_kurz", "")
        oe_bez = emp_data.get("oe_bez", "")
        
        if oe_kurz:
            oe_chain.append((oe_kurz, oe_bez))
        
        # Gehe zum Vorgesetzten
        vg_pn = emp_data.get("vg_pn")
        if not vg_pn or vg_pn not in emp_index:
            break
        current_pn = vg_pn
    
    # Baue Organisations-Ebenen auf (von unten nach oben)
    for i, (oe_kurz, oe_bez) in enumerate(oe_chain):
        oe_levels[f"org_level_{i}"] = oe_kurz
        oe_levels[f"org_bez_level_{i}"] = oe_bez
    
    return oe_levels


def _build_oe_hierarchy(oe_kurz: str, emp_index: Dict) -> str:
    """
    Baut OE-Hierarchie basierend auf Vorgesetzten-OE.
    """
    if not oe_kurz:
        return ""
    
    # Finde alle Mitarbeiter in derselben OE
    oe_members = [pn for pn, data in emp_index.items() 
                  if data.get("oe_kurz") == oe_kurz]
    
    if not oe_members:
        return oe_kurz
    
    # Finde den höchsten Vorgesetzten in dieser OE
    top_vg = None
    for pn in oe_members:
        emp_data = emp_index.get(pn, {})
        vg_pn = emp_data.get("vg_pn")
        if vg_pn and vg_pn in emp_index:
            vg_data = emp_index[vg_pn]
            if vg_data.get("oe_kurz") != oe_kurz:  # VG ist in anderer OE
                top_vg = vg_pn
                break
    
    if top_vg:
        vg_data = emp_index[top_vg]
        vg_oe = vg_data.get("oe_kurz", "")
        return f"{vg_oe} > {oe_kurz}" if vg_oe else oe_kurz
    
    return oe_kurz


def _build_oe_chain(emp_pn: str, emp_index: Dict, sep: str = " < ") -> tuple[str, str]:
    """
    Liefert die komplette OE-Kette von oben nach unten als String,
    sowohl für Kurz- als auch Langbezeichnungen.
    Beispiel: "GL < A1 < T1"
    """
    # Sammle Kette bottom-up über Vorgesetzte
    chain_kurz = []
    chain_bez = []
    visited = set()
    current_pn = emp_pn

    while current_pn and current_pn not in visited:
        visited.add(current_pn)
        emp_data = emp_index.get(current_pn, {})
        chain_kurz.append(str(emp_data.get("oe_kurz", "")).strip())
        chain_bez.append(str(emp_data.get("oe_bez", "")).strip())
        vg_pn = emp_data.get("vg_pn")
        if not vg_pn or vg_pn not in emp_index:
            break
        current_pn = vg_pn

    # Top-Down Reihenfolge und Verbinden
    chain_kurz = [c for c in reversed(chain_kurz) if c]
    chain_bez = [c for c in reversed(chain_bez) if c]
    return sep.join(chain_kurz), sep.join(chain_bez)


def export_org_structure(org_df: pd.DataFrame, output_path: Path) -> None:
    """
    Exportiert Organisationsstruktur in Excel-Datei.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Haupttabelle
        org_df.to_excel(writer, sheet_name="Organisationsstruktur", index=False)
        
        # OE-Übersicht
        oe_summary = _create_oe_summary(org_df)
        oe_summary.to_excel(writer, sheet_name="OE_Übersicht", index=False)
        
        # Hierarchie-Übersicht
        hierarchy_summary = _create_hierarchy_summary(org_df)
        hierarchy_summary.to_excel(writer, sheet_name="Hierarchie_Übersicht", index=False)
    
    print(f"Organisationsstruktur exportiert nach: {output_path}")


def _create_oe_summary(org_df: pd.DataFrame) -> pd.DataFrame:
    """Erstellt OE-Übersicht mit Mitarbeiteranzahl pro OE."""
    oe_stats = org_df.groupby("OE_Kurz").agg({
        "Personalnummer": "count",
        "Hierarchie_Ebene": ["min", "max"],
        "Status": lambda x: (x == "Aktiv").sum()
    }).round(2)
    
    oe_stats.columns = ["Mitarbeiter_Anzahl", "Min_Ebene", "Max_Ebene", "Aktive_MA"]
    oe_stats = oe_stats.reset_index()
    oe_stats = oe_stats.sort_values("Mitarbeiter_Anzahl", ascending=False)
    
    return oe_stats


def _create_hierarchy_summary(org_df: pd.DataFrame) -> pd.DataFrame:
    """Erstellt Hierarchie-Übersicht mit Statistiken pro Ebene."""
    hierarchy_stats = org_df.groupby("Hierarchie_Ebene").agg({
        "Personalnummer": "count",
        "OE_Kurz": "nunique",
        "Status": lambda x: (x == "Aktiv").sum()
    }).round(2)
    
    hierarchy_stats.columns = ["Mitarbeiter_Anzahl", "OE_Anzahl", "Aktive_MA"]
    hierarchy_stats = hierarchy_stats.reset_index()
    hierarchy_stats = hierarchy_stats.sort_values("Hierarchie_Ebene")
    
    return hierarchy_stats


def analyze_org_structure(org_df: pd.DataFrame) -> Dict:
    """
    Analysiert die Organisationsstruktur und gibt Statistiken zurück.
    """
    total_employees = len(org_df)
    active_employees = len(org_df[org_df["Status"] == "Aktiv"])
    
    # Hierarchie-Statistiken
    max_level = org_df["Hierarchie_Ebene"].max()
    level_distribution = org_df["Hierarchie_Ebene"].value_counts().sort_index()
    
    # OE-Statistiken
    unique_oes = org_df["OE_Kurz"].nunique()
    oe_distribution = org_df["OE_Kurz"].value_counts()
    
    # Top-Manager (Level 0)
    top_managers = org_df[org_df["Hierarchie_Ebene"] == 0]
    
    # Mitarbeiter ohne Vorgesetzten
    no_supervisor = org_df[org_df["Vorgesetzter_PN"] == ""]
    
    return {
        "total_employees": total_employees,
        "active_employees": active_employees,
        "max_hierarchy_level": max_level,
        "level_distribution": level_distribution.to_dict(),
        "unique_oes": unique_oes,
        "oe_distribution": oe_distribution.head(10).to_dict(),
        "top_managers_count": len(top_managers),
        "top_managers": top_managers[["Personalnummer", "Name", "OE_Kurz"]].to_dict("records"),
        "no_supervisor_count": len(no_supervisor),
        "no_supervisor": no_supervisor[["Personalnummer", "Name", "OE_Kurz"]].to_dict("records")
    }


if __name__ == "__main__":
    # Test der Organisationsstruktur-Funktionen
    from data_loader import load_employees
    
    print("Lade SAP Stammdaten...")
    sap_df = load_employees()
    
    print("Baue Organisationsstruktur auf...")
    org_df = build_org_structure(sap_df)
    
    print("Analysiere Struktur...")
    analysis = analyze_org_structure(org_df)
    
    print(f"\nOrganisationsstruktur-Analyse:")
    print(f"- Gesamt Mitarbeiter: {analysis['total_employees']}")
    print(f"- Aktive Mitarbeiter: {analysis['active_employees']}")
    print(f"- Maximale Hierarchie-Ebene: {analysis['max_hierarchy_level']}")
    print(f"- Anzahl OE: {analysis['unique_oes']}")
    print(f"- Top-Manager: {analysis['top_managers_count']}")
    print(f"- Ohne Vorgesetzten: {analysis['no_supervisor_count']}")
    
    # Export
    output_path = Path("tracking/org_structure/org_structure.xlsx")
    export_org_structure(org_df, output_path)
