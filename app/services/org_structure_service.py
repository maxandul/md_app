from __future__ import annotations

"""
Organisationsstruktur-Tool (Service)

Baut aus SAP Stammdaten eine Organisationsstruktur-Tabelle auf und liefert
Zusammenfassungen/Analysen (für DS-Export genutzt).
"""

import pandas as pd
from pathlib import Path
from typing import Dict


def build_org_structure(sap_df: pd.DataFrame) -> pd.DataFrame:
    if "ID_NO_ZERO" not in sap_df.columns or "Dir. Vorgesetzter (PN)" not in sap_df.columns:
        raise ValueError("Erforderliche Spalten 'ID_NO_ZERO' oder 'Dir. Vorgesetzter (PN)' fehlen")

    df = sap_df.copy()
    df["ID_NO_ZERO"] = df["ID_NO_ZERO"].astype(str).str.strip()
    df["Dir. Vorgesetzter (PN)"] = df["Dir. Vorgesetzter (PN)"].astype(str).str.strip()

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
            "austritt": row.get("Austritt"),
        }

    hierarchy_levels = _calculate_hierarchy_levels(emp_index)

    org_rows = []
    for pn, emp_data in emp_index.items():
        vg_pn = emp_data["vg_pn"]
        vg_data = emp_index.get(vg_pn, {})
        level = hierarchy_levels.get(pn, -1)
        org_path = _build_org_path(pn, emp_index, hierarchy_levels)
        oe_chain_kurz, oe_chain_bez = _build_oe_chain(pn, emp_index)
        oe_levels = _build_oe_hierarchy_levels(pn, emp_index)

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
            "OE_Hierarchie": _build_oe_hierarchy(emp_data["oe_kurz"], emp_index),
            "OE_Kette": oe_chain_kurz,
            "OE_Bez_Kette": oe_chain_bez,
            "Eintritt": emp_data["eintritt"],
            "Austritt": emp_data["austritt"],
            "Status": "Aktiv" if pd.isna(emp_data["austritt"]) else "Ausgeschieden",
        }
        row_data.update(oe_levels)
        org_rows.append(row_data)

    return pd.DataFrame(org_rows)


def _calculate_hierarchy_levels(emp_index: Dict) -> Dict[str, int]:
    levels = {}

    def calculate_level(pn: str, path: set) -> int:
        if pn in path:
            return 0
        if pn in levels:
            return levels[pn]
        emp_data = emp_index.get(pn)
        if not emp_data or not emp_data["vg_pn"] or emp_data["vg_pn"] not in emp_index:
            levels[pn] = 0
            return 0
        vg_level = calculate_level(emp_data["vg_pn"], path | {pn})
        levels[pn] = vg_level + 1
        return levels[pn]

    for pn in emp_index:
        calculate_level(pn, set())
    return levels


def _build_org_path(pn: str, emp_index: Dict, levels: Dict[str, int]) -> str:
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
    oe_levels = {}
    current_pn = emp_pn
    visited = set()
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
        vg_pn = emp_data.get("vg_pn")
        if not vg_pn or vg_pn not in emp_index:
            break
        current_pn = vg_pn
    for i, (oe_kurz, oe_bez) in enumerate(oe_chain):
        oe_levels[f"org_level_{i}"] = oe_kurz
        oe_levels[f"org_bez_level_{i}"] = oe_bez
    return oe_levels


def _build_oe_hierarchy(oe_kurz: str, emp_index: Dict) -> str:
    if not oe_kurz:
        return ""
    oe_members = [pn for pn, data in emp_index.items() if data.get("oe_kurz") == oe_kurz]
    if not oe_members:
        return oe_kurz
    top_vg = None
    for pn in oe_members:
        emp_data = emp_index.get(pn, {})
        vg_pn = emp_data.get("vg_pn")
        if vg_pn and vg_pn in emp_index:
            vg_data = emp_index[vg_pn]
            if vg_data.get("oe_kurz") != oe_kurz:
                top_vg = vg_pn
                break
    if top_vg:
        vg_data = emp_index[top_vg]
        vg_oe = vg_data.get("oe_kurz", "")
        return f"{vg_oe} > {oe_kurz}" if vg_oe else oe_kurz
    return oe_kurz


def _build_oe_chain(emp_pn: str, emp_index: Dict, sep: str = " < ") -> tuple[str, str]:
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
    chain_kurz = [c for c in reversed(chain_kurz) if c]
    chain_bez = [c for c in reversed(chain_bez) if c]
    return sep.join(chain_kurz), sep.join(chain_bez)


