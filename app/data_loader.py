from pathlib import Path
import pandas as pd
import yaml

def load_config():
    cfg_path = Path(__file__).parent / "config.yaml"
    with open(cfg_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

CFG = load_config()

def load_employees():
    xlsx_path = Path(__file__).parent / CFG["paths"]["sap_stammdaten"]
    df = pd.read_excel(
        xlsx_path,
        dtype={"ID_NO_ZERO": str, "Dir. Vorgesetzter (PN)": str, "Ans.": str},
    )
    df.columns = [c.strip() for c in df.columns]

    # Nur noch neutrale Textspalten
    text_cols = ["Rufname","Nachname","OE Bez.","OE Kurzb.","Plans. Bez.","lange ID/Nummer"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Datumsfelder
    date_cols = CFG.get("excel", {}).get("date_cols", [])
    for c in date_cols + ["Eintritt", "Austritt"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Personalnummern normalisieren
    if "ID_NO_ZERO" in df.columns:
        df["ID_NO_ZERO"] = df["ID_NO_ZERO"].astype(str).str.strip()
    if "Dir. Vorgesetzter (PN)" in df.columns:
        df["Dir. Vorgesetzter (PN)"] = df["Dir. Vorgesetzter (PN)"].astype(str).str.strip()

    # --- IDs mit führenden Nullen angleichen ---
    if "ID_NO_ZERO" in df.columns and "Dir. Vorgesetzter (PN)" in df.columns:
        max_len = max(
            df["ID_NO_ZERO"].str.len().max(),
            df["Dir. Vorgesetzter (PN)"].str.len().max()
        )
        df["ID_NO_ZERO"] = df["ID_NO_ZERO"].str.zfill(max_len)
        df["Dir. Vorgesetzter (PN)"] = df["Dir. Vorgesetzter (PN)"].str.zfill(max_len)

    return df

def build_manager_index(df: pd.DataFrame):
    """vg_pn -> {'manager': row, 'subs': DataFrame}; ignoriert vg_pn leer/'0'/NaN."""
    index = {}
    by_pn = {pn: r for pn, r in df.set_index("ID_NO_ZERO").iterrows()}

    if "Dir. Vorgesetzter (PN)" not in df.columns:
        return {}

    for vg_pn, subs in df.groupby("Dir. Vorgesetzter (PN)"):
        if not vg_pn or vg_pn in ("0", "nan", "NaN", "None"):
            continue  # Personen ohne gültige VG-PN nicht versenden
        mgr_row = by_pn.get(vg_pn)
        if mgr_row is None:
            continue
        # Nur Manager mit mindestens 1 Direct
        if len(subs) > 0:
            index[vg_pn] = {"manager": mgr_row, "subs": subs.copy()}
    return index

def _manager_fields(vg_pn: str, managers: dict | None):
    """Liefert (Name, E-Mail) des/der Vorgesetzten über den managers-Index."""
    vg_name, vg_email = "", ""
    if managers and vg_pn in managers:
        vg_row = managers[vg_pn]["manager"]
        vg_name = f"{str(vg_row.get('Rufname','')).strip()} {str(vg_row.get('Nachname','')).strip()}".strip()
        vg_email = str(vg_row.get("lange ID/Nummer","")).strip()
    return vg_name, vg_email

def row_to_context(row, typ: str, managers: dict | None = None) -> dict:
    """Tag->Wert Mapping gemäss Vorgabe. Vorgesetzten-Daten kommen ausschliesslich über VG-PN."""
    get = lambda k: "" if row.get(k) is None else str(row.get(k))
    vorname = get("Rufname")
    nachname = get("Nachname")
    pn = get("ID_NO_ZERO")
    funktion = get("Plans. Bez.")
    oe = get("OE Bez.")
    vg_pn = get("Dir. Vorgesetzter (PN)")
    vg_name, _vg_email = _manager_fields(vg_pn, managers)

    if typ == "Ausblick":  # ab_*
        return {
            "ab_name": f"{vorname} {nachname}".strip(),
            "ab_funktion": funktion,
            "ab_name_vg": vg_name,
            "ab_pn": pn,
            "ab_oe": oe,
            "ab_pn_vg": vg_pn,
        }

    if typ == "Rückblick":  # rb_*
        return {
            "rb_name": f"{vorname} {nachname}".strip(),
            "rb_funktion": funktion,
            "rb_name_vg": vg_name,
            "rb_pn": pn,
            "rb_oe": oe,
            "rb_pn_vg": vg_pn,
        }

    if typ == "Rückblick_Probezeit":  # pz_*
        return {
            "pz_name": f"{vorname} {nachname}".strip(),
            "pz_funktion": funktion,
            "pz_name_vg": vg_name,
            "pz_pn": pn,
            "pz_oe": oe,
            "pz_pn_vg": vg_pn,
        }

    if typ == "Feedback":  # fb_* (für Manager selbst)
        return {
            "fb_name": "",  # Feld nicht nutzen oder leer lassen
            "fb_name_vg": f"{vorname} {nachname}".strip(),  # Name des Managers selbst
            "fb_pn_vg": get("ID_NO_ZERO"),   # PN des Managers selbst
        }

    return {}