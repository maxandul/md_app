# app/utils.py
from datetime import date
from calendar import monthrange
import pandas as pd

def date_in_range(d, start: date, end: date) -> bool:
    if pd.isna(d):
        return False
    ts = pd.to_datetime(d, errors="coerce")
    if pd.isna(ts):
        return False
    dd = ts.date()
    return start <= dd <= end

def last_day_of_month(y: int, m: int) -> int:
    return monthrange(y, m)[1]

def fixed_filename(typ: str, jahr: int | None, nachname_ma: str, vorname_ma: str, pn_ma: str,
                   nachname_vg: str | None = None, vorname_vg: str | None = None, pn_vg: str | None = None) -> str:
    """Erstellt Dateinamen exakt nach Vorgabe (Umlaute bleiben erhalten, Leerzeichen -> _)."""
    def clean(s: str) -> str:
        return (s or "").strip().replace(" ", "_")

    if typ == "Ausblick":
        return f"Ausblick_{jahr}_{clean(nachname_ma)}_{clean(vorname_ma)}_{pn_ma}"
    if typ == "Rückblick":
        return f"Rückblick_{jahr}_{clean(nachname_ma)}_{clean(vorname_ma)}_{pn_ma}"
    if typ == "Rückblick_Probezeit":
        return f"Rückblick_Probezeit_{clean(nachname_ma)}_{clean(vorname_ma)}_{pn_ma}"
    if typ == "Feedback":
        return f"Feedback_{clean(nachname_vg)}_{clean(vorname_vg)}_{pn_vg}"
    raise ValueError(f"Unbekannter Typ: {typ}")
