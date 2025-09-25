# app/dispatch.py
from pathlib import Path
from datetime import date
import pandas as pd

from .data_loader import load_config, row_to_context
from .word_tools import fill_word_template
from .mail_send import send_mail
from .utils import date_in_range, last_day_of_month, fixed_filename

CFG = load_config()

def determine_docset(row: pd.Series, today: date) -> list[str]:
    """
    Regeln:
    1) Austritt Okt (Y) – Jan (Y+1)         -> ["Rückblick"]
    2) Ende Probezeit Okt (Y) – Jan (Y+1)   -> ["Rückblick_Probezeit", "Ausblick"]
    3) Ende Probezeit Jun–Sep (Y)           -> ["Ausblick"]
    4) Sonst                                -> ["Rückblick", "Ausblick"]
    """
    y = today.year
    austritt = row.get("Austritt")
    ende_pz = row.get("Ende Probezeit")

    oct1 = date(y, 10, 1)
    jan_next = date(y + 1, 1, 31)
    jun1 = date(y, 6, 1)
    sep_end = date(y, 9, last_day_of_month(y, 9))

    if date_in_range(austritt, oct1, jan_next):
        return ["Rückblick"]

    if date_in_range(ende_pz, oct1, jan_next):
        return ["Rückblick_Probezeit", "Ausblick"]

    if date_in_range(ende_pz, jun1, sep_end):
        return ["Ausblick"]

    return ["Rückblick", "Ausblick"]

def build_and_send_for_manager(mgr_row: pd.Series, subs_df: pd.DataFrame,
                               rb_year: int, ab_year: int,
                               today: date, out_root: Path, managers_index: dict | None = None):
    tp = CFG["paths"]["templates"]
    tpl_paths = {
        "Ausblick": Path(__file__).parent / tp["ausblick"],
        "Rückblick": Path(__file__).parent / tp["rueckblick"],
        "Rückblick_Probezeit": Path(__file__).parent / tp["rueckblick_probezeit"],
        "Feedback": Path(__file__).parent / tp["feedback"],
    }

    attachments: list[Path] = []

    # Ordner je VG
    mgr_pn = str(subs_df["Dir. Vorgesetzter (PN)"].iloc[0]).strip()
    out_dir = out_root / f"VG_{mgr_pn}"
    out_dir.mkdir(parents=True, exist_ok=True)

    # Directs
    for _, r in subs_df.iterrows():
        for typ in determine_docset(r, today):
            ctx = row_to_context(r, typ, managers=managers_index)

            # Jahr pro Typ
            jahr_for_name = None
            if typ == "Ausblick":
                jahr_for_name = ab_year
            elif typ == "Rückblick":
                jahr_for_name = rb_year
            elif typ == "Rückblick_Probezeit":
                jahr_for_name = None  # ohne Jahr

            fname = fixed_filename(
                typ=typ,
                jahr=jahr_for_name,
                nachname_ma=r.get("Nachname",""),
                vorname_ma=r.get("Rufname",""),
                pn_ma=r.get("ID_NO_ZERO",""),
            ) + ".docx"
            out_path = out_dir / fname
            fill_word_template(tpl_paths[typ], ctx, out_path)
            attachments.append(out_path)

    # Feedback (für Manager selbst)
    fb_ctx = row_to_context(mgr_row, "Feedback", managers=managers_index)
    fb_ctx["fb_pn_vg"] = mgr_pn   

    fb_name = (
        "Vorlage_Feedback_an_"
        f"{mgr_row.get('Nachname','')}_{mgr_row.get('Rufname','')}_{mgr_pn}.docx"
    )
    print("Feedback-Kontext:", fb_ctx)
    fb_path = out_dir / fb_name
    fill_word_template(tpl_paths["Feedback"], fb_ctx, fb_path)
    attachments.append(fb_path)

    # Empfänger
    to = str(mgr_row.get("lange ID/Nummer", "")).strip()
    if not to:
        raise ValueError("E-Mail des/der Vorgesetzten nicht gefunden (Spalte 'lange ID/Nummer').")

    # Mailtexte aus config
    subject_tpl = (CFG.get("mail", {}).get("subject_template") or "Unterlagen")
    subject = subject_tpl.format(rb_year=rb_year, ab_year=ab_year)

    vg_vorname = str(mgr_row.get("Rufname", "")).strip()
    anrede = f"Hallo {vg_vorname}" if vg_vorname else "Hallo"

    subject_tpl = CFG.get("mail", {}).get("subject_template", "MD-Unterlagen Durchlauf {rb_year}/{ab_year}")
    subject = subject_tpl.format(rb_year=rb_year, ab_year=ab_year)

    body_tpl = CFG.get("mail", {}).get("body_html_template", "")
    body = body_tpl.format(anrede=anrede, rb_year=rb_year, ab_year=ab_year)

    # Versand
    send_mail(
        to=to, subject=subject, html_body=body, attachments=attachments
    )