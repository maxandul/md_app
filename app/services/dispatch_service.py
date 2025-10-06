from pathlib import Path
from datetime import date
import pandas as pd

from data_loader import load_config, row_to_context
from word_tools import fill_word_template
from adapters.mail_outlook import send_mail
from utils import date_in_range, last_day_of_month, fixed_filename
from constants import MDConstants, DocType

CFG = load_config()

def determine_docset(row: pd.Series, today: date) -> list[str]:
    from utils import date_in_range, last_day_of_month
    y = today.year
    austritt = row.get("Austritt")
    ende_pz = row.get("Ende Probezeit")

    oct1 = date(y, MDConstants.PROBEZEIT_MONTHS[0], 1)
    jan_next = date(y + 1, MDConstants.PROBEZEIT_MONTHS[3], 31)
    jun1 = date(y, 6, 1)
    sep_end = date(y, 9, last_day_of_month(y, 9))

    if date_in_range(austritt, oct1, jan_next):
        return [DocType.RUECKBLICK.value]
    if date_in_range(ende_pz, oct1, jan_next):
        return [f"{DocType.RUECKBLICK.value}_Probezeit", DocType.AUSBLiCK.value]
    if date_in_range(ende_pz, jun1, sep_end):
        return [DocType.AUSBLiCK.value]
    return [DocType.RUECKBLICK.value, DocType.AUSBLiCK.value]


def build_and_send_for_manager(
    mgr_row: pd.Series,
    subs_df: pd.DataFrame,
    rb_year: int,
    ab_year: int,
    today: date,
    out_root: Path,
    managers_index: dict | None = None,
    doc_types_override: list[str] | None = None,
    include_feedback: bool = True,
    send_mode: str | None = None,
):
    tp = CFG["paths"]["templates"]
    tpl_paths = {
        DocType.AUSBLiCK.value: Path(__file__).parent.parent / tp["ausblick"],
        DocType.RUECKBLICK.value: Path(__file__).parent.parent / tp["rueckblick"],
        f"{DocType.RUECKBLICK.value}_Probezeit": Path(__file__).parent.parent / tp["rueckblick_probezeit"],
        DocType.FEEDBACK.value: Path(__file__).parent.parent / tp["feedback"],
    }

    attachments: list[Path] = []
    mgr_pn = str(subs_df["Dir. Vorgesetzter (PN)"].iloc[0]).strip()
    out_dir = out_root / f"VG_{mgr_pn}"
    out_dir.mkdir(parents=True, exist_ok=True)

    allowed_types = {"Rückblick", "Ausblick", "Rückblick_Probezeit"}

    for _, r in subs_df.iterrows():
        types_for_row = list(doc_types_override) if doc_types_override else determine_docset(r, today)
        types_for_row = [t for t in types_for_row if t in allowed_types]
        for typ in types_for_row:
            ctx = row_to_context(r, typ, managers=managers_index)

            jahr_for_name = None
            if typ == "Ausblick":
                jahr_for_name = ab_year
            elif typ == "Rückblick":
                jahr_for_name = rb_year

            fname = fixed_filename(
                typ=typ,
                jahr=jahr_for_name,
                nachname_ma=r.get("Nachname",""),
                vorname_ma=r.get("Rufname",""),
                pn_ma=r.get("ID_NO_ZERO",""),
            ) + MDConstants.ALLOWED_EXTENSIONS[0]
            out_path = out_dir / fname
            fill_word_template(tpl_paths[typ], ctx, out_path)
            attachments.append(out_path)

    if include_feedback:
        fb_ctx = row_to_context(mgr_row, "Feedback", managers=managers_index)
        fb_ctx["fb_pn_vg"] = mgr_pn
        fb_name = (
            f"{DocType.FEEDBACK_VORLAGE_PREFIX}"
            f"{mgr_row.get('Nachname','')}_{mgr_row.get('Rufname','')}_{mgr_pn}{MDConstants.ALLOWED_EXTENSIONS[0]}"
        )
        fb_path = out_dir / fb_name
        fill_word_template(tpl_paths[DocType.FEEDBACK.value], fb_ctx, fb_path)
        attachments.append(fb_path)

    to = str(mgr_row.get("lange ID/Nummer", "")).strip()
    if not to:
        raise ValueError("E-Mail des/der Vorgesetzten nicht gefunden (Spalte 'lange ID/Nummer').")

    vg_vorname = str(mgr_row.get("Rufname", "")).strip()
    anrede = f"Hallo {vg_vorname}" if vg_vorname else "Hallo"

    if include_feedback:
        subject_tpl = CFG.get("mail", {}).get("subject_template", "MD-Unterlagen Durchlauf {rb_year}/{ab_year}")
        body_tpl = CFG.get("mail", {}).get("body_html_template", "")
        subject = subject_tpl.format(rb_year=rb_year, ab_year=ab_year)
        body = body_tpl.format(anrede=anrede, rb_year=rb_year, ab_year=ab_year)
    else:
        subject_tpl = CFG.get("mail_underjaehrig", {}).get("subject_template", "MD-Unterlagen")
        body_tpl = CFG.get("mail_underjaehrig", {}).get("body_html_template", "")
        subject = subject_tpl.format(rb_year=rb_year, ab_year=ab_year)
        body = body_tpl.format(anrede=anrede, rb_year=rb_year, ab_year=ab_year)

    send_mail(to=to, subject=subject, html_body=body, attachments=attachments, mode_override=send_mode)


