"""
Dashboard-Service für das MD-Prozess-Tool.

Dieser Service kapselt alle Funktionen zum Dashboard-Management,
einschließlich Datenaktualisierung und Export.
"""

from __future__ import annotations
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime

from constants import MDConstants, ProcStatus, DashTag


def refresh_dashboard(app) -> None:
    """Lädt Dashboard-Daten gemäß aktueller Filter und befüllt den Treeview."""
    for item in app.tree_dashboard.get_children():
        app.tree_dashboard.delete(item)

    try:
        name_search = app.dash_name_search.get().strip().lower()
        status_filter = app.dash_status_filter.get().strip()

        df = app.tracking.get_dashboard_data(filter_status=status_filter)
        if df.empty:
            return

        if name_search:
            name_mask = (
                df["vg_name"].astype(str).str.lower().str.contains(name_search, na=False)
                | df["ma_name"].astype(str).str.lower().str.contains(name_search, na=False)
            )
            df = df[name_mask]

        def safe_value(val):
            try:
                if pd.isna(val) or val == "nan" or val == "NaN":
                    return ""
            except Exception:
                pass
            return str(val) if val is not None else ""

        def format_pn(val):
            try:
                if isinstance(val, float) and float(val).is_integer():
                    return str(int(val))
                s = safe_value(val)
                if s.endswith('.0') and s.replace('.', '', 1).isdigit():
                    return s[:-2]
                return s
            except Exception:
                return safe_value(val)

        def status_tag_for(status: str) -> str:
            s = (status or "").lower().strip()
            if s == ProcStatus.AUSSTEHEND.value:
                return DashTag.AUSSTEHEND.value
            if s == ProcStatus.ERHALTEN.value:
                return DashTag.ERHALTEN.value
            if s == ProcStatus.ERUEBRIGT.value:
                return DashTag.ERUEBRIGT.value
            if s == ProcStatus.PRUEFUNG_NOETIG.value:
                return DashTag.PRUEFUNG_NOETIG.value
            if s == ProcStatus.OK.value:
                return DashTag.OK.value
            if s == ProcStatus.MANUELL.value:
                return DashTag.MANUELL.value
            return ""

        for _, row in df.iterrows():
            status = safe_value(row.get("status", ""))
            tag = status_tag_for(status)
            app.tree_dashboard.insert(
                "",
                "end",
                values=(
                    safe_value(row.get("log_id", "")),
                    format_pn(row.get("vg_pn", "")),
                    safe_value(row.get("vg_name", "")),
                    format_pn(row.get("ma_pn", "")),
                    safe_value(row.get("ma_name", "")),
                    safe_value(row.get("doc_type", "")),
                    safe_value(row.get("erwartet", "")),
                    safe_value(row.get("erhalten", "")),
                    status,
                    safe_value(row.get("status_grund", "")),
                    safe_value(row.get("versendet_am", "")),
                    safe_value(row.get("zuletzt_erinnert_am", "")),
                ),
                tags=(tag,) if tag else (),
            )
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Laden der Dashboard-Daten: {e}")


def export_dashboard(app) -> None:
    """Exportiert die aktuell gefilterten Dashboard-Daten als CSV."""
    try:
        from tkinter import filedialog

        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not filename:
            return

        status_filter = app.dash_status_filter.get().strip()
        df = app.tracking.get_dashboard_data(filter_status=status_filter)
        df.to_csv(filename, sep=";", index=False, encoding="utf-8-sig")
        messagebox.showinfo(MDConstants.MSG_SUCCESS, f"Dashboard-Daten exportiert nach: {filename}")
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Export fehlgeschlagen: {e}")


def manual_adjustment(app) -> None:
    """Öffnet einen Dialog zur manuellen Anpassung eines ausgewählten Dashboard-Eintrags."""
    selection = app.tree_dashboard.selection()
    if not selection:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie einen Eintrag aus.")
        return

    item = app.tree_dashboard.item(selection[0])
    values = item.get("values", [])
    if len(values) < 12:
        messagebox.showerror(MDConstants.MSG_ERROR, "Ungültiger Eintrag ausgewählt.")
        return

    log_id = values[0]

    dialog = tk.Toplevel(app)
    dialog.title("Manuelle Anpassung")
    dialog.geometry("700x600")
    dialog.transient(app)
    dialog.grab_set()

    canvas = tk.Canvas(dialog)
    scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    ttk.Label(scrollable_frame, text=f"Log-ID: {log_id}", font=("Arial", 12, "bold")).pack(pady=8)

    columns = [
        ("vg_pn", "VG PN", True),
        ("vg_name", "VG Name", True),
        ("ma_pn", "MA PN", True),
        ("ma_name", "MA Name", True),
        ("doc_type", "Dokument-Typ", False),
        ("erwartet", "Erwartet", True),
        ("erhalten", "Erhalten", True),
        ("status", "Status", True),
        ("status_grund", "Status Grund", True),
        ("versendet_am", "Versendet am", True),
        ("zuletzt_erinnert_am", "Zuletzt erinnert am", True),
    ]

    entry_vars: dict[str, tk.StringVar] = {}

    for i, (col_key, col_label, editable) in enumerate(columns):
        frame = ttk.Frame(scrollable_frame)
        frame.pack(fill="x", padx=8, pady=2)

        ttk.Label(frame, text=f"{col_label}:", width=20, anchor="w").pack(side="left")

        if not editable:
            ttk.Label(frame, text=str(values[i + 1]), relief="sunken", width=40).pack(side="left", padx=(4, 0))
        else:
            if col_key == "status":
                var = tk.StringVar(value=values[i + 1])
                entry = ttk.Combobox(
                    frame,
                    textvariable=var,
                    width=37,
                    values=[
                        ProcStatus.AUSSTEHEND.value,
                        ProcStatus.ERHALTEN.value,
                        ProcStatus.PRUEFUNG_NOETIG.value,
                        ProcStatus.ERUEBRIGT.value,
                    ],
                )
            elif col_key == "status_grund":
                var = tk.StringVar(value=values[i + 1])
                entry = ttk.Combobox(
                    frame,
                    textvariable=var,
                    width=37,
                    values=[
                        "",
                        "Grund_Prüfung (aus Verarbeitung)",
                        "Krankheit/Unfall",
                        "anderer VG",
                        "Austritt",
                        "sonstiges",
                    ],
                )
            else:
                var = tk.StringVar(value=str(values[i + 1]))
                entry = ttk.Entry(frame, textvariable=var, width=40)

            entry.pack(side="left", padx=(4, 0))
            entry_vars[col_key] = var

    button_frame = ttk.Frame(scrollable_frame)
    button_frame.pack(pady=16)

    def apply_adjustment() -> None:
        try:
            updates: dict[str, str] = {}
            for col_key, var in entry_vars.items():
                new_value = var.get().strip()
                col_index = None
                for idx, (key, _, _) in enumerate(columns):
                    if key == col_key:
                        col_index = idx + 1
                        break
                if col_index is not None:
                    old_value = str(values[col_index])
                    if new_value != old_value:
                        updates[col_key] = new_value

            if not updates:
                messagebox.showinfo(MDConstants.MSG_INFO, "Keine Änderungen vorgenommen.")
                dialog.destroy()
                return

            if app.tracking.update_entry(log_id, updates):
                messagebox.showinfo(MDConstants.MSG_SUCCESS, f"Anpassung gespeichert.\nGeändert: {', '.join(updates.keys())}")
                dialog.destroy()
                refresh_dashboard(app)
            else:
                messagebox.showerror(MDConstants.MSG_ERROR, "Anpassung fehlgeschlagen - Eintrag nicht gefunden.")

        except Exception as e:
            messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Speichern: {e}")

    ttk.Button(button_frame, text="Anwenden", command=apply_adjustment).pack(side="left", padx=(0, 8))
    ttk.Button(button_frame, text="Abbrechen", command=dialog.destroy).pack(side="left")

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")


def get_dashboard_data(filter_status: str = "") -> pd.DataFrame:
    """
    Lädt Dashboard-Daten aus dem Tracking-System.
    
    Args:
        filter_status: Optionaler Status-Filter
        
    Returns:
        DataFrame mit Dashboard-Daten
    """
    # Optional: Kann später implementiert werden
    pass
