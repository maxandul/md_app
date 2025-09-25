# app/main.py
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
import os
import win32com.client

from .data_loader import load_employees, load_config, build_manager_index
from .dispatch import build_and_send_for_manager
from .doc_processing import process_docx_folder, export_sap_massenupload, export_ds_csv, move_after_processing, export_ds_csv, process_pdfs

CFG = load_config()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MD-Prozess-Tool")
        self.geometry("1000x640")

        # Jahr-Variable ZENTRAL anlegen (wichtig, sonst None im Callback)
        self.jahr_var = tk.IntVar(value=date.today().year)

        # Notebook mit Tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.frame_versand = ttk.Frame(self.notebook)
        self.frame_ruecklauf = ttk.Frame(self.notebook)
        self.frame_verarbeitung = ttk.Frame(self.notebook)

        self.notebook.add(self.frame_versand, text="Versand")
        self.notebook.add(self.frame_ruecklauf, text="Rücklauf")
        self.notebook.add(self.frame_verarbeitung, text="Verarbeitung")

        # Tabs aufbauen
        self.build_versand()
        self.build_ruecklauf()
        self.build_verarbeitung()

    # ---------------------------
    # VERSAND
    # ---------------------------
    def build_versand(self):
        # Daten laden
        self.df = load_employees()
        self.mgr_index = build_manager_index(self.df)

        # --- Toolbar Frame für Jahr + Suche ---
        toolbar = ttk.Frame(self.frame_versand)
        toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

        # Jahr wählen (Rückblick-Jahr, Ausblick-Jahr = +1)
        self.rb_year_var = tk.IntVar(value=date.today().year)
        self.ab_year_var = tk.IntVar(value=self.rb_year_var.get() + 1)

        ttk.Label(toolbar, text="Jahr:").pack(side="left", padx=(0, 4))
        jahr_box = ttk.Combobox(
            toolbar,
            textvariable=self.rb_year_var,
            values=[date.today().year-1, date.today().year, date.today().year+1],
            state="readonly",
            width=8
        )
        jahr_box.pack(side="left", padx=(0, 8))

        self.year_label = ttk.Label(toolbar, text="")
        self.year_label.pack(side="left", padx=(0, 20))

        def _update_year_label(*_):
            rb = self.rb_year_var.get()
            self.ab_year_var.set(rb + 1)
            self.year_label.config(text=f"Rückblick: {rb} / Ausblick: {rb+1}")

        jahr_box.bind("<<ComboboxSelected>>", _update_year_label)
        _update_year_label()

        # Suche
        ttk.Label(toolbar, text="Suche:").pack(side="left", padx=(0, 4))
        self.filter_var = tk.StringVar()
        entry = ttk.Entry(toolbar, textvariable=self.filter_var, width=36)
        entry.pack(side="left", padx=(0, 8))
        entry.bind("<KeyRelease>", lambda e: self._refresh_mgr_table())

        # Info-Button (Popup)
        hinweis = (
            "So funktioniert der Versand:\n\n"
            "1. Wähle oben das Jahr.\n"
            "2. Suche optional nach einer Person oder OE.\n"
            "3. Markiere in der Tabelle eine oder mehrere Vorgesetzte.\n"
            "4. Klicke auf den Button unten, um automatisch die MD-Unterlagen\n"
            "   für alle direktunterstellten Mitarbeitenden zu erzeugen und\n"
            "   per E-Mail zu versenden."
        )
        def show_hint():
            messagebox.showinfo("Info Versand", hinweis)

        ttk.Button(toolbar, text="ℹ Info", command=show_hint).pack(side="right", padx=8)

        # --- Tabelle der Vorgesetzten (sortierbar) ---
        cols = ["VG_PN", "Nachname", "Vorname", "OE Kurzb.", "E-Mail", "#Directs"]
        self.tree = ttk.Treeview(self.frame_versand, columns=cols, show="headings", height=20)

        for c in cols:
            self.tree.heading(c, text=c, command=lambda col=c: self._sort_by(col, False))
            self.tree.column(c, width=120 if c not in ("E-Mail", "Nachname") else 200, anchor="w")

        self.tree.grid(row=2, column=0, columnspan=6, padx=8, pady=8, sticky="nsew")

        # Button
        ttk.Button(
            self.frame_versand,
            text="Unterlagen erzeugen & E-Mail an ausgewählte Vorgesetzte senden",
            command=self.on_send_managers
        ).grid(row=3, column=0, padx=8, pady=8, sticky="w")

        # Layout-Weights
        self.frame_versand.grid_rowconfigure(2, weight=1)
        self.frame_versand.grid_columnconfigure(5, weight=1)

        # --- Refresh-Logik für Tabelle (unverändert) ---
        def _rows_for_table():
            rows = []
            q = (self.filter_var.get() or "").lower()
            for vg_pn, pack in self.mgr_index.items():
                mgr = pack["manager"]
                subs = pack["subs"]
                vorname = str(mgr.get("Rufname", "") or "")
                nachname = str(mgr.get("Nachname", "") or "")
                email = str(mgr.get("lange ID/Nummer", "") or "")
                oe_kurz = str(mgr.get("OE Kurzb.", "") or "")
                hay = f"{vg_pn} {vorname} {nachname} {email} {oe_kurz}".lower()

                if q and q not in hay:
                    found = any(
                        (q in f"{str(r.get('ID_NO_ZERO','')).lower()} "
                            f"{str(r.get('Rufname','')).lower()} "
                            f"{str(r.get('Nachname','')).lower()}")
                        for _, r in subs.iterrows()
                    )
                    if not found:
                        continue

                rows.append((vg_pn, nachname, vorname, oe_kurz, email, len(subs)))
            return rows

        def _refresh_mgr_table():
            for i in self.tree.get_children():
                self.tree.delete(i)
            for row in _rows_for_table():
                self.tree.insert("", "end", iid=str(row[0]), values=list(row))

        self._refresh_mgr_table = _refresh_mgr_table
        self._refresh_mgr_table()


    def _sort_by(self, col, descending):
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children('')]
        # numerisch sortieren, wenn möglich
        def _to_key(v):
            try:
                return float(v)
            except:
                return v.lower() if isinstance(v, str) else v
        data.sort(key=lambda t: _to_key(t[0]), reverse=descending)
        for index, item in enumerate(data):
            self.tree.move(item[1], '', index)
        # Toggle für nächste Sortierung
        self.tree.heading(col, command=lambda _col=col: self._sort_by(_col, not descending))


    def on_send_managers(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Hinweis", "Bitte mindestens eine/n Vorgesetzte/n auswählen.")
            return

        rb_year = self.rb_year_var.get()
        ab_year = self.ab_year_var.get()
        out_root = Path(__file__).parent.parent / "tracking" / "versand"
        out_root.mkdir(parents=True, exist_ok=True)

        errors = []
        for vg_pn in sel:
            pack = self.mgr_index.get(vg_pn)
            if not pack:
                errors.append(f"Kein Paket für VG_PN {vg_pn}")
                continue
            mgr = pack["manager"]
            subs = pack["subs"]
            if mgr is None:
                errors.append(f"Vorgesetzte/r mit PN {vg_pn} nicht in EXPORT.xlsx gefunden.")
                continue

            try:
                build_and_send_for_manager(
                    mgr_row=mgr,
                    subs_df=subs,
                    rb_year=rb_year,
                    ab_year=ab_year,
                    today=date.today(),
                    out_root=out_root,
                    managers_index=self.mgr_index,
                )
            except Exception as e:
                errors.append(f"{mgr.get('Rufname','')} {mgr.get('Nachname','')} ({vg_pn}): {e}")

        if errors:
            messagebox.showerror("Abschluss mit Fehlern", "\n".join(errors))
        else:
            messagebox.showinfo("Fertig", "Versand ausgefuehrt (siehe Outbox/gesendete Elemente).")

    # ---------------------------
    # RÜCKLAUF (Prototyp GUI)
    # ---------------------------
            
    def build_ruecklauf(self):
        # --- Toolbar Frame ---
        toolbar = ttk.Frame(self.frame_ruecklauf)
        toolbar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

        # Infobox-Button
        hinweis = (
            "So funktioniert der Rücklauf:\n\n"
            "1. Klicke auf 'Posteingang scannen', um neue Mails im Gruppenpostfach "
            "'VD-GS HR' zu prüfen.\n"
            "2. Anhänge mit Rückblick, Ausblick oder Feedback werden automatisch kopiert.\n"
            "3. Enthält eine Mail NUR diese Anhänge → Dokumente werden kopiert und die Mail "
            "in den Ordner '12 Mitarbeitenden-Dialog' verschoben.\n"
            "4. Enthält eine Mail zusätzlich Probezeit-Dokumente oder fremde Anhänge → "
            "die MD-Anhänge werden kopiert, die Mail bleibt im Posteingang, Hinweis unter "
            "'Prüfen erforderlich'.\n"
            "5. Mails ohne MD-Anhänge werden übersprungen."
        )
        def show_hint():
            messagebox.showinfo("Info Rücklauf", hinweis)

        ttk.Button(toolbar, text="ℹ Info", command=show_hint).pack(side="right", padx=8)

        # Scan-Button
        ttk.Button(toolbar, text="Posteingang scannen", command=self.on_scan_real).pack(side="left", padx=(0, 8))

        # Status-Zeile
        self.ruecklauf_status = ttk.Label(self.frame_ruecklauf, text="Noch kein Scan durchgeführt.", foreground="gray")
        self.ruecklauf_status.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

        # Notebook für Ergebnislisten
        self.ruecklauf_nb = ttk.Notebook(self.frame_ruecklauf)
        self.ruecklauf_nb.grid(row=2, column=0, columnspan=6, sticky="nsew", padx=8, pady=8)

        # Tabs
        self.tab_ok = ttk.Frame(self.ruecklauf_nb)
        self.tab_pruefen = ttk.Frame(self.ruecklauf_nb)
        self.tab_skip = ttk.Frame(self.ruecklauf_nb)

        self.ruecklauf_nb.add(self.tab_ok, text="Kopiert & verschoben")
        self.ruecklauf_nb.add(self.tab_pruefen, text="Prüfen erforderlich")
        self.ruecklauf_nb.add(self.tab_skip, text="Übersprungen")

        # Treeviews für jede Kategorie
        self.tree_ok = self._make_tree(self.tab_ok, ["Datei", "Zielordner", "Absender", "Betreff"])
        self.tree_pruefen = self._make_tree(self.tab_pruefen, ["Grund", "Zu prüfende Dokumente", "Absender", "Betreff", "Rückblick/Ausblick/Feedback kopiert?"])
        self.tree_skip = self._make_tree(self.tab_skip, ["Absender", "Betreff", "Grund"])

        # Layout-Weights
        self.frame_ruecklauf.grid_rowconfigure(2, weight=1)
        self.frame_ruecklauf.grid_columnconfigure(5, weight=1)



    def _make_tree(self, parent, cols):
        """Helper: erstellt Treeview mit Scrollbar"""
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=180 if c != "Betreff" else 300, anchor="w")
        tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        return tree


    def on_scan_real(self):
        """Scan der Shared Mailbox 'VD-GS HR' und Verarbeitung nach Regeln"""
        for t in [self.tree_ok, self.tree_pruefen, self.tree_skip]:
            for i in t.get_children():
                t.delete(i)

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        mailbox = outlook.Folders["VD-GS HR"]
        inbox = mailbox.Folders["Posteingang"]
        target_folder = inbox.Folders["12 Mitarbeitenden-Dialog"]

        md_keywords = ["rückblick", "rueckblick", "ausblick", "feedback"]
        kw_probezeit = "probezeit"
        allowed_exts = [".docx", ".pdf"]

        base_path = Path(__file__).parent.parent / "ruecklauf"
        base_path.mkdir(parents=True, exist_ok=True)

        found, copied, moved, to_check, skipped = 0, 0, 0, 0, 0

        for mail in inbox.Items:
            found += 1
            sender = self._get_sender_address(mail)
            subject = str(mail.Subject or "")

            # Anhänge einsammeln
            files = []
            for att in mail.Attachments:
                try:
                    fname = str(att.FileName or "").strip()
                except Exception:
                    continue
                if not fname or "." not in fname:
                    continue
                files.append((fname, att))

            if not files:
                self.tree_skip.insert("", "end", values=[sender, subject, "Keine Anhänge"])
                skipped += 1
                continue

            # Klassifizieren
            md_files = [f for f, _ in files if any(k in f.lower() for k in md_keywords) and os.path.splitext(f)[1].lower() in allowed_exts]
            probezeit_files = [f for f, _ in files if kw_probezeit in f.lower() and os.path.splitext(f)[1].lower() in allowed_exts]
            other_files = [f for f, _ in files if f not in md_files and f not in probezeit_files]

            if md_files and not probezeit_files and not other_files:
                # Sauber → kopieren & verschieben
                for fname, att in files:
                    if fname in md_files:
                        save_path = base_path / fname
                        att.SaveAsFile(str(save_path))
                        self.tree_ok.insert("", "end", values=[fname, str(base_path), sender, subject])
                        copied += 1
                mail.Move(target_folder)
                moved += 1

            elif probezeit_files:
                # Probezeit → MD kopieren, Mail bleibt
                copied_names = []
                for fname, att in files:
                    if fname in md_files:
                        save_path = base_path / fname
                        att.SaveAsFile(str(save_path))
                        copied += 1
                        copied_names.append(fname)
                grund = "Probezeit"
                self.tree_pruefen.insert("", "end", values=[grund, ", ".join(probezeit_files) or "Keine", sender, subject, ", ".join(copied_names) or "Keine"])
                to_check += 1

            elif md_files and other_files:
                # Gemischt → MD kopieren, Mail bleibt
                copied_names = []
                for fname, att in files:
                    if fname in md_files:
                        save_path = base_path / fname
                        att.SaveAsFile(str(save_path))
                        copied += 1
                        copied_names.append(fname)
                grund = "Fremde Anhänge"
                self.tree_pruefen.insert("", "end", values=[grund, ", ".join(other_files) or "Keine", sender, subject, ", ".join(copied_names) or "Keine"])
                to_check += 1

            elif probezeit_files and other_files:
                # Probezeit + Fremd
                copied_names = []
                for fname, att in files:
                    if fname in md_files:
                        save_path = base_path / fname
                        att.SaveAsFile(str(save_path))
                        copied += 1
                        copied_names.append(fname)
                grund = "Probezeit + Fremde Anhänge"
                zu_pruefen = probezeit_files + other_files
                self.tree_pruefen.insert("", "end", values=[grund, ", ".join(zu_pruefen) or "Keine", sender, subject, ", ".join(copied_names) or "Keine"])
                to_check += 1

            else:
                # Keine MD-Anhänge → skip
                self.tree_skip.insert("", "end", values=[sender, subject, "Keine MD-Anhänge"])
                skipped += 1

        self.ruecklauf_status.config(
            text=f"Scan abgeschlossen: {found} Mails • {copied} Anhänge kopiert • {moved} verschoben • {to_check} prüfen • {skipped} übersprungen",
            foreground="black"
        )

    def _get_sender_address(self, mail):
        """Versucht, eine saubere SMTP-Adresse für den Absender zurückzugeben."""
        try:
            sender = mail.Sender
            if sender and sender.AddressEntryUserType == 0:  # 0 = ExchangeUser
                ex_user = sender.GetExchangeUser()
                if ex_user:
                    return ex_user.PrimarySmtpAddress
        except Exception:
            pass
        try:
            return mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        except Exception:
            pass
        return str(mail.SenderEmailAddress or "")

    # ---------------------------
    # VERARBEITUNG
    # ---------------------------

    def build_verarbeitung(self):
        # Toolbar
        bar = ttk.Frame(self.frame_verarbeitung)
        bar.grid(row=0, column=0, columnspan=6, sticky="ew", padx=8, pady=8)

        ttk.Button(bar, text="DOCX prüfen & extrahieren",
                command=self.on_process_docx).pack(side="left", padx=(0, 8))
        ttk.Button(bar, text="Export (SAP+DS) & verschieben",
                command=self.on_export_and_move).pack(side="left", padx=(0, 8))
        # 👉 neuer Button für PDFs
        ttk.Button(bar, text="PDFs verarbeiten",
                command=self.on_process_pdfs).pack(side="left")

        # Statuslabel
        self.proc_status = ttk.Label(self.frame_verarbeitung, text="Noch kein Lauf.", foreground="gray")
        self.proc_status.grid(row=1, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

        # Ergebnis-Tabelle DOCX
        cols = ["Datei", "Typ", "PN", "Name", "Status", "Grund", "Gesamteindruck (RB)"]
        self.tree_proc = ttk.Treeview(self.frame_verarbeitung, columns=cols, show="headings", height=12)
        for c in cols:
            self.tree_proc.heading(c, text=c)
            self.tree_proc.column(c, width=180 if c not in ("Datei", "Grund") else 260, anchor="w")
        self.tree_proc.grid(row=2, column=0, columnspan=6, sticky="nsew", padx=8, pady=8)

        # Ergebnis-Tabelle PDFs
        pdf_cols = ["Datei", "Typ", "PN", "Ziel", "Status"]
        self.tree_pdfs = ttk.Treeview(self.frame_verarbeitung, columns=pdf_cols, show="headings", height=6)
        for c in pdf_cols:
            self.tree_pdfs.heading(c, text=c)
            self.tree_pdfs.column(c, width=200 if c == "Ziel" else 140, anchor="w")
        self.tree_pdfs.grid(row=3, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0, 8))

        # Grid-Konfiguration
        self.frame_verarbeitung.grid_rowconfigure(2, weight=3)  # DOCX-Tabelle größer
        self.frame_verarbeitung.grid_rowconfigure(3, weight=1)  # PDFs kleiner
        self.frame_verarbeitung.grid_columnconfigure(5, weight=1)


    def on_process_docx(self):
        # Quelle: /ruecklauf
        input_dir = Path(__file__).parent.parent / "ruecklauf"

        # Tabelle leeren
        for i in self.tree_proc.get_children():
            self.tree_proc.delete(i)

        # Stammdaten laden
        sap_df = load_employees()

        # Prozess laufen lassen
        results = process_docx_folder(input_dir, sap_df)

        ok_count = 0
        man_count = 0

        for r in results:
            gi = r["extras"].get("rb_gesamteindruck", "") if isinstance(r.get("extras"), dict) else ""
            self.tree_proc.insert("", "end", values=[
                r.get("file",""), r.get("typ",""), r.get("pn",""), r.get("name",""),
                r.get("status",""), r.get("reason",""), gi
            ])
            if r.get("status") == "ok":
                ok_count += 1
            elif r.get("status") == "manuell":
                man_count += 1

        self.proc_status.config(
            text=f"DOCX geprüft: {len(results)} Dateien • OK: {ok_count} • Manuell: {man_count}",
            foreground="black"
        )

        self._last_docx_results = results


    def on_export_and_move(self):
        input_dir = Path(__file__).parent.parent / "ruecklauf"
        sap_out = Path(__file__).parent.parent / "sap_massenupload" / "massenupload.xlsx"
        ds_out  = Path(__file__).parent.parent / "tracking" / "ds_export" / "docx_extract.csv"

        # 1) Wenn keine letzte Prüfung im UI vorhanden ist, einmal neu prüfen
        if not hasattr(self, "_last_docx_results"):
            sap_df = load_employees()
            self._last_docx_results = process_docx_folder(input_dir, sap_df)

        results = self._last_docx_results

        # 2) Exporte schreiben
        try:
            sap_df = load_employees()
            export_sap_massenupload(results, sap_df, sap_out)
            export_ds_csv(results, ds_out)
        except Exception as e:
            messagebox.showerror("Export-Fehler", f"Export fehlgeschlagen:\n{e}")
            return

        # 3) Verschieben
        try:
            moved_ok, moved_man = move_after_processing(input_dir, results)
        except Exception as e:
            messagebox.showerror("Verschiebefehler", f"Verschieben fehlgeschlagen:\n{e}")
            return

        messagebox.showinfo(
            "Fertig",
            f"Export geschrieben:\n- SAP: {sap_out}\n- DS:  {ds_out}\n\n"
            f"Verschoben:\n- OK → archiv: {moved_ok}\n- manuell → manuell: {moved_man}"
        )

    def on_export_ds(self):
        if not hasattr(self, "_last_docx_results"):
            messagebox.showwarning("Hinweis", "Bitte zuerst DOCX prüfen.")
            return
        out_csv = Path("export") / "ds_export.csv"
        export_ds_csv(self._last_docx_results, out_csv)
        messagebox.showinfo("Erfolg", f"DS-Export geschrieben: {out_csv}")

    def on_process_pdfs(self):
        in_dir = Path(__file__).parent.parent / "ruecklauf"
        out_root = Path(__file__).parent.parent  # oder wohin du verschiebst

        if hasattr(self, "sap_df"):
            sap_df = self.sap_df
        else:
            from app.data_loader import load_employees
            sap_df = load_employees()
            self.sap_df = sap_df

        results = process_pdfs(in_dir, out_root, sap_df)

        # Treeview leeren
        for item in self.tree_pdfs.get_children():
            self.tree_pdfs.delete(item)

        # Ergebnisse einfüllen
        for r in results:
            self.tree_pdfs.insert("", "end", values=(
                r.get("file", ""),
                r.get("typ", ""),
                r.get("pn", ""),
                r.get("target", ""),
                r.get("status", "")
            ))

        from tkinter import messagebox
        messagebox.showinfo("Fertig", f"{len(results)} PDFs verarbeitet.")


if __name__ == "__main__":
    App().mainloop()