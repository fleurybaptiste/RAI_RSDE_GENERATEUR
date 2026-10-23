# -*- coding: utf-8 -*-
"""
Application unifiee SANDRE RSDE (FCT_ASSAIN v4).

Enchaine l'etape 1 (extraction Tetraed -> fichier a plat) et l'etape 2
(fichier a plat -> XML FCT_ASSAIN v4) dans une seule interface moderne
(CustomTkinter), pensee pour etre utilisee par n'importe quel operateur
sans connaitre le fonctionnement interne.

Ne duplique aucune logique metier : tout est reutilise depuis les modules
deja valides sur des campagnes reelles, sandre_etape1_gui.py (etape 1) et
xlsx_to_sandre_xml.py (etape 2).
"""
import os
import re
import sys
import datetime as dt
import threading
import traceback
from collections import Counter
from tkinter import filedialog, messagebox

import customtkinter as ctk

import sandre_etape1_gui as etape1
import xlsx_to_sandre_xml as etape2


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

APP_DIR = etape1.app_dir()
LOGS_DIR = os.path.join(APP_DIR, "logs")


def resource_path(relative):
    """Chemin vers une ressource embarquee (icone...), que l'app tourne depuis
    le script source ou depuis l'exe compile (PyInstaller --add-data)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def log_error(exc=None):
    os.makedirs(LOGS_DIR, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    path = os.path.join(LOGS_DIR, f"erreur_{stamp}.log")
    with open(path, "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())
    return path


def show_error(title, message, exc=None):
    full = message
    if exc is not None:
        path = log_error(exc)
        full += f"\n\nDetail technique enregistre dans :\n{path}"
    messagebox.showerror(title, full)


# ---------------------------------------------------------------------------
# Ecran de base
# ---------------------------------------------------------------------------
class BaseScreen(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color="transparent")
        self.app = app

    def on_show(self):
        pass


# ---------------------------------------------------------------------------
# Gestion des profils de station (equivalent de ProfileDialog en CTk)
# ---------------------------------------------------------------------------
class ProfileDialogCTk(ctk.CTkToplevel):
    def __init__(self, parent, initial=None, on_saved=None):
        super().__init__(parent)
        self.title("Profil de station")
        self.geometry("580x640")
        self.on_saved = on_saved
        self.entries = {}
        self.transient(parent)
        self.grab_set()

        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(fill="both", expand=True, padx=14, pady=14)

        initial = initial or {}
        for key, label in etape1.PROFILE_FIELDS:
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=4)
            ctk.CTkLabel(row, text=label, anchor="w", wraplength=260, width=260,
                         justify="left").pack(side="left", padx=(0, 10))
            entry = ctk.CTkEntry(row)
            entry.insert(0, initial.get(key, etape1.PROFILE_DEFAULTS.get(key, "")))
            entry.pack(side="left", fill="x", expand=True)
            self.entries[key] = entry

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkButton(btns, text="Enregistrer", command=self.on_save).pack(side="right", padx=4)
        ctk.CTkButton(btns, text="Annuler", fg_color="transparent", border_width=1,
                      command=self.destroy).pack(side="right", padx=4)

    def on_save(self):
        data = {k: e.get().strip() for k, e in self.entries.items()}
        if not data.get("nom_profil"):
            messagebox.showwarning("Nom manquant", "Donnez un nom au profil.", parent=self)
            return
        try:
            etape1.save_profile(data)
        except Exception as e:
            show_error("Erreur", str(e), e)
            return
        if self.on_saved:
            self.on_saved()
        self.destroy()


# ---------------------------------------------------------------------------
# Accueil
# ---------------------------------------------------------------------------
class HomeScreen(BaseScreen):
    def __init__(self, parent, app):
        super().__init__(parent, app)

        wrap = ctk.CTkFrame(self, fg_color="transparent")
        wrap.place(relx=0.5, rely=0.42, anchor="center")

        ctk.CTkLabel(wrap, text="Generation RSDE", font=ctk.CTkFont(size=30, weight="bold")).pack(pady=(0, 4))
        ctk.CTkLabel(wrap, text="Scenario SANDRE FCT_ASSAIN v4 — Autosurveillance RSDE",
                     font=ctk.CTkFont(size=14), text_color=("gray30", "gray70")).pack(pady=(0, 36))

        ctk.CTkButton(wrap, text="Nouvelle generation complete", width=360, height=56,
                      font=ctk.CTkFont(size=16, weight="bold"),
                      command=self.start_full).pack(pady=8)
        ctk.CTkButton(wrap, text="Etape 2 seule (fichier a plat existant)", width=360, height=46,
                      font=ctk.CTkFont(size=14), fg_color="transparent", border_width=2,
                      text_color=("gray10", "gray90"),
                      command=self.start_step2_only).pack(pady=8)

        theme_row = ctk.CTkFrame(self, fg_color="transparent")
        theme_row.place(relx=0.97, rely=0.04, anchor="ne")
        ctk.CTkLabel(theme_row, text="Theme sombre").pack(side="left", padx=(0, 8))
        self.theme_switch = ctk.CTkSwitch(theme_row, text="", width=40, command=self.toggle_theme)
        self.theme_switch.pack(side="left")

        ctk.CTkLabel(self, text="IANESCO SAS — outil interne RSDE",
                     font=ctk.CTkFont(size=11),
                     text_color=("gray50", "gray50")).place(relx=0.5, rely=0.96, anchor="center")

    def toggle_theme(self):
        ctk.set_appearance_mode("dark" if self.theme_switch.get() else "light")

    def start_full(self):
        self.app.app_state.clear()
        self.app.frames["Etape1Screen"].reset_for_new_run()
        self.app.show_frame("Etape1Screen")

    def start_step2_only(self):
        self.app.app_state.clear()
        self.app.frames["Etape2Screen"].reset_for_direct_entry()
        self.app.show_frame("Etape2Screen")


# ---------------------------------------------------------------------------
# Etape 1 : extraction -> fichier a plat
# ---------------------------------------------------------------------------
class Etape1Screen(BaseScreen):
    def __init__(self, parent, app):
        super().__init__(parent, app)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(20, 8))
        ctk.CTkButton(header, text="< Retour", width=90, fg_color="transparent", border_width=1,
                      command=lambda: self.app.show_frame("HomeScreen")).pack(side="left")
        ctk.CTkLabel(header, text="Etape 1 — Extraction vers fichier a plat",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", padx=16)

        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=24, pady=8)

        files_box = ctk.CTkFrame(body)
        files_box.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(files_box, text="Fichiers", font=ctk.CTkFont(size=15, weight="bold")).pack(
            anchor="w", padx=14, pady=(10, 4))

        self.path_rows = {}
        self._file_row(files_box, "Extraction brute du labo (.xlsx/.xlsm)", "extraction", True)
        self._file_row(files_box, "Parametres.xlsx", "parametres", True)
        self._file_row(files_box, "Unites.xlsx", "unites", True)
        self._file_row(files_box, "Fractions.xlsx", "fractions", True)
        self._file_row(files_box, "Methodes_Labo_Sandre.xlsx", "methodes", True,
                        default=os.path.join(APP_DIR, "Methodes_Labo_Sandre.xlsx"))
        self._file_row(files_box, "Referentiel LQ/Incertitude (optionnel)", "referentiel", False,
                        default=os.path.join(APP_DIR, "Referentiel_LQ_Incertitude_IANESCO.xlsx"))
        self._file_row(files_box, "Reference multi-stations (optionnel)", "points_stations", False)

        station_box = ctk.CTkFrame(body)
        station_box.pack(fill="x", pady=(0, 14))
        ctk.CTkLabel(station_box, text="Station / periode", font=ctk.CTkFont(size=15, weight="bold")).pack(
            anchor="w", padx=14, pady=(10, 4))

        prof_row = ctk.CTkFrame(station_box, fg_color="transparent")
        prof_row.pack(fill="x", padx=14, pady=6)
        ctk.CTkLabel(prof_row, text="Profil :", width=150, anchor="w").pack(side="left")
        self.profile_combo = ctk.CTkComboBox(prof_row, values=[], width=260)
        self.profile_combo.pack(side="left", padx=(0, 8))
        ctk.CTkButton(prof_row, text="Nouveau...", width=90, command=self.new_profile).pack(side="left", padx=4)
        ctk.CTkButton(prof_row, text="Modifier...", width=90, command=self.edit_profile).pack(side="left", padx=4)
        ctk.CTkButton(prof_row, text="Actualiser", width=90, command=self.refresh_profiles).pack(side="left", padx=4)

        date_row = ctk.CTkFrame(station_box, fg_color="transparent")
        date_row.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkLabel(date_row, text="Date debut (AAAA-MM-JJ) :", width=190, anchor="w").pack(side="left")
        self.date_debut = ctk.CTkEntry(date_row, width=140)
        self.date_debut.pack(side="left", padx=(0, 24))
        ctk.CTkLabel(date_row, text="Date fin (AAAA-MM-JJ) :", width=180, anchor="w").pack(side="left")
        self.date_fin = ctk.CTkEntry(date_row, width=140)
        self.date_fin.pack(side="left")

        action = ctk.CTkFrame(self, fg_color="transparent")
        action.pack(fill="x", padx=24, pady=(4, 6))
        self.gen_btn = ctk.CTkButton(action, text="Generer le fichier a plat", height=44,
                                      font=ctk.CTkFont(size=14, weight="bold"), command=self.on_generate)
        self.gen_btn.pack(side="left")
        self.status_label = ctk.CTkLabel(action, text="", text_color=("gray30", "gray70"))
        self.status_label.pack(side="left", padx=16)

        self.progress = ctk.CTkProgressBar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=24, pady=(0, 16))
        self.progress.set(0)

        self.refresh_profiles()

    def _file_row(self, parent, label, key, required, default=""):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=4)
        ctk.CTkLabel(row, text=label, width=280, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(row)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        if default:
            entry.insert(0, default)
        ctk.CTkButton(row, text="Parcourir...", width=100,
                      command=lambda e=entry: self._browse(e)).pack(side="left")
        self.path_rows[key] = (entry, required)

    def _browse(self, entry):
        path = filedialog.askopenfilename(title="Choisir un fichier",
                                           filetypes=[("Classeur Excel", "*.xlsx *.xlsm")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def refresh_profiles(self):
        names = etape1.list_profiles()
        self.profile_combo.configure(values=names)
        if names and not self.profile_combo.get():
            self.profile_combo.set(names[0])

    def new_profile(self):
        ProfileDialogCTk(self, on_saved=self.refresh_profiles)

    def edit_profile(self):
        name = self.profile_combo.get()
        if not name:
            messagebox.showwarning("Aucun profil", "Choisissez d'abord un profil.")
            return
        data = etape1.load_profile(name)
        ProfileDialogCTk(self, initial=data, on_saved=self.refresh_profiles)

    def reset_for_new_run(self):
        self.status_label.configure(text="")

    def on_generate(self):
        paths = {}
        for key, (entry, required) in self.path_rows.items():
            val = entry.get().strip()
            if required and (not val or not os.path.isfile(val)):
                messagebox.showerror("Fichier manquant", f"Fichier introuvable ou non renseigne pour : {key}")
                return
            if val and not os.path.isfile(val):
                messagebox.showerror("Fichier manquant", f"Fichier introuvable : {val}")
                return
            paths[key] = val or None

        profile_name = self.profile_combo.get()
        if not profile_name:
            messagebox.showwarning("Profil manquant", "Choisissez ou creez un profil de station.")
            return
        date_debut = self.date_debut.get().strip()
        date_fin = self.date_fin.get().strip()
        if not date_debut or not date_fin:
            messagebox.showwarning("Dates manquantes", "Renseignez les dates de debut et de fin de periode.")
            return

        profile = etape1.load_profile(profile_name)

        self.gen_btn.configure(state="disabled")
        self.status_label.configure(text="Generation en cours...")
        self.progress.start()

        thread = threading.Thread(target=self._run, args=(paths, profile, date_debut, date_fin), daemon=True)
        thread.start()

    def _run(self, paths, profile, date_debut, date_fin):
        try:
            rows, issues = etape1.run_generation(
                paths["extraction"], paths["parametres"], paths["unites"], paths["fractions"],
                paths["methodes"], profile, date_debut, date_fin, lambda msg: None,
                referentiel_path=paths.get("referentiel"),
                points_stations_path=paths.get("points_stations"),
            )
        except Exception as e:
            self.after(0, lambda: self._on_error(e))
            return
        self.after(0, lambda: self._on_done(rows, issues, profile, date_debut, date_fin))

    def _on_error(self, exc):
        self.progress.stop()
        self.gen_btn.configure(state="normal")
        self.status_label.configure(text="Echec.")
        show_error("Erreur", "La generation a echoue.", exc)

    def _on_done(self, rows, issues, profile, date_debut, date_fin):
        self.progress.stop()
        self.gen_btn.configure(state="normal")
        self.status_label.configure(text=f"{len(rows)} ligne(s) generee(s).")

        self.app.app_state["rows"] = rows
        self.app.app_state["issues"] = issues
        self.app.app_state["profile"] = profile
        self.app.app_state["date_debut"] = date_debut
        self.app.app_state["date_fin"] = date_fin

        self.app.frames["Etape1ResultScreen"].on_show()
        self.app.show_frame("Etape1ResultScreen")


# ---------------------------------------------------------------------------
# Resultat etape 1 : enregistrement + enchainement
# ---------------------------------------------------------------------------
class Etape1ResultScreen(BaseScreen):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        wrap = ctk.CTkFrame(self, fg_color="transparent")
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(wrap, text="Fichier a plat genere", font=ctk.CTkFont(size=24, weight="bold")).pack(
            pady=(0, 16))
        self.summary = ctk.CTkLabel(wrap, text="", justify="left", font=ctk.CTkFont(size=14))
        self.summary.pack(pady=(0, 20))

        self.save_btn = ctk.CTkButton(wrap, text="Enregistrer le fichier a plat...", width=320, height=48,
                                       font=ctk.CTkFont(size=14, weight="bold"), command=self.save_file)
        self.save_btn.pack(pady=6)

        self.saved_label = ctk.CTkLabel(wrap, text="", text_color=("green4", "green3"))
        self.saved_label.pack(pady=(4, 20))

        nav = ctk.CTkFrame(wrap, fg_color="transparent")
        nav.pack(pady=10)
        self.continue_btn = ctk.CTkButton(nav, text="Continuer vers l'Etape 2 ->", width=220, height=44,
                                           state="disabled", command=self.go_step2)
        self.continue_btn.pack(side="left", padx=8)
        ctk.CTkButton(nav, text="Terminer ici", width=180, height=44, fg_color="transparent", border_width=1,
                      command=lambda: self.app.show_frame("HomeScreen")).pack(side="left", padx=8)

        self.saved_path = None

    def on_show(self):
        rows = self.app.app_state.get("rows", [])
        issues = self.app.app_state.get("issues", [])
        nb_a_completer = sum(1 for r in rows if r.get("LQAna") == "" or r.get("IncertAna") == "")
        text = (f"{len(rows)} ligne(s) d'analyse generee(s)\n"
                f"{len(issues)} point(s) signale(s) (voir le fichier _a_verifier)\n"
                f"{nb_a_completer} ligne(s) avec LQ/Incertitude a completer")
        self.summary.configure(text=text)
        self.saved_label.configure(text="")
        self.continue_btn.configure(state="disabled")
        self.saved_path = None

    def save_file(self):
        profile = self.app.app_state["profile"]
        date_debut = self.app.app_state["date_debut"]
        date_fin = self.app.app_state["date_fin"]
        safe_name = re.sub(r"[^A-Za-z0-9_\-]", "_", profile.get("nom_profil", "Station"))
        default_name = f"Resultats_{safe_name}_{date_debut}_{date_fin}.xlsx"
        path = filedialog.asksaveasfilename(title="Enregistrer le fichier a plat",
                                             defaultextension=".xlsx", initialfile=default_name,
                                             filetypes=[("Classeur Excel", "*.xlsx")])
        if not path:
            return
        try:
            rows = self.app.app_state["rows"]
            issues = self.app.app_state["issues"]
            etape1.write_resultats_xlsx(rows, path)
            base, _ = os.path.splitext(path)
            if issues:
                etape1.write_issues_xlsx(issues, base + "_a_verifier.xlsx")
            etape1.write_completion_request_xlsx(rows, base + "_LQ_Incertitude_a_completer.xlsx")
        except Exception as e:
            show_error("Erreur", "Echec de l'enregistrement.", e)
            return
        self.saved_path = path
        self.saved_label.configure(text=f"Enregistre : {path}")
        self.continue_btn.configure(state="normal")

    def go_step2(self):
        self.app.frames["Etape2Screen"].set_source_file(self.saved_path)
        self.app.show_frame("Etape2Screen")


# ---------------------------------------------------------------------------
# Etape 2 : fichier a plat -> XML
# ---------------------------------------------------------------------------
class Etape2Screen(BaseScreen):
    def __init__(self, parent, app):
        super().__init__(parent, app)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(20, 8))
        ctk.CTkButton(header, text="< Retour", width=90, fg_color="transparent", border_width=1,
                      command=lambda: self.app.show_frame("HomeScreen")).pack(side="left")
        ctk.CTkLabel(header, text="Etape 2 — Fichier a plat vers XML FCT_ASSAIN v4",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", padx=16)

        body = ctk.CTkFrame(self)
        body.pack(fill="x", padx=24, pady=20)
        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=14)
        ctk.CTkLabel(row, text="Fichier a plat (Resultats) :", width=220, anchor="w").pack(side="left")
        self.source_entry = ctk.CTkEntry(row)
        self.source_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(row, text="Parcourir...", width=100, command=self.browse_source).pack(side="left")

        action = ctk.CTkFrame(self, fg_color="transparent")
        action.pack(fill="x", padx=24, pady=(4, 6))
        self.gen_btn = ctk.CTkButton(action, text="Generer le XML", height=44,
                                      font=ctk.CTkFont(size=14, weight="bold"), command=self.on_generate)
        self.gen_btn.pack(side="left")
        self.status_label = ctk.CTkLabel(action, text="", text_color=("gray30", "gray70"))
        self.status_label.pack(side="left", padx=16)

        self.progress = ctk.CTkProgressBar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=24, pady=(0, 16))
        self.progress.set(0)

    def set_source_file(self, path):
        self.source_entry.delete(0, "end")
        self.source_entry.insert(0, path or "")

    def reset_for_direct_entry(self):
        self.source_entry.delete(0, "end")
        self.status_label.configure(text="")

    def browse_source(self):
        path = filedialog.askopenfilename(title="Choisir le fichier a plat",
                                           filetypes=[("Classeur Excel", "*.xlsx")])
        if path:
            self.set_source_file(path)

    def on_generate(self):
        path = self.source_entry.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Fichier manquant", "Choisissez un fichier a plat valide.")
            return
        self.gen_btn.configure(state="disabled")
        self.status_label.configure(text="Generation en cours...")
        self.progress.start()
        thread = threading.Thread(target=self._run, args=(path,), daemon=True)
        thread.start()

    def _run(self, path):
        try:
            rows = etape2.read_rows(path)
            problems = etape2.check_missing(rows)
            root = etape2.build_xml(rows)
        except Exception as e:
            self.after(0, lambda: self._on_error(e))
            return
        self.after(0, lambda: self._on_done(path, rows, problems, root))

    def _on_error(self, exc):
        self.progress.stop()
        self.gen_btn.configure(state="normal")
        self.status_label.configure(text="Echec.")
        show_error("Erreur", "La generation du XML a echoue.", exc)

    def _on_done(self, source_path, rows, problems, root):
        self.progress.stop()
        self.gen_btn.configure(state="normal")
        self.status_label.configure(text=f"{len(rows)} ligne(s) traitee(s).")

        self.app.app_state["source_flat_file"] = source_path
        self.app.app_state["xml_root"] = root
        self.app.app_state["xml_problems"] = problems
        self.app.app_state["xml_rows_count"] = len(rows)

        self.app.frames["Etape2ResultScreen"].on_show()
        self.app.show_frame("Etape2ResultScreen")


# ---------------------------------------------------------------------------
# Resultat etape 2 : enregistrement du XML
# ---------------------------------------------------------------------------
class Etape2ResultScreen(BaseScreen):
    def __init__(self, parent, app):
        super().__init__(parent, app)
        wrap = ctk.CTkFrame(self, fg_color="transparent")
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        self.title_label = ctk.CTkLabel(wrap, text="", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(pady=(0, 12))

        self.problems_box = ctk.CTkTextbox(wrap, width=580, height=170, wrap="word")
        self.problems_box.pack(pady=(0, 16))

        self.save_btn = ctk.CTkButton(wrap, text="Enregistrer le XML...", width=320, height=48,
                                       font=ctk.CTkFont(size=14, weight="bold"), command=self.save_xml)
        self.save_btn.pack(pady=6)
        self.saved_label = ctk.CTkLabel(wrap, text="", text_color=("green4", "green3"))
        self.saved_label.pack(pady=(4, 10))

        reminder = ("Avant depot sur Verseau : tester ce fichier sur le certificateur\n"
                    "SANDRE en ligne (sandre.eaufrance.fr) et viser 0 erreur.")
        ctk.CTkLabel(wrap, text=reminder, justify="center",
                     text_color=("gray30", "gray70")).pack(pady=(10, 20))

        ctk.CTkButton(wrap, text="Nouvelle generation", width=240, height=44,
                      fg_color="transparent", border_width=1,
                      command=lambda: self.app.show_frame("HomeScreen")).pack()

    def on_show(self):
        problems = self.app.app_state.get("xml_problems", [])
        n_rows = self.app.app_state.get("xml_rows_count", 0)
        self.problems_box.configure(state="normal")
        self.problems_box.delete("1.0", "end")
        if not problems:
            self.title_label.configure(text="XML genere — tout est complet")
            self.problems_box.insert(
                "1.0", f"{n_rows} ligne(s) traitee(s), aucun champ obligatoire manquant detecte localement.")
        else:
            counts = Counter(col for _, col in problems)
            self.title_label.configure(text="XML genere — des champs manquent")
            lines = [f"{n_rows} ligne(s) traitees, {len(problems)} champ(s) obligatoire(s) manquant(s) :"]
            for col, n in counts.most_common():
                lines.append(f"  - {col} : {n} occurrence(s)")
            lines.append("\nLe XML est quand meme genere ci-dessous : a corriger avant depot.")
            self.problems_box.insert("1.0", "\n".join(lines))
        self.problems_box.configure(state="disabled")
        self.saved_label.configure(text="")

    def save_xml(self):
        source_path = self.app.app_state.get("source_flat_file", "")
        base = os.path.splitext(os.path.basename(source_path))[0] if source_path else "RAI"
        default_name = f"SANDRE_Autosurveillance_RSDE_v4_{base}.xml"
        path = filedialog.asksaveasfilename(title="Enregistrer le XML", defaultextension=".xml",
                                             initialfile=default_name, filetypes=[("Fichier XML", "*.xml")])
        if not path:
            return
        try:
            root = self.app.app_state["xml_root"]
            xml_text = etape2.prettify(root)
            with open(path, "w", encoding="utf-8") as f:
                f.write(xml_text)
        except Exception as e:
            show_error("Erreur", "Echec de l'enregistrement du XML.", e)
            return
        self.saved_label.configure(text=f"Enregistre : {path}")


# ---------------------------------------------------------------------------
# Fenetre principale
# ---------------------------------------------------------------------------
class RSDEApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RSDE - IANESCOV4")
        self.geometry("960x720")
        self.minsize(860, 640)
        try:
            self.iconbitmap(resource_path("ianesco_rsde.ico"))
        except Exception:
            pass  # pas bloquant si l'icone est absente (ex: lancement .py sans le fichier a cote)

        self.app_state = {}

        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True)

        self.frames = {}
        for screen_cls in (HomeScreen, Etape1Screen, Etape1ResultScreen, Etape2Screen, Etape2ResultScreen):
            frame = screen_cls(container, self)
            self.frames[screen_cls.__name__] = frame
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)

        self.show_frame("HomeScreen")

    def show_frame(self, name):
        self.frames[name].tkraise()


def main():
    app = RSDEApp()
    app.mainloop()


if __name__ == "__main__":
    main()
