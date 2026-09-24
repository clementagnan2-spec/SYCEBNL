# -*- coding: utf-8 -*-
"""
main.py
=======
Application de bureau (Tkinter) : "Générateur de liasse SYCEBNL".
Point d'entrée utilisé aussi bien en développement (`python src/main.py`)
que pour la compilation en .exe via PyInstaller.
"""

import os
import sys
import traceback
from datetime import date

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Permet les imports "plats" (mapping_sycebnl, engine, ...) que l'app soit
# lancée depuis n'importe quel répertoire, en dev ou une fois compilée.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import importer
import liasse_association
import liasse_projet
import excel_export
import templates

APP_TITLE = "Générateur de liasse SYCEBNL"
APP_VERSION = "2.0.0"


def resource_path(relative_path):
    """Résout un chemin de ressource, que l'app tourne en script ou en .exe PyInstaller."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base_path, relative_path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} — v{APP_VERSION}")
        self.geometry("880x760")
        self.minsize(820, 680)

        self.mode = tk.StringVar(value="association")
        self.entite_nom = tk.StringVar()
        self.exercice = tk.StringVar(value=str(date.today().year - 1) + "-12-31")
        self.chemin_balance = tk.StringVar()
        self.chemin_balance_n1 = tk.StringVar()
        self.chemin_budget = tk.StringVar()
        self.tresorerie_ouverture = tk.StringVar(value="0")
        self.solde_releves = tk.StringVar()
        self.ajustements = []  # liste de {"libelle": str, "montant": float}

        self._build_menu()
        self._build_layout()
        self._on_mode_change()

    # ------------------------------------------------------------------
    # Construction de l'interface
    # ------------------------------------------------------------------
    def _build_menu(self):
        menubar = tk.Menu(self)
        fichier = tk.Menu(menubar, tearoff=0)
        fichier.add_command(label="Générer le modèle de balance…", command=self._gen_modele_balance)
        fichier.add_command(label="Générer le modèle de budget…", command=self._gen_modele_budget)
        fichier.add_separator()
        fichier.add_command(label="Quitter", command=self.destroy)
        menubar.add_cascade(label="Fichier", menu=fichier)

        aide = tk.Menu(menubar, tearoff=0)
        aide.add_command(label="À propos", command=self._show_about)
        menubar.add_cascade(label="Aide", menu=aide)
        self.config(menu=menubar)

    def _build_layout(self):
        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)

        title = ttk.Label(outer, text=APP_TITLE, font=("Arial", 16, "bold"))
        title.pack(anchor="w")
        subtitle = ttk.Label(
            outer,
            text="Utilise le template SYCEBNL fourni comme modèle de la liasse et le remplit à partir de la balance comptable.",
            font=("Arial", 9), foreground="#555555")
        subtitle.pack(anchor="w", pady=(0, 10))

        # --- Bloc mode ---
        mode_frame = ttk.LabelFrame(outer, text="1. Type d'entité", padding=10)
        mode_frame.pack(fill="x", pady=6)
        ttk.Radiobutton(mode_frame, text="Association / ONG classique", variable=self.mode,
                         value="association", command=self._on_mode_change).pack(side="left", padx=10)
        ttk.Radiobutton(mode_frame, text="Projet de développement financé par bailleurs",
                         variable=self.mode, value="projet",
                         command=self._on_mode_change).pack(side="left", padx=10)

        # --- Bloc identification ---
        id_frame = ttk.LabelFrame(outer, text="2. Identification", padding=10)
        id_frame.pack(fill="x", pady=6)
        id_frame.columnconfigure(1, weight=1)
        id_frame.columnconfigure(3, weight=1)
        ttk.Label(id_frame, text="Nom de l'entité :").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(id_frame, textvariable=self.entite_nom).grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Label(id_frame, text="Date de clôture (ex. 2025-12-31) :").grid(row=0, column=2, sticky="w", padx=4, pady=4)
        ttk.Entry(id_frame, textvariable=self.exercice, width=16).grid(row=0, column=3, sticky="w", padx=4, pady=4)

        # --- Bloc fichiers ---
        self.files_frame = ttk.LabelFrame(outer, text="3. Données comptables", padding=10)
        self.files_frame.pack(fill="x", pady=6)
        self.files_frame.columnconfigure(1, weight=1)

        self._file_row(self.files_frame, 0, "Balance N (obligatoire) :", self.chemin_balance,
                        self._pick_balance)
        self.row_balance_n1 = self._file_row(self.files_frame, 1, "Balance N-1 (optionnel — flux de trésorerie) :",
                                              self.chemin_balance_n1, self._pick_balance_n1)
        self.row_budget = self._file_row(self.files_frame, 2, "Fichier budget (optionnel) :",
                                          self.chemin_budget, self._pick_budget)

        # --- Bloc spécifique Projet (empaqueté/dépaqueté dynamiquement par _on_mode_change) ---
        self.projet_frame = ttk.LabelFrame(outer, text="4. Informations complémentaires (mode Projet)", padding=10)
        self.projet_frame.columnconfigure(1, weight=1)
        self.projet_frame.columnconfigure(3, weight=1)

        ttk.Label(self.projet_frame, text="Trésorerie disponible à l'ouverture :").grid(
            row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(self.projet_frame, textvariable=self.tresorerie_ouverture, width=18).grid(
            row=0, column=1, sticky="w", padx=4, pady=4)
        ttk.Label(self.projet_frame, text="Solde relevés bancaires (optionnel) :").grid(
            row=0, column=2, sticky="w", padx=4, pady=4)
        ttk.Entry(self.projet_frame, textvariable=self.solde_releves, width=18).grid(
            row=0, column=3, sticky="w", padx=4, pady=4)

        ttk.Label(self.projet_frame, text="Ajustements de réconciliation (chèques non débités, etc.) :").grid(
            row=1, column=0, columnspan=4, sticky="w", padx=4, pady=(10, 2))

        adj_row = ttk.Frame(self.projet_frame)
        adj_row.grid(row=2, column=0, columnspan=4, sticky="ew", padx=4)
        self.adj_libelle = tk.StringVar()
        self.adj_montant = tk.StringVar()
        ttk.Entry(adj_row, textvariable=self.adj_libelle, width=40).pack(side="left", padx=(0, 6))
        ttk.Entry(adj_row, textvariable=self.adj_montant, width=14).pack(side="left", padx=(0, 6))
        ttk.Button(adj_row, text="Ajouter", command=self._add_ajustement).pack(side="left", padx=(0, 6))
        ttk.Button(adj_row, text="Retirer la sélection", command=self._remove_ajustement).pack(side="left")

        self.adj_list = tk.Listbox(self.projet_frame, height=4)
        self.adj_list.grid(row=3, column=0, columnspan=4, sticky="ew", padx=4, pady=(4, 0))

        # --- Bouton génération ---
        self.action_frame = ttk.Frame(outer)
        self.action_frame.pack(fill="x", pady=12)
        self.generate_btn = ttk.Button(self.action_frame, text="Générer la liasse Excel",
                                        command=self._generate)
        self.generate_btn.pack(side="left")
        ttk.Button(self.action_frame, text="Modèle de balance", command=self._gen_modele_balance).pack(side="left", padx=8)
        ttk.Button(self.action_frame, text="Modèle de budget", command=self._gen_modele_budget).pack(side="left")

        # --- Journal ---
        log_frame = ttk.LabelFrame(outer, text="Journal", padding=6)
        log_frame.pack(fill="both", expand=True, pady=(4, 0))
        self.log = scrolledtext.ScrolledText(log_frame, height=10, state="disabled",
                                              font=("Consolas", 9))
        self.log.pack(fill="both", expand=True)

        self._log("Bienvenue. Choisissez le type d'entité, importez la balance, puis "
                   "cliquez sur « Générer la liasse Excel ».")
        self._log("⚠ Ce logiciel produit une aide au classement comptable automatisée ; "
                   "faites toujours valider le résultat par un professionnel avant tout "
                   "dépôt officiel.")

    def _file_row(self, parent, row, label, var, command):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=4, pady=4)
        entry = ttk.Entry(parent, textvariable=var, state="readonly")
        entry.grid(row=row, column=1, sticky="ew", padx=4, pady=4)
        btn = ttk.Button(parent, text="Parcourir…", command=command)
        btn.grid(row=row, column=2, sticky="w", padx=4, pady=4)
        return row

    # ------------------------------------------------------------------
    # Actions UI
    # ------------------------------------------------------------------
    def _on_mode_change(self):
        if self.mode.get() == "association":
            self._set_row_state(self.files_frame, self.row_balance_n1, True)
            self._set_row_state(self.files_frame, self.row_budget, False)
            self.projet_frame.pack_forget()
        else:
            self._set_row_state(self.files_frame, self.row_balance_n1, False)
            self._set_row_state(self.files_frame, self.row_budget, True)
            self.projet_frame.pack(fill="x", pady=6, before=self.action_frame)

    def _set_row_state(self, frame, row, visible):
        widgets = [w for w in frame.grid_slaves() if int(w.grid_info()["row"]) == row]
        for w in widgets:
            if visible:
                w.grid()
            else:
                w.grid_remove()

    def _pick_balance(self):
        path = filedialog.askopenfilename(title="Sélectionner la balance N",
                                           filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Tous", "*.*")])
        if path:
            self.chemin_balance.set(path)

    def _pick_balance_n1(self):
        path = filedialog.askopenfilename(title="Sélectionner la balance N-1",
                                           filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Tous", "*.*")])
        if path:
            self.chemin_balance_n1.set(path)

    def _pick_budget(self):
        path = filedialog.askopenfilename(title="Sélectionner le fichier budget",
                                           filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Tous", "*.*")])
        if path:
            self.chemin_budget.set(path)

    def _add_ajustement(self):
        libelle = self.adj_libelle.get().strip()
        montant_raw = self.adj_montant.get().strip().replace(",", ".")
        if not libelle or not montant_raw:
            messagebox.showwarning(APP_TITLE, "Renseignez un libellé et un montant pour l'ajustement.")
            return
        try:
            montant = float(montant_raw)
        except ValueError:
            messagebox.showerror(APP_TITLE, "Le montant de l'ajustement doit être un nombre.")
            return
        self.ajustements.append({"libelle": libelle, "montant": montant})
        self.adj_list.insert("end", f"{libelle} : {montant:,.0f}")
        self.adj_libelle.set("")
        self.adj_montant.set("")

    def _remove_ajustement(self):
        sel = list(self.adj_list.curselection())
        for idx in reversed(sel):
            self.adj_list.delete(idx)
            del self.ajustements[idx]

    def _gen_modele_balance(self):
        path = filedialog.asksaveasfilename(title="Enregistrer le modèle de balance",
                                             defaultextension=".xlsx",
                                             initialfile="modele_balance_sycebnl.xlsx",
                                             filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            templates.create_balance_template(path)
            self._log(f"Modèle de balance créé : {path}")
            messagebox.showinfo(APP_TITLE, "Modèle de balance créé avec succès.")
        except Exception as e:
            self._log(f"Erreur lors de la création du modèle : {e}")
            messagebox.showerror(APP_TITLE, str(e))

    def _gen_modele_budget(self):
        path = filedialog.asksaveasfilename(title="Enregistrer le modèle de budget",
                                             defaultextension=".xlsx",
                                             initialfile="modele_budget_sycebnl.xlsx",
                                             filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            templates.create_budget_template(path)
            self._log(f"Modèle de budget créé : {path}")
            messagebox.showinfo(APP_TITLE, "Modèle de budget créé avec succès.")
        except Exception as e:
            self._log(f"Erreur lors de la création du modèle : {e}")
            messagebox.showerror(APP_TITLE, str(e))

    def _show_about(self):
        messagebox.showinfo(
            "À propos",
            f"{APP_TITLE} — v{APP_VERSION}\n\n"
            "Outil d'aide à la préparation des états financiers SYCEBNL "
            "(Système Comptable des Entités à But Non Lucratif — OHADA).\n\n"
            "Ce logiciel n'est pas un service officiel de télédéclaration : "
            "vérifiez toujours le résultat avec votre expert-comptable avant "
            "tout dépôt réglementaire.")

    # ------------------------------------------------------------------
    # Génération
    # ------------------------------------------------------------------
    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def _generate(self):
        if not self.chemin_balance.get():
            messagebox.showwarning(APP_TITLE, "Veuillez sélectionner le fichier de balance N.")
            return

        self.generate_btn.configure(state="disabled")
        try:
            self._log("—" * 60)
            self._log("Import de la balance…")
            balance_df = importer.load_balance(self.chemin_balance.get())
            self._log(f"{len(balance_df)} ligne(s) importée(s).")

            entite = self.entite_nom.get().strip() or "Entité"
            exercice = self.exercice.get().strip()

            if self.mode.get() == "association":
                balance_n1_df = None
                if self.chemin_balance_n1.get():
                    self._log("Import de la balance N-1…")
                    balance_n1_df = importer.load_balance(self.chemin_balance_n1.get())
                self._log("Calcul du Bilan, du Compte de résultat et des flux de trésorerie…")
                result = liasse_association.generate(balance_df, balance_n1_df)
            else:
                budget_df = None
                if self.chemin_budget.get():
                    self._log("Import du budget…")
                    budget_df = importer.load_budget(self.chemin_budget.get())
                try:
                    tresorerie_ouverture = float(self.tresorerie_ouverture.get().replace(",", ".") or 0)
                except ValueError:
                    tresorerie_ouverture = 0.0
                solde_releves = None
                if self.solde_releves.get().strip():
                    try:
                        solde_releves = float(self.solde_releves.get().replace(",", "."))
                    except ValueError:
                        solde_releves = None
                self._log("Calcul du Bilan, du Compte de résultat, du Ressources-Emplois, "
                           "de l'exécution budgétaire et de la réconciliation de trésorerie…")
                result = liasse_projet.generate(balance_df, budget_df, tresorerie_ouverture,
                                                 solde_releves, list(self.ajustements))

            if not result["non_classes"].empty:
                self._log(f"⚠ {len(result['non_classes'])} compte(s) non classé(s) automatiquement "
                           "(voir la feuille dédiée dans le fichier généré).")

            default_name = f"Liasse_SYCEBNL_{entite}_{exercice or 'exercice'}.xlsx".replace(" ", "_")
            out_path = filedialog.asksaveasfilename(
                title="Enregistrer la liasse générée", defaultextension=".xlsx",
                initialfile=default_name, filetypes=[("Excel", "*.xlsx")])
            if not out_path:
                self._log("Génération annulée par l'utilisateur (aucun emplacement choisi).")
                return

            excel_export.export_liasse(result, out_path, entite, exercice)
            self._log(f"✓ Liasse SYCEBNL générée dans le template officiel : {out_path}")

            bilan_ok = result["bilan"]["equilibre"]
            if not bilan_ok:
                self._log(f"⚠ ATTENTION : le bilan n'est pas équilibré (écart = "
                           f"{result['bilan']['ecart']:,.0f}). Vérifiez la balance importée.")
            messagebox.showinfo(APP_TITLE, "Liasse SYCEBNL générée dans le template officiel :\n" + out_path)

        except importer.ImportError_ as e:
            self._log(f"✗ Erreur d'import : {e}")
            messagebox.showerror(APP_TITLE, str(e))
        except Exception as e:
            self._log(f"✗ Erreur inattendue : {e}")
            self._log(traceback.format_exc())
            messagebox.showerror(APP_TITLE, f"Une erreur inattendue est survenue :\n{e}")
        finally:
            self.generate_btn.configure(state="normal")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
