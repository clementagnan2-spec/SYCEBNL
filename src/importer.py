# -*- coding: utf-8 -*-
"""
importer.py
============
Lecture et normalisation de la balance des comptes (obligatoire) et,
en mode "Projet de développement", du budget (optionnel) depuis un
fichier Excel (.xlsx) ou CSV.

La balance attendue comporte, dans n'importe quel ordre de colonnes,
et avec une orthographe/casse tolérante :
    - Compte        : numéro de compte SYCEBNL
    - Intitulé       : libellé du compte
    - Débit          : somme des mouvements débit (ou solde débiteur)
    - Crédit         : somme des mouvements crédit (ou solde créditeur)

Une balance N-1 (exercice précédent), optionnelle, permet de calculer
le tableau des flux de trésorerie. Elle doit avoir la même structure.
"""

import re
import unicodedata
from pathlib import Path

import pandas as pd

REQUIRED_COLUMNS = {
    "compte": ["compte", "numero de compte", "numéro de compte", "n° compte", "account"],
    "intitule": ["intitule", "intitulé", "libelle", "libellé", "designation", "désignation", "label"],
    "debit": ["debit", "débit", "solde debiteur", "solde débiteur", "debit balance"],
    "credit": ["credit", "crédit", "solde crediteur", "solde créditeur", "credit balance"],
}


class ImportError_(Exception):
    """Erreur métier lisible, affichée telle quelle dans l'interface."""
    pass


def _strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s))
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def _normalize_col(col: str) -> str:
    col = _strip_accents(str(col)).lower().strip()
    col = re.sub(r"\s+", " ", col)
    return col


def _find_column(columns, candidates):
    norm_cols = {_normalize_col(c): c for c in columns}
    for cand in candidates:
        cand_n = _normalize_col(cand)
        if cand_n in norm_cols:
            return norm_cols[cand_n]
    # recherche partielle en secours
    for norm, original in norm_cols.items():
        for cand in candidates:
            if _normalize_col(cand) in norm:
                return original
    return None


def _read_any(path: Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise ImportError_(f"Fichier introuvable : {path}")
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        # Les modèles générés par l'application contiennent un titre et une
        # notice avant la ligne d'en-têtes. On recherche donc automatiquement
        # la vraie ligne des colonnes au lieu de supposer qu'elle est en ligne 1.
        preview = pd.read_excel(path, header=None, dtype=str, nrows=15)
        header_row = None
        for idx in range(len(preview)):
            values = {_normalize_col(v) for v in preview.iloc[idx].dropna().tolist()}
            if "compte" in values and ("debit" in values or "débit" in values) and ("credit" in values or "crédit" in values):
                header_row = idx
                break
        return pd.read_excel(path, header=header_row if header_row is not None else 0, dtype=str)
    elif path.suffix.lower() in (".csv", ".txt"):
        # tente plusieurs séparateurs courants
        for sep in [";", ",", "\t"]:
            try:
                df = pd.read_csv(path, sep=sep, dtype=str, engine="python")
                if df.shape[1] > 1:
                    return df
            except Exception:
                continue
        raise ImportError_("Impossible de lire le fichier CSV (séparateur non reconnu).")
    else:
        raise ImportError_(f"Format de fichier non supporté : {path.suffix}")


def _to_float(series: pd.Series) -> pd.Series:
    def conv(v):
        if v is None:
            return 0.0
        s = str(v).strip()
        if s == "" or s.lower() in ("nan", "none"):
            return 0.0
        s = s.replace(" ", "").replace("\xa0", "")
        # gère les formats "1 234,56" et "1,234.56"
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0
    return series.apply(conv)


def load_balance(path: str) -> pd.DataFrame:
    """
    Charge une balance des comptes et renvoie un DataFrame normalisé avec
    les colonnes : compte (str), intitule (str), debit (float), credit (float),
    solde (float, positif si débiteur, négatif si créditeur).
    """
    raw = _read_any(Path(path))
    if raw.empty:
        raise ImportError_("Le fichier de balance est vide.")

    col_compte = _find_column(raw.columns, REQUIRED_COLUMNS["compte"])
    col_intitule = _find_column(raw.columns, REQUIRED_COLUMNS["intitule"])
    col_debit = _find_column(raw.columns, REQUIRED_COLUMNS["debit"])
    col_credit = _find_column(raw.columns, REQUIRED_COLUMNS["credit"])

    missing = []
    if col_compte is None:
        missing.append("Compte")
    if col_debit is None:
        missing.append("Débit")
    if col_credit is None:
        missing.append("Crédit")
    if missing:
        raise ImportError_(
            "Colonnes manquantes dans le fichier de balance : " + ", ".join(missing) +
            "\nColonnes trouvées : " + ", ".join(str(c) for c in raw.columns) +
            "\nUtilisez le modèle fourni (Fichier > Générer le modèle de balance)."
        )

    df = pd.DataFrame()
    df["compte"] = raw[col_compte].astype(str).str.strip()
    df["intitule"] = raw[col_intitule].astype(str).str.strip() if col_intitule else ""
    df["debit"] = _to_float(raw[col_debit])
    df["credit"] = _to_float(raw[col_credit])

    # supprime les lignes totalement vides / lignes de total éventuelles
    df = df[df["compte"].notna() & (df["compte"].str.strip() != "") & (df["compte"].str.lower() != "nan")]
    df = df[~df["compte"].str.lower().str.contains("total", na=False)]

    df["solde"] = df["debit"] - df["credit"]
    df = df.reset_index(drop=True)
    return df


def load_budget(path: str) -> pd.DataFrame:
    """
    Charge un fichier budget (mode Projet) avec les colonnes :
        - Rubrique          : libellé de la ligne budgétaire
        - Comptes            : préfixes de comptes associés, séparés par ";"
        - Budget             : montant budgétisé pour l'exercice
    """
    raw = _read_any(Path(path))
    if raw.empty:
        raise ImportError_("Le fichier de budget est vide.")

    col_rubrique = _find_column(raw.columns, ["rubrique", "ligne budgetaire", "ligne budgétaire", "poste"])
    col_comptes = _find_column(raw.columns, ["comptes", "comptes associes", "comptes associés", "prefixes", "préfixes"])
    col_budget = _find_column(raw.columns, ["budget", "montant budgetise", "montant budgétisé", "budget n"])

    missing = []
    if col_rubrique is None:
        missing.append("Rubrique")
    if col_comptes is None:
        missing.append("Comptes")
    if col_budget is None:
        missing.append("Budget")
    if missing:
        raise ImportError_(
            "Colonnes manquantes dans le fichier budget : " + ", ".join(missing) +
            "\nUtilisez le modèle fourni (Fichier > Générer le modèle de budget)."
        )

    df = pd.DataFrame()
    df["rubrique"] = raw[col_rubrique].astype(str).str.strip()
    df["comptes"] = raw[col_comptes].astype(str).str.strip()
    df["budget"] = _to_float(raw[col_budget])
    df = df[df["rubrique"].notna() & (df["rubrique"].str.strip() != "")]
    df = df.reset_index(drop=True)
    return df
