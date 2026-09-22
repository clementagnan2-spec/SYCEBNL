# -*- coding: utf-8 -*-
"""
engine.py
=========
Classe chaque ligne de la balance dans une rubrique de la liasse et calcule
les totaux du Bilan et du Compte de résultat.
"""

from collections import OrderedDict

import pandas as pd

from mapping_sycebnl import (
    RUBRIQUES_ACTIF, RUBRIQUES_PASSIF, RUBRIQUES_CHARGES,
    RUBRIQUES_PRODUITS, RUBRIQUES_HAO, match_prefix,
)


def _best_match(compte: str, rubriques) -> "Rubrique|None":
    """Renvoie la rubrique dont le préfixe correspondant est le plus long
    (la plus spécifique) parmi celles qui matchent le compte."""
    candidates = [r for r in rubriques if match_prefix(compte, r.prefixes)]
    if not candidates:
        return None
    return max(candidates, key=lambda r: max(len(p) for p in r.prefixes if compte.startswith(p)))


def classify_balance(df: pd.DataFrame):
    """
    Prend le DataFrame de balance normalisé (colonnes: compte, intitule,
    debit, credit, solde) et renvoie :
        - lines: DataFrame enrichi avec rubrique_code, rubrique_libelle,
                 section, montant (toujours positif)
        - non_classes: DataFrame des comptes n'ayant trouvé aucune rubrique
    """
    rows = []
    non_classes = []

    for _, row in df.iterrows():
        compte = str(row["compte"]).strip()
        solde = float(row["solde"])
        if solde == 0:
            continue
        classe1 = compte[0] if compte else ""

        rubrique = None
        section = None
        montant = abs(solde)

        if classe1 == "1":
            rubrique = _best_match(compte, RUBRIQUES_PASSIF)
            section = "PASSIF"
        elif classe1 == "2":
            rubrique = _best_match(compte, RUBRIQUES_ACTIF)
            section = "ACTIF"
        elif classe1 == "3":
            rubrique = _best_match(compte, RUBRIQUES_ACTIF)
            section = "ACTIF"
        elif classe1 == "4":
            if solde >= 0:
                rubrique = _best_match(compte, [r for r in RUBRIQUES_ACTIF if r.code.startswith("AC")])
                section = "ACTIF"
            else:
                rubrique = _best_match(compte, [r for r in RUBRIQUES_PASSIF if r.code.startswith("DT")])
                section = "PASSIF"
        elif classe1 == "5":
            if solde >= 0:
                rubrique = _best_match(compte, [r for r in RUBRIQUES_ACTIF if r.code.startswith("TA")])
                section = "ACTIF"
            else:
                rubrique = _best_match(compte, [r for r in RUBRIQUES_PASSIF if r.code.startswith("TP")])
                section = "PASSIF"
        elif classe1 == "6":
            rubrique = _best_match(compte, RUBRIQUES_CHARGES)
            section = "CHARGES"
        elif classe1 == "7":
            rubrique = _best_match(compte, RUBRIQUES_PRODUITS)
            section = "PRODUITS"
        elif classe1 == "8":
            rubrique = _best_match(compte, RUBRIQUES_HAO)
            section = rubrique.section if rubrique else None
        else:
            # classe 9 (engagements / comptabilité analytique) : hors liasse
            non_classes.append({**row.to_dict(), "raison": "Classe 9 (engagements) — hors liasse"})
            continue

        if rubrique is None:
            non_classes.append({**row.to_dict(), "raison": "Aucune rubrique correspondante"})
            continue

        rows.append({
            "compte": compte,
            "intitule": row.get("intitule", ""),
            "solde": solde,
            "montant": montant,
            "rubrique_code": rubrique.code,
            "rubrique_libelle": rubrique.libelle,
            "section": section,
        })

    lines = pd.DataFrame(rows)
    non_classes_df = pd.DataFrame(non_classes)
    return lines, non_classes_df


def totals_by_rubrique(lines: pd.DataFrame, section: str, rubriques_ref) -> "OrderedDict":
    """Renvoie un OrderedDict {libelle: montant} pour une section donnée,
    dans l'ordre de référence défini dans mapping_sycebnl, en incluant
    les rubriques à 0 pour que la présentation reste stable."""
    result = OrderedDict()
    if lines.empty:
        sub = lines
    else:
        sub = lines[lines["section"] == section]
    totals = sub.groupby("rubrique_code")["montant"].sum().to_dict() if not sub.empty else {}
    seen_codes = set()
    for r in rubriques_ref:
        if r.section != section or r.code in seen_codes:
            continue
        seen_codes.add(r.code)
        result[r.libelle] = totals.get(r.code, 0.0)
    return result


def resultat_net(lines: pd.DataFrame) -> float:
    """Résultat net = (Produits AO + Produits HAO) - (Charges AO + Charges HAO)."""
    if lines.empty:
        return 0.0
    produits = lines[lines["section"].isin(["PRODUITS", "HAO_PRODUITS"])]["montant"].sum()
    charges = lines[lines["section"].isin(["CHARGES", "HAO_CHARGES"])]["montant"].sum()
    return produits - charges
