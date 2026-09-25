# -*- coding: utf-8 -*-
"""
liasse_association.py
======================
Génère les 4 états financiers obligatoires SYCEBNL pour une entité de type
"Association / ONG classique" :
    1. Bilan
    2. Compte de résultat
    3. Tableau des flux de trésorerie
    4. Notes annexes (synthèse de base)
"""

from collections import OrderedDict

from mapping_sycebnl import RUBRIQUES_ACTIF, RUBRIQUES_PASSIF, RUBRIQUES_CHARGES, RUBRIQUES_PRODUITS
from engine import classify_balance, totals_by_rubrique, resultat_net


def build_bilan(lines, resultat):
    actif = totals_by_rubrique(lines, "ACTIF", RUBRIQUES_ACTIF)
    passif = totals_by_rubrique(lines, "PASSIF", RUBRIQUES_PASSIF)

    # Les rubriques d'amortissement/dépréciation viennent en déduction de l'actif brut
    deductions = ["(-) Amortissements et dépréciations des immobilisations",
                  "(-) Dépréciations des comptes de tiers"]
    total_actif = 0.0
    for lib, montant in actif.items():
        if lib in deductions:
            total_actif -= montant
        else:
            total_actif += montant

    # Injecte le résultat net calculé dans la rubrique "Résultat net de l'exercice"
    for lib in list(passif.keys()):
        if lib == "Résultat net de l'exercice":
            passif[lib] = resultat
    total_passif = sum(passif.values())

    return {
        "actif": actif,
        "passif": passif,
        "total_actif": total_actif,
        "total_passif": total_passif,
        "equilibre": round(total_actif - total_passif, 2) == 0,
        "ecart": round(total_actif - total_passif, 2),
    }


def build_compte_resultat(lines):
    charges = totals_by_rubrique(lines, "CHARGES", RUBRIQUES_CHARGES)
    produits = totals_by_rubrique(lines, "PRODUITS", RUBRIQUES_PRODUITS)
    total_charges_ao = sum(charges.values())
    total_produits_ao = sum(produits.values())
    resultat_ao = total_produits_ao - total_charges_ao

    haos = lines[lines["section"].isin(["HAO_CHARGES", "HAO_PRODUITS"])] if not lines.empty else lines
    total_charges_hao = haos[haos["section"] == "HAO_CHARGES"]["montant"].sum() if not haos.empty else 0.0
    total_produits_hao = haos[haos["section"] == "HAO_PRODUITS"]["montant"].sum() if not haos.empty else 0.0
    resultat_hao = total_produits_hao - total_charges_hao

    resultat_net_ex = resultat_ao + resultat_hao

    return {
        "charges": charges,
        "produits": produits,
        "total_charges_ao": total_charges_ao,
        "total_produits_ao": total_produits_ao,
        "resultat_ao": resultat_ao,
        "total_charges_hao": total_charges_hao,
        "total_produits_hao": total_produits_hao,
        "resultat_hao": resultat_hao,
        "resultat_net": resultat_net_ex,
    }


def _classe2_brut(lines):
    if lines.empty:
        return 0.0
    mask = lines["compte"].str.startswith(("20", "21", "22", "23", "24", "26", "27"))
    return lines[mask]["montant"].sum()


def _classe1_hors_resultat(lines, resultat_courant):
    """Ressources durables hors résultat de l'exercice (pour le flux de financement)."""
    if lines.empty:
        return 0.0
    mask = lines["compte"].str.startswith("1") & ~lines["compte"].str.startswith("13")
    return lines[mask]["montant"].sum()


def _tresorerie_nette(lines):
    if lines.empty:
        return 0.0
    actif_tre = lines[(lines["compte"].str.startswith(("5",))) & (lines["section"] == "ACTIF")]["montant"].sum()
    passif_tre = lines[(lines["compte"].str.startswith(("5",))) & (lines["section"] == "PASSIF")]["montant"].sum()
    return actif_tre - passif_tre


def build_flux_tresorerie(lines_n, resultat_n, lines_n1=None, resultat_n1=None):
    """
    Tableau des flux de trésorerie (méthode indirecte, simplifiée).
    Nécessite la balance de l'exercice précédent (N-1) pour calculer les
    variations. Si elle n'est pas fournie, renvoie disponible=False.
    """
    if lines_n1 is None:
        return {"disponible": False}

    tresorerie_n = _tresorerie_nette(lines_n)
    tresorerie_n1 = _tresorerie_nette(lines_n1)

    dotations = lines_n[lines_n["rubrique_code"] == "CH_DOT"]["montant"].sum() if not lines_n.empty else 0.0
    cafg = resultat_n + dotations

    stocks_n = lines_n[lines_n["rubrique_code"] == "AC_STK"]["montant"].sum() if not lines_n.empty else 0.0
    stocks_n1 = lines_n1[lines_n1["rubrique_code"] == "AC_STK"]["montant"].sum() if not lines_n1.empty else 0.0
    creances_n = lines_n[lines_n["rubrique_code"] == "AC_CRE"]["montant"].sum() if not lines_n.empty else 0.0
    creances_n1 = lines_n1[lines_n1["rubrique_code"] == "AC_CRE"]["montant"].sum() if not lines_n1.empty else 0.0
    dettes_n = lines_n[lines_n["rubrique_code"] == "DT_CIR"]["montant"].sum() if not lines_n.empty else 0.0
    dettes_n1 = lines_n1[lines_n1["rubrique_code"] == "DT_CIR"]["montant"].sum() if not lines_n1.empty else 0.0

    variation_bfr = (stocks_n - stocks_n1) + (creances_n - creances_n1) - (dettes_n - dettes_n1)
    flux_activites = cafg - variation_bfr

    immo_n = _classe2_brut(lines_n)
    immo_n1 = _classe2_brut(lines_n1)
    flux_investissement = -(immo_n - immo_n1)

    ressources_n = _classe1_hors_resultat(lines_n, resultat_n)
    ressources_n1 = _classe1_hors_resultat(lines_n1, resultat_n1 or 0.0)
    flux_financement = ressources_n - ressources_n1

    variation_tresorerie_calculee = flux_activites + flux_investissement + flux_financement

    return {
        "disponible": True,
        "tresorerie_ouverture": tresorerie_n1,
        "cafg": cafg,
        "variation_bfr": variation_bfr,
        "flux_activites": flux_activites,
        "flux_investissement": flux_investissement,
        "flux_financement": flux_financement,
        "variation_tresorerie_calculee": variation_tresorerie_calculee,
        "tresorerie_cloture": tresorerie_n,
        "ecart_controle": round((tresorerie_n1 + variation_tresorerie_calculee) - tresorerie_n, 2),
    }


def build_notes_annexes(lines, non_classes, resultat):
    """Notes annexes de base : liste des principales masses + alertes de classement."""
    notes = []
    notes.append("1. Référentiel comptable : Système Comptable des Entités à But Non "
                  "Lucratif (SYCEBNL - Acte uniforme OHADA, en vigueur depuis le 1er "
                  "janvier 2024).")
    notes.append(f"2. Résultat net de l'exercice : {resultat:,.0f} (Produits des activités "
                  "moins Charges des activités, y compris éléments HAO).")
    if non_classes is not None and not non_classes.empty:
        notes.append(f"3. ATTENTION : {len(non_classes)} ligne(s) de la balance n'ont pas pu "
                      "être classées automatiquement (voir feuille 'Comptes non classés'). "
                      "Elles ne sont PAS incluses dans le Bilan ni le Compte de résultat.")
    else:
        notes.append("3. Toutes les lignes de la balance ont été classées.")
    notes.append("4. Ce document est généré automatiquement à partir de la balance importée. "
                  "Il doit être relu et validé par un professionnel comptable avant tout "
                  "dépôt réglementaire.")
    return notes


def generate(balance_df, balance_n1_df=None):
    lines, non_classes = classify_balance(balance_df)
    resultat = resultat_net(lines)
    bilan = build_bilan(lines, resultat)
    compte_resultat = build_compte_resultat(lines)

    flux = {"disponible": False}
    if balance_n1_df is not None:
        lines_n1, _ = classify_balance(balance_n1_df)
        resultat_n1 = resultat_net(lines_n1)
        flux = build_flux_tresorerie(lines, resultat, lines_n1, resultat_n1)

    notes = build_notes_annexes(lines, non_classes, resultat)

    return {
        "mode": "association",
        "lines": lines,
        "lines_n1": lines_n1 if balance_n1_df is not None else None,
        "non_classes": non_classes,
        "bilan": bilan,
        "compte_resultat": compte_resultat,
        "flux_tresorerie": flux,
        "notes": notes,
        "resultat_net": resultat,
    }
