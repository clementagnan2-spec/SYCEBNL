# -*- coding: utf-8 -*-
"""
liasse_projet.py
=================
Génère les états financiers du mode "Projet de développement financé par des
bailleurs" (Acte uniforme SYCEBNL) :
    1. Bilan (réutilise le moteur du mode association)
    2. Compte de résultat (idem)
    3. Tableau Ressources - Emplois
    4. Tableau d'exécution budgétaire (nécessite un fichier budget)
    5. Tableau de réconciliation de trésorerie
"""

from collections import OrderedDict

from engine import classify_balance, resultat_net
from liasse_association import build_bilan, build_compte_resultat, build_notes_annexes


def build_ressources_emplois(lines, tresorerie_ouverture=0.0):
    if lines.empty:
        ressources = OrderedDict()
        emplois = OrderedDict()
    else:
        ressources = OrderedDict()
        ressources["Subventions et financements de bailleurs"] = \
            lines[lines["rubrique_code"] == "PR_SUB"]["montant"].sum()
        ressources["Autres produits (cotisations, dons, prestations, financiers)"] = \
            lines[lines["section"] == "PRODUITS"]["montant"].sum() - ressources["Subventions et financements de bailleurs"]

        emplois = OrderedDict()
        emplois["Charges par nature (achats, services, personnel, etc.)"] = \
            lines[lines["section"] == "CHARGES"]["montant"].sum()
        emplois["Acquisitions d'immobilisations"] = \
            lines[lines["compte"].str.startswith(("20", "21", "22", "23", "24", "26", "27"))]["montant"].sum()

    total_ressources = tresorerie_ouverture + sum(ressources.values())
    total_emplois = sum(emplois.values())
    solde_periode = total_ressources - total_emplois

    return {
        "tresorerie_ouverture": tresorerie_ouverture,
        "ressources": ressources,
        "emplois": emplois,
        "total_ressources": total_ressources,
        "total_emplois": total_emplois,
        "solde_tresorerie_fin": solde_periode,
    }


def build_execution_budgetaire(lines, budget_df):
    """
    Pour chaque ligne du budget (rubrique, préfixes de comptes associés,
    montant budgétisé), calcule le réalisé en sommant les montants des
    comptes de la balance dont le numéro commence par l'un des préfixes.
    """
    rows = []
    if budget_df is None or budget_df.empty:
        return {"disponible": False, "lignes": rows, "total_budget": 0.0, "total_realise": 0.0}

    for _, b in budget_df.iterrows():
        prefixes = [p.strip() for p in str(b["comptes"]).split(";") if p.strip()]
        if lines.empty or not prefixes:
            realise = 0.0
        else:
            mask = lines["compte"].str.startswith(tuple(prefixes))
            realise = lines[mask]["montant"].sum()
        budget = float(b["budget"])
        ecart = budget - realise
        taux = (realise / budget * 100.0) if budget else None
        rows.append({
            "rubrique": b["rubrique"],
            "budget": budget,
            "realise": realise,
            "ecart": ecart,
            "taux_execution": taux,
        })

    total_budget = sum(r["budget"] for r in rows)
    total_realise = sum(r["realise"] for r in rows)
    return {
        "disponible": True,
        "lignes": rows,
        "total_budget": total_budget,
        "total_realise": total_realise,
        "ecart_total": total_budget - total_realise,
        "taux_execution_global": (total_realise / total_budget * 100.0) if total_budget else None,
    }


def _tresorerie_comptable(lines):
    if lines.empty:
        return 0.0
    actif_tre = lines[(lines["compte"].str.startswith("5")) & (lines["section"] == "ACTIF")]["montant"].sum()
    passif_tre = lines[(lines["compte"].str.startswith("5")) & (lines["section"] == "PASSIF")]["montant"].sum()
    return actif_tre - passif_tre


def build_reconciliation_tresorerie(lines, solde_releves_bancaires=None, ajustements=None):
    """
    Rapproche le solde de trésorerie comptable (classe 5 de la balance) avec
    le solde des relevés bancaires/mobile money saisi manuellement par
    l'utilisateur, et une liste d'ajustements (chèques non débités, dépôts
    non crédités, etc.), chacun avec un libellé et un montant (signé).
    """
    solde_compta = _tresorerie_comptable(lines)
    ajustements = ajustements or []
    total_ajustements = sum(a["montant"] for a in ajustements)

    result = {
        "solde_comptable": solde_compta,
        "ajustements": ajustements,
        "total_ajustements": total_ajustements,
        "solde_rapproche": solde_compta + total_ajustements,
    }
    if solde_releves_bancaires is not None:
        result["solde_releves_bancaires"] = solde_releves_bancaires
        result["ecart_non_explique"] = round(
            result["solde_rapproche"] - solde_releves_bancaires, 2)
    return result


def generate(balance_df, budget_df=None, tresorerie_ouverture=0.0,
             solde_releves_bancaires=None, ajustements=None, balance_n1_df=None):
    lines, non_classes = classify_balance(balance_df)
    resultat = resultat_net(lines)
    bilan = build_bilan(lines, resultat)
    compte_resultat = build_compte_resultat(lines)
    ressources_emplois = build_ressources_emplois(lines, tresorerie_ouverture)
    execution_budgetaire = build_execution_budgetaire(lines, budget_df)
    reconciliation = build_reconciliation_tresorerie(lines, solde_releves_bancaires, ajustements)
    lines_n1 = None
    if balance_n1_df is not None:
        lines_n1, _ = classify_balance(balance_n1_df)
    notes = build_notes_annexes(lines, non_classes, resultat)
    notes.append("5. Mode 'Projet de développement' : les états spécifiques (Ressources-"
                 "Emplois, exécution budgétaire, réconciliation de trésorerie) complètent "
                 "le Bilan et le Compte de résultat conformément aux dispositions SYCEBNL "
                 "applicables aux entités gérant des projets financés par des bailleurs.")

    return {
        "mode": "projet",
        "balance_source": balance_df.copy(),
        "balance_n1_source": balance_n1_df.copy() if balance_n1_df is not None else None,
        "lines": lines,
        "lines_n1": lines_n1,
        "non_classes": non_classes,
        "bilan": bilan,
        "compte_resultat": compte_resultat,
        "ressources_emplois": ressources_emplois,
        "execution_budgetaire": execution_budgetaire,
        "reconciliation_tresorerie": reconciliation,
        "notes": notes,
        "resultat_net": resultat,
    }
