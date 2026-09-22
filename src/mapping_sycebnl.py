# -*- coding: utf-8 -*-
"""
mapping_sycebnl.py
===================
Référentiel de correspondance "numéro de compte -> rubrique de la liasse"
pour le Système Comptable des Entités à But Non Lucratif (SYCEBNL - OHADA).

IMPORTANT (à lire avant tout dépôt officiel) :
Ce mapping reprend la structure générale des 9 classes du plan comptable
OHADA telle qu'adaptée par l'Acte uniforme SYCEBNL (entré en vigueur le
1er janvier 2024). Il constitue une base de travail raisonnable pour un
premier classement automatique, mais :
  - les intitulés exacts de rubriques de l'imprimé officiel de la liasse
    fiscale peuvent varier selon les administrations fiscales nationales ;
  - un expert-comptable ou un commissaire aux comptes doit valider le
    classement avant tout dépôt réglementaire.
Le fichier RUBRIQUES ci-dessous est volontairement séparé du code pour
que l'utilisateur puisse l'ajuster sans toucher au reste du programme.
"""

from dataclasses import dataclass
from typing import List


@dataclass
class Rubrique:
    code: str            # code interne court (ex: "AI_INC")
    libelle: str          # libellé affiché dans la liasse
    section: str          # "ACTIF", "PASSIF", "CHARGES", "PRODUITS", "HAO_CHARGES", "HAO_PRODUITS"
    prefixes: List[str]   # préfixes de numéros de compte (ex: ["20","21"])
    sens_normal: str      # "D" (débit) ou "C" (crédit) - sens normal du solde


# ---------------------------------------------------------------------------
# BILAN - ACTIF
# ---------------------------------------------------------------------------
RUBRIQUES_ACTIF: List[Rubrique] = [
    Rubrique("AI_INC", "Immobilisations incorporelles", "ACTIF", ["20"], "D"),
    Rubrique("AI_COR", "Immobilisations corporelles (terrains, bâtiments, matériel)", "ACTIF",
             ["21", "22", "23", "24"], "D"),
    Rubrique("AI_FIN", "Immobilisations financières", "ACTIF", ["26", "27"], "D"),
    Rubrique("AI_AMO", "(-) Amortissements et dépréciations des immobilisations", "ACTIF",
             ["28", "29"], "C"),  # vient en déduction (compte de sens créditeur)
    Rubrique("AC_STK", "Stocks et en-cours", "ACTIF", ["3"], "D"),
    Rubrique("AC_CRE", "Créances et emplois assimilés (bénéficiaires, fondateurs, État, débiteurs divers)",
             "ACTIF", ["40", "41", "42", "43", "44", "45", "46", "47", "48"], "D"),
    Rubrique("AC_DEP", "(-) Dépréciations des comptes de tiers", "ACTIF", ["49"], "C"),
    Rubrique("TA_TRE", "Trésorerie et équivalents de trésorerie (banques, caisse, placements)",
             "ACTIF", ["50", "51", "52", "53", "54", "55", "56", "57", "58"], "D"),
]

# ---------------------------------------------------------------------------
# BILAN - PASSIF
# ---------------------------------------------------------------------------
RUBRIQUES_PASSIF: List[Rubrique] = [
    Rubrique("RD_FPR", "Fonds propres / Capital / Dotations", "PASSIF", ["10"], "C"),
    Rubrique("RD_ECA", "Écarts de réévaluation", "PASSIF", ["105", "106"], "C"),
    Rubrique("RD_RES", "Réserves", "PASSIF", ["11"], "C"),
    Rubrique("RD_RAN", "Report à nouveau", "PASSIF", ["12"], "C"),
    Rubrique("RD_RES_EX", "Résultat net de l'exercice", "PASSIF", ["13"], "C"),
    Rubrique("RD_SUB", "Subventions d'investissement", "PASSIF", ["14"], "C"),
    Rubrique("RD_PRO", "Provisions réglementées", "PASSIF", ["15"], "C"),
    Rubrique("RD_FDE", "Fonds dédiés / Fonds affectés (dotations consomptibles et non consomptibles)",
             "PASSIF", ["17"], "C"),
    Rubrique("RD_DFI", "Dettes financières et ressources assimilées (emprunts)", "PASSIF",
             ["16", "18", "19"], "C"),
    Rubrique("DT_CIR", "Dettes circulantes (fournisseurs, fiscales, sociales, autres créditeurs)",
             "PASSIF", ["40", "41", "42", "43", "44", "45", "46", "47", "48"], "C"),
    Rubrique("TP_TRE", "Trésorerie passif (banques créditrices, concours bancaires)", "PASSIF",
             ["50", "51", "52", "53", "54", "55", "56", "57", "58"], "C"),
]

# ---------------------------------------------------------------------------
# COMPTE DE RÉSULTAT - CHARGES (classe 6) et PRODUITS (classe 7)
# ---------------------------------------------------------------------------
RUBRIQUES_CHARGES: List[Rubrique] = [
    Rubrique("CH_ACH", "Achats et variations de stocks", "CHARGES", ["60"], "D"),
    Rubrique("CH_SER", "Services extérieurs et autres charges externes", "CHARGES",
             ["61", "62"], "D"),
    Rubrique("CH_IMP", "Impôts et taxes", "CHARGES", ["64"], "D"),
    Rubrique("CH_PER", "Charges de personnel", "CHARGES", ["66"], "D"),
    Rubrique("CH_FIN", "Charges financières", "CHARGES", ["67"], "D"),
    Rubrique("CH_DOT", "Dotations aux amortissements et provisions", "CHARGES", ["68"], "D"),
    Rubrique("CH_AUT", "Autres charges des activités", "CHARGES", ["63", "65"], "D"),
]

RUBRIQUES_PRODUITS: List[Rubrique] = [
    Rubrique("PR_COT", "Cotisations des membres", "PRODUITS", ["70"], "C"),
    Rubrique("PR_DON", "Dons, legs et générosité du public", "PRODUITS", ["71"], "C"),
    Rubrique("PR_SUB", "Subventions et financements de bailleurs", "PRODUITS", ["72", "74"], "C"),
    Rubrique("PR_PRE", "Ventes de prestations / production de services", "PRODUITS", ["73"], "C"),
    Rubrique("PR_FIN", "Produits financiers", "PRODUITS", ["77"], "C"),
    Rubrique("PR_REP", "Reprises sur amortissements et provisions", "PRODUITS", ["78", "79"], "C"),
    Rubrique("PR_AUT", "Autres produits des activités", "PRODUITS", ["75", "76"], "C"),
]

RUBRIQUES_HAO: List[Rubrique] = [
    Rubrique("HAO_CH", "Autres charges (Hors Activités Ordinaires)", "HAO_CHARGES", ["81", "83", "85", "87"], "D"),
    Rubrique("HAO_PR", "Autres produits (Hors Activités Ordinaires)", "HAO_PRODUITS", ["82", "84", "86", "88"], "C"),
]

ALL_RUBRIQUES = (RUBRIQUES_ACTIF + RUBRIQUES_PASSIF + RUBRIQUES_CHARGES +
                  RUBRIQUES_PRODUITS + RUBRIQUES_HAO)


def match_prefix(compte: str, prefixes: List[str]) -> bool:
    """Renvoie True si le numéro de compte commence par l'un des préfixes donnés."""
    compte = str(compte).strip()
    return any(compte.startswith(p) for p in prefixes)


def classe(compte: str) -> str:
    """Renvoie la classe (1er chiffre) d'un numéro de compte."""
    compte = str(compte).strip()
    return compte[0] if compte else "?"
