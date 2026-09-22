# -*- coding: utf-8 -*-
"""
excel_export.py
================
Écrit la liasse SYCEBNL calculée dans un classeur Excel (.xlsx) mis en forme,
avec une feuille par état financier et des formules de total pour permettre
la vérification directe dans Excel.
"""

from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "Arial"
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="1F3864")
HEADER_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
SUBHEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TOTAL_FONT = Font(name=FONT_NAME, size=11, bold=True)
NORMAL_FONT = Font(name=FONT_NAME, size=10)
WARNING_FONT = Font(name=FONT_NAME, size=10, bold=True, color="C00000")
NUM_FMT = "#,##0;(#,##0);\"-\""
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _sheet_header(ws, title, entite, exercice, subtitle=None):
    ws["A1"] = entite
    ws["A1"].font = Font(name=FONT_NAME, size=12, bold=True)
    ws["A2"] = title
    ws["A2"].font = TITLE_FONT
    ws["A3"] = f"Exercice clos le {exercice}" if exercice else ""
    ws["A3"].font = Font(name=FONT_NAME, size=10, italic=True)
    if subtitle:
        ws["A4"] = subtitle
        ws["A4"].font = Font(name=FONT_NAME, size=9, italic=True, color="808080")
    ws["A5"] = "Montants exprimés en Franc CFA (XOF), sauf indication contraire."
    ws["A5"].font = Font(name=FONT_NAME, size=9, italic=True, color="808080")
    return 7  # première ligne libre


def _write_table(ws, start_row, headers, rows, total_label=None, total_values=None,
                  col_widths=None):
    """Écrit un petit tableau (en-têtes + lignes + ligne de total optionnelle).
    Renvoie la ligne suivante disponible."""
    r = start_row
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center" if c > 1 else "left", vertical="center")
    r += 1
    first_data_row = r
    for row_vals in rows:
        for c, v in enumerate(row_vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.font = NORMAL_FONT
            if c > 1 and isinstance(v, (int, float)):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right")
        r += 1
    last_data_row = r - 1
    if total_label is not None:
        ws.cell(row=r, column=1, value=total_label).font = TOTAL_FONT
        if total_values:
            for c, _ in enumerate(total_values, start=2):
                col_letter = get_column_letter(c)
                if last_data_row >= first_data_row:
                    formula = f"=SUM({col_letter}{first_data_row}:{col_letter}{last_data_row})"
                else:
                    formula = 0
                cell = ws.cell(row=r, column=c, value=formula)
                cell.font = TOTAL_FONT
                cell.number_format = NUM_FMT
                cell.border = BORDER
        r += 1
    if col_widths:
        for c, w in enumerate(col_widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
    return r + 1


def _write_bilan_sheet(wb, result, entite, exercice):
    ws = wb.create_sheet("Bilan")
    r = _sheet_header(ws, "BILAN", entite, exercice)
    bilan = result["bilan"]

    deductions = ["(-) Amortissements et dépréciations des immobilisations",
                  "(-) Dépréciations des comptes de tiers"]
    ws.cell(row=r, column=1, value="ACTIF").font = Font(name=FONT_NAME, bold=True, size=12)
    r += 1
    actif_rows = [[lib, -montant if lib in deductions else montant]
                  for lib, montant in bilan["actif"].items()]
    r = _write_table(ws, r, ["Rubrique", "Montant net"], actif_rows,
                      total_label="TOTAL ACTIF", total_values=[None], col_widths=[55, 20])

    r += 1
    ws.cell(row=r, column=1, value="PASSIF").font = Font(name=FONT_NAME, bold=True, size=12)
    r += 1
    passif_rows = [[lib, montant] for lib, montant in bilan["passif"].items()]
    r = _write_table(ws, r, ["Rubrique", "Montant"], passif_rows,
                      total_label="TOTAL PASSIF", total_values=[None], col_widths=[55, 20])

    r += 1
    ok = bilan["equilibre"]
    msg = "✓ Le bilan est équilibré." if ok else f"⚠ ÉCART Actif - Passif = {bilan['ecart']:,.0f} — vérifier le classement des comptes."
    cell = ws.cell(row=r, column=1, value=msg)
    cell.font = NORMAL_FONT if ok else WARNING_FONT
    return ws


def _write_compte_resultat_sheet(wb, result, entite, exercice):
    ws = wb.create_sheet("Compte de résultat")
    r = _sheet_header(ws, "COMPTE DE RÉSULTAT", entite, exercice)
    cr = result["compte_resultat"]

    ws.cell(row=r, column=1, value="CHARGES DES ACTIVITÉS ORDINAIRES").font = Font(name=FONT_NAME, bold=True, size=12)
    r += 1
    rows = [[lib, m] for lib, m in cr["charges"].items()]
    r = _write_table(ws, r, ["Rubrique", "Montant"], rows,
                      total_label="Total charges des activités ordinaires", total_values=[None],
                      col_widths=[55, 20])

    r += 1
    ws.cell(row=r, column=1, value="PRODUITS DES ACTIVITÉS ORDINAIRES").font = Font(name=FONT_NAME, bold=True, size=12)
    r += 1
    rows = [[lib, m] for lib, m in cr["produits"].items()]
    r = _write_table(ws, r, ["Rubrique", "Montant"], rows,
                      total_label="Total produits des activités ordinaires", total_values=[None],
                      col_widths=[55, 20])

    r += 1
    ws.cell(row=r, column=1, value="Résultat des activités ordinaires").font = TOTAL_FONT
    ws.cell(row=r, column=2, value=round(cr["resultat_ao"], 2)).font = TOTAL_FONT
    ws.cell(row=r, column=2).number_format = NUM_FMT
    r += 2

    ws.cell(row=r, column=1, value="Charges HAO").font = NORMAL_FONT
    ws.cell(row=r, column=2, value=round(cr["total_charges_hao"], 2)).number_format = NUM_FMT
    r += 1
    ws.cell(row=r, column=1, value="Produits HAO").font = NORMAL_FONT
    ws.cell(row=r, column=2, value=round(cr["total_produits_hao"], 2)).number_format = NUM_FMT
    r += 1
    ws.cell(row=r, column=1, value="Résultat HAO").font = NORMAL_FONT
    ws.cell(row=r, column=2, value=round(cr["resultat_hao"], 2)).number_format = NUM_FMT
    r += 2

    ws.cell(row=r, column=1, value="RÉSULTAT NET DE L'EXERCICE").font = Font(name=FONT_NAME, bold=True, size=12)
    ws.cell(row=r, column=2, value=round(cr["resultat_net"], 2)).font = Font(name=FONT_NAME, bold=True, size=12)
    ws.cell(row=r, column=2).number_format = NUM_FMT
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 20
    return ws


def _write_flux_sheet(wb, result, entite, exercice):
    ws = wb.create_sheet("Flux de trésorerie")
    r = _sheet_header(ws, "TABLEAU DES FLUX DE TRÉSORERIE", entite, exercice,
                       subtitle="Méthode indirecte simplifiée")
    flux = result["flux_tresorerie"]
    if not flux.get("disponible"):
        ws.cell(row=r, column=1,
                value="Non calculable : la balance de l'exercice précédent (N-1) n'a pas été fournie.").font = WARNING_FONT
        ws.column_dimensions["A"].width = 70
        return ws

    lignes = [
        ("Trésorerie nette à l'ouverture", flux["tresorerie_ouverture"], False),
        ("Capacité d'autofinancement globale (CAFG)", flux["cafg"], False),
        ("Variation du besoin en fonds de roulement", -flux["variation_bfr"], False),
        ("Flux de trésorerie des activités opérationnelles", flux["flux_activites"], True),
        ("Flux de trésorerie des activités d'investissement", flux["flux_investissement"], False),
        ("Flux de trésorerie des activités de financement", flux["flux_financement"], False),
        ("Variation de trésorerie de l'exercice", flux["variation_tresorerie_calculee"], True),
        ("Trésorerie nette à la clôture (calculée)", flux["tresorerie_ouverture"] + flux["variation_tresorerie_calculee"], True),
        ("Trésorerie nette à la clôture (balance)", flux["tresorerie_cloture"], True),
        ("Écart de contrôle", flux["ecart_controle"], False),
    ]
    for lib, montant, bold in lignes:
        ws.cell(row=r, column=1, value=lib).font = TOTAL_FONT if bold else NORMAL_FONT
        c = ws.cell(row=r, column=2, value=round(montant, 2))
        c.number_format = NUM_FMT
        c.font = TOTAL_FONT if bold else NORMAL_FONT
        r += 1
    if abs(flux["ecart_controle"]) > 1:
        r += 1
        ws.cell(row=r, column=1,
                value="⚠ Écart de contrôle non nul : vérifier le classement des comptes "
                      "entre les deux exercices.").font = WARNING_FONT
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 20
    return ws


def _write_ressources_emplois_sheet(wb, result, entite, exercice):
    ws = wb.create_sheet("Ressources-Emplois")
    r = _sheet_header(ws, "TABLEAU RESSOURCES - EMPLOIS", entite, exercice,
                       subtitle="Entités gérant des projets de développement")
    re_ = result["ressources_emplois"]

    ws.cell(row=r, column=1, value="RESSOURCES").font = Font(name=FONT_NAME, bold=True, size=12)
    r += 1
    rows = [["Trésorerie disponible à l'ouverture", re_["tresorerie_ouverture"]]] + \
           [[lib, m] for lib, m in re_["ressources"].items()]
    r = _write_table(ws, r, ["Rubrique", "Montant"], rows,
                      total_label="TOTAL RESSOURCES", total_values=[None], col_widths=[55, 20])
    ws.cell(row=r - 1, column=2, value=round(re_["total_ressources"], 2))

    r += 1
    ws.cell(row=r, column=1, value="EMPLOIS").font = Font(name=FONT_NAME, bold=True, size=12)
    r += 1
    rows = [[lib, m] for lib, m in re_["emplois"].items()]
    r = _write_table(ws, r, ["Rubrique", "Montant"], rows,
                      total_label="TOTAL EMPLOIS", total_values=[None], col_widths=[55, 20])

    r += 1
    ws.cell(row=r, column=1, value="SOLDE DE TRÉSORERIE EN FIN DE PÉRIODE").font = Font(name=FONT_NAME, bold=True, size=12)
    ws.cell(row=r, column=2, value=round(re_["solde_tresorerie_fin"], 2)).font = Font(name=FONT_NAME, bold=True, size=12)
    ws.cell(row=r, column=2).number_format = NUM_FMT
    return ws


def _write_execution_budgetaire_sheet(wb, result, entite, exercice):
    ws = wb.create_sheet("Exécution budgétaire")
    r = _sheet_header(ws, "TABLEAU D'EXÉCUTION BUDGÉTAIRE", entite, exercice)
    eb = result["execution_budgetaire"]
    if not eb.get("disponible"):
        ws.cell(row=r, column=1,
                value="Non calculable : aucun fichier budget n'a été importé.").font = WARNING_FONT
        ws.column_dimensions["A"].width = 70
        return ws

    rows = []
    for l in eb["lignes"]:
        taux = f"{l['taux_execution']:.1f} %" if l["taux_execution"] is not None else "n/a"
        rows.append([l["rubrique"], round(l["budget"], 2), round(l["realise"], 2),
                     round(l["ecart"], 2), taux])
    r = _write_table(ws, r, ["Rubrique budgétaire", "Budget", "Réalisé", "Écart", "Taux d'exécution"],
                      rows, col_widths=[45, 18, 18, 18, 16])
    r += 1
    ws.cell(row=r, column=1, value="TOTAL").font = TOTAL_FONT
    ws.cell(row=r, column=2, value=round(eb["total_budget"], 2)).number_format = NUM_FMT
    ws.cell(row=r, column=3, value=round(eb["total_realise"], 2)).number_format = NUM_FMT
    ws.cell(row=r, column=4, value=round(eb["ecart_total"], 2)).number_format = NUM_FMT
    taux_g = f"{eb['taux_execution_global']:.1f} %" if eb.get("taux_execution_global") is not None else "n/a"
    ws.cell(row=r, column=5, value=taux_g)
    for c in range(1, 6):
        ws.cell(row=r, column=c).font = TOTAL_FONT
    return ws


def _write_reconciliation_sheet(wb, result, entite, exercice):
    ws = wb.create_sheet("Réconciliation trésorerie")
    r = _sheet_header(ws, "TABLEAU DE RÉCONCILIATION DE TRÉSORERIE", entite, exercice)
    rec = result["reconciliation_tresorerie"]

    ws.cell(row=r, column=1, value="Solde de trésorerie selon la comptabilité (classe 5)").font = NORMAL_FONT
    ws.cell(row=r, column=2, value=round(rec["solde_comptable"], 2)).number_format = NUM_FMT
    r += 2

    if rec["ajustements"]:
        ws.cell(row=r, column=1, value="Ajustements de rapprochement").font = Font(name=FONT_NAME, bold=True)
        r += 1
        rows = [[a["libelle"], a["montant"]] for a in rec["ajustements"]]
        r = _write_table(ws, r, ["Libellé", "Montant"], rows,
                          total_label="Total des ajustements", total_values=[None], col_widths=[45, 18])

    ws.cell(row=r, column=1, value="Solde de trésorerie rapproché").font = TOTAL_FONT
    ws.cell(row=r, column=2, value=round(rec["solde_rapproche"], 2)).font = TOTAL_FONT
    ws.cell(row=r, column=2).number_format = NUM_FMT
    r += 2

    if "solde_releves_bancaires" in rec:
        ws.cell(row=r, column=1, value="Solde selon relevés bancaires / mobile money").font = NORMAL_FONT
        ws.cell(row=r, column=2, value=round(rec["solde_releves_bancaires"], 2)).number_format = NUM_FMT
        r += 1
        ok = abs(rec["ecart_non_explique"]) < 1
        msg = "✓ Rapprochement conforme." if ok else f"⚠ Écart non expliqué : {rec['ecart_non_explique']:,.0f}"
        ws.cell(row=r, column=1, value=msg).font = NORMAL_FONT if ok else WARNING_FONT
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 20
    return ws


def _write_notes_sheet(wb, result, entite, exercice):
    ws = wb.create_sheet("Notes annexes")
    r = _sheet_header(ws, "NOTES ANNEXES", entite, exercice)
    for note in result["notes"]:
        ws.cell(row=r, column=1, value=note).font = NORMAL_FONT
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
        r += 1
    ws.column_dimensions["A"].width = 100
    return ws


def _write_non_classes_sheet(wb, result):
    non_classes = result["non_classes"]
    if non_classes is None or non_classes.empty:
        return
    ws = wb.create_sheet("Comptes non classés")
    ws.cell(row=1, column=1, value="Comptes n'ayant pas pu être classés automatiquement").font = TITLE_FONT
    headers = ["compte", "intitule", "debit", "credit", "solde", "raison"]
    r = 3
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    r += 1
    for _, row in non_classes.iterrows():
        for c, h in enumerate(headers, start=1):
            ws.cell(row=r, column=c, value=row.get(h, ""))
        r += 1
    for c, w in enumerate([14, 40, 14, 14, 14, 45], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w


def export_liasse(result, output_path, entite_nom="Entité", exercice=""):
    """Écrit le classeur Excel complet et le sauvegarde à output_path."""
    wb = Workbook()
    wb.remove(wb.active)

    _write_bilan_sheet(wb, result, entite_nom, exercice)
    _write_compte_resultat_sheet(wb, result, entite_nom, exercice)

    if result["mode"] == "association":
        _write_flux_sheet(wb, result, entite_nom, exercice)
    else:
        _write_ressources_emplois_sheet(wb, result, entite_nom, exercice)
        _write_execution_budgetaire_sheet(wb, result, entite_nom, exercice)
        _write_reconciliation_sheet(wb, result, entite_nom, exercice)

    _write_notes_sheet(wb, result, entite_nom, exercice)
    _write_non_classes_sheet(wb, result)

    wb.save(output_path)
    return output_path
