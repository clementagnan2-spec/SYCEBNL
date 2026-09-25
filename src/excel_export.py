# -*- coding: utf-8 -*-
"""Export SYCEBNL dans les deux templates Excel officiels fournis."""
from pathlib import Path
from shutil import copyfile
from copy import copy
import sys
import os
import datetime as dt

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


def _resource_path(name):
    base = Path(getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
    return base / name


def _set_if_exists(ws, cell, value):
    if not isinstance(ws[cell], MergedCell):
        ws[cell] = value


def _copy_row_style(ws, src_row, dst_row, max_col=None):
    max_col = max_col or ws.max_column
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    for c in range(1, max_col + 1):
        src = ws.cell(src_row, c)
        dst = ws.cell(dst_row, c)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)


def _write_headers(wb, identification, exercice=""):
    """Alimente la fiche IDENTIFICATION et les en-têtes liés au dossier."""
    ident = wb["IDENTIFICATION"] if "IDENTIFICATION" in wb.sheetnames else None
    if ident is None:
        return

    data = identification or {}
    def val(key, default=""):
        return data.get(key, default)

    # Cellules de la fiche officielle fournie dans les deux templates.
    mapping = {
        "B2": "denomination", "B3": "denomination_suite", "B4": "sigle",
        "B5": "adresse_postale", "B7": "ifu", "B8": "adresse_geographique",
        "G4": "nes", "B10": "exercice_precedent", "B11": "date_debut",
        "F11": "date_fin", "H11": "duree_mois", "B12": "centre_impots",
        "H12": "pays", "B13": "recepisse", "E13": "cnss", "H13": "telephone",
        "B17": "projet_designation", "B18": "projet_suite",
        "B19": "projet_sigle", "B20": "projet_code",
        "B21": "projet_date_debut", "G21": "projet_date_fin",
    }
    for cell, key in mapping.items():
        if key in data and data[key] not in (None, "") and not isinstance(ident[cell], MergedCell):
            ident[cell] = data[key]

    # Année d'établissement des états financiers.
    if val("date_fin"):
        ident["H7"] = val("date_fin").year

    # Les autres feuillets officiels utilisent souvent les cellules ci-dessous
    # pour reprendre automatiquement l'identification. On renseigne seulement
    # les cellules libres, sans écraser les formules du template.
    end_date = val("date_fin")
    common = {
        "denomination": val("denomination"),
        "sigle": val("sigle"),
        "ifu": val("ifu"),
        "nes": val("nes"),
        "adresse": val("adresse_postale"),
    }
    for ws in wb.worksheets:
        if ws.title == "IDENTIFICATION":
            continue
        # Certaines feuilles AOP/projet utilisent des liens/formules vers IDENTIFICATION.
        # Ne jamais remplacer une formule existante.
        for cell, value in (("C2", common["denomination"]), ("C3", val("denomination_suite")),
                            ("C4", common["sigle"]), ("C5", common["adresse"]),
                            ("C6", common["ifu"]), ("C7", common["nes"])):
            if value not in (None, "") and cell in ws and not isinstance(ws[cell], MergedCell):
                if not (isinstance(ws[cell].value, str) and ws[cell].value.startswith("=")):
                    ws[cell] = value
        if end_date:
            for cell in ("G6", "F6", "G13", "H13"):
                if cell in ws and not isinstance(ws[cell], MergedCell):
                    if ws[cell].value is None:
                        ws[cell] = end_date

    for sheet, cells in {
        "BILAN": ("D15", "G15", "I15", "K15"),
        "COMPTE-RESULTAT": ("E16",), "TFT": ("E16",),
        "CPTE EXPLOITATION": ("E16",), "TER": ("E16",),
    }.items():
        if end_date and sheet in wb.sheetnames:
            ws = wb[sheet]
            for cell in cells:
                if not isinstance(ws[cell], MergedCell):
                    ws[cell] = end_date


def _clear_data_rows(ws, start_row, end_row, start_col=1, end_col=None):
    end_col = end_col or ws.max_column
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(r, c).value = None


def _balance_rows(ws, df):
    """Injecte la balance IMPORTÉE dans le format officiel à 8 colonnes.

    Les fichiers d'entrée de l'application sont des balances débit/crédit.
    Elles alimentent donc directement les colonnes de solde de clôture de
    l'imprimé fiscal, sans réinterpréter les montants comme des mouvements.
    Les colonnes d'ouverture et de mouvements restent à zéro faute de détail
    fourni dans la balance source.
    """
    start = 4
    old_end = max(ws.max_row, start)
    # Conserve les lignes d'en-tête et la mise en page de la première ligne de données.
    _clear_data_rows(ws, start, old_end, 1, 8)

    rows = list(df.iterrows()) if df is not None else []
    needed_end = start + len(rows) - 1
    for r in range(start, max(old_end, needed_end) + 1):
        if r > start:
            _copy_row_style(ws, start, r, 8)

    for i, (_, row) in enumerate(rows, start=start):
        debit = float(row.get("debit", 0) or 0)
        credit = float(row.get("credit", 0) or 0)
        ws.cell(i, 1).value = str(row.get("compte", ""))
        ws.cell(i, 2).value = row.get("intitule", "")
        # Balance source = soldes débit/crédit de clôture.
        ws.cell(i, 3).value = 0.0
        ws.cell(i, 4).value = 0.0
        ws.cell(i, 5).value = 0.0
        ws.cell(i, 6).value = 0.0
        ws.cell(i, 7).value = debit
        ws.cell(i, 8).value = credit


def _sum_prefix(lines, prefixes, positive=None):
    if lines is None or lines.empty:
        return 0.0
    mask = lines["compte"].astype(str).str.startswith(tuple(str(p) for p in prefixes))
    x = lines.loc[mask]
    if positive is True:
        return float(x.loc[x["solde"] >= 0, "montant"].sum())
    if positive is False:
        return float(x.loc[x["solde"] < 0, "montant"].sum())
    return float(x["montant"].sum())


def _value_by_prefixes(lines, prefixes):
    """Retourne un solde signé: débit positif, crédit négatif."""
    if lines is None or lines.empty:
        return 0.0
    mask = lines["compte"].astype(str).str.startswith(tuple(str(p) for p in prefixes))
    return float(lines.loc[mask, "solde"].sum())


def _fill_aop_bilan(ws, lines, lines_n1=None):
    # Actif N : brut, amortissements/dépréciations, net.
    n = {
        20: _sum_prefix(lines, ["20"], True),
        21: _sum_prefix(lines, ["21"], True),
        22: _sum_prefix(lines, ["22"], True),
        23: _sum_prefix(lines, ["23"], True),
        24: _sum_prefix(lines, ["24"], True),
        26: _sum_prefix(lines, ["26"], True),
        27: _sum_prefix(lines, ["27"], True),
        3: _sum_prefix(lines, ["3"], True),
        40: _sum_prefix(lines, ["40"], True),
        41: _sum_prefix(lines, ["41"], True),
        42: _sum_prefix(lines, ["42"], True),
        43: _sum_prefix(lines, ["43"], True),
        44: _sum_prefix(lines, ["44"], True),
        45: _sum_prefix(lines, ["45"], True),
        46: _sum_prefix(lines, ["46"], True),
        47: _sum_prefix(lines, ["47"], True),
        48: _sum_prefix(lines, ["48"], True),
        50: _sum_prefix(lines, ["50"], True),
        51: _sum_prefix(lines, ["51"], True),
        52: _sum_prefix(lines, ["52"], True),
        53: _sum_prefix(lines, ["53"], True),
        54: _sum_prefix(lines, ["54"], True),
        55: _sum_prefix(lines, ["55"], True),
        56: _sum_prefix(lines, ["56"], True),
        57: _sum_prefix(lines, ["57"], True),
        58: _sum_prefix(lines, ["58"], True),
    }
    # Dépréciations/amortissements (créditeurs) en colonne E.
    deductions = _sum_prefix(lines, ["28", "29", "49"], False)

    # lignes officielles principales
    ws["D18"] = n[20]
    ws["D24"] = n[21] + n[22] + n[23] + n[24]
    ws["D31"] = n[26] + n[27]
    ws["D38"] = n[3]
    ws["D40"] = n[41]
    ws["D41"] = n[40] + n[42] + n[43] + n[44] + n[45] + n[46] + n[47] + n[48]
    ws["D43"] = n[50]
    ws["D44"] = n[51] + n[52] + n[53] + n[54] + n[55] + n[56] + n[57] + n[58]
    ws["E18"] = _sum_prefix(lines, ["28"], False)
    ws["E24"] = _sum_prefix(lines, ["28"], False)
    ws["E38"] = _sum_prefix(lines, ["49"], False)
    ws["F18"] = ws["D18"].value + ws["E18"].value
    ws["F24"] = ws["D24"].value + ws["E24"].value
    ws["F31"] = ws["D31"].value
    ws["F38"] = ws["D38"].value + ws["E38"].value
    ws["F40"] = ws["D40"].value
    ws["F41"] = ws["D41"].value
    ws["F43"] = ws["D43"].value
    ws["F44"] = ws["D44"].value
    ws["F46"] = ws["F43"].value + ws["F44"].value
    ws["F42"] = ws["F38"].value + ws["F40"].value + ws["F41"].value
    ws["F36"] = ws["F18"].value + ws["F24"].value + ws["F31"].value
    ws["F48"] = ws["F36"].value + ws["F42"].value + ws["F46"].value

    # Passif N.
    passif = {
        10: _sum_prefix(lines, ["10"], False), 105: _sum_prefix(lines, ["105"], False),
        106: _sum_prefix(lines, ["106"], False), 11: _sum_prefix(lines, ["11"], False),
        12: _sum_prefix(lines, ["12"], False), 13: _sum_prefix(lines, ["13"], False),
        14: _sum_prefix(lines, ["14"], False), 15: _sum_prefix(lines, ["15"], False),
        17: _sum_prefix(lines, ["17"], False), 16: _sum_prefix(lines, ["16"], False),
        18: _sum_prefix(lines, ["18"], False), 19: _sum_prefix(lines, ["19"], False),
        40: _sum_prefix(lines, ["40"], False), 41: _sum_prefix(lines, ["41"], False),
        42: _sum_prefix(lines, ["42"], False), 43: _sum_prefix(lines, ["43"], False),
        44: _sum_prefix(lines, ["44"], False), 45: _sum_prefix(lines, ["45"], False),
        46: _sum_prefix(lines, ["46"], False), 47: _sum_prefix(lines, ["47"], False),
        48: _sum_prefix(lines, ["48"], False), 50: _sum_prefix(lines, ["50"], False),
        51: _sum_prefix(lines, ["51"], False), 52: _sum_prefix(lines, ["52"], False),
        53: _sum_prefix(lines, ["53"], False), 54: _sum_prefix(lines, ["54"], False),
        55: _sum_prefix(lines, ["55"], False), 56: _sum_prefix(lines, ["56"], False),
        57: _sum_prefix(lines, ["57"], False), 58: _sum_prefix(lines, ["58"], False),
    }
    for cell, val in {
        "K17": passif[10], "K21": passif[105] + passif[106],
        "K22": passif[11], "K23": passif[12], "K24": passif[13],
        "K25": passif[14], "K26": passif[15], "K28": passif[17],
        "K32": passif[16] + passif[18] + passif[19],
        "K38": passif[40], "K39": passif[41],
        "K40": passif[42] + passif[43] + passif[44] + passif[45] + passif[46] + passif[47] + passif[48],
        "K43": passif[50] + passif[51] + passif[52] + passif[53] + passif[54] + passif[55] + passif[56] + passif[57] + passif[58],
    }.items():
        ws[cell] = val
    ws["K27"] = sum(ws.cell(r, 11).value or 0 for r in (17, 21, 22, 23, 24, 25, 26))
    ws["K30"] = ws["K28"].value or 0
    ws["K31"] = ws["K27"].value + ws["K30"].value
    ws["K35"] = ws["K32"].value
    ws["K36"] = ws["K35"].value + ws["K31"].value
    ws["K42"] = (ws["K38"].value or 0) + (ws["K39"].value or 0) + (ws["K40"].value or 0)
    ws["K46"] = ws["K43"].value or 0
    ws["K48"] = ws["K36"].value + ws["K42"].value + ws["K46"].value

    # N-1 net columns when supplied.
    if lines_n1 is not None:
        # Reuse same main-line aggregation without changing N values.
        n1 = {
            "G18": _sum_prefix(lines_n1, ["20"], True),
            "G24": sum(_sum_prefix(lines_n1, [p], True) for p in ("21","22","23","24")),
            "G31": sum(_sum_prefix(lines_n1, [p], True) for p in ("26","27")),
            "G38": _sum_prefix(lines_n1, ["3"], True),
            "G40": _sum_prefix(lines_n1, ["41"], True),
            "G41": sum(_sum_prefix(lines_n1, [p], True) for p in ("40","42","43","44","45","46","47","48")),
            "G43": sum(_sum_prefix(lines_n1, [p], True) for p in ("50",)),
            "G44": sum(_sum_prefix(lines_n1, [p], True) for p in ("51","52","53","54","55","56","57","58")),
        }
        for cell,val in n1.items(): ws[cell]=val
        ws["G46"] = (ws["G43"].value or 0)+(ws["G44"].value or 0)
        ws["G42"] = (ws["G38"].value or 0)+(ws["G40"].value or 0)+(ws["G41"].value or 0)
        ws["G36"] = (ws["G18"].value or 0)+(ws["G24"].value or 0)+(ws["G31"].value or 0)
        ws["G48"] = (ws["G36"].value or 0)+(ws["G42"].value or 0)+(ws["G46"].value or 0)
        for cell,prefixes in {
            "L17":["10"],"L21":["105","106"],"L22":["11"],"L23":["12"],"L24":["13"],
            "L25":["14"],"L26":["15"],"L28":["17"],"L32":["16","18","19"],
            "L38":["40"],"L39":["41"],"L40":["42","43","44","45","46","47","48"],
            "L43":["50","51","52","53","54","55","56","57","58"],
        }.items():
            ws[cell]=_sum_prefix(lines_n1,prefixes,False)
        ws["L27"]=sum(ws.cell(r,12).value or 0 for r in (17,21,22,23,24,25,26))
        ws["L30"]=ws["L28"].value or 0
        ws["L31"]=ws["L27"].value+ws["L30"].value
        ws["L35"]=ws["L32"].value or 0
        ws["L36"]=ws["L35"].value+ws["L31"].value
        ws["L42"]=(ws["L38"].value or 0)+(ws["L39"].value or 0)+(ws["L40"].value or 0)
        ws["L46"]=ws["L43"].value or 0
        ws["L48"]=ws["L36"].value+ws["L42"].value+ws["L46"].value


def _fill_resultat(ws, lines, project=False):
    if lines is None:
        return
    if project:
        rows = {
            70:19, 71:20, 72:20, 74:20, 73:19, 75:21, 76:21, 78:22, 79:22,
            60:24, 61:28, 62:28, 63:30, 64:29, 65:30, 66:31, 67:32, 68:33,
            81:34,83:34,85:34,87:34,82:34,84:34,86:34,88:34
        }
        ws["E23"] = 0
        ws["E36"] = 0
        for p,r in rows.items():
            pos = p < 70 or p in (60,61,62,63,64,65,66,67,68,81,83,85,87)
            # Products/HAO products are credit balances, therefore negative signed balances.
            if p >= 70 and p not in (81,83,85,87):
                val = _sum_prefix(lines,[str(p)],False)
            elif p in (82,84,86,88):
                val = _sum_prefix(lines,[str(p)],False)
            else:
                val = _sum_prefix(lines,[str(p)],True)
            ws.cell(r,5).value = (ws.cell(r,5).value or 0) + val
        ws["E23"] = sum(ws.cell(r,5).value or 0 for r in range(19,23))
        ws["E36"] = sum(ws.cell(r,5).value or 0 for r in range(24,35))
        ws["E37"] = ws["E23"].value + ws["E36"].value
    else:
        maps = {70:18,71:20,72:23,74:23,73:22,75:24,76:24,78:25,79:25,
                60:27,61:33,62:33,63:35,64:34,65:35,66:36,67:37,68:38,
                81:42,83:42,85:42,87:42,82:41,84:41,86:41,88:41}
        for p,r in maps.items():
            if p >= 70:
                val = _sum_prefix(lines,[str(p)],False)
            else:
                val = _sum_prefix(lines,[str(p)],True)
            ws.cell(r,5).value = (ws.cell(r,5).value or 0) + val
        ws["E26"] = sum(ws.cell(r,5).value or 0 for r in range(18,26))
        ws["E39"] = sum(ws.cell(r,5).value or 0 for r in range(27,39))
        ws["E40"] = ws["E26"].value - ws["E39"].value
        ws["E43"] = (ws["E41"].value or 0) - (ws["E42"].value or 0)
        ws["E44"] = ws["E40"].value + ws["E43"].value


def _fill_tft(ws, flux):
    if not flux or not flux.get("disponible"):
        return
    vals = {
        17: flux.get("tresorerie_ouverture", 0),
        28: flux.get("flux_activites", 0),
        34: flux.get("flux_investissement", 0),
        39: flux.get("flux_financement", 0),
        43: 0,
        45: flux.get("variation_tresorerie_calculee", 0),
        46: flux.get("tresorerie_cloture", 0),
    }
    for r,v in vals.items():
        ws.cell(r,5).value = v


def _fill_project_bilan(ws, lines):
    # Correspondance avec les rubriques du template projet.
    vals = {
        17: _sum_prefix(lines, ["20"], True),
        18: _sum_prefix(lines, ["21","22"], True),
        19: _sum_prefix(lines, ["23"], True),
        20: _sum_prefix(lines, ["24"], True),
        21: _sum_prefix(lines, ["24"], True),
        22: _sum_prefix(lines, ["40"], True),
        27: _sum_prefix(lines, ["3"], True),
        28: _sum_prefix(lines, ["40"], True),
        29: sum(_sum_prefix(lines,[p],True) for p in ("41","42","43","44","45","46","47","48")),
        33: sum(_sum_prefix(lines,[p],True) for p in ("50","51","52","53","54","55","56","57","58")),
        22: _sum_prefix(lines, ["26","27"], True),
    }
    for r,v in vals.items(): ws.cell(r,4).value=v
    ws["D25"] = sum(ws.cell(r,4).value or 0 for r in range(17,25))
    ws["D31"] = sum(ws.cell(r,4).value or 0 for r in range(27,31))
    ws["D34"] = ws["D33"].value or 0
    ws["D36"] = ws["D25"].value + ws["D31"].value + ws["D34"].value
    # Passif
    p = {
        17: _sum_prefix(lines,["10"],False),
        18: _sum_prefix(lines,["12"],False),
        19: _sum_prefix(lines,["13"],False),
        20: _sum_prefix(lines,["14"],False),
        22: sum(_sum_prefix(lines,[x],False) for x in ("16","18","19")),
        23: _sum_prefix(lines,["15"],False),
        27: _sum_prefix(lines,["17"],False),
        28: _sum_prefix(lines,["40"],False),
        29: sum(_sum_prefix(lines,[x],False) for x in ("41","42","43","44","45","46","47","48")),
        33: sum(_sum_prefix(lines,[x],False) for x in ("50","51","52","53","54","55","56","57","58")),
    }
    for r,v in p.items(): ws.cell(r,9).value=v
    ws["I21"] = sum(ws.cell(r,9).value or 0 for r in (17,18,19,20))
    ws["I24"] = (ws["I22"].value or 0)+(ws["I23"].value or 0)
    ws["I25"] = ws["I21"].value + ws["I24"].value
    ws["I31"] = (ws["I28"].value or 0)+(ws["I29"].value or 0)
    ws["I34"] = ws["I33"].value or 0
    ws["I36"] = ws["I25"].value + ws["I31"].value + ws["I34"].value


def _fill_project_ter(ws, data):
    if not data:
        return
    # Exercice column E; cumulative beginning/end are optional approximations.
    ws["E21"] = data.get("total_ressources", 0)
    ws["E40"] = data.get("total_emplois", 0)
    ws["E41"] = data.get("solde_tresorerie_fin", 0)
    ws["E45"] = data.get("tresorerie_ouverture", 0)
    ws["E46"] = data.get("solde_tresorerie_fin", 0)
    ws["E50"] = data.get("solde_tresorerie_fin", 0)
    ws["E51"] = (ws["E46"].value or 0) - (ws["E50"].value or 0)


def _fill_project_teb(ws, data):
    if not data or not data.get("disponible"):
        return
    start = 19
    for idx, row in enumerate(data.get("lignes", []), start=start):
        ws.cell(idx, 1).value = idx - start + 1
        ws.cell(idx, 2).value = row["rubrique"]
        ws.cell(idx, 3).value = row["budget"]
        ws.cell(idx, 6).value = row["realise"]
        ws.cell(idx, 7).value = row["ecart"]
        ws.cell(idx, 8).value = row["taux_execution"] if row["taux_execution"] is not None else 0


def _fill_project_trt(ws, data):
    if not data:
        return
    # Montants de rapprochement disponibles à partir des données saisies.
    ws["G18"] = data.get("solde_comptable", 0)
    ws["G41"] = data.get("solde_rapproche", 0)
    if "solde_releves_bancaires" in data:
        ws["G46"] = data.get("solde_releves_bancaires", 0)


def export_liasse(result, output_path, entite_nom="Entité", exercice="", identification=None):
    """Copie le template officiel correspondant au mode puis le remplit."""
    mode = result.get("mode", "association")
    template_name = "EtaFi_SYCEBNL_PROJET.xlsx" if mode == "projet" else "EtaFi_SYCEBNL_AOP.xlsx"
    template = _resource_path("templates") / template_name
    if not template.exists():
        raise FileNotFoundError(
            "Template SYCEBNL introuvable : " + str(template) +
            f"\nMode demandé : {mode}\nVérifiez l'installation de l'application."
        )

    copyfile(template, output_path)
    wb = load_workbook(output_path, keep_links=True)

    _write_headers(wb, identification or {"denomination": entite_nom}, exercice)
    lines = result.get("lines")
    lines_n1 = result.get("lines_n1")
    balance_source = result.get("balance_source")
    balance_n1_source = result.get("balance_n1_source")

    if mode == "projet":
        if "ANNEE N" in wb.sheetnames:
            _balance_rows(wb["ANNEE N"], balance_source if balance_source is not None else [])
        if "ANNEE N-1" in wb.sheetnames and lines_n1 is not None:
            _balance_rows(wb["ANNEE N-1"], balance_n1_source)
        if "BILAN" in wb.sheetnames:
            _fill_project_bilan(wb["BILAN"], lines)
        if "CPTE EXPLOITATION" in wb.sheetnames:
            _fill_resultat(wb["CPTE EXPLOITATION"], lines, project=True)
        if "TER" in wb.sheetnames:
            _fill_project_ter(wb["TER"], result.get("ressources_emplois"))
        if "TEB" in wb.sheetnames:
            _fill_project_teb(wb["TEB"], result.get("execution_budgetaire"))
        if "TRT" in wb.sheetnames:
            _fill_project_trt(wb["TRT"], result.get("reconciliation_tresorerie"))
    else:
        if "BALANCE N" in wb.sheetnames:
            _balance_rows(wb["BALANCE N"], balance_source if balance_source is not None else [])
        if "FeuiBALANCE N-1" in wb.sheetnames and lines_n1 is not None:
            _balance_rows(wb["FeuiBALANCE N-1"], balance_n1_source)
        if "BILAN" in wb.sheetnames:
            _fill_aop_bilan(wb["BILAN"], lines, lines_n1)
        if "COMPTE-RESULTAT" in wb.sheetnames:
            _fill_resultat(wb["COMPTE-RESULTAT"], lines, project=False)
        if "TFT" in wb.sheetnames:
            _fill_tft(wb["TFT"], result.get("flux_tresorerie"))

    # Ajoute la liste des comptes non classés sans toucher aux feuilles officielles.
    non_classes = result.get("non_classes")
    if non_classes is not None and not getattr(non_classes, "empty", True):
        name = "Comptes non classes"
        if name in wb.sheetnames:
            del wb[name]
        ws = wb.create_sheet(name)
        cols = list(non_classes.columns)
        for c,h in enumerate(cols,1):
            ws.cell(1,c).value=h
        for r,(_,row) in enumerate(non_classes.iterrows(),2):
            for c,h in enumerate(cols,1):
                ws.cell(r,c).value=row.get(h)

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(output_path)
    return output_path
