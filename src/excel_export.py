# -*- coding: utf-8 -*-
"""Export SYCEBNL dans le template Excel officiel fourni par l'utilisateur."""
from pathlib import Path
from shutil import copyfile
import sys
import os
from openpyxl import load_workbook


def _resource_path(name):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return Path(base) / name


def _set_if_exists(ws, cell, value):
    if cell in ws:
        ws[cell] = value


def _fill_headers(wb, entite, exercice):
    """Remplit les zones d'identification communes sans modifier la mise en page."""
    import datetime as dt
    try:
        end_date = dt.datetime.strptime(exercice, '%Y-%m-%d') if exercice else None
    except ValueError:
        end_date = None
    for ws in wb.worksheets:
        # Structures de liasse DGI/SYCEBNL fournies dans le template.
        if ws.title == 'IDENTIFICATION':
            _set_if_exists(ws, 'B2', entite)
            if end_date: _set_if_exists(ws, 'H7', end_date.year)
            if end_date: _set_if_exists(ws, 'B10', end_date)
            if end_date: _set_if_exists(ws, 'F11', end_date)
        else:
            # En-têtes récurrents du modèle : nom de l'entité et date de clôture.
            for cell in ('C2','F5','B2'):
                if ws[cell].value is not None and isinstance(ws[cell].value, str) and ('ATL2E' in ws[cell].value or ws[cell].value.strip() == ''):
                    ws[cell] = entite
            # cellules connues de l'imprimé
            if ws.max_row >= 6:
                for cell in ('G6','H6','F6','G13','H13'):
                    if cell in ws and end_date:
                        # uniquement si la cellule contenait déjà une date ou est vide dans un en-tête
                        old = ws[cell].value
                        if old is None or hasattr(old, 'year'):
                            ws[cell] = end_date


def _balance_rows(ws, df):
    """Injecte la balance dans les colonnes normalisées du template."""
    # Conserve les 3 premières lignes du template.
    if ws.max_row >= 4:
        # Nettoyer uniquement la zone de données existante.
        for r in range(4, ws.max_row + 1):
            for c in range(1, 9):
                ws.cell(r, c).value = None
    for i, row in df.iterrows():
        r = 4 + i
        ws.cell(r,1).value = str(row.get('compte',''))
        ws.cell(r,2).value = row.get('intitule','')
        # Le moteur travaille sur débit/crédit; le template de balance attend aussi les mouvements.
        ws.cell(r,3).value = 0
        ws.cell(r,4).value = 0
        ws.cell(r,5).value = float(row.get('debit',0) or 0)
        ws.cell(r,6).value = float(row.get('credit',0) or 0)
        solde = float(row.get('solde',0) or 0)
        ws.cell(r,7).value = max(solde,0)
        ws.cell(r,8).value = max(-solde,0)


def _fill_bilan(ws, lines):
    """Remplit les lignes détaillées du bilan officiel à partir des comptes."""
    if lines is None or lines.empty: return
    # Sommes par préfixe de compte et sens.
    def s(prefixes, positive=True):
        mask = lines['compte'].astype(str).str.startswith(tuple(prefixes))
        x = lines.loc[mask]
        if positive:
            return float(x.loc[x['solde'] >= 0, 'montant'].sum())
        return float(x.loc[x['solde'] < 0, 'montant'].sum())
    # Actif N (col F = net), passif N (col K = net)
    actif = {
        20: s(['20']), 21: s(['21']), 22: s(['22']), 23: s(['23']), 24: s(['24']),
        26: s(['26']), 27: s(['27']), 3: s(['3']), 40: s(['40']), 41: s(['41']),
        42: s(['42']), 43: s(['43']), 44: s(['44']), 45: s(['45']), 46: s(['46']),
        47: s(['47']), 48: s(['48']), 50: s(['50']), 51: s(['51']), 52: s(['52']),
        53: s(['53']), 54: s(['54']), 55: s(['55']), 56: s(['56']), 57: s(['57']), 58: s(['58'])}
    # Lignes officielles du bilan
    rowmap = {20:20,21:25,22:26,23:27,24:28,26:32,27:33,3:38,40:39,41:40,42:41,43:41,44:41,45:41,46:41,47:41,48:41,50:45,51:45,52:45,53:45,54:45,55:45,56:45,57:45,58:45}
    for k,v in actif.items():
        if v:
            r=rowmap[k]; ws.cell(r,6).value = (ws.cell(r,6).value or 0) + v
    # Passif classes 1-5: crédit net
    prow={10:17,105:21,106:21,11:22,12:23,13:24,14:25,15:26,17:28,18:32,19:34,16:32,40:39,41:40,42:40,43:40,44:40,45:40,46:40,47:40,48:40,50:43,51:43,52:43,53:43,54:43,55:43,56:43,57:43,58:43}
    for p,r in prow.items():
        mask=lines['compte'].astype(str).str.startswith(str(p))
        val=float(lines.loc[mask & (lines['solde']<0),'montant'].sum()) if not lines.empty else 0
        if val: ws.cell(r,11).value=(ws.cell(r,11).value or 0)+val


def _fill_resultat(ws, lines):
    if lines is None or lines.empty: return
    maps={70:18,71:20,72:23,74:23,73:22,77:24,75:24,76:24,78:25,79:25,
          60:27,61:33,62:33,63:35,64:34,65:35,66:36,67:37,68:38,
          81:42,83:42,85:42,87:42,82:41,84:41,86:41,88:41}
    for prefix,r in maps.items():
        mask=lines['compte'].astype(str).str.startswith(str(prefix))
        # produits = comptes créditeurs; charges = comptes débiteurs
        if prefix>=70 and prefix not in (81,83,85,87):
            val=float(lines.loc[mask & (lines['solde']<0),'montant'].sum())
        elif prefix in (82,84,86,88):
            val=float(lines.loc[mask & (lines['solde']<0),'montant'].sum())
        else:
            val=float(lines.loc[mask & (lines['solde']>=0),'montant'].sum())
        if val: ws.cell(r,5).value=(ws.cell(r,5).value or 0)+val


def _fill_flux(ws, flux):
    if not flux or not flux.get('disponible'): return
    vals={17:flux.get('tresorerie_ouverture',0),28:flux.get('flux_activites',0),34:flux.get('flux_investissement',0),39:0,
          43:flux.get('flux_financement',0),45:flux.get('variation_tresorerie_calculee',0),46:flux.get('tresorerie_cloture',0),47:flux.get('ecart_controle',0)}
    for r,v in vals.items(): ws.cell(r,5).value=v


def export_liasse(result, output_path, entite_nom='Entité', exercice=''):
    """Produit la liasse en copiant strictement le template utilisateur puis en le remplissant."""
    template = _resource_path('templates/EtaFi_SYCEBNL_AOP.xlsx')
    if not template.exists():
        raise FileNotFoundError('Template SYCEBNL introuvable : ' + str(template))
    copyfile(template, output_path)
    wb=load_workbook(output_path)
    _fill_headers(wb, entite_nom, exercice)
    lines=result.get('lines')
    if 'BALANCE N' in wb.sheetnames:
        _balance_rows(wb['BALANCE N'], lines if lines is not None else [])
    if 'Feuil12' in wb.sheetnames:
        _fill_bilan(wb['Feuil12'], lines)
    if 'Feuil14' in wb.sheetnames:
        _fill_resultat(wb['Feuil14'], lines)
    if 'Feuil16' in wb.sheetnames:
        _fill_flux(wb['Feuil16'], result.get('flux_tresorerie'))
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = 'auto'
    wb.save(output_path)
    return output_path
