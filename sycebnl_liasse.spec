# -*- mode: python ; coding: utf-8 -*-
# Fichier de configuration PyInstaller.
# Compilation : pyinstaller sycebnl_liasse.spec
#
# Produit un exécutable Windows unique, sans fenêtre de console, nommé
# "GenerateurLiasseSYCEBNL.exe", dans le dossier dist/.

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=['src'],
    binaries=[],
    datas=[
        ('templates/EtaFi_SYCEBNL_AOP.xlsx', 'templates'),
        ('templates/EtaFi_SYCEBNL_PROJET.xlsx', 'templates'),
        ('templates/modele_balance_sycebnl.xlsx', 'templates'),
        ('templates/modele_budget_sycebnl.xlsx', 'templates'),
    ],
    hiddenimports=[
        'pandas', 'openpyxl', 'openpyxl.cell._writer',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'PyQt5', 'PySide2', 'PySide6', 'notebook'],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GenerateurLiasseSYCEBNL',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
