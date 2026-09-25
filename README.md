# Application Liasse SYCEBNL — templates officiels intégrés

Version **2.1.0**.

Cette version corrige la génération Windows et intègre les deux classeurs officiels fournis :
- `templates/EtaFi_SYCEBNL_AOP.xlsx` pour **Association / ONG classique** ;
- `templates/EtaFi_SYCEBNL_PROJET.xlsx` pour **Projet de développement financé par bailleurs**.

## Corrections principales

1. Les deux templates Excel sont embarqués dans l'application compilée avec PyInstaller.
2. Le générateur sélectionne automatiquement le template correspondant au type d'entité.
3. Les anciens noms de feuilles (`Feuil12`, `Feuil14`, etc.) ne sont plus utilisés : le remplissage cible les feuilles réelles des templates (`BILAN`, `COMPTE-RESULTAT`, `TFT`, `CPTE EXPLOITATION`, `TER`, `TEB`, `TRT`, etc.).
4. La balance N et, pour le mode Association, la balance N-1 sont injectées dans les feuilles officielles.
5. Les comptes non classés sont reportés dans une feuille `Comptes non classes`.
6. Le classeur source est copié avant remplissage afin de préserver la structure, les formats, les validations et les liens externes du template.

## Lancement

```text
python src/main.py
```

## Compilation Windows

```text
pyinstaller sycebnl_liasse.spec --noconfirm --clean
```

Le fichier `sycebnl_liasse.spec` embarque désormais les quatre fichiers Excel du dossier `templates/`.

Le résultat doit être vérifié par un professionnel comptable avant tout dépôt réglementaire.
