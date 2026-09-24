# Application Liasse SYCEBNL — Template intégré

Version 2.0 : l'application utilise **EtaFi_SYCEBNL_AOP.xlsx comme template de sortie**. Elle ne reconstruit plus une nouvelle liasse avec des feuilles inventées.

Flux :
1. Importer la balance N.
2. Optionnel : importer N-1 et le budget.
3. Le moteur classe les comptes SYCEBNL.
4. L'application copie le template utilisateur et remplit les zones prévues : identification, balance N, bilan, compte de résultat et flux lorsque disponibles.
5. Le classeur final conserve la structure, les feuilles, les formats et les formules du template.

Lancement : `python src/main.py`

Compilation Windows : `pyinstaller sycebnl_liasse.spec --noconfirm --clean`
