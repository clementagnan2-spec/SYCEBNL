# Générateur de liasse SYCEBNL

Application de bureau (Windows) qui aide à préparer les états financiers du
**Système Comptable des Entités à But Non Lucratif (SYCEBNL)**, référentiel
OHADA en vigueur depuis le 1er janvier 2024 pour les associations, ONG,
fondations et projets de développement.

À partir d'une **balance des comptes** (Excel ou CSV), le logiciel génère
automatiquement un classeur Excel contenant :

- **Mode Association / ONG classique** : Bilan, Compte de résultat, Tableau
  des flux de trésorerie (si la balance N-1 est fournie), Notes annexes.
- **Mode Projet de développement** (financé par des bailleurs) : Bilan,
  Compte de résultat, Tableau Ressources-Emplois, Tableau d'exécution
  budgétaire (si un fichier budget est fourni), Tableau de réconciliation
  de trésorerie.

> ⚠️ **Avertissement important** : ce logiciel automatise un premier
> classement comptable pour gagner du temps. Ce n'est **pas** un service
> officiel de télédéclaration et le mapping compte → rubrique
> (`src/mapping_sycebnl.py`) est une base de travail générale. **Faites
> toujours valider les états produits par un expert-comptable ou un
> commissaire aux comptes avant tout dépôt réglementaire.**

---

## 1. Récupérer le programme (.exe) sans rien installer

Un exécutable Windows est **compilé automatiquement par GitHub** à chaque
mise à jour du code (voir `.github/workflows/build-windows-exe.yml`).

Deux façons de le récupérer une fois le dépôt publié sur GitHub (étape 2) :

- **Onglet "Actions"** du dépôt → cliquez sur la dernière exécution du
  workflow *"Compiler le .exe Windows"* → section *Artifacts* en bas de
  page → téléchargez `GenerateurLiasseSYCEBNL-windows.zip`, qui contient
  `GenerateurLiasseSYCEBNL.exe`.
- **Onglet "Releases"** (uniquement si vous avez créé un tag de version,
  voir §3) → téléchargez directement le fichier `.exe` attaché à la
  release.

Aucun compte GitHub n'est nécessaire pour télécharger : le dépôt peut rester
public en lecture.

## 2. Publier ce projet sur GitHub

Depuis ce dossier, en local (ou en important le zip fourni) :

```bash
git init
git add .
git commit -m "Initial commit - Générateur de liasse SYCEBNL"
git branch -M main
git remote add origin https://github.com/<votre-compte>/<votre-repo>.git
git push -u origin main
```

Dès le push, l'onglet **Actions** du dépôt GitHub lance automatiquement la
compilation sur une machine Windows fournie par GitHub (`windows-latest`) et
produit `GenerateurLiasseSYCEBNL.exe` (voir §1). Aucune configuration
supplémentaire n'est nécessaire : le workflow est déjà inclus dans le dépôt
(`.github/workflows/build-windows-exe.yml`).

## 3. Publier une "Release" avec le .exe en pièce jointe (optionnel)

```bash
git tag v1.0.0
git push origin v1.0.0
```

Le workflow détecte le tag `v1.0.0`, recompile et crée automatiquement une
Release GitHub avec l'exécutable attaché, prête à partager par un simple
lien.

## 4. Lancer le programme en développement (sans compiler)

Nécessite Python 3.10+ installé.

```bash
pip install -r requirements.txt
python src/main.py
```

## 5. Compiler le .exe soi-même (sous Windows, sans passer par GitHub)

```powershell
pip install -r requirements.txt
pyinstaller sycebnl_liasse.spec --noconfirm --clean
```

L'exécutable est produit dans `dist\GenerateurLiasseSYCEBNL.exe`.

---

## 6. Utilisation du logiciel

1. **Choisir le type d'entité** : Association/ONG classique, ou Projet de
   développement financé par des bailleurs.
2. **Renseigner** le nom de l'entité et la date de clôture de l'exercice.
3. **Générer un modèle de balance** (menu *Fichier*) si vous n'avez pas
   encore de fichier au bon format, le remplir, puis l'importer.
   - Colonnes attendues : `Compte`, `Intitulé`, `Débit`, `Crédit` (une
     ligne par compte de la balance, sans ligne de total).
4. *(Mode Association, optionnel)* Importer la balance de l'exercice N-1
   pour obtenir le tableau des flux de trésorerie.
5. *(Mode Projet, optionnel)* Importer un fichier budget (modèle disponible
   dans le même menu) pour obtenir le tableau d'exécution budgétaire, et
   saisir la trésorerie d'ouverture / le solde des relevés bancaires pour
   la réconciliation de trésorerie.
6. Cliquer sur **"Générer la liasse Excel"** et choisir où enregistrer le
   fichier produit.

Le fichier généré contient une feuille par état financier, une feuille
**"Comptes non classés"** listant les lignes de balance que le logiciel n'a
pas su rattacher automatiquement (à vérifier manuellement), et une feuille
**"Notes annexes"**.

## 7. Structure du projet

```
sycebnl-liasse/
├── src/
│   ├── main.py               # Interface graphique (Tkinter) - point d'entrée
│   ├── importer.py           # Lecture souple des fichiers balance / budget
│   ├── mapping_sycebnl.py    # Référentiel compte -> rubrique (à ajuster si besoin)
│   ├── engine.py             # Classement des comptes de la balance
│   ├── liasse_association.py # Calcul Bilan / CR / Flux de trésorerie
│   ├── liasse_projet.py      # Calcul Ressources-Emplois / Exéc. budgétaire / Réconciliation
│   ├── excel_export.py       # Génération du classeur Excel final
│   └── templates.py          # Génération des modèles Excel à remplir
├── .github/workflows/
│   └── build-windows-exe.yml # Compilation automatique du .exe par GitHub
├── requirements.txt
├── sycebnl_liasse.spec       # Configuration PyInstaller
└── README.md
```

## 8. Adapter le classement comptable à votre pays / administration

Le fichier `src/mapping_sycebnl.py` liste, rubrique par rubrique, les
préfixes de comptes SYCEBNL qui y sont rattachés. Si l'administration
fiscale de votre pays impose un intitulé ou un regroupement différent,
modifiez simplement les listes `RUBRIQUES_ACTIF`, `RUBRIQUES_PASSIF`,
`RUBRIQUES_CHARGES`, `RUBRIQUES_PRODUITS`, `RUBRIQUES_HAO` : aucune autre
partie du code n'a besoin d'être touchée.

## Licence

MIT — voir `LICENSE`. Logiciel fourni "en l'état", sans garantie de
conformité réglementaire (voir avertissement en tête de ce document).
