# smartPrintPay

![Python](https://img.shields.io/badge/Python-3.12%20%7C%203.14-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-0078D4?style=for-the-badge&logo=windows&logoColor=white)
![UI](https://img.shields.io/badge/UI-CustomTkinter-00A1FF?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Prototype%20actif-success?style=for-the-badge)
![Packaging](https://img.shields.io/badge/Packaging-PyInstaller-4EAA25?style=for-the-badge)
![School](https://img.shields.io/badge/Use-Print%20Management%20School-FF6F61?style=for-the-badge)

Application desktop Python destinée à la gestion et au calcul du coût d’impression scolaire, avec interface moderne, impression directe, historique local et synchronisation vers Google Sheets.

## Sommaire

- [Présentation](#présentation)
- [Fonctionnalités](#fonctionnalités)
- [Technologies utilisées](#technologies-utilisées)
- [Arborescence du projet](#arborescence-du-projet)
- [Installation](#installation)
- [Lancement](#lancement)
- [Utilisation](#utilisation)
- [Tarification](#tarification)
- [Déploiement](#déploiement)
- [Détails techniques](#détails-techniques)
- [Points de vigilance](#points-de-vigilance)
- [Améliorations futures](#améliorations-futures)

## Présentation

smartPrintPay est une application de bureau créée pour simplifier l’impression dans un contexte scolaire ou de reprographie. Elle permet à un enseignant ou à un agent de :

- sélectionner un document PDF, Word ou Excel ;
- choisir un enseignant ;
- définir la plage de pages à imprimer ;
- choisir le mode Recto ou Recto-Verso ;
- activer ou non l’impression couleur ;
- sélectionner l’imprimante disponible ;
- calculer le coût total automatiquement ;
- lancer l’impression ;
- enregistrer l’opération localement dans une base SQLite ;
- synchroniser les logs vers Google Sheets.

Le projet a été pensé pour fonctionner à la fois sur Windows et sur macOS, tout en gardant une interface simple et accessible pour des utilisateurs non techniques.

## Fonctionnalités

### Gestion des enseignants

- ajout d’un enseignant depuis l’interface ;
- suppression d’un enseignant ;
- sélection dans une liste déroulante ;
- stockage en base SQLite locale.

### Sélection des documents

- support des fichiers PDF, DOCX, DOC et XLSX ;
- sélection via un explorateur système ;
- historique des derniers fichiers utilisés.

### Calcul de coût

- calcul automatique selon le nombre de pages ;
- mode Recto Simple : 15 FCFA par feuille ;
- mode Recto-Verso : 25 FCFA par feuille ;
- impression couleur : 25 FCFA supplémentaire par feuille ;
- nombre de copies configurable ;
- affichage du total avant impression.

### Impression

- détection automatique des imprimantes selon le système ;
- mode test sans imprimante pour simuler un traitement ;
- impression native sur macOS via la commande `lp` ;
- impression standard sur Windows via le système d’impression natif.

### Aperçu de pages

- aperçu visuel de PDF ;
- navigation entre pages sélectionnées ;
- calcul de la plage de pages demandée.

### Historique et synchronisation

- enregistrement des impressions dans une base SQLite locale ;
- journalisation des données d’impression ;
- synchronisation par lots vers Google Sheets via un webhook ;
- protection du thread UI pour éviter les blocages graphiques.

## Technologies utilisées

- Python 3.12 / 3.14
- CustomTkinter pour l’interface graphique
- SQLite pour la base locale
- PyPDF pour la lecture des PDF
- Pillow pour l’aperçu des pages PDF
- PyMuPDF (optionnel selon l’environnement) pour la génération d’aperçu
- Requests pour la synchronisation HTTP
- PyInstaller pour le packaging en exécutable
- win32print côté Windows
- `lp` / `lpstat` côté macOS

## Arborescence du projet

```text
smartPrintPay/
├── .git/                  # données Git du projet
├── .kilo/                 # éventuels fichiers de config/usage Kilo
├── .python-version        # version Python ciblée par le projet
├── .venv/                 # environnement virtuel local
├── app.py                 # application principale
├── printpay.db            # base SQLite locale (enseignants + logs)
├── pyproject.toml         # configuration du projet et dépendances
├── README.md              # documentation du projet
├── requirements.txt       # dépendances Python
├── uv.lock                # lockfile UV
└── .gitignore             # (si présent selon l’environnement local)
```

## Installation

### Avec UV (recommandé)

```bash
uv sync
```

### Avec pip

```bash
python -m venv .venv
source .venv/bin/activate   # macOS / Linux
# ou .venv\Scripts\activate  # Windows
pip install -r requirements.txt
```

### Dépendances importantes

```bash
pip install customtkinter pypdf Pillow pymupdf requests pyinstaller openpyxl pandas psutil docx2pdf
```

Sur Windows, il est utile d’ajouter :

```bash
pip install pywin32
```

## Lancement

### Depuis le terminal

```bash
uv run python app.py
```

ou

```bash
source .venv/bin/activate
python app.py
```

### En mode simulation

L’application fournit un mode “Mode Test (Simulation sans imprimante)” dans la liste des imprimantes. Ce mode sert à tester le comportement sans action réelle d’impression.

## Utilisation

1. Choisir un enseignant.
2. Sélectionner un document PDF, Word ou Excel.
3. Définir la plage de pages à imprimer.
4. Choisir le type d’impression : Recto ou Recto-Verso.
5. Cocher “Couleur” si nécessaire.
6. Définir le nombre de copies.
7. Sélectionner l’imprimante disponible.
8. Vérifier le résumé du coût.
9. Cliquer sur “IMPRIMER”.
10. Vérifier l’historique et la synchronisation dans la base locale.

## Tarification

Le calcul est intégré directement dans la logique de l’application.

| Mode | Prix unitaire | Note |
| --- | --- | --- |
| Recto Simple, Noir & Blanc | 15 FCFA | par feuille |
| Recto-Verso, Noir & Blanc | 25 FCFA | par feuille imprimée |
| Recto Simple, Couleur | 25 FCFA | par feuille |
| Recto-Verso, Couleur | 50 FCFA | par feuille |

Exemple :

- 10 pages
- 20 exemplaires
- Recto, noir & blanc

Total calculé :

$$
10 \times 20 \times 15 = 3000 \text{ FCFA}
$$

Autre exemple :

- 10 pages
- 20 exemplaires
- Recto-Verso, noir & blanc

Le système convertit automatiquement le nombre de feuilles pour calculer le bon coût.

## Déploiement

### Windows

Le projet peut être empaqueté en exécutable unique :

```bash
pyinstaller --noconsole --onefile app.py
```

### macOS

```bash
pyinstaller --noconsole --windowed app.py
```

Cette logique est pensée pour livrer une application utilisable sans dépendances Python côté utilisateur final.

## Détails techniques

### Base de données locale

Le fichier `printpay.db` contient au moins deux tables :

- `teachers` : liste des enseignants
- `print_logs` : historique des impressions avec :
  - nom de l’enseignant ;
  - nom du fichier ;
  - chemin du fichier ;
  - plage sélectionnée ;
  - nombre de pages ;
  - copies ;
  - total de feuilles ;
  - mode d’impression ;
  - type noir/blanc ou couleur ;
  - coût total ;
  - date et heure ;
  - état de synchronisation.

### Synchronisation

Les lignes non synchronisées sont envoyées via une requête HTTP vers un webhook Google Apps Script. Cette logique est exécutée dans un thread distinct pour ne pas bloquer l’interface utilisateur.

### Sécurité / robustesse

- éviter les segfaults UI par l’utilisation d’un queue de messages côté thread principal ;
- validation de la plage de pages ;
- contrôle des valeurs de copies ;
- détection des erreurs d’impression ;
- message d’alerte en cas de problème de connexion ou d’imprimantes indisponibles.

## Points de vigilance

### Système d’exploitation

- Windows : dépend de l’environnement, des pilotes et du système d’impression native.
- macOS : repose sur `lp` et les outils nativement fournis par le système.

### Versions Python

Le projet est configuré avec une version Python ciblée dans `.python-version` et dans `pyproject.toml`. Vérifiez la compatibilité avant d’installer les dépendances.

### Dépendances optionnelles

Certains paquets sont conditionnels selon l’OS (par exemple `pywin32`). Cela évite les erreurs d’installation sur macOS.

## Améliorations futures

- ajouter un vrai système de gestion des quotas par enseignant ;
- exporter les historiques en CSV ou Excel ;
- gérer les fichiers images et les formats supplémentaires ;
- améliorer le design avec une vraie zone de prévisualisation plus avancée ;
- sécuriser le webhook avec authentification ou proxy propre ;
- générer un packaging final plus stable pour la production.

## Conclusion

smartPrintPay est un projet de gestion d’impression scolaire pensé pour être simple, rapide et fiable, avec un usage réel dans un environnement de classe ou de reprographie. L’architecture actuelle associe interface utilisateur, base locale, calcul tarifaire automatique et synchronisation réseau, ce qui en fait un outil pratique pour des usages quotidiens.

---

Projet : smartPrintPay  
Type : application desktop Python  
Objectif : gestion et calcul des impressions scolaires  
Compatibilité : Windows et macOS
Docker: docker run --rm -v "$(pwd):/src" --platform linux/amd64 cdrx/pyinstaller-windows:python3 \
  "pip install -r requirements.txt && pyinstaller --noconsole --onefile --collect-all customtkinter --name PrintPay app.py"