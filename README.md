# 🖨️ PrintPay — Station d'Impression & Comptabilité Scolaire Intelligente

<p align="center">
  <img src="https://img.shields.io/badge/Version-1.0.0-0284C7?style=for-the-badge" alt="Version">
  <img src="https://img.shields.io/badge/Python-3.12-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/Interface-CustomTkinter-blue?style=for-the-badge" alt="CustomTkinter">
  <img src="https://img.shields.io/badge/Moteur_PDF-PyMuPDF_300DPI-FF6F00?style=for-the-badge" alt="PyMuPDF">
  <img src="https://img.shields.io/badge/Base_de_Donn%C3%A9es-SQLite3-003B57?style=for-the-badge&logo=sqlite&logoColor=white" alt="SQLite">
  <img src="https://img.shields.io/badge/Cloud-Google_Sheets_API-34A853?style=for-the-badge&logo=googlesheets&logoColor=white" alt="Google Sheets">
</p>

> **PrintPay** réinvente le dialogue d'impression classique (*Ctrl + P*) pour les établissements scolaires et centres de reprographie.  
> C'est une application de bureau multiplateforme (**Windows & macOS**) autonome, **100 % opérationnelle hors-ligne**, combinant prévisualisation interactive, routage d'imprimante matériel, calcul tarifaire automatisé et synchronisation cloud discrète vers un tableau de bord décisionnel.

---

## 📌 Sommaire
1. [Problématique & Solution](#-problématique--solution)
2. [Fonctionnalités Clés](#-fonctionnalités-clés)
3. [Architecture Technique & Innovations](#-architecture-technique--innovations)
4. [Règles de Calcul & Modèle Économique](#-règles-de-calcul--modèle-économique)
5. [Tableau de Bord & Gestion des Stocks (Google Sheets)](#-tableau-de-bord--gestion-des-stocks-google-sheets)
6. [Installation & Développement](#-installation--développement)
7. [Génération des Exécutables (Build)](#-génération-des-exécutables-build)
8. [Guide d'Utilisation](#-guide-dutilisation)
9. [Auteur](#-auteur)

---

## 🎯 Problématique & Solution

### Le problème en établissement scolaire :
* **Gaspillage massif de papier** : Un clic accidentel sur *Imprimer* lance souvent l'intégralité d'un document (ex: 149 pages) au lieu d'une fiche précise.
* **Dérive des coûts** : L'encre couleur est utilisée à l'insu de l'administration sur des documents qui ne nécessitent que du noir et blanc.
* **Comptabilité manuelle fastidieuse** : Noter sur un cahier qui a imprimé quoi engendre des oublis, des contestations et une perte de revenus pour le centre de reprographie.

### La solution PrintPay :
* **Pré-traitement physique des fichiers** : L'application découpe le document avant impression ; si l'utilisateur demande `1-1`, le bac d'impression ne reçoit physiquement qu'une seule page.
* **Double verrouillage Noir & Blanc / Couleur** : Conversion à la volée en niveaux de gris purs (0 % de composante couleur) pour interdire le tirage couleur accidentel.
* **Traçabilité totale en temps réel** : Chaque feuille tirée est débitée au compte du professeur, enregistrée en base locale SQLite et synchronisée sur Google Sheets pour les bilans mensuels.

---

## ✨ Fonctionnalités Clés

* 👁️ **Aperçu Dynamique en Direct** :
  * L'aperçu réagit instantanément à la sélection de plage (`1-1`, `1-5`, etc.).
  * **Contrôle visuel de la couleur** : Par défaut, l'aperçu s'affiche en **Noir & Blanc `[N&B]`**. Dès que la case couleur est cochée, il bascule en direct en **Couleur `[Couleur]`**.
* 📐 **Orientation Automatique (façon Foxit Reader)** :
  * Détection automatique des formats paysage (`largeur > hauteur`).
  * Rotation physique de 90° pour exploiter 100 % de la feuille A4 (fini les bandes blanches géantes ou les rétrécissements à 70 %).
* 🖨️ **Routage Strict d'Imprimante** :
  * Détection à chaud des imprimantes connectées (USB, réseau, ticket POS).
  * Envoi direct à l'imprimante choisie (supprime le bug de forçage de l'imprimante système par défaut).
* 🧪 **Mode Test & Simulation** :
  * Permet de simuler un tirage complet (calcul, débit, enregistrement et synchronisation) même si aucune imprimante n'est physiquement branchée.
* ☁️ **Synchronisation Cloud Asynchrone (Zero Lag)** :
  * Fonctionnement garanti hors-ligne (stockage local immédiat).
  * Démon multithread sécurisé avec file d'attente (`queue.Queue`) qui pousse les données par lot vers Google Sheets dès qu'une connexion internet est détectée.

---

## 🏗️ Architecture Technique & Innovations

```
┌─────────────────────────────────────────────────────────────┐
│                   PrintPay (Interface Utilisateur)          │
└───────────────┬─────────────────────────────┬───────────────┘
                │                             │
    [Pré-traitement PDF]             [Base Locale SQLite]
    • Découpe de plage               • Table teachers
    • Rotation physique 90°          • Table print_logs (is_synced)
    • Rendu 300 DPI (csRGB / csGRAY)          │
                │                             ▼
                ▼                [Démon Multithread Sécurisé]
    [Moteur d'Impression]                     │
    • macOS : CUPS (lp -d ...)                ▼
    • Windows : SumatraPDF CLI       [Webhook Google Apps Script]
                │                             │
                ▼                             ▼
    [Sortie Papier Physique]         [Google Sheets : smartprintPay]
    (Canon, HP, Epson...)            • Journal détaillé
                                     • Bilan par Enseignant
                                     • Tableau de Bord & Donut Dépenses
```

### 1. Pourquoi le "Slicing" résout le bug des 149 pages :
Au lieu de déléguer le filtrage au pilote Windows/Mac (qui l'ignore souvent en ligne de commande), PrintPay utilise **PyMuPDF** pour extraire uniquement les pages demandées dans un fichier temporaire :
$$\text{Fichier source (149 pages)} \xrightarrow{\text{Filtre 1-1}} \text{Fichier temporaire (1 page physique)}$$
L'imprimante est physiquement incapable d'imprimer le reste du document.

### 2. Pourquoi le double verrouillage N&B bloque l'encre couleur :
Sur imprimante jet d'encre (ex: Canon G3010), un document RVB imprimé sans ordre strict mélange les encres Cyan, Magenta et Jaune pour fabriquer un noir composite.  
PrintPay neutralise cela à la racine :
1. **Niveau Fichier** : Rendu raster 300 DPI converti en `csGRAY` (1 seul canal, 0 octet de couleur).
2. **Niveau Matériel** : Envoi de l'instruction matérielle `monochrome` (Windows) ou `-o print-color-mode=monochrome` (macOS).

---

## 💰 Règles de Calcul & Modèle Économique

Le barème intégré applique une tarification transparente et équitable :

| Mode d'impression | Tarif unitaire | Formule de calcul |
| :--- | :---: | :--- |
| **Recto Simple (Noir & Blanc)** | **15 FCFA** | $\text{Feuilles} = \text{Pages} \times \text{Copies}$ |
| **Recto-Verso (Noir & Blanc)** | **25 FCFA** | $\text{Feuilles} = \lceil\text{Pages} / 2\rceil \times \text{Copies}$ |
| **Recto Simple (Couleur)** | **25 FCFA** | $\text{Feuilles} = \text{Pages} \times \text{Copies}$ |
| **Recto-Verso (Couleur)** | **50 FCFA** | $\text{Feuilles} = \lceil\text{Pages} / 2\rceil \times \text{Copies}$ |

### Comptes Spéciaux intégrés :
* **`photocopie`** : Dédié aux tirages comptoir (élèves, secrétariat, usagers externes).
* **`papier gache`** : Dédié au suivi des bourrages et ratés techniques. Les feuilles sont décomptées du stock réel sans être facturées à un enseignant.

---

## 📊 Tableau de Bord & Gestion des Stocks (Google Sheets)

Le classeur Google Sheets connecté (`smartprintPay`) est structuré en 4 onglets :

1. **`Journal`** : Historique brut exhaustif alimenté en temps réel par les postes d'impression (ID unique, Date, Enseignant, Fichier, Plage, Feuilles, Mode, Type, Coût).
2. **`Bilan Enseignants`** : Synthèse mensuelle dynamique (`YYYY-MM`) indiquant le total des feuilles et le montant dû par chaque professeur.
3. **`Dépenses`** : Registre d'achats pour la reprographie (Cartons de papier, Toners, Boîtes d'agrafes, Transport, Achat machines).
4. **`Tableau de Bord`** : Vue exécutive avec calcul automatique de la rentabilité :
   * **Consommation de papier** : $\text{Cartons consommés} = \frac{\text{Total Feuilles}}{2\,500}$ *(base : 1 carton = 5 rames de 500 feuilles)*.
   * **Bénéfice Net d'Exploitation** : $\text{Recettes Tirages} - \text{Dépenses Consommables}$.
   * **Graphique 1 (Barres)** : Classement des enseignants selon le volume de papier consommé.
   * **Graphique 2 (Donut)** : Répartition analytique des dépenses (Papier vs Encre vs Fournitures).

---

## 🚀 Installation & Développement

### Prérequis
* Python **3.11** ou **3.12**
* Gestionnaire de paquets **`uv`** (recommandé) ou `pip`

### 1. Cloner le projet
```bash
git clone https://github.com/votre-compte/smartPrintPay.git
cd smartPrintPay
```

### 2. Installer les dépendances
```bash
uv venv --python 3.12
source .venv/bin/activate   # Sur Windows : .venv\Scripts\activate
uv pip install -r requirements.txt
```

### Contenu de `requirements.txt` :
```text
customtkinter>=5.2.0
pypdf>=4.0.0
Pillow
pymupdf
requests
pywin32>=306; sys_platform == 'win32'
```

### 3. Lancer l'application en développement
```bash
python app.py
```

---

## 📦 Génération des Exécutables (Build)

Aucun des postes finaux de l'école n'a besoin d'avoir Python installé. L'application est compilée en fichier autonome.

### 🍏 Compilation pour macOS (`PrintPay.app`)
À exécuter directement sur votre Mac :
```bash
.venv/bin/pyinstaller --noconsole --windowed --collect-all customtkinter --name "PrintPay" app.py
```
👉 L'application **`PrintPay.app`** est générée dans le dossier `dist/`.

### 🪟 Compilation pour Windows (`PrintPay.exe`)

#### Méthode A : Directement depuis votre Mac via Docker (Recommandée)
Nul besoin d'allumer un PC Windows pour compiler l'exécutable Windows. Lancez :
```bash
docker run --rm -v "$(pwd):/src" -w /src --platform linux/amd64 tobix/pywine:3.11 \
  sh -c "wine python -m pip install pyinstaller -r requirements.txt && wine pyinstaller --noconsole --onefile --collect-all customtkinter --name PrintPay app.py"
```
👉 Le fichier **`PrintPay.exe`** est généré directement dans votre dossier `dist/`.

#### Méthode B : Directement sur un PC Windows
1. Installez Python 3.12 sur le PC (en cochant **☑️ Add python.exe to PATH**).
2. Lancez :
   ```cmd
   pip install -r requirements.txt
   pip install pyinstaller
   pyinstaller --noconsole --onefile --collect-all customtkinter --name PrintPay app.py
   ```

> 💡 **Important pour Windows** : Déposez le petit fichier portable **`SumatraPDF.exe`** (fourni dans le dossier `win/`) dans le même répertoire que `PrintPay.exe` pour garantir l'impression silencieuse sans dépendre de logiciels tiers.

---

## 📖 Guide d'Utilisation

1. **Lancement** : Double-cliquez sur l'icône `PrintPay`.
2. **Sélection de l'enseignant** : Choisissez dans la liste déroulante (ou cliquez sur `➕` pour en ajouter un, `🗑️` pour en supprimer).
3. **Choix du document** : Cliquez sur `📂 Parcourir...` ou sélectionnez l'un des fichiers récents.
4. **Options** :
   * Choisissez le mode : *Recto* ou *Recto-Verso*.
   * Cochez *Couleur* si nécessaire (l'aperçu à droite confirme immédiatement la couleur).
   * Indiquez la plage (ex : `1-1` pour la première page, `1-3` pour les 3 premières).
   * Ajustez le nombre d'exemplaires avec `+` et `-`.
5. **Impression** : Sélectionnez l'imprimante et cliquez sur **`🖨️ IMPRIMER`**.

---

## 👨‍💻 Auteur

**Edjabrou Ulrich Blanchard KABLAM**  
*Ingénieur en Intelligence Artificielle & Data Science*  
*Enseignant de Mathématiques (ENS Abidjan)*  

* **LinkedIn** : [blanchard-kablam](https://www.linkedin.com/in/blanchard-kablam)
* **ORCID** : [0009-0009-1464-5049](https://orcid.org/0009-0009-1464-5049)
* **Email** : bkablam11@gmail.com
* **Localisation** : Abidjan, Côte d'Ivoire

---
*Projet développé avec passion pour moderniser et fiabiliser la gestion matérielle et financière des établissements éducatifs.*
