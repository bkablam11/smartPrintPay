# Installation de PrintPay sur Windows pour débutant

Ce guide explique comment installer PrintPay sur un PC Windows qui n’a presque rien installé. L’objectif est de pouvoir lancer l’application facilement, même sans connaitre Python ni le terminal.

---

## 1. Ce qu’il faut préparer

Sur votre clé USB ou dans un dossier de travail, préparez un dossier nommé par exemple :

```text
PrintPay_Setup
```

Dedans, placez les fichiers suivants :

- `python-3.12.0-amd64.exe`
- `SumatraPDF.exe`
- `app.py`
- `requirements.txt`

> Si tu n’as pas encore le fichier `SumatraPDF.exe`, il faut le télécharger puis l’extraire. Le fichier final doit être directement exécutable, pas un ZIP.

---

## 2. Installer Python sur Windows

### Étape 1 : lancer l’installateur

Double-cliquez sur :

```text
python-3.12.0-amd64.exe
```

### Étape 2 : attention à cette case importante

Au début de l’installation, il y a souvent une case à cocher :

```text
Add python.exe to PATH
```

Il faut la cocher.

> Sans cette case, le terminal Windows ne saura pas où trouver Python et les commandes comme `python` ou `pip` ne fonctionneront pas.

### Étape 3 : continuer l’installation

Cliquez sur :

```text
Install Now
```

Laissez terminer l’installation. Cela peut prendre 1 à 3 minutes.

### Étape 4 : vérifier que Python est installé

Ouvrez le terminal Windows :

- appuyez sur la touche Windows
- tapez `cmd`
- appuyez sur Entrée

Dans le terminal, tapez :

```cmd
python --version
```

Si tout va bien, le terminal doit afficher quelque chose comme :

```text
Python 3.12.x
```

Si cela ne marche pas, il faut relancer l’installation et vérifier que la case “Add python.exe to PATH” a bien été cochée.

---

## 3. Créer un dossier pour le projet

Sur le disque C, créez un dossier :

```text
C:\PrintPay
```

Ensuite, copiez dans ce dossier :

- `app.py`
- `requirements.txt`
- `SumatraPDF.exe`

Le dossier doit ressembler à ceci :

```text
C:\PrintPay\
│
├── app.py
├── requirements.txt
├── SumatraPDF.exe
└── ...
```

> Il est important que le fichier `SumatraPDF.exe` soit dans le même dossier que le projet, sinon l’application ne retrouvera pas l’imprimante PDF de test ou ne pourra pas fonctionner correctement.

---

## 4. Ouvrir le terminal Windows dans le bon dossier

Dans le terminal Windows :

```cmd
cd C:\PrintPay
```

Tu dois maintenant être dans le dossier du projet.

Pour vérifier, tape :

```cmd
dir
```

Tu devrais voir les fichiers `app.py`, `requirements.txt` et `SumatraPDF.exe`.

---

## 5. Installer les dépendances du projet

Dans le terminal, exécute :

```cmd
pip install -r requirements.txt
```

Cela installe les bibliothèques Python nécessaires à l’application.

### Exemple de résultat attendu

Le terminal affichera plusieurs lignes de téléchargement et d’installation. Si tout se passe bien, il n’y a pas d’erreur à la fin.

> Si tu vois une erreur du type “pip not recognized”, alors Python n’a pas été ajouté au PATH. Il faut réinstaller Python et cocher la bonne case.

---

## 6. Lancer l’application directement

Tu peux maintenant lancer le programme avec :

```cmd
python app.py
```

Si tout fonctionne, la fenêtre de PrintPay doit s’ouvrir.

---

## 7. Créer le fichier exécutable Windows (.exe)

Pour que l’école puisse lancer le logiciel sans ouvrir le terminal, il faut créer un .exe.

### Étape 1 : installer PyInstaller

Dans le terminal, tape :

```cmd
pip install pyinstaller
```

### Étape 2 : créer le .exe

Exécute :

```cmd
pyinstaller --noconsole --onefile --collect-all customtkinter --add-binary "SumatraPDF.exe;." --name "PrintPay" app.py
```

### Explication rapide

- `--noconsole` : la fenêtre ne montre pas le terminal noir
- `--onefile` : un seul fichier exécutable est généré
- `--collect-all customtkinter` : permet de bien inclure l’interface graphique
- `--add-binary "SumatraPDF.exe;."` : ajoute le logiciel PDF au bon endroit dans le paquet final
- `--name "PrintPay"` : donne le nom du programme

---

## 8. Où trouver le fichier final

Après la commande, PyInstaller crée un dossier :

```text
dist
```

Dans ce dossier, tu trouveras :

```text
PrintPay.exe
```

C’est ce fichier qu’il faut donner aux utilisateurs.

---

## 9. Ce qu’il faut faire pour l’utilisateur final

Tu peux copier `PrintPay.exe` sur le bureau de l’école ou dans un dossier dédié :

```text
C:\PrintPay\
```

L’utilisateur peut ensuite :

1. double-cliquer sur `PrintPay.exe`
2. utiliser le logiciel normalement
3. lancer l’impression sans ouvrir le terminal

---

## 10. Astuces importantes

### Vérifier que le logiciel fonctionne bien

Avant de livrer le logiciel à l’école, fais un test rapide :

- ouvre `PrintPay.exe`
- choisi un enseignant
- sélectionne un PDF
- vérifie le calcul du coût
- clique sur l’impression de test

### Résoudre les erreurs fréquentes

#### Erreur : `python` n’est pas reconnu

- réinstaller Python
- cocher `Add python.exe to PATH`
- relancer le terminal

#### Erreur : `pip install -r requirements.txt` ne marche pas

- vérifier que tu es bien dans le bon dossier :

```cmd
cd C:\PrintPay
```

- vérifier que le fichier `requirements.txt` est présent

#### Erreur : le fichier exe ne se crée pas

- vérifier que le dossier contient bien `app.py`
- vérifier que `pyinstaller` est installé
- relancer la commande en étant dans le bon dossier

---

## 11. Résumé simple

Pour installer PrintPay sur un PC Windows sans rien :

1. installer Python 3.12 ;
2. cocher `Add python.exe to PATH` ;
3. copier `app.py`, `requirements.txt` et `SumatraPDF.exe` dans un dossier ;
4. ouvrir le terminal ;
5. lancer :

```cmd
cd C:\PrintPay
pip install -r requirements.txt
python app.py
```

6. pour livrer le logiciel à l’école :

```cmd
pip install pyinstaller
pyinstaller --noconsole --onefile --collect-all customtkinter --add-binary "SumatraPDF.exe;." --name "PrintPay" app.py
```

7. prendre le fichier `PrintPay.exe` dans le dossier `dist`

---

## 12. Objectif final

Le but est que la secrétaire ou la personne qui s’occupe de la reprographie puisse simplement :

- ouvrir le fichier `.exe`
- utiliser l’application
- imprimer sans avoir à installer Python ni à utiliser la ligne de commande

Si tu veux, je peux aussi te faire un second fichier plus court, prêt à imprimer pour “installer en 5 minutes sur PC Windows”, avec un style très pratique et ultra simple.
