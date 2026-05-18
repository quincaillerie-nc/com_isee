# 📊 COM ISEE - Génération et envoi automatique du relevé de prix HT

## 📌 Présentation

`com_isee` est un script Python permettant de générer automatiquement un fichier Excel contenant une sélection d’articles à transmettre à l’ISEE.

Le script :

- charge les données articles depuis un fichier `.DBF`,
- filtre uniquement les références suivies par l’ISEE,
- génère un fichier Excel propre et formaté,
- envoie automatiquement le fichier par email aux destinataires définis,
- conserve une logique modulaire en réutilisant les modules internes :
  - `file-manager`,
  - `dbf-loader`,
  - `quincaillerie-mailer`.

---

## 🎯 Objectif du script

Ce script a pour objectif d’automatiser l’envoi mensuel du relevé des prix HT à l’ISEE.

Il permet d’éviter :

- les extractions manuelles,
- les erreurs de copier-coller,
- les oublis d’envoi,
- les problèmes de format Excel,
- la perte de temps sur les traitements répétitifs.

---

## 🧩 Modules utilisés

Le script repose sur trois modules internes.

### 1. `file-manager`

Utilisé pour générer automatiquement le chemin du fichier Excel de sortie.

Fonctions utilisées :

```python
generer_chemin_fichier()
generer_nom_fichier()
```

---

### 2. `dbf-loader`

Utilisé pour charger les données du fichier DBF article.

Fonction utilisée :

```python
get_dbf()
```

Exemple utilisé dans le script :

```python
df = get_dbf("qc/article.dbf")
```

---

### 3. `quincaillerie-mailer`

Utilisé pour envoyer le fichier Excel généré par email.

Fonction utilisée :

```python
envoyer_email()
```

---

## 🏗️ Structure attendue du projet

Le script est conçu pour fonctionner dans l’architecture suivante :

```bash
dev/
│
├── modules/
│   │
│   ├── file-manager/
│   │   └── file_manager.py
│   │
│   ├── dbf-loader/
│   │   └── dbf_loader.py
│   │
│   └── quincaillerie-mailer/
│       └── mailer.py
│
└── scripts/
    │
    └── com_isee/
        ├── com_isee.py
        └── README.md
```

---

## ⚙️ Configuration principale

### Société concernée

```python
SOC = "QC"
```

Cette variable est utilisée dans :

- le nom du fichier Excel,
- le sujet de l’email,
- l’identification du rapport.

---

## 📧 Destinataires des emails

Le script utilise deux listes de destinataires.

### Maintenance

```python
MAIL_MAINTENANCE = [
    "support@quincaillerie.nc"
]
```

Cette liste permet d’envoyer une copie au support informatique ou au service technique.

---

### Destinataires principaux

```python
MAIL_TARGET = [
    "support@quincaillerie.nc",
    "exploitation@quincaillerie.nc",
    "n.leroux@quincaillerie.nc",
    "indices@isee.nc"
]
```

Les doublons sont automatiquement supprimés avant l’envoi grâce à :

```python
list(set(MAIL_MAINTENANCE + MAIL_TARGET))
```

---

## 📦 Articles suivis

La liste `NARTS` contient les références articles à transmettre à l’ISEE.

Exemple :

```python
NARTS = [
    "710092",
    "760043",
    "760041"
]
```

Seules les lignes dont la référence article correspond à cette liste sont conservées.

---

## 🔄 Fonctionnement global du script

Le script suit les étapes suivantes :

1. démarrage du script,
2. calcul automatique du mois de rapport,
3. chargement du fichier `qc/article.dbf`,
4. filtrage des articles suivis,
5. sélection des colonnes utiles,
6. génération du fichier Excel,
7. envoi automatique du fichier par email,
8. affichage des logs de réussite ou d’erreur.

---

## 📅 Calcul automatique du mois concerné

Le script génère le rapport pour le mois précédent.

Exemple :

Si le script est lancé en mai 2026, le rapport généré concerne avril 2026.

Code utilisé :

```python
today = datetime.now()
report_date = today.replace(day=1) - timedelta(days=1)
```

---

## 📥 Chargement des données DBF

Le fichier chargé est :

```python
qc/article.dbf
```

Code utilisé :

```python
df = get_dbf("qc/article.dbf")
```

Le module `dbf-loader` retourne un DataFrame pandas.

---

## 🔎 Filtrage des articles

Le script filtre uniquement les références présentes dans `NARTS`.

```python
df_filtered = df[df["NART"].astype(str).isin(NARTS)]
```

---

## 📊 Colonnes exportées

Le script conserve uniquement trois colonnes :

| Colonne DBF | Colonne Excel | Description |
|---|---|---|
| `NART` | `Réf_Mag` | Référence magasin |
| `DESIGN` | `Produit` | Désignation produit |
| `PVTE` | `PVTE_HT` | Prix de vente hors taxe |

Code utilisé :

```python
df_result = df_filtered[["NART", "DESIGN", "PVTE"]].rename(
    columns={
        "NART": "Réf_Mag",
        "DESIGN": "Produit",
        "PVTE": "PVTE_HT"
    }
)
```

---

## 📁 Génération du fichier Excel

Le fichier généré suit ce format :

```bash
QC_comISEE_mois-année.xlsx
```

Exemple :

```bash
QC_comISEE_avr-26.xlsx
```

Le fichier est généré avec `pandas` et `openpyxl`.

---

## 🎨 Mise en forme Excel

Le fichier Excel généré contient :

- une feuille nommée `ISEE`,
- une ligne d’en-tête en bleu,
- du texte blanc en gras,
- un alignement centré,
- une largeur de colonne automatique.

---

## 📧 Envoi automatique des emails

Le sujet de l’email suit ce format :

```text
[QC] - Relevé ISEE prix HT - AVRIL 2026
```

Le corps de l’email est généré en HTML.

Le fichier Excel est ajouté en pièce jointe.

---

## 📝 Logs

Le script affiche des logs horodatés.

Exemple :

```bash
[08:30:15] Début du script ComISEE
[08:30:16] Rapport pour avril 2026
[08:30:20] 22 articles sélectionnés
[08:30:21] 📁 Fichier généré : ...
[08:30:22] 📤 Envoi à indices@isee.nc...
[08:30:23] ✅ Mail envoyé à indices@isee.nc
[08:30:24] ✅ Script terminé avec succès
```

---

## ▶️ Lancer le script

Depuis le dossier du script :

```bash
python com_isee.py
```

Ou sous Windows :

```powershell
py com_isee.py
```

---

## 📦 Dépendances Python

Installer les dépendances nécessaires :

```bash
pip install pandas openpyxl dbfread
```

Les modules internes doivent également être disponibles dans le dossier `modules`.

---

## ✅ Pré-requis

Avant d’exécuter le script, vérifier que :

- le fichier DBF est accessible,
- le module `dbf-loader` fonctionne,
- le module `file-manager` fonctionne,
- le module `quincaillerie-mailer` fonctionne,
- les emails destinataires sont corrects,
- le serveur SMTP est configuré dans le module mailer,
- les références `NARTS` sont à jour.

---

## 🚨 Gestion des erreurs

En cas d’erreur, le script :

- affiche un message d’erreur,
- affiche le détail technique avec `traceback`,
- arrête l’exécution avec `sys.exit(1)`.

Exemple :

```python
except Exception as e:
    log(f"❌ ERREUR : {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
```

---

## 🔐 Sécurité

Attention : ce script utilise des modules internes qui peuvent contenir des accès réseau, des identifiants SMB ou SMTP.

Il est recommandé de :

- ne jamais stocker les mots de passe dans le code,
- utiliser un fichier `.env`,
- ignorer `.env` dans Git,
- ne pas publier les fichiers sensibles sur GitHub,
- utiliser un `.gitignore`.

---

## 📄 Exemple de `.gitignore`

Créer un fichier `.gitignore` :

```gitignore
__pycache__/
*.pyc
.env
.venv/
venv/
*.log
*.xlsx
```

Les fichiers Excel générés ne doivent généralement pas être versionnés.

---

## 🚀 Initialisation Git

Depuis le dossier du script :

```powershell
git init
git add .
git commit -m "initial commit"
```

---

## 🔗 Ajouter le dépôt distant

Version HTTPS :

```powershell
git remote add origin https://github.com/quincaillerie-nc/com_isee.git
git branch -M main
git push -u origin main
```

---

## 🔐 Version recommandée avec SSH

Si HTTPS pose un problème d’authentification ou de permission, utiliser SSH :

```powershell
git remote set-url origin git@github.com:quincaillerie-nc/com_isee.git
git push -u origin main
```

Si le remote n’existe pas encore :

```powershell
git remote add origin git@github.com:quincaillerie-nc/com_isee.git
git branch -M main
git push -u origin main
```

---

## 🧪 Vérifier le remote Git

```powershell
git remote -v
```

Résultat attendu en SSH :

```text
origin  git@github.com:quincaillerie-nc/com_isee.git (fetch)
origin  git@github.com:quincaillerie-nc/com_isee.git (push)
```

---

## 🧼 Nettoyer les fichiers inutiles avant push

Si un dossier `__pycache__` ou un fichier `.pyc` a été ajouté par erreur :

```powershell
git rm -r --cached __pycache__
git add .
git commit -m "clean repository files"
git push
```

---

## 🛠️ Améliorations possibles

Évolutions recommandées :

- ajout d’un fichier `.env`,
- ajout d’un fichier `config.json`,
- ajout d’un mode test sans envoi email,
- ajout d’un argument CLI pour choisir le mois,
- ajout d’un fichier log,
- archivage automatique des rapports générés,
- gestion d’un mode simulation,
- envoi d’un email maintenance uniquement en cas d’erreur.

---

## 📌 Exemple de mode test à ajouter plus tard

```python
SEND_EMAILS = False
```

Puis :

```python
if SEND_EMAILS:
    send_mails(file_path, file_name, report_date)
else:
    log("Mode test : aucun email envoyé")
```

---

## 👨‍💻 Maintenance

Pour modifier les articles suivis, mettre à jour la liste :

```python
NARTS = [...]
```

Pour modifier les destinataires, mettre à jour :

```python
MAIL_TARGET = [...]
MAIL_MAINTENANCE = [...]
```

Pour modifier le format Excel, adapter la fonction :

```python
generate_excel()
```

Pour modifier le contenu de l’email, adapter la fonction :

```python
send_mails()
```

---

## 📜 Conclusion

`com_isee` est un script métier simple, fiable et automatisé.

Il permet de produire chaque mois un relevé de prix HT propre, standardisé et envoyé automatiquement aux destinataires concernés.

Il s’intègre parfaitement dans une architecture Python modulaire pour les traitements internes, les exports métiers et les automatisations de reporting.
