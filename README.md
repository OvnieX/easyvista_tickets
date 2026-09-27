# EasyVista — Tickets uniques & Dashboard Excel

Ce projet transforme l'export **« Mes actions »** d'EasyVista en un fichier Excel plus exploitable.

Dans EasyVista, la vue **Mes actions** contient une ligne par action. Un même ticket peut donc apparaître plusieurs fois si plusieurs actions ont été réalisées dessus.

Le script `easyvista_tickets.py` déduplique cet export pour obtenir **une seule ligne par ticket**, puis génère automatiquement un classeur Excel contenant la liste des tickets uniques et un dashboard avec plusieurs graphiques.

> Le script ne crée volontairement **pas** de colonne `Nombre d'actions`.

---

## Fonctionnalités

Le script permet de :

- lire un export EasyVista au format **CSV** ou **XLSX** ;
- gérer les particularités de certains CSV EasyVista, notamment les lignes contenant une colonne vide supplémentaire ;
- ignorer la ligne technique éventuellement présente dans l'export ;
- conserver **une seule ligne par numéro de ticket** ;
- choisir la ligne la plus récente lorsqu'un ticket apparaît plusieurs fois ;
- trier les tickets récents en premier ;
- générer un fichier Excel final avec :
  - un onglet **Tickets uniques** ;
  - un onglet **Dashboard** ;
- créer automatiquement des graphiques par :
  - statut ;
  - priorité ;
  - type de ticket ;
  - mois ;
  - groupe ;
- afficher un mode `--verbose` pour diagnostiquer la lecture du fichier.

---

## Fonctionnement du dédoublonnage

La colonne utilisée comme identifiant unique est :

```text
N°
```

Lorsqu'un même ticket apparaît plusieurs fois, le script conserve la ligne considérée comme la plus récente selon cet ordre :

1. `Date de dernière modification`
2. `Date de fin réelle`
3. `Date d'émission`

Cela permet d'éviter de conserver arbitrairement la première action trouvée dans le fichier.

---

## Détection Incident / Demande de service

Pour alimenter le graphique correspondant, le script cherche le type de ticket dans plusieurs informations :

1. la colonne `Type` ;
2. la colonne `Sujet (complet)` ;
3. le préfixe du numéro de ticket :
   - `I...` → Incident ;
   - `S...` → Demande de service.

---

## Prérequis

- Python 3
- `xlsxwriter`
- `openpyxl` uniquement si l'entrée est un fichier `.xlsx`

### Installation

Depuis un terminal dans le dossier du projet :

```bash
python -m pip install -r requirements.txt
```

Ou directement :

```bash
python -m pip install xlsxwriter openpyxl
```

> La version actuelle du script n'utilise pas `pandas`.

---

## Utilisation

### Traiter un export CSV EasyVista

```bash
python easyvista_tickets.py "Mes actions.csv"
```

### Traiter un export Excel

```bash
python easyvista_tickets.py "Mes actions.xlsx"
```

### Afficher les détails du traitement

```bash
python easyvista_tickets.py "Mes actions.csv" --verbose
```

Exemple de sortie :

```text
[Lecture CSV] Encodage  : utf-8-sig
[Lecture CSV] Séparateur: ';'
[Lecture CSV] Colonnes  : 22
[Lecture CSV] Lignes    : 92
[Dédoublonnage] Lignes valides : 92
[Dédoublonnage] Tickets uniques: 37
[Dédoublonnage] Doublons retirés: 55

Traitement terminé avec succès.
Lignes/actions analysées : 92
Tickets uniques          : 37
Doublons retirés         : 55
Fichier créé             : ...\Mes actions_tickets_uniques_dashboard.xlsx
```

### Choisir le nom du fichier de sortie

```bash
python easyvista_tickets.py "Mes actions.csv" -o "Tickets_EasyVista.xlsx"
```

### Choisir une feuille précise dans un fichier XLSX

```bash
python easyvista_tickets.py "Mes actions.xlsx" --sheet "Mes actions"
```

---

## Fichier généré

Sans option `-o`, le fichier de sortie est automatiquement créé à côté du fichier source.

Exemple :

```text
Mes actions.csv
```

donne :

```text
Mes actions_tickets_uniques_dashboard.xlsx
```

### Onglet `Tickets uniques`

Cet onglet contient les données de l'export EasyVista avec **une seule ligne par ticket**.

Il conserve notamment, lorsque ces colonnes existent dans l'export :

- numéro du ticket ;
- titre ;
- statut ;
- priorité ;
- impact ;
- urgence ;
- groupe ;
- demandeur ;
- bénéficiaire ;
- sujet ;
- type ;
- dates ;
- informations Jira éventuelles.

Un filtre Excel est automatiquement ajouté et la ligne d'en-tête reste figée pendant le défilement.

### Onglet `Dashboard`

Le dashboard est construit uniquement à partir des **tickets uniques**.

Il contient notamment :

- le nombre total de tickets uniques ;
- les tickets par statut ;
- les tickets par priorité ;
- la répartition Incident / Demande de service ;
- les tickets par mois ;
- les principaux groupes.

---

## Particularité des exports CSV EasyVista

Certains exports EasyVista peuvent avoir, par exemple :

- **22 colonnes dans l'en-tête** ;
- mais **23 champs dans certaines lignes**, le dernier champ étant vide.

Une lecture CSV classique peut alors produire une erreur similaire à :

```text
Expected 22 fields in line 3, saw 23
```

Le script gère explicitement ce cas :

- les champs supplémentaires **vides** à droite sont supprimés ;
- une ligne trop courte est complétée avec des valeurs vides ;
- si un champ supplémentaire contient réellement une donnée, le traitement s'arrête avec un message d'erreur afin d'éviter toute perte silencieuse d'information.

---

## Options disponibles

Afficher l'aide :

```bash
python easyvista_tickets.py --help
```

| Option | Description |
|---|---|
| `fichier` | Fichier EasyVista `.csv` ou `.xlsx` à traiter |
| `-o`, `--output` | Choisit le nom du fichier Excel de sortie |
| `--sheet` | Choisit la feuille source si l'entrée est un XLSX |
| `-v`, `--verbose` | Affiche les informations techniques du traitement |

---

## Dépannage

### `ModuleNotFoundError: No module named 'xlsxwriter'`

```bash
python -m pip install xlsxwriter
```

### `ModuleNotFoundError: No module named 'openpyxl'`

Cette dépendance est nécessaire pour lire un fichier XLSX :

```bash
python -m pip install openpyxl
```

### Le fichier de sortie ne peut pas être créé

Fermer le fichier Excel s'il est déjà ouvert puis relancer le script. Sous Windows, Excel peut verrouiller le fichier et empêcher son remplacement.

### `Colonne 'N°' absente`

Le script s'attend à retrouver la colonne EasyVista `N°`. Vérifier que le fichier utilisé est bien un export compatible de la vue **Mes actions**.

### Le mauvais onglet XLSX est lu

```bash
python easyvista_tickets.py "Mes actions.xlsx" --sheet "Nom de la feuille"
```

---

## Confidentialité

Les exports EasyVista peuvent contenir des noms de collaborateurs, des numéros de tickets, des descriptions d'incidents, des groupes internes ou d'autres informations liées au système d'information.

Il est donc fortement recommandé de **ne jamais publier les exports réels sur un dépôt GitHub public**.

Le fichier `.gitignore` fourni ignore notamment :

```text
*.csv
*.xlsx
*.xls
```

Si un fichier sensible a déjà été ajouté à Git avant l'ajout du `.gitignore` :

```bash
git rm --cached "Mes actions.csv"
git commit -m "Remove EasyVista export from repository"
```

---

# GitHub

Le dépôt GitHub est publié sur :

```text
main
```

Pour une modification classique :

```bash
git status
git add .
git commit -m "Description de la modification"
git push
```
