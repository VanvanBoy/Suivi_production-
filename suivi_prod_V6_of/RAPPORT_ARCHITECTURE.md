# Rapport — Suivi de production (architecture & fonctionnement)

Ce document explique comment l'application est construite après refactorisation,
comment elle fonctionne en interne, ce qui a été corrigé/optimisé, et ce qu'il
resterait à améliorer. Il s'adresse à toi (ou à quiconque reprendra ce code plus tard).

---

## 1. Vue d'ensemble

L'application est un **IHM de suivi de production de batteries** (Tkinter),
connectée à une base **MySQL** (`suivi_production`, `produit_voltr`, `cellule`, ...)
et à un fichier **Excel de configuration** (`Suivi_prod_par_modele.xlsx`) qui décrit,
pour chaque référence de batterie, quelles étapes de fabrication s'appliquent et
dans quel ordre.

Au démarrage, l'utilisateur choisit une référence batterie. L'app lit alors la
feuille Excel `Flux test modele`, en déduit la liste et l'ordre des étapes
actives pour cette référence (picking, pack, nappe, bms, wrap, fermeture, capa,
fin de ligne, emballage, expédition, recherche, recyclage, tri test, banc
somfy...), et construit **dynamiquement un onglet par étape active**.

Chaque onglet permet à un opérateur de :
- scanner/saisir un numéro de série de cellule ou de batterie,
- consulter la liste des batteries encore "à faire" pour cette étape,
- valider l'étape (écrit `1` + date + visa dans la colonne correspondante de
  `suivi_production`) ou déclarer une non-conformité (incrémente un compteur
  d'échec `xxx_fail` et/ou ouvre un formulaire Google de déclaration).

Un mécanisme de **rafraîchissement automatique** (toutes les 10 s, cf.
`_tab_tick` dans `mixin_common.py`) recharge la liste des batteries à traiter
de l'onglet actif, pour que plusieurs postes puissent travailler en //
sur la même base sans redémarrer l'appli.

---

## 2. Architecture des fichiers

```
suivi_prod/
├── main.py                # Point d'entrée. Assemble StockApp. À viser pour PyInstaller.
├── config.py              # Constantes : chemins réseau, mapping étapes<->colonnes SQL
├── db_manager.py          # Connexion MySQL (demande identifiants au lancement)
├── mixin_common.py        # Logique transverse à tous les onglets (voir §3)
├── tab_picking.py         # Onglet "picking" (contrôle tension + remplacement cellule)
├── tab_pack.py            # Onglet "soudure pack"
├── tab_nappe.py           # Onglet "soudure nappe"
├── tab_bms.py             # Onglet "soudure BMS"
├── tab_wrap.py            # Onglet "wrap"
├── tab_fermeture.py       # Onglet "fermeture batterie"
├── tab_emballage.py       # Onglet "emballage"
├── tab_capa.py            # Onglet "test de capacité / cyclage" (le plus complexe)
├── tab_finligne.py        # Onglet "fin de ligne"
├── tab_expedition.py      # Onglet "expédition"
├── tab_recherche.py       # Onglet "recherche de batterie"
├── tab_recyclage.py       # Onglet "recyclage"
├── tab_tri_test.py        # Onglet "tri test"
├── tab_banc_somfy.py      # Onglet "banc Somfy"
└── requirements.txt
```

### Pourquoi des "Mixins" et pas des classes séparées ?

Chaque `tab_*.py` définit une classe `TabXxxMixin` qui **ne fonctionne pas
seule** : elle suppose l'existence de `self.db_manager`, `self.selected_model`,
etc., fournis par les autres mixins. `main.py` compose tout via héritage
multiple :

```python
class StockApp(ThemedTk, CommonMixin, TabPickingMixin, TabPackMixin, ...):
    ...
```

**Avantage** : c'est un découpage *mécanique*, donc sûr — chaque méthode garde
exactement le même corps et peut continuer à appeler `self.autre_methode()`
comme avant, peu importe dans quel fichier `autre_methode` vit réellement.
**Limite** : ce n'est pas un découpage "propre" au sens objet (pas
d'encapsulation, tout partage le même espace `self`). C'est un compromis
volontaire pour refactoriser un gros script Tkinter existant sans réécrire sa
logique métier (voir §7 pour la suite possible).

---

## 3. `mixin_common.py` — le socle commun

C'est le fichier le plus important à comprendre. Il regroupe :

| Fonction | Rôle |
|---|---|
| `_show_model_selector_and_build` | Popup de choix du modèle au lancement, lit `Flux test modele` dans l'Excel, construit `self.stage_order` (dict étape → rang, 0 = inactif) |
| `_create_widgets_with_order` | Crée un onglet Tkinter par étape active, dans l'ordre, en appelant le `setup_xxx` correspondant |
| `verif_etape_act_non_ok` | Vérifie qu'une étape n'est pas déjà validée pour une batterie donnée (empêche de revalider 2x) |
| `_required_previous_dbcols` / `_check_prereqs_and_warn` | Vérifie que **toutes les étapes précédentes** (selon `stage_order`) sont bien validées avant d'autoriser l'étape courante — c'est le garde-fou anti "on saute une étape" |
| `build_stage_query` / `build_stage_query_EOP` | Génère dynamiquement la requête SQL "liste des batteries prêtes pour cette étape" à partir de `stage_order` (la variante `_EOP` gère le cas particulier de la gamme `PPTR018A`, avec un `LIKE` plutôt qu'un `=`) |
| `_on_tab_changed` / `_tab_tick` / `_schedule_next_tab_tick` | Le rafraîchissement auto toutes les 10 s + jitter aléatoire (pour éviter que tous les postes tapent la BDD en même temps) |
| `make_tab_chain` / `register_focus_target` / `_focus_active_tab` | Gestion de l'ordre de tabulation (touche Tab) et du focus automatique sur le bon champ à l'ouverture d'un onglet |
| `set_photo` / `convert_comma_to_dot` | Utilitaires d'affichage (photo produit) et de saisie (accepte "3,72" comme "3.72") |
| `verfier_coherence_ref` / `changer_ref_batterie` | Vérifie que le numéro de série scanné correspond bien à la référence produit sélectionnée |
| **`_lookup_batterie_from_cellule`** (nouveau) | Helper factorisé : retrouve la batterie associée à une cellule scannée |
| **`_fill_entry_from_listbox_selection`** (nouveau) | Helper factorisé : recopie la sélection d'une listbox dans le champ de saisie |

---

## 4. Anatomie d'un onglet "type" (picking, pack, nappe, bms, wrap, fermeture, emballage, fin de ligne)

Ces 8 onglets suivent quasiment tous le même schéma (c'est ce qui a permis de
factoriser leur code — voir §6) :

1. **`setup_xxx(frame)`** : construit l'UI (champs, listbox, boutons), appelle
   `display_model_list_xxx()` une première fois, retourne le widget à focus.
2. **`xxx_check_entry_length`** : dès que le champ "n° série cellule" atteint
   12 caractères, cherche en base à quelle batterie cette cellule est affectée
   (table `cellule`, colonne `affectation_produit`) et pré-remplit le champ
   "n° série produit". → délègue maintenant à `_lookup_batterie_from_cellule`.
3. **`display_model_list_xxx`** : interroge `suivi_production` /
   `produit_voltr` pour lister les batteries qui ont terminé l'étape
   précédente mais pas encore celle-ci, alimente la listbox.
4. **`xxx_on_select_batt`** : clic sur la listbox → recopie le n° série
   sélectionné dans le champ produit. → délègue à `_fill_entry_from_listbox_selection`.
5. **`valider_xxx`** : vérifie les prérequis (`verif_etape_act_non_ok`,
   `_check_prereqs_and_warn`), fait les contrôles spécifiques à l'étape
   (impédance, tension, écarts...), puis `UPDATE suivi_production SET
   xxx = 1, date_xxx = NOW(), visa_xxx = %s ...`.
6. **`add_non_conf_batterie_xxx`** : bouton "❌ Non conforme" → ouvre (ou
   propose d'ouvrir) le formulaire Google de non-conformité et incrémente un
   compteur d'échec dédié (`xxx_fail`).

`tab_capa.py` (test de capacité / cyclage) est le seul onglet qui **ne suit
pas** ce schéma : au lieu d'une saisie manuelle, il **dépouille automatiquement
des fichiers Excel** déposés dans un dossier réseau (résultats de banc de
cyclage), voir §5.

---

## 5. Cas particulier : `tab_capa.py`

C'est le fichier le plus gros (~1000 lignes) et le seul qui n'a pas été
restructuré en profondeur, par prudence (voir §7). Sa méthode `_on_click`
(~760 lignes, désormais documentée par une docstring d'orientation) :

1. Récupère la liste des batteries en attente de test de capacité.
2. Parcourt tous les fichiers Excel du dossier réseau `CAPA_DOSSIER_ENTRANT`
   (nom de fichier = numéro de série + emplacement + modèle + type de cyclage,
   ex. `MC0002031068-4-3-7-INR18650MH1_A.0.1`).
3. Pour chaque fichier reconnu, va chercher les seuils attendus (feuille
   `Cyclage` du fichier Excel de config) selon le couple modèle/référence
   cellule, lit les mesures dans le fichier, calcule des indicateurs, décide
   OK/KO, met à jour la base (`_update_ok_in_db`).
4. Déplace le fichier traité vers `CAPA_DOSSIER_EXPLOITES` (si OK) ou
   `CAPA_DOSSIER_KO` (si KO), via `_move_processed_files`.

Il existe deux blocs de code très proches dans cette méthode (un pour la
gamme `PPTR018A`, un pour les autres modèles) qui répètent le même
enchaînement "récupérer ref_cell → lire les seuils → lire les mesures →
calculer les indicateurs". C'est le meilleur candidat à une factorisation
future (cf. §8), mais cela nécessite des tests avant d'y toucher : c'est une
grosse pièce logique, sans filet de sécurité actuellement.

---

## 5 bis. Fonctionnalité "Ordre de Fabrication" (OF)

Ajoutée lors d'une session ultérieure, à partir d'une seconde version de ton
script qui intégrait ce principe. Elle permet de suivre la progression de
fabrication **par lot (OF)** en plus du suivi par batterie individuelle.

### Modèle de données attendu côté MySQL
- `produit_voltr.n_of` : colonne indiquant à quel OF appartient chaque
  batterie.
- Table `ref_of` : une ligne par OF, avec au minimum `n_of`,
  `quantite_batterie` (quantité totale prévue dans l'OF) et
  `etat_fabrication` (0 = en cours, utilisé pour peupler les listes d'OF
  disponibles).

> ⚠️ Si ces éléments n'existent pas encore dans ta base, la fonctionnalité OF
> plantera au premier appel (erreurs SQL "colonne/table inconnue").

### Où la batterie est-elle affectée à un OF ?
**Uniquement à l'étape "pack"**. C'est le seul onglet où l'opérateur choisit
manuellement un OF (combobox `s_numero_of`, alimentée par `show_of_in_process`
= liste des OF avec `etat_fabrication = 0`). À la validation
(`valider_soudure_pack`), en plus des contrôles habituels, la batterie est
affectée à cet OF : `UPDATE produit_voltr SET n_of = %s WHERE
numero_serie_produit = %s`.

### Comment les autres onglets (nappe, bms, wrap, fermeture, fin de ligne) s'en servent
Ces onglets **affichent** la progression de l'OF mais n'en changent pas
l'affectation :
- Un combobox "N° of" est présent mais purement informatif (rempli
  automatiquement via `update_avancement`, pas de sélection manuelle
  effective).
- Un label "Avancement of : X/Y" affiche combien de batteries de l'OF en
  cours ont déjà validé cette étape, sur le total attendu
  (`ref_of.quantite_batterie`).
- Dès que le champ "n° série produit" atteint 9 caractères
  (`check_entry_length_batt`), ou après sélection d'une batterie dans la
  listbox, ou après une validation réussie, `of_avancement()` est appelé : il
  déduit l'OF de la batterie scannée (`produit_voltr.n_of`) et rafraîchit le
  label + le combobox pour cet OF.
- **`picking`** reçoit aussi le binding `check_entry_length_batt`, mais n'a pas
  de label/combobox associés : l'appel ne fait rien de visible (juste un
  `print` de diagnostic dans la console). C'est fidèle au comportement de ta
  version source ; à toi de voir si tu veux ajouter l'affichage sur cet onglet
  aussi.

### Vérification de cohérence OF, pour l'instant sur `bms` uniquement
`valider_bms` est le seul endroit qui appelle `check_of(num_batt, of_saisi)` :
il compare l'OF sélectionné dans le combobox de l'onglet à l'OF réellement
enregistré pour cette batterie en base. En cas d'écart, une boîte de dialogue
propose d'échanger cette batterie avec une autre (`handle_of_mismatch` →
`swap_of`), qui demande le numéro de série de remplacement, vérifie qu'il
appartient bien au bon OF, puis échange les deux affectations `n_of` en une
seule requête.

Ce contrôle n'existe **que sur bms** dans la version source que tu m'as
fournie — je l'ai porté tel quel, sans le dupliquer sur nappe/wrap/fermeture/
fin de ligne, pour ne pas inventer un comportement que tu n'as pas
explicitement validé ailleurs. Si tu veux ce même contrôle sur les autres
onglets, dis-le moi : c'est un ajout mécanique de 2 lignes par onglet une fois
qu'on sait où les insérer.

### ⚠️ Point à tester avec de vraies données
`check_of` compare `of_saisi` (une chaîne, issue de `combobox.get()`) à
`of_bdd` (la valeur brute retournée par MySQL pour `n_of`). Si cette colonne
est un entier en base, la comparaison `"12" != 12` sera **toujours vraie**,
et le message "OF différent" apparaîtra à chaque validation, même quand tout
est cohérent. Je n'ai pas pu vérifier le typage réel de ta base depuis cet
environnement (pas d'accès MySQL ici) : à tester en priorité. Si besoin, la
correction est d'une ligne (caster les deux côtés en `str(...)` avant de
comparer).

### Fichiers impactés par cet ajout
`mixin_common.py` (7 nouvelles méthodes + construction de `self.entry_widgets`),
`tab_picking.py`, `tab_pack.py`, `tab_nappe.py`, `tab_bms.py`, `tab_wrap.py`,
`tab_fermeture.py`, `tab_finligne.py`.

### Réglages fins apportés (cycle de vie du statut de l'OF)

`ref_of.etat_fabrication` est traité comme un statut texte à 3 valeurs :
**"en attente"** → **"en cours"** → **"terminé"**.

1. **Liste des OF filtrée par référence, OF terminés masqués**
   `show_of_in_process()` ne renvoie plus que les OF dont
   `ref_of.reference_batterie` correspond à la référence actuellement
   sélectionnée (`self.selected_model`), et exclut ceux déjà `'terminé'`.
   Cas particulier **PPTR018A** (Bosch/Makita/Ryobi/...) : comme la variante
   finale (AA/AB/AC/AD) n'est choisie qu'à la validation du pack, la liste des
   OF utilise ici aussi un `LIKE 'PPTR018A%'` plutôt qu'une égalité stricte —
   cohérent avec ce que fait déjà `display_model_list_pack` pour la liste des
   batteries de cette famille.

2. **Passage automatique à "en cours"**
   Dans `valider_soudure_pack` (seul endroit où une batterie est affectée à un
   OF), juste après l'affectation, une requête ne bascule le statut que s'il
   était encore `'en attente'` :
   ```sql
   UPDATE ref_of SET etat_fabrication = 'en cours'
   WHERE n_of = %s AND etat_fabrication = 'en attente'
   ```
   Sans effet si l'OF est déjà "en cours" ou "terminé" — donc sûr à exécuter
   à chaque batterie affectée, pas seulement la première.

3. **Passage automatique à "terminé"**
   Ajout d'un helper `_last_of_tracked_stage()` qui identifie, pour le modèle
   en cours, la dernière étape suivie par OF (parmi pack/nappe/bms/wrap/
   fermeture_batt/fin_ligne) dans l'ordre réel du flux (`self.ordered_keys`) —
   c'est-à-dire celle juste avant l'onglet recherche. `of_avancement()` (déjà
   appelée après chaque validation réussie) compare, pour cette étape
   uniquement, le nombre de batteries ayant validé à `ref_of.quantite_batterie` ;
   si tout est fait, l'OF passe à `'terminé'`.
   > Comme `_check_prereqs_and_warn` interdit de valider une étape tant que
   > les précédentes ne le sont pas, valider la dernière étape suffit à
   > garantir que toutes les étapes précédentes le sont aussi — pas besoin de
   > revérifier chaque étape individuellement.

---

## 6. Ce qui a été fait dans la session d'optimisation précédente

### 6.1 Factorisation de code dupliqué
- 8 onglets dupliquaient à l'identique (à la nomenclature des widgets près) la
  logique "retrouver la batterie à partir d'une cellule scannée" et "recopier
  une sélection de listbox dans un champ". Cette logique est désormais
  centralisée dans `mixin_common.py` (`_lookup_batterie_from_cellule`,
  `_fill_entry_from_listbox_selection`). **Gain : ~250 lignes en moins, un
  seul endroit à corriger/faire évoluer pour ces 16 méthodes.**
- Chemins réseau codés en dur et dupliqués (logo VoltR répété dans 3 fichiers,
  chemins de dossiers du module capa, chemin du template de remplacement
  cellule) → centralisés dans `config.py`.
- URL du formulaire Google de non-conformité, dupliquée **11 fois** dans le
  code → centralisée en une seule constante `NON_CONFORMITE_FORM_URL`.

### 6.2 Suppression de code mort
- `register_focus_target` était défini deux fois dans le même fichier : la
  deuxième définition écrasait silencieusement la première (même
  comportement au final, mais source de confusion). Doublon supprimé.
- `_move_processed_files` (dans `tab_capa.py`) avait une ancienne version
  laissée en commentaire (`"""..."""`) juste au-dessus de la version active.
  Supprimée.
- D'anciens chemins de dossiers de test (`C:/Users/User/Desktop/MAJ_TEST/...`)
  laissés en commentaire dans `tab_capa.py`. Supprimés.

### 6.3 Sécurité
- `db_manager.py` contenait, **en clair dans le code source**, deux jeux
  d'identifiants MySQL (un compte cloud et un compte `root` local) laissés en
  commentaire mais toujours présents dans le fichier. **Supprimés.** Seule
  reste la saisie interactive des identifiants au lancement (`simpledialog`),
  qui était déjà le comportement réellement actif.
  > ⚠️ Recommandation : si ces mots de passe (`VoltR99!`, compte `Vanvan`)
  > étaient encore valides sur le serveur `34.77.226.40`, il serait prudent de
  > les changer maintenant qu'ils ont pu circuler dans l'historique du fichier.

### 6.4 Bugs corrigés
Deux bugs de comportement réel (pas juste de structure) ont été identifiés et
corrigés, car les deux avaient un niveau de confiance élevé :

1. **`tab_nappe.py`** — la requête de non-conformité était :
   ```sql
   UPDATE suivi_production SET soudure_pack_fail = soudure_nappe_fail + 1 ...
   ```
   c'est-à-dire qu'elle lisait le compteur d'échec **nappe** pour l'incrémenter
   dans le compteur d'échec **pack** (copier-coller depuis `tab_pack.py` non
   corrigé). Corrigé en auto-incrément cohérent :
   ```sql
   UPDATE suivi_production SET soudure_nappe_fail = soudure_nappe_fail + 1 ...
   ```
   (conforme au schéma utilisé par tous les autres onglets).

2. **`tab_capa.py`** — deux appels à `messagebox.shoxinfo(...)`, méthode qui
   **n'existe pas** dans `tkinter.messagebox` (typo). Si ce chemin d'erreur
   avait été déclenché (échec de récupération de la liste des batteries),
   l'application aurait planté avec une `AttributeError` **au lieu** d'afficher
   le message d'erreur prévu. Corrigé en `messagebox.showerror(...)`.

### 6.5 Bug détecté mais **non corrigé** — à ta décision

Dans `tab_finligne.py`, le bouton "❌ Non conforme" (`add_non_conf_batterie_fl`)
exécute :
```sql
UPDATE suivi_production SET fermeture_fail = fermeture_fail + 1 ...
```
Il incrémente donc le compteur d'échec de l'étape **fermeture**, pas celui de
la **fin de ligne**. Contrairement au cas nappe/pack, je n'ai pas trouvé dans
le code de colonne du type `fin_ligne_fail` qui prouverait quelle est la
"bonne" cible — seulement `charge_fail` et `fonction_fail`, qui sont déjà
utilisés ailleurs dans le même fichier pour des échecs spécifiques (charge
KO / fonctionnel KO). Je n'ai donc **pas deviné** de nom de colonne au hasard,
pour ne pas introduire une erreur SQL sur une colonne inexistante.

→ **Action à faire de ton côté** : dis-moi quelle colonne `suivi_production`
doit recevoir ce compteur (`fin_ligne_fail` si elle existe, ou un autre nom),
et je corrige en 30 secondes.

---

## 7. Vérifications effectuées

- Tous les fichiers compilent sans erreur (`python -m py_compile`), à
  l'exception d'un avertissement bénin préexistant (regex avec échappement,
  `tab_capa.py`, sans impact fonctionnel).
- Recherche exhaustive d'autres noms de méthode incorrects sur les objets
  `messagebox`, `cursor`, `conn` (les plus utilisés dans tout le fichier) :
  aucune autre anomalie trouvée.
- **Limite de cet environnement** : ce sandbox n'a ni `tkinter`, ni
  `mysql-connector-python`, ni accès réseau — je n'ai donc pas pu lancer
  l'application "pour de vrai" avec une vraie base. Les vérifications ci-dessus
  sont donc **statiques** (compilation + relecture attentive + comparaison
  ligne-à-ligne des méthodes factorisées). **Recommandation : fais un test
  fumée rapide sur ton poste** (lancer l'app, tester un onglet de chaque
  "famille" — un onglet simple type wrap, et l'onglet capa qui est
  différent) avant de déployer en prod.

---

## 8. Dette technique restante (pour une prochaine session)

Par ordre d'impact décroissant :

1. **`tab_capa.py::_on_click`** (~760 lignes) : les deux branches
   PPTR018A / autres modèles sont structurellement très proches. Une
   factorisation ferait gagner ~300 lignes, mais nécessite d'écrire des tests
   (au moins quelques cas avec de vrais fichiers Excel de cyclage) avant d'y
   toucher, pour ne pas casser une logique de production sans filet.
2. **Absence de tests automatisés** sur l'ensemble du projet. Même un jeu de
   tests minimal sur `mixin_common.py` (règles de prérequis entre étapes,
   génération des requêtes SQL) sécuriserait beaucoup les évolutions futures.
3. **`add_non_conf_batterie_xxx`** : le comportement diffère d'un onglet à
   l'autre (emballage n'incrémente aucun compteur ; nappe/fin de ligne avaient
   des bugs de colonne ; les autres sont cohérents). Ça vaudrait le coup de
   décider d'un comportement standard une bonne fois pour toutes.
4. Les mots de passe MySQL sont demandés par une boîte de dialogue standard
   Tkinter à chaque lancement (`simpledialog.askstring`) : fonctionnel, mais
   sans gestion d'erreur si l'utilisateur annule/se trompe (l'app tente quand
   même la connexion avec des valeurs vides ou `None`).
5. Le `requirements.txt` liste les dépendances mais pas leurs versions
   exactes ; pour un exe reproductible, il serait plus sûr de figer les
   versions (`pip freeze` sur le poste où l'exe fonctionne actuellement).

---

## 9. Comment étendre l'application (ajouter une nouvelle étape)

1. Créer `tab_ma_nouvelle_etape.py` sur le modèle d'un onglet existant proche
   (ex. `tab_wrap.py`, le plus simple) avec une classe `TabMaNouvelleEtapeMixin`.
2. Ajouter la colonne SQL correspondante dans `STAGE_TO_DBCOL` (`config.py`).
3. Dans `mixin_common.py` :
   - ajouter l'entrée dans `column_to_stage` (méthode `_show_model_selector_and_build`),
   - ajouter l'entrée dans `stage_defs` (méthode `_create_widgets_with_order`),
   - si l'onglet doit se rafraîchir automatiquement, ajouter sa fonction
     d'affichage dans `stage_to_func` et `self.stage_refreshers`.
4. Dans `main.py`, importer `TabMaNouvelleEtapeMixin` et l'ajouter à la liste
   des classes parentes de `StockApp`.
5. Ajouter la colonne correspondante côté base MySQL si elle n'existe pas déjà,
   et une ligne dans la feuille Excel `Flux test modele` pour activer l'étape
   sur les références concernées.

---

## 10. Générer l'exécutable (.exe)

Le point d'entrée reste `main.py`. Depuis le dossier `suivi_prod/` :

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --noconsole main.py
```

L'exécutable sera généré dans `dist/main.exe`. Points d'attention :
- PyInstaller doit être lancé **sur Windows** pour produire un `.exe` Windows.
- `ttkthemes`, `PIL`/`Pillow`, `openpyxl` et `mysql-connector-python`
  embarquent parfois des données non détectées automatiquement par
  PyInstaller ; si l'exe se lance mais plante sur le thème `xpnative` ou sur
  une image, ajoute `--collect-all ttkthemes` et/ou `--collect-all PIL` à la
  commande.
- Les chemins réseau (`config.py`) supposent que le lecteur `G:` est monté sur
  le poste qui exécute l'exe (lecteur partagé Google Drive / Drive partagés).

---

## 11. Résumé chiffré

| Avant | Après |
|---|---|
| 1 fichier, 5 373 lignes | 18 fichiers, ~5 820 lignes (dont commentaires/docstrings et fonctionnalité OF ajoutés) |
| 2 fonctions dupliquées silencieusement écrasées | 0 |
| ~250 lignes dupliquées à l'identique sur 16 méthodes | Centralisées en 2 helpers communs |
| Mots de passe MySQL en clair dans le fichier | Supprimés |
| 1 URL et plusieurs chemins réseau dupliqués (jusqu'à 11×) | Centralisés dans `config.py` |
| 2 bugs de comportement silencieux (mauvaise colonne, méthode inexistante) | Corrigés |
| 1 bug de comportement suspect supplémentaire | Identifié, en attente de ta confirmation |
| Pas de suivi par lot | Suivi d'avancement par OF sur 6 onglets (pack, nappe, bms, wrap, fermeture, fin de ligne) |
