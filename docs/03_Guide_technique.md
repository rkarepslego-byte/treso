# 3. Guide technique (administrateur)

## 3.1 Architecture

```
Relevés CSV/XLSX          Power Query                 VBA                         Formules Excel
Imports\Banque1  ──►  Import_Banque1 ─► tbl_Stage_B1 ─┐
Imports\Banque2  ──►  Import_Banque2 ─► tbl_Stage_B2 ─┤  modImport : ajoute les    ┌► soldes réels
                      (onglets masqués, recalculés    │  clés inconnues au        │  soldes prévisionnels
                       à chaque actualisation)        ▼  journal permanent        │  jour par jour, KPI
                                                   tbl_Banque (jamais vidé) ──────┤  suivi chronologique
Saisie / Charges fixes ─► tbl_Previsions ◄── modPrevisions (échéances)            │
                               ▲  │                                               │
                               └──┴── modRapprochement / modATraiter (liens) ─────┘
```

Principes :

* **Power Query ne fait que lire et normaliser** les relevés (tables de transit recalculées à chaque
  actualisation). Il ne détient aucune donnée durable.
* **Le journal bancaire `tbl_Banque` est permanent** : le VBA y ajoute uniquement les lignes dont la
  clé technique est inconnue. Il n'est jamais vidé (sauf `Admin_ViderDonnees`). Supprimer un relevé du
  dossier ne supprime donc aucune opération ; une erreur de lecture n'efface rien.
* **Les rapprochements sont stockés dans les tables** (`ID prévision` dans `tbl_Banque`,
  `ID opération` dans `tbl_Previsions`) : ils sont conservés d'une actualisation à l'autre.
* **Tous les soldes sont calculés par des formules** (et non par macro) : ils sont toujours à jour,
  vérifiables, et indépendants du VBA.
* **Les paramètres d'import sont transmis à Power Query par le VBA** sous forme d'un record M
  littéral (requête `Parametres_Import`). Les requêtes n'utilisent donc qu'une seule source (le
  dossier) : aucun conflit de niveaux de confidentialité (« Formula.Firewall ») ni question de
  confidentialité.
* Aucune connexion bancaire, API ou service cloud : tout reste sur le poste.

## 3.2 Onglets et tables

| Onglet | Visible | Contenu |
|---|---|---|
| Tableau de bord | oui | cartes KPI, graphique, projection, soldes jour par jour (`Horizon_Jours`) |
| Saisie | oui | formulaire (`Form_*`) + `tbl_Previsions` |
| Charges fixes | oui | `tbl_Charges` |
| Suivi prévisionnel | oui | vue chronologique (1 000 lignes de formules) + filtres `Suivi_*` |
| Synthèse mensuelle | oui | flux par catégorie × 12 mois (réel / réel + prévu / prévu), solde de fin de mois, filtre `Synth_Banque` |
| Rapprochement | oui | `tbl_Propositions` (écrite par le moteur) |
| À traiter | oui | `tbl_ATraiter` (écrite par le moteur) |
| Opérations bancaires | oui | `tbl_Banque` (journal permanent) |
| Paramètres | oui | paramètres nommés, `tbl_Categories`, `tbl_Regles` |
| Aide | oui | mode d'emploi |
| Import_B1 / Import_B2 | très masqué | `tbl_Stage_B1/B2` (sorties Power Query) |
| Historique | très masqué | `tbl_Historique` (une ligne par fichier et par import) |
| Calculs | très masqué | KPI nommés, série journalière J+0..J+90 (`Calc_*`) |
| Listes | très masqué | listes déroulantes `L_*`, `tbl_Rejets`, noms précédents des banques |
| Param_Import | très masqué | `tbl_ParamImport` (formats bancaires), `tbl_CodeM` (code M de secours) |

« Très masqué » : invisible depuis Excel ; `Admin_AfficherOngletsTechniques` les affiche.

### Colonnes principales

**tbl_Previsions** — saisie : Date prévue, Banque, Libellé, Catégorie, Type (Dépense / Recette /
Virement sortant / Virement entrant), Montant (positif), Commentaire, Statut (Prévu / Annulé ;
« Pointé » est posé par l'outil). Gérées par l'outil : ID (P-000001), Origine, Réf charge, ID opération,
ID proposé, Date réelle, Montant réel, Lien virement, Empreinte. Calculées : Montant signé, Suivi,
Contrôle, Montant à venir, Date effective, Clé tri, Inclus banque, Inclus vue, Rang vue.

**tbl_Banque** — ID (OB-000001), Banque, Date, Date valeur, Libellé, Montant, Référence, Catégorie,
Commentaire, Statut (À traiter / Proposé / Pointé / Non prévue / Virement interne), ID prévision,
Mode (Validé / Automatique / Manuel), ID import, Fichier, Date import, Clé ; calculées : Inclus,
Clé tri, Inclus banque, Inclus vue, Rang vue.

**tbl_Charges** — ID (CF-001), Libellé, Banque, Catégorie, Type, Montant, Première échéance,
Fréquence, Date de fin, Actif, Commentaire ; calculées : Pas, Prochaine échéance, Équivalent mensuel,
Échéances à venir.

Les tables des onglets protégés ne peuvent pas s'agrandir d'elles-mêmes (limite d'Excel) : 30 lignes
de saisie libres sont maintenues en bas de `tbl_Previsions` et `tbl_Charges`, et les macros
agrandissent les tables au besoin.

## 3.3 Règles de calcul

* **Solde réel** d'une banque = solde initial + Σ montants de `tbl_Banque` dont la date est
  **strictement postérieure** à la date du solde initial. Le solde initial est réputé inclure toutes
  les opérations jusqu'à sa date incluse : ces opérations sont ignorées à l'import (comptées
  « antérieures » dans l'historique) et exclues des formules.
* **Montant à venir** d'une prévision = montant signé si elle n'est ni pointée, ni annulée, ni liée à
  une opération, ni « proposée », et datée après le solde initial de sa banque ; sinon 0.
* **Solde prévisionnel à J+n** = solde réel + Σ montants à venir datés jusqu'à J+n (les prévisions
  échues non pointées restent dues : elles sont comptées et signalées « À contrôler »).
* **Virements internes** : deux prévisions liées (sortant / entrant). Ils sont exclus des totaux
  « dépenses à venir » / « recettes attendues » et s'annulent dans la trésorerie globale.
* **Pas de double comptage** : dès qu'une correspondance est proposée, la prévision la plus probable
  est neutralisée (`ID proposé`) ; après validation elle est « Pointé ». L'opération bancaire, elle,
  est dans le réel dès son import.
* **Suivi prévisionnel** : opérations bancaires incluses puis prévisions restantes, triées par date
  effective (prévision échue → date du jour), solde cumulé depuis le solde initial (filtre banque
  appliqué), indépendamment du filtre de période.
* `Date_Reference` = `Paramètres > Date de calcul` si renseignée, sinon AUJOURDHUI().

## 3.4 Import bancaire (Power Query)

Requêtes (sources dans `src/powerquery/`) :

| Requête | Rôle |
|---|---|
| `Parametres_Import` | record généré par le VBA (`modImport.FormuleParametresM`) depuis `tbl_ParamImport` |
| `fnTexteNormalise` | majuscules, sans accents, A-Z/0-9 |
| `fnLireMontant` | « 1 234,56 », « -1.234,56 », « (12,00) », « 12,00 € », nombres Excel |
| `fnLireDate` | dates Excel, numéros de série, textes au format paramétré |
| `fnNormaliserReleve` | **un fichier** → Date, DateValeur, Libelle, Montant, Reference, Ligne |
| `fnImporterBanque` | **un dossier** → table de transit, occurrence, clé technique, erreurs par fichier |
| `Import_Banque1` / `Import_Banque2` | `fnImporterBanque(Parametres_Import[BanqueN])`, chargées dans `tbl_Stage_B1/B2` |

### Adapter un vrai format d'export (quand les fichiers des banques seront fournis)

0. **Analyse automatique** (poste de développement) :
   ```
   python -m outils.analyser_releve chemin\vers\export_banque1.csv --banque "Nom de la banque"
   ```
   L'outil détecte l'encodage, le séparateur, la ligne d'en-tête, les colonnes, le format de date et
   le séparateur décimal, propose les valeurs de `tbl_ParamImport`, relit le fichier avec ces
   paramètres (même logique que Power Query) et liste les points à vérifier : lignes écartées, montants
   tous positifs, référence instable (simple numéro de ligne), lignes identiques dans le fichier.
1. `Admin_AfficherOngletsTechniques`, onglet **Param_Import**, table `tbl_ParamImport` (colonne de la
   banque concernée) :

| Paramètre | Exemple | Remarque |
|---|---|---|
| Extensions acceptées | `csv;txt` | |
| Séparateur CSV | `;` `,` ou `TAB` | |
| Encodage CSV | `1252` (Windows) / `65001` (UTF-8) | accents illisibles → changer l'encodage |
| Lignes à ignorer avant l'en-tête | `3` | lignes « Compte n°… », « Période… » |
| Feuille Excel | vide = 1re feuille | |
| Colonne date d'opération | `Date opération` | casse, accents et espaces indifférents |
| Colonne date de valeur | facultatif | |
| Format de date | `dd/MM/yyyy`, `dd/MM/yy`, `yyyy-MM-dd` | si les dates sont du texte |
| Colonnes libellé | `Libellé;Détail` | plusieurs colonnes concaténées |
| Mode montant | `Montant signé` / `Débit/Crédit` / `Montant + Sens` | |
| Colonne montant / débit / crédit / sens | | selon le mode |
| Valeurs signifiant un débit | `D;DEBIT` | mode Montant + Sens |
| Inverser le signe | `Oui` si les débits sont exportés en positif | mode Montant signé |
| Séparateur décimal | `,` ou `.` | |
| Colonne référence | facultatif | n° d'opération unique : renforce la clé anti-doublon |

2. Déposer un relevé réel dans le dossier et cliquer **Actualiser les banques**. Un fichier
   incompatible produit un message explicite (« La colonne « Date opération » est absente du fichier
   X. Colonnes trouvées : … ») sans bloquer l'autre banque ni modifier les données existantes.
3. Si un format sort du cadre de ces paramètres (en-têtes sur deux lignes, montants dans une colonne
   texte composite…), adapter la fonction `fnNormaliserReleve` (Données > Requêtes et connexions),
   reporter la modification dans `src/powerquery/fnNormaliserReleve.pq` et dans la table `tbl_CodeM`.

> **Attention** : la clé technique dépend du libellé normalisé. Changer les colonnes de libellé
> après des imports peut faire réapparaître d'anciennes opérations encore présentes dans le dossier.
> Avant de modifier le paramétrage, déplacer les relevés déjà importés hors du dossier. L'outil
> signale de toute façon un fichier déjà importé qui produit de nouvelles lignes.

### Détection des doublons

Clé technique = `code banque | date | montant en centimes | libellé normalisé | référence | occurrence`.
`occurrence` = rang de la ligne parmi les lignes **strictement identiques d'un même fichier**.
Deux paiements identiques le même jour donnent les occurrences 1 et 2 (deux opérations) ; un relevé
qui chevauche le précédent redonne les mêmes clés (aucun ajout). Le code banque (`B1`/`B2`) et non le
nom figure dans la clé : renommer une banque ne provoque pas de réimport. Aucune opération n'est
jamais supprimée automatiquement.

## 3.5 Rapprochement

Arête candidate entre une opération (statut À traiter / Proposé, postérieure au solde initial) et une
prévision (non pointée, non annulée, non liée) si : **même banque**, **même montant au centime**,
**|écart de dates| ≤ tolérance** (7 jours), paire non rejetée (`tbl_Rejets`).

* Similarité des libellés (0 à 1) = max(Dice sur les bigrammes de caractères des mots significatifs,
  part des mots du libellé le plus court retrouvés dans l'autre). Mots ignorés : PRLV, SEPA, VIR, CB…
* Score = 0,5 × (1 − écart / (tolérance + 1)) + 0,5 × similarité.
* Composantes connexes du graphe des arêtes :
  * 1 opération – 1 prévision : **Certaine** si écart ≤ 3 jours **ou** similarité ≥ 40 %, sinon
    Incertaine ;
  * plusieurs candidats : **Incertaine** ; le moteur choisit l'appariement le plus probable (score
    décroissant) marqué « Suggérée ».
* `Paramètres > Valider automatiquement les correspondances certaines` (Non par défaut). Les
  correspondances incertaines ne sont **jamais** validées automatiquement.
* Validation : un même ID ne peut être validé deux fois dans un même lot (sinon rien n'est enregistré).
* Rapprochement manuel (montants différents) depuis l'onglet À traiter.

## 3.6 Charges fixes (sans effet rétroactif)

À chaque génération (actualisation ou bouton) :

* une échéance **figée** (pointée, annulée, proposée, passée, ou modifiée à la main — détectée par
  l'`Empreinte` date|banque|libellé|type|montant enregistrée à la création) n'est jamais modifiée ni
  supprimée et « couvre » son créneau (± 14 j en mensuel, ± 44 j en trimestriel…) ;
* une échéance non figée est conservée (même ID) si elle correspond exactement à la définition
  actuelle, sinon supprimée ;
* les échéances manquantes sont créées de max(première échéance, solde initial + 1, aujourd'hui −
  tolérance) jusqu'à min(date de fin, aujourd'hui + horizon). Les numéros d'ID ne sont jamais réutilisés.

## 3.7 Code VBA (`src/vba/`)

| Module | Contenu |
|---|---|
| `modConfig` | constantes : onglets, tables, statuts, **mot de passe** (`MOT_DE_PASSE`) |
| `modOutils` | accès aux tables (lecture / réécriture en bloc), protection, normalisation, similarité, dossiers, mode silencieux |
| `modImport` | `Actualiser`, sauvegarde, paramètres Power Query, intégration des relevés, historique |
| `modPrevisions` | formulaire de saisie, virements, suppression, tri, génération des échéances |
| `modRapprochement` | moteur, validation, rejet, annulation |
| `modATraiter` | liste des opérations non prévues et actions |
| `modTableauDeBord` | horizon 30/60/90, navigation, filtres |
| `modAdmin` | fonctions d'administration, installation des requêtes, intégrité, remise à zéro |
| `modRecette` | recette automatique (`install/Recette.ps1`) |
| `ThisWorkbook` | à l'ouverture : protection, boutons, horizon, contrôle des requêtes |

Conventions : toutes les macros publiques encadrent leur travail par `DebutTraitement` /
`FinTraitement` (calcul manuel, écran figé, feuilles déprotégées) et un gestionnaire d'erreur qui
rétablit Excel puis affiche un message en français. Les tables sont lues en tableau mémoire et
réécrites en bloc (`EcrireTable`), les colonnes calculées étant reconstituées.

### Macros d'administration (Alt+F8, mot de passe = `MOT_DE_PASSE`, « Treso2026 » par défaut)

| Macro | Effet |
|---|---|
| `Admin_AfficherOngletsTechniques` / `Admin_MasquerOngletsTechniques` | affiche / masque les onglets techniques |
| `Admin_DeprotegerFeuilles` / `Admin_ProtegerFeuilles` | protection des feuilles |
| `Admin_ReinstallerRequetes` | recrée les requêtes Power Query depuis `tbl_CodeM` et relie les tables de transit |
| `Admin_VerifierIntegrite` | identifiants en double, liens incohérents, virements sans contrepartie |
| `Admin_ViderDonnees` | supprime les données (après sauvegarde) ; conserve les paramètres |

**Changer le mot de passe** : modifier `MOT_DE_PASSE` dans `src/vba/modConfig.bas` **et**
`MOT_DE_PASSE` dans `build/schema.py`, puis régénérer le classeur (ou, dans un classeur déjà en
service : `Admin_DeprotegerFeuilles` avec l'ancien mot de passe, modifier la constante dans l'éditeur
VBA, puis `Admin_ProtegerFeuilles`).

## 3.8 Mise en service chez le client

1. Installer (docs/01_Installation.md) et passer la recette (`install\Recette.cmd`).
2. `Admin_ViderDonnees`.
3. Paramètres : société, banques, **soldes initiaux = solde du dernier relevé, date = date de ce
   relevé** (les opérations jusqu'à cette date incluse ne seront pas importées), catégories.
4. Adapter `tbl_ParamImport` aux exports réels (§ 3.4), déposer un premier relevé, actualiser,
   contrôler le solde réel affiché par rapport au relevé bancaire.
5. Saisir les charges fixes et les prévisions (ou les coller depuis l'ancien fichier dans les lignes
   libres de l'onglet Saisie, puis « Valider les saisies du tableau »).

## 3.9 Génération du classeur et tests (poste de développement)

Le classeur est **produit par programme** (Python 3.11+, XlsxWriter, openpyxl) :

```
python -m build.generer_classeur --date 2026-11-25      # -> dist/
python -m pytest tests                                    # tests (LibreOffice requis pour certains)
```

* `build/schema.py` : structure (tables, colonnes, formules, paramètres) — source unique de vérité ;
* `build/generer_classeur.py` : onglets, mise en forme, formules, graphique, boutons ;
* `build/vba_project.py` : compile `src/vba` en `vbaProject.bin` ([MS-CFB], [MS-OVBA] ; code source
  seul, Excel le compile à la première ouverture) ;
* `build/datamashup.py` + `build/integration.py` : intègre les requêtes Power Query ([MS-QDEFF]) et
  relie les tables de transit (connexions, queryTables) ;
* `build/donnees_demo.py` : jeu de données fictif et relevés de démonstration ;
* `outils/analyser_releve.py` : analyse d'un export bancaire réel et proposition des paramètres d'import ;
* `reference/moteur.py` : **moteur de référence** Python, miroir exact du VBA / Power Query /
  formules, qui sert de spécification exécutable.

Toute modification de la logique doit être faite à l'identique dans le VBA **et** dans
`reference/moteur.py`, puis vérifiée par les tests et la recette Windows.

## 3.10 Limites connues

* Windows et Excel bureau uniquement (pas Mac, pas Excel Online).
* Volumétrie mesurée : avec 3 ans d'activité (6 000 opérations bancaires, 3 000 prévisions), un
  recalcul complet prend environ 2,5 s dans LibreOffice (Excel, multi-cœur, est en général plus
  rapide). Les calculs coûteux (rang et solde cumulé de la vue Suivi) ne portent que sur la période
  filtrée. Au-delà de plusieurs années, archiver : nouveau classeur avec solde initial = solde du
  dernier relevé.
* Les tables sont dimensionnées pour quelques dizaines de milliers de lignes ; la vue Suivi affiche
  1 000 lignes (un message invite à réduire la période au-delà). Clé de tri prévue pour < 50 000
  lignes par table.
* Une échéance « divisée » en deux paiements se traite en scindant la prévision, puis en rapprochant
  chaque part (le rapprochement est un-pour-un).
* Les exports `.xls` nécessitent le moteur Access Database Engine.
* La protection des feuilles évite les erreurs de manipulation ; ce n'est pas une sécurité forte.
