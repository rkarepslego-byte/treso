# 2. Guide utilisateur

Ce guide s'adresse à la personne qui utilise l'outil au quotidien. Vous n'avez que trois choses à
faire : **saisir** les opérations prévues, **déposer** les relevés bancaires, **actualiser**.

---

## Les onglets

| Onglet | À quoi il sert |
|---|---|
| **Tableau de bord** | Soldes réels et prévisionnels des deux banques, graphique, alertes, soldes jour par jour |
| **Saisie** | Saisir les dépenses et recettes prévues (formulaire + tableau) |
| **Charges fixes** | Loyer, salaires, abonnements... générés automatiquement chaque mois / trimestre / an |
| **Suivi prévisionnel** | Chronologie des opérations avec le solde après chaque opération, filtrable |
| **Synthèse mensuelle** | Flux par catégorie et par mois (réel passé, prévu à venir), solde de fin de mois |
| **Rapprochement** | Correspondances trouvées entre vos prévisions et la banque, à valider |
| **À traiter** | Opérations bancaires que vous n'aviez pas prévues |
| **Opérations bancaires** | Toutes les opérations importées des relevés (consultation) |
| **Paramètres** | Banques, soldes initiaux, catégories, réglages |
| **Aide** | Rappel de la procédure |

Les **cellules jaunes** se saisissent. Les **en-têtes bleus** signalent les colonnes modifiables,
les **en-têtes gris** les colonnes calculées par l'outil (protégées).

---

## 1. Saisir une opération prévue

Onglet **Saisie**, formulaire « Nouvelle opération » :

1. Date prévue, Banque, Libellé, Catégorie.
2. Type : **Dépense**, **Recette** ou **Virement interne**.
3. Montant **toujours positif** : l'outil enregistre lui-même les dépenses en négatif.
4. Pour un virement interne, choisir la **banque de destination** : deux lignes liées sont créées
   (sortie de la banque 1, entrée dans la banque 2). Un virement interne n'est compté **ni en dépense
   ni en recette** et ne change pas la trésorerie globale.
5. Cliquer **＋ Ajouter l'opération**.

Vous pouvez aussi écrire directement dans les **lignes vides en bas du tableau** (ou y coller
plusieurs lignes), puis cliquer **Valider les saisies du tableau** : l'outil attribue les
identifiants et complète les valeurs par défaut. Le tableau garde toujours 30 lignes vides : pour
coller davantage de lignes, collez-en 30, validez, puis recommencez.

**Reprendre un fichier existant** (ancien prévisionnel, liste de factures…) : bouton **Importer des
prévisions…**, choisir le fichier (Excel ou CSV), puis la banque par défaut. L'outil reconnaît les
colonnes d'après leurs titres (Date / Échéance, Libellé / Intitulé, Montant ou Débit / Crédit ou
Dépenses / Recettes, Banque, Catégorie / Rubrique, Type, Commentaire), affiche un récapitulatif et
demande confirmation. Les lignes déjà présentes, sans montant ou antérieures au solde initial sont
ignorées : l'import peut être relancé sans créer de doublon.

Autres boutons :
* **Trier par date** ;
* **Supprimer la ligne sélectionnée** (une échéance de charge fixe est marquée « Annulé » au lieu
  d'être supprimée ; une prévision pointée ne peut pas être supprimée) ;
* **Annuler le rapprochement** de la ligne sélectionnée.

La colonne **Suivi** indique : *Prévu*, *À contrôler* (date dépassée sans opération bancaire, ou
rapprochement en attente de validation), *Pointé* (réalisé) ou *Annulé*. La colonne **Contrôle**
signale en rouge une saisie incomplète (banque inconnue, date manquante...).

Pour annuler une prévision qui n'aura pas lieu : mettre son **Statut** à « Annulé ».

---

## 2. Charges fixes

Onglet **Charges fixes** : une ligne par charge récurrente (libellé, banque, catégorie, montant,
première échéance, fréquence mensuelle / trimestrielle / semestrielle / annuelle, date de fin
éventuelle, actif Oui / Non). Le type « Recette » permet aussi les recettes récurrentes.

Les échéances sont créées automatiquement dans l'onglet Saisie (12 mois à l'avance par défaut) à
chaque actualisation ou avec le bouton **Générer les échéances**.

Modifier une charge (hausse du loyer, changement de date) **ne modifie jamais** les échéances déjà
passées, pointées, annulées ou modifiées à la main : seules les échéances futures sont recalculées.
Pour arrêter une charge : Actif = « Non » ou date de fin.

---

## 3. Déposer les relevés bancaires

1. Sur le site de chaque banque, exporter les opérations au format **CSV ou Excel**.
2. Enregistrer le fichier dans le dossier **Imports\Banque1** (ou **Imports\Banque2**) situé à côté
   du classeur.

Inutile de supprimer les anciens relevés ni de faire attention aux périodes : **une opération déjà
importée n'est jamais ajoutée une seconde fois**, même si deux relevés se chevauchent. Deux paiements
identiques le même jour (deux péages, par exemple) restent bien deux opérations.

---

## 4. Actualiser

Tableau de bord, bouton **⟳ Actualiser les banques**. L'outil :

1. fait une **sauvegarde** du classeur (dossier Sauvegardes) ;
2. lit les relevés des deux banques et ajoute les **nouvelles** opérations ;
3. génère les échéances des charges fixes ;
4. recherche les **correspondances** entre prévisions et opérations bancaires.

Un message récapitule le résultat. Si un relevé est illisible, il est signalé, les autres sont
importés, et **rien de ce qui existait n'est supprimé**. Pensez à **enregistrer** le classeur.

---

## 5. Valider les rapprochements

Onglet **Rapprochement**. Une correspondance réunit une opération bancaire et une prévision de la
**même banque**, du **même montant**, avec au plus **7 jours** d'écart.

* **Certaine** : un seul candidat possible. Cliquer **Tout valider (certaines)**, ou écrire « Oui »
  dans la colonne *Valider*.
* **Incertaine** : plusieurs candidats (même numéro de *Groupe*). Écrire « Oui » sur **la** bonne
  ligne (la ligne « Suggérée » est la plus probable : libellé le plus proche, date la plus proche).
* Écrire « Non » pour rejeter une proposition (elle ne sera plus proposée).

Puis cliquer **Valider les choix**. La prévision passe à « Pointé » : elle n'est plus comptée dans le
prévisionnel, l'opération bancaire est dans le réel — **jamais de double comptage**.

En attendant votre validation, la prévision la plus probable est déjà neutralisée dans le
prévisionnel, pour que les soldes restent justes.

Pour corriger une erreur : sélectionner la ligne (onglet Saisie ou Opérations bancaires) puis
**Annuler le rapprochement**.

---

## 6. Traiter les opérations non prévues

Onglet **À traiter** : opérations bancaires sans prévision correspondante. Pour chacune :

1. vérifier la **Catégorie** (proposée automatiquement d'après le libellé) ;
2. ajouter un **Commentaire** si utile ;
3. choisir l'**Action** :
   * *Conserver comme non prévue* — dépense ou recette réelle non prévue ;
   * *Virement interne* — la colonne Suggestion signale les virements probables entre vos deux banques ;
   * *Rapprocher avec la prévision indiquée* — quand le montant diffère de la prévision (frais,
     arrondi) : saisir l'ID de la prévision (colonne ID de l'onglet Saisie) ;
   * *Créer charge mensuelle / trimestrielle / semestrielle / annuelle* — l'opération est conservée
     et une charge fixe est créée pour les prochaines échéances ;
4. cliquer **Appliquer les traitements**.

### Ce que l'outil fait tout seul

* **Il apprend vos traitements.** Quand vous conservez une opération non prévue avec une catégorie
  (ou la classez en virement interne), une règle est mémorisée (onglet Paramètres, « Règles
  automatiques ») : les opérations semblables des relevés suivants (frais de tenue de compte,
  abonnement, commission…) sont classées automatiquement et n'apparaissent plus dans « À traiter ».
  Vous pouvez modifier ou supprimer ces règles (vider la colonne « Action automatique » pour ne garder
  que la suggestion de catégorie).
* **Il ne prend jamais de risque** : si une prévision proche (même banque, libellé ressemblant, montant
  différent) existe, l'opération reste à traiter et la colonne Suggestion indique « Prévision proche »,
  avec son numéro déjà pré-rempli : il suffit de choisir l'action « Rapprocher avec la prévision indiquée ».
* **Il reconnaît les virements entre vos deux banques** non prévus (montants opposés, libellés de
  virement, quelques jours d'écart) et les classe automatiquement.
* **Il vous prévient** : un bandeau rouge en haut du tableau de bord signale si une banque va passer
  en négatif ou si la trésorerie globale va passer sous le seuil d'alerte (Paramètres) dans les
  90 jours ; le message d'actualisation le rappelle, ainsi que les prévisions échues depuis longtemps.
* **Il propose d'actualiser** dès l'ouverture du classeur quand de nouveaux relevés ont été déposés.

Chaque automatisme se désactive dans **Paramètres > Automatisation**. Une opération classée
automatiquement porte le mode « Automatique » (onglet Opérations bancaires) et peut être remise à
traiter avec **Annuler le rapprochement / traitement**.

---

## 7. Lire le tableau de bord

* **Soldes réels** : solde initial + opérations importées, par banque et global, avec la date du
  dernier relevé.
* **Soldes prévisionnels** : solde réel + prévisions non encore réalisées jusqu'à J+30, J+60 ou J+90
  (boutons **30 j / 60 j / 90 j**). Les prévisions échues non pointées restent comptées : vérifiez-les.
* **À surveiller** : dépenses à venir, recettes attendues (hors virements internes), **point bas**
  de trésorerie, opérations en attente, opérations non prévues, prévisions à contrôler.
* **Graphique** et **soldes jour par jour** sur l'horizon choisi ; tableau de projection
  30 / 60 / 90 jours.

Onglet **Synthèse mensuelle** : 5 mois passés (réel), le mois en cours (réel + prévu) et 6 mois à
venir (prévu), par catégorie, avec le flux net et le solde de fin de chaque mois ; filtre par banque.

Onglet **Suivi prévisionnel** : filtrer par banque, période et type (tout / prévisions / réel). Les
lignes en italique bleu sont des prévisions ; les soldes négatifs sont en rouge.

---

## Questions fréquentes

**Une opération n'est pas proposée alors que je l'avais prévue.** Vérifiez la banque, le montant
exact et la date (écart de 7 jours au plus). Sinon, utilisez l'action « Rapprocher avec la prévision
indiquée » dans l'onglet À traiter.

**J'ai modifié le nom d'une banque.** À la prochaine actualisation, l'outil propose de mettre à jour
toutes les opérations existantes.

**Les boutons ne réagissent pas.** Les macros sont désactivées : fermez et rouvrez le classeur en
cliquant « Activer le contenu ».

**Revenir en arrière.** Chaque actualisation crée une copie datée dans le dossier Sauvegardes.
