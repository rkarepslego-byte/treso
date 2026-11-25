# Outil de gestion de trésorerie (Excel 365 · Power Query · VBA)

Outil de trésorerie pour une entreprise disposant de **deux comptes bancaires** : saisie des
prévisions, charges fixes récurrentes, import des relevés bancaires (CSV / Excel) **sans connexion
bancaire**, rapprochement automatique, soldes réels et prévisionnels jour par jour.

> Version 1.0 — première version fonctionnelle sur **données fictives**. Les formats d'import sont
> paramétrables et seront ajustés quand les vrais exports des deux banques seront fournis
> (voir [guide technique § 3.4](docs/03_Guide_technique.md#34-import-bancaire-power-query)).

## Livrables

| # | Livrable | Emplacement |
|---|---|---|
| 1 | Fichier Excel opérationnel | [`dist/Tresorerie.xlsm`](dist/Tresorerie.xlsm) |
| 2 | Requêtes Power Query | [`src/powerquery/`](src/powerquery) (intégrées au classeur) |
| 3 | Code VBA | [`src/vba/`](src/vba) (intégré au classeur) |
| 4 | Instructions d'installation | [`docs/01_Installation.md`](docs/01_Installation.md), [`install/`](install) |
| 5 | Guide utilisateur | [`docs/02_Guide_utilisateur.md`](docs/02_Guide_utilisateur.md) |
| 6 | Guide technique (administrateur) | [`docs/03_Guide_technique.md`](docs/03_Guide_technique.md) |
| 7 | Jeu de données fictives | dans le classeur + [`dist/Imports/`](dist/Imports) ; générateur [`build/donnees_demo.py`](build/donnees_demo.py) |
| + | Plan de tests et résultats | [`docs/04_Plan_de_tests.md`](docs/04_Plan_de_tests.md), [`tests/`](tests) |

## Démarrage rapide

1. Copier le dossier `dist` sur le poste Windows, lancer `install\Installer.cmd`.
2. Lancer `install\Recette.cmd` : vérifie dans Excel toute la chaîne (doit afficher RECETTE RÉUSSIE).
3. Ouvrir `Tresorerie.xlsm`, activer les macros, cliquer **Actualiser les banques**.

## Fonctionnalités

* **Tableau de bord** : soldes réels et prévisionnels par banque et consolidés, horizon 30/60/90 jours,
  graphique, soldes jour par jour, point bas, dépenses à venir, recettes attendues, opérations en
  attente, dépenses non prévues.
* **Synthèse mensuelle** par catégorie (réel passé / prévu à venir, solde de fin de mois).
* **Saisie** par formulaire ou tableau ; dépenses enregistrées en négatif ; virements internes liés,
  neutres sur la trésorerie globale.
* **Charges fixes** mensuelles / trimestrielles / semestrielles / annuelles, générées automatiquement,
  sans effet rétroactif des modifications.
* **Import Power Query** de deux banques indépendantes, formats paramétrables (séparateur, encodage,
  colonnes, débit/crédit, format de date…), fichiers illisibles signalés sans bloquer le reste.
* **Anti-doublons** par clé technique avec numéro d'occurrence : relevés répétés ou qui se chevauchent
  sans double import, paiements identiques légitimes conservés, historique des imports.
* **Rapprochement** banque / montant exact / tolérance de 7 jours / similarité de libellé : correspondances
  certaines (validation en un clic, automatique en option), incertaines (choix manuel obligatoire),
  opérations non prévues à traiter (catégorie, commentaire, virement interne, rapprochement manuel,
  création de charge récurrente). Rapprochements conservés entre deux actualisations.
* **Automatisations** : règles apprises des traitements manuels (frais, abonnements… classés tout
  seuls), virements internes non prévus détectés, suggestion de prévision proche, alertes de
  trésorerie (banque négative, seuil global), actualisation proposée à l'ouverture si nouveaux relevés.
* **Sécurité** : sauvegarde avant chaque actualisation, journal bancaire jamais vidé, feuilles protégées,
  onglets techniques masqués, messages d'erreur en français.

## Organisation du dépôt

```
dist/         classeur livré + relevés de démonstration + valeurs de recette
src/          sources Power Query (.pq) et VBA (.bas/.cls)
build/        générateur du classeur (Python)
reference/    moteur de référence (spécification exécutable de la logique)
outils/       analyse d'un export bancaire réel -> paramètres d'import proposés
tests/        tests automatisés (scénarios, formules via LibreOffice, intégrité)
install/      installation et recette Windows (PowerShell)
docs/         documentation
```

Régénérer le classeur : `python -m build.generer_classeur` · lancer les tests : `python -m pytest tests`
(Python 3.11+, `xlsxwriter`, `openpyxl`, `olefile`, `oletools`, `pytest` ; LibreOffice pour les tests
de formules).
