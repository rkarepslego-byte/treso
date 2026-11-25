# 4. Plan de tests et résultats

## 4.1 Stratégie

L'outil associe trois technologies qu'aucun environnement hors Windows ne sait exécuter ensemble
(Excel, VBA, Power Query). La validation est donc organisée en couches :

| Couche | Comment elle est vérifiée | Où |
|---|---|---|
| Logique métier (import, doublons, charges fixes, rapprochement, soldes) | **Moteur de référence** Python exécuté sur les scénarios ci-dessous, montants attendus écrits à la main | `tests/test_scenarios.py` (26 tests) |
| Formules du classeur | Classeur généré dans un état réaliste (import + rapprochements + traitements), **recalculé par LibreOffice**, chaque solde comparé au moteur de référence (au centime) | `tests/test_formules_classeur.py` (8 tests) |
| Projet VBA | Contrôle statique (blocs, variables non déclarées, appels, colonnes et noms utilisés vs structure réelle, limites VBA) ; `vbaProject.bin` relu par 3 lecteurs indépendants (olefile, olevba, LibreOffice) avec code identique aux sources | `tests/test_integrite_classeur.py` |
| Power Query | Syntaxe validée par l'analyseur officiel de Microsoft (`@microsoft/powerquery-parser`) sur les sources et sur la section M intégrée au classeur ; logique de lecture reproduite par le moteur de référence | `tests/test_integrite_classeur.py` |
| Fichier OpenXML | Parties, relations, types de contenu, connexions et queryTables cohérents | `tests/test_integrite_classeur.py` |
| **Chaîne complète dans Excel** | **Recette automatique Windows** : Power Query + VBA exécutés dans Excel sur les relevés de démonstration, 18 indicateurs comparés aux valeurs attendues du moteur de référence, contrôle de l'absence de doublon au réimport | `install/Recette.ps1` (à lancer sur le poste) |

Commande (poste de développement) : `python -m pytest tests` → **46 tests réussis** (dont 5 pour l'analyse automatique des exports bancaires et 1 pour le classeur vide de mise en service).

## 4.2 Scénarios métier (cahier des charges, phase 7)

Banque 1 : solde initial 10 000 € au 31/10/2026 ; Banque 2 : 5 000 € ; date de calcul 25/11/2026.

| # | Scénario | Vérifications | Résultat |
|---|---|---|---|
| 1 | Paiement prévu puis débité | Loyer prévu 1 500 € le 05/11, prélèvement « PRLV SEPA SCI DU PARC LOYER NOV » le 06/11 : correspondance **certaine** ; tant qu'elle n'est pas validée la prévision est neutralisée (prévisionnel 8 500 € et non 7 000 €) ; après validation : prévision « Pointé », catégorie reprise, 8 500 € ; réimport du relevé : rien d'ajouté ni de reproposé ; annulation du rapprochement possible | ✅ |
| 2 | Dépense non prévue | CB Amazon 89,90 € → liste « à traiter » ; « Conserver comme non prévue » avec catégorie et commentaire ; solde réel 9 910,10 € ; variante : création d'une charge mensuelle (12 échéances, 1re le mois suivant) | ✅ |
| 3 | Paiement identique à deux dates | Abonnement 200 € prévu le 03/11 et le 03/12, débit le 04/11 : seule l'échéance de novembre est proposée ; variante : deux prévisions dans la tolérance → **incertaine**, jamais validée automatiquement, une seule prévision neutralisée (pas de double comptage) | ✅ |
| 4 | Deux paiements identiques légitimes | Deux « CARTE TOTAL STATION » 45 € le même jour dans le même relevé : 2 opérations (occurrences 1 et 2) ; réimport : 0 ajout, 2 doublons ignorés ; une seule prévision → correspondance incertaine, l'autre opération reste à traiter | ✅ |
| 5 | Virement entre les deux banques | Virement prévu 5 000 € : 2 lignes liées (−/+), exclu des dépenses et recettes, trésorerie globale inchangée (15 000 €) avant et après import des deux côtés ; variante non prévue : suggestion « virement interne probable » et traitement | ✅ |
| 6 | Charge fixe mensuelle | Loyer le 31 de chaque mois : 30/11, 31/12, 31/01, **28/02**… (12 échéances) ; génération idempotente (mêmes ID) ; trimestrielle, semestrielle, annuelle, date de fin, charge inactive | ✅ |
| 7 | Modification d'une charge fixe | Échéance de novembre pointée ; hausse de 1 200 € à 1 300 € : novembre inchangé, décembre à 1 300 €, échéance modifiée à la main conservée, échéance annulée non recréée ; changement du jour d'échéance : pas de doublon sur le mois déjà réalisé ; désactivation : seules les échéances figées restent | ✅ |
| 8 | Import répété du même relevé | 4 imports successifs du même fichier : 3 opérations, solde 8 622,50 €, historique de 4 imports | ✅ |
| 9 | Relevés qui se chevauchent | Relevés du 02 au 13/11 et du 12 au 20/11 (avec deux péages identiques dans la zone commune) : 5 opérations, solde exact, quel que soit l'ordre ; un 3e péage identique dans le second relevé est ajouté (1 seul) | ✅ |
| 10 | Libellé bancaire différent | « Assurance locaux » vs « PRLV SEPA AXA FRANCE IARD CONTRAT 1234 » : similarité < 40 % mais écart de 2 jours → certaine ; la similarité départage deux prévisions de même montant (« Abonnement Orange » suggérée pour « PRLV SEPA ORANGE SA ») ; correspondance unique mais lointaine et dissemblable → incertaine | ✅ |

Règles transverses également testées : opérations antérieures à la date du solde initial ignorées,
prévision échue restant due et « à contrôler », rejet d'une proposition, validation contradictoire
refusée sans rien enregistrer, validation automatique des seules correspondances certaines,
rapprochement manuel d'un montant différent, fichier au mauvais format signalé sans bloquer les autres.

## 4.3 Formules (LibreOffice)

Le classeur de démonstration, après import des relevés, validation de correspondances et traitement
d'opérations, est recalculé par LibreOffice puis comparé au moteur de référence :

* aucune cellule en erreur dans tout le classeur ;
* soldes réels B1 / B2 / global ; soldes prévisionnels à 30, 60 et 90 jours ;
* 91 soldes journaliers par banque et global ; point bas et sa date ;
* dépenses à venir, recettes attendues, opérations à traiter / proposées, prévisions à contrôler,
  dépenses non prévues ;
* colonnes calculées de chaque prévision (montant signé, suivi, montant à venir), lignes libres vides ;
* **vue chronologique** : lignes, ordre et solde après chaque opération identiques au calcul attendu ;
* prochaine échéance de chaque charge fixe.

## 4.4 Recette sur le poste Windows du client

À exécuter une fois après installation (et après toute mise à jour du code) :

```
install\Recette.cmd
```

Résultat attendu : tableau « Indicateur / Attendu / Obtenu » entièrement « OK », message
« Réimport des mêmes relevés : aucun doublon ajouté », puis **RECETTE RÉUSSIE**.

Recette manuelle complémentaire (5 minutes, sur la démonstration) :

| Étape | Attendu |
|---|---|
| Ouvrir le classeur, activer les macros | Tableau de bord affiché, aucun message d'erreur |
| Cliquer **Actualiser les banques** | Message : Banque Alpha 10 nouvelles opérations (4 doublons, 1 antérieure), Banque Beta 9 ; propositions certaines et incertaines ; une sauvegarde créée dans `Sauvegardes` |
| Onglet Rapprochement, **Tout valider (certaines)** | Les prévisions concernées passent à « Pointé » dans l'onglet Saisie |
| Groupe « Carburant flotte » : écrire Oui sur une ligne, **Valider les choix** | Plus aucune proposition |
| Onglet À traiter : choisir des actions, **Appliquer** | Les opérations disparaissent de la liste ; une charge créée apparaît dans Charges fixes |
| Saisie : ajouter un virement interne | Deux lignes liées ; trésorerie globale prévisionnelle inchangée |
| Boutons **30 j / 60 j / 90 j** | Le graphique et le tableau jour par jour changent d'horizon |
| Cliquer de nouveau **Actualiser** | « 0 nouvelle opération », doublons ignorés, soldes inchangés |
| Renommer la Banque 1 dans Paramètres, actualiser | Proposition de renommage, opérations mises à jour |
