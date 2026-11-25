# 1. Installation

## Prérequis

| Élément | Exigence |
|---|---|
| Système | Windows 10 ou 11 |
| Excel | Microsoft 365 (ou Excel 2019 / 2021) en version **bureau** — Excel Online et Excel pour Mac ne sont pas pris en charge (macros VBA, Power Query local, `Scripting.Dictionary`) |
| Fichiers bancaires | Exports CSV / TXT / XLSX téléchargés depuis les sites des banques (aucune connexion bancaire) |
| Droits | Aucun droit administrateur : tout se passe dans le profil de l'utilisateur |

Les relevés `.xls` (ancien format Excel 97-2003) nécessitent le composant gratuit « Microsoft Access Database Engine » ; préférer les exports CSV ou XLSX.

## Contenu livré

```
dist/
  Tresorerie.xlsm            le classeur (données de démonstration)
  Imports/Banque1/           relevés de démonstration Banque 1 (CSV)
  Imports/Banque2/           relevés de démonstration Banque 2 (XLSX)
  Sauvegardes/               sauvegardes automatiques (créées avant chaque actualisation)
  recette_attendue.json      valeurs attendues pour la recette automatique
install/
  Installer.cmd / .ps1       installation (déblocage, dossiers, emplacement approuvé, reconstruction)
  Recette.cmd / .ps1         recette automatique dans Excel
```

## Étapes

1. **Copier le dossier `dist`** à son emplacement définitif sur le poste (par exemple
   `C:\Users\<nom>\Documents\Tresorerie`). Le classeur cherche les relevés dans les sous-dossiers
   `Imports\Banque1` et `Imports\Banque2` **situés à côté de lui**.
2. **Débloquer les fichiers et autoriser les macros** : lancer `install\Installer.cmd`.
   Pour un dossier autre que `dist` :
   ```
   powershell -ExecutionPolicy Bypass -File install\Installer.ps1 -Dossier "C:\...\Tresorerie" -EmplacementApprouve
   ```
   * le script retire la marque « fichier provenant d'Internet » (sinon Windows bloque les macros) ;
   * `-EmplacementApprouve` ajoute le dossier aux **emplacements approuvés** d'Excel : les macros
     s'exécutent sans bandeau d'avertissement. Sans cette option, cliquer sur « Activer le contenu »
     à l'ouverture.
3. **Recette automatique (recommandée, 1 minute)** : lancer `install\Recette.cmd`.
   Le script ouvre une *copie* du dossier dans Excel, exécute toute la chaîne (Power Query, import,
   rapprochement), vérifie qu'un second import n'ajoute aucun doublon et compare 18 indicateurs aux
   valeurs attendues. Il doit afficher **RECETTE RÉUSSIE**. Le classeur livré n'est pas modifié.
4. **Ouvrir `Tresorerie.xlsm`**, activer les macros si demandé, puis explorer la démonstration :
   bouton **Actualiser les banques** sur le tableau de bord.
5. **Mise en service pour le client** (voir le guide technique, § Mise en service) :
   * macro `Admin_ViderDonnees` (Alt+F8, mot de passe administrateur) pour supprimer les données de
     démonstration ;
   * onglet **Paramètres** : nom de la société, noms des banques, **soldes initiaux et leur date** ;
   * supprimer les relevés de démonstration des dossiers `Imports` ;
   * adapter l'onglet technique `Param_Import` aux vrais formats d'export des banques.

## Si Excel signale un problème à l'ouverture

Le classeur est généré par programme (voir guide technique). Si Excel propose de « réparer » le
fichier ou si les boutons ne fonctionnent pas :

* **Macros désactivées** : vérifier l'étape 2 (déblocage / emplacement approuvé).
* **Projet VBA ou requêtes retirés par Excel** : relancer l'installation avec reconstruction ; Excel
  réimporte alors lui-même le code VBA et les requêtes Power Query depuis les sources `src\` :
  ```
  powershell -ExecutionPolicy Bypass -File install\Installer.ps1 -Reconstruire
  ```
  (une copie `Tresorerie.xlsm.avant_reconstruction` est conservée). Cette option active
  temporairement « Accès approuvé au modèle d'objet du projet VBA » puis remet le réglage initial.
* **Requêtes absentes mais macros actives** : à l'ouverture, le classeur propose de réinstaller les
  requêtes à partir du code M stocké dans l'onglet technique `Param_Import` (macro
  `Admin_ReinstallerRequetes`).

## Classeur synchronisé OneDrive / SharePoint

Les chemins OneDrive sont convertis automatiquement en chemin local. Si la conversion échoue, un
message le signale : indiquer alors un dossier local dans **Paramètres > Dossier racine**.
