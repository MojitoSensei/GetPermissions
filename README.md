# Analyse des Permissions de Répertoires

Ce script PowerShell permet d'analyser les permissions d'accès pour les dossiers et sous-dossiers d'un répertoire sélectionné par l'utilisateur. Les résultats sont exportés dans un fichier Excel avec une mise en forme claire et des légendes pour faciliter l'interprétation.

## Fonctionnalités

- Permet de sélectionner un répertoire à analyser via une interface graphique.
- Analyse les permissions d'accès pour chaque utilisateur/groupe sur les dossiers et sous-dossiers (jusqu'à 3 niveaux).
- Catégorise les permissions en :
  - **RO** : Lecture seule
  - **RW** : Lecture et écriture
  - **ALL** : Contrôle total
  - **X** : Pas de permission significative
  - **DENY** : Accès refusé
- Exporte les résultats dans un fichier Excel :
  - Mise en forme conditionnelle des cellules basée sur les permissions.
  - Tableau pivoté par utilisateur/groupe.
  - Feuille de légende expliquant les catégories de permissions.

## Prérequis

- Windows PowerShell avec le module `ImportExcel` installé.
  - Pour installer le module : 
    ```powershell
    Install-Module -Name ImportExcel -Force
    ```

## Utilisation

1. Exécutez le script dans PowerShell.
2. Une fenêtre s'affiche pour sélectionner le répertoire à analyser.
3. Une autre fenêtre s'affiche pour choisir où sauvegarder le fichier Excel (avec un nom par défaut `Permissions.xlsx`).
4. Le fichier Excel est généré et ouvert automatiquement à la fin de l'analyse.

## Résultats

Le fichier Excel contient :
- **Feuille principale** : Un tableau des permissions des dossiers et sous-dossiers, avec :
  - Les colonnes **Niveau1**, **Niveau2**, **Niveau3** représentant la hiérarchie des dossiers.
  - Une colonne par utilisateur/groupe avec leur niveau d'accès.
- **Feuille de légende** : Une explication des niveaux d'accès (RO, RW, ALL, X, DENY).

## Exemple de Scénario

- Vous souhaitez vérifier les permissions d'un répertoire partagé en réseau.
- Vous exécutez ce script pour générer un rapport Excel clair et lisible des droits des utilisateurs.

## Avertissement

- Ce script n'altère aucune permission, il se contente d'effectuer une analyse.
- Veillez à disposer des droits d'accès nécessaires pour analyser les dossiers sélectionnés.
