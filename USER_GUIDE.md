# Guide utilisateur simplifié

## 1. Ouvrir l'application

Sous Windows, double-cliquez sur `Lancer_PLEX2.bat`.

## 2. Définir vos facteurs

Dans la table de gauche :

- sélectionnez un facteur existant pour le modifier,
- ou saisissez un nouveau **nom de facteur**,
- puis entrez ses **niveaux** dans la zone dédiée,
- cliquez sur **Ajouter**.

Vous pouvez aussi :

- **Supprimer** le facteur sélectionné,
- **Vider** l'éditeur,
- générer rapidement des facteurs uniformes avec **Création rapide**.

## 3. Choisir le fichier Excel de sortie

Dans le panneau **Génération** :

- indiquez le nom du fichier `.xlsx`,
- définissez si besoin le seuil maximum du plan factoriel complet.

## 4. Générer le DOE

Cliquez sur **Générer le fichier Excel**.

Le fichier produit contient :

- une feuille `00_RESUME`,
- les onglets de plans générés dans le bon ordre,
- les avertissements `⚠` si l'outil a dû adapter ou omettre un plan.

## 5. Comprendre les mentions d'avertissement

- `⚠ REFORMATAGE` : le plan a nécessité une adaptation des niveaux ou des facteurs.
- `⚠ OMIS` : le plan n'a pas été généré.
- `⚠ LHS` : le Latin Hypercube a été calculé automatiquement.

## 6. Sauvegarder / recharger vos paramètres

- **Exporter les facteurs en JSON** pour sauvegarder votre configuration.
- **Importer des facteurs JSON** pour la recharger.

## 7. Bonnes pratiques

- utilisez des noms de facteurs courts et explicites,
- évitez des niveaux non numériques si vous voulez profiter pleinement des plans de surface de réponse et du LHS,
- consultez toujours `00_RESUME` avant d'exploiter les onglets détaillés.
