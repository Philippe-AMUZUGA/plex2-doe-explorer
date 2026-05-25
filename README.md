# PLEX2 DOE Explorer

Application Windows pour explorer et exporter des plans d'expériences vers Excel, avec interface graphique.

- **Repository** : `plex2-doe-explorer`
- **Auteur** : Philipp AMUZUGA
- **Version stable** : 1.0
- **Licence** : AGPLv3

## Fonctionnalités

- Interface graphique Tkinter orientée utilisateurs non-techniques.
- Génération multi-plans dans un seul fichier Excel formaté.
- Plans pris en charge : factoriel complet, fractionnaires, Plackett-Burman, DSD, OFAT, Taguchi, GSD, CCD, Box-Behnken, Doehlert, LHS automatique.
- Plans space-filling discrets sur les niveaux utilisateur : maximin rapide et projection maximin inspirée MaxPro.
- Visualisation Excel optionnelle avec repères facteurs/niveaux et cellules d'essais colorées selon le séquencement.
- Tri automatique du résumé Excel :
  1. plans nominaux,
  2. plans avec avertissement `⚠`,
  3. plans `⚠ OMIS` en fin de liste.
- Mise en évidence systématique des adaptations : `⚠ REFORMATAGE`, `⚠ OMIS`, `⚠ LHS`.
- Limite de génération par défaut à 99 essais par plan pour garder des exports rapides et lisibles.
- Export par défaut dans le dossier `examples/`.

## Lancement rapide Windows

### Option 1 — usage direct

1. Télécharger le ZIP du projet.
2. Décompresser.
3. Double-cliquer sur `Lancer_PLEX2.bat`.
4. Au premier lancement, un environnement local `.runtime` est créé automatiquement si Python 3 est déjà installé sur le poste.

### Option 2 — création d'un exécutable Windows

```bat
build\build_windows.bat
```

Le binaire PyInstaller est ensuite généré dans `dist\PLEX2\`.

## Dépendances runtime

Installées automatiquement par `Lancer_PLEX2.bat` si Python est présent :

- `pyDOE3>=1.6`
- `pandas>=2.0`
- `openpyxl>=3.1`
- `numpy>=1.24`

## Structure du dépôt

```text
app/
  PLEX2_Launcher.py
  plex2_gui.py
  plex2_core.py
  assets/
build/
examples/
Lancer_PLEX2.bat
requirements.txt
README.md
USER_GUIDE.md
LICENSE
```

## Notes techniques

- `plex2_core.py` contient le moteur DOE, les plans space-filling et la logique Excel.
- `plex2_gui.py` contient l'interface Windows, la visualisation optionnelle et la gestion ergonomique du redimensionnement.
- La feuille `00_RESUME` suit strictement l'ordre des onglets générés.
- Les plans omis n'ont pas d'onglet et sont volontairement listés en fin de résumé avec `Feuille = —`.

## Publication GitHub

Contenu recommandé pour la première publication :

- ce code source,
- le `README.md`,
- le `LICENSE`,
- le `USER_GUIDE.md`,
- un ZIP de release prêt à l'emploi pour Windows.
