# RAI RSDE Générateur IANESCO

Outil de génération de fichiers XML **FCT_ASSAIN RSDE v4** (SANDRE) à partir
d'une extraction Tetraed Builder, pour la télétransmission aux agences de l'eau.

## Structure du projet

```
sandre_etape1_gui.py      Source principal (interface + traitement)
Parametres.xlsx           Table CODE_PARAMETRE → CdParametre SANDRE
Methodes.xlsx             Table CODE_METHODE → CdMethode SANDRE
Fractions.xlsx            Table CODE_Produit → CdFractionAnalysee
Unités.xlsx               Table libellé unité → CdUniteMesure
LQ_Incertitude.xlsx       Référentiel LQ / incertitude par paramètre + méthode
profils_station/          Profils JSON (SIRET, ouvrage, dates…) par station
CHANGELOG.md              Historique des versions
```

## Compilation de l'exécutable

Prérequis : Python 3.12, PyInstaller 6.17.0

```bash
pyinstaller --onefile sandre_etape1_gui.py
```

L'exe produit est copié manuellement dans ce dossier sous le nom `SANDRE_RSDE_App.exe`
(non versionné dans git — voir `.gitignore`).

## Versioning

Schéma `MAJEUR.MINEUR.PATCH` :
- **MAJEUR** : changement de format SANDRE ou refonte majeure
- **MINEUR** : nouvelle fonctionnalité ou correction de règle métier
- **PATCH** : correction de bug sans impact sur les sorties XML

La version courante est définie par `__version__` dans `sandre_etape1_gui.py`
et taguée dans git (`git tag v4.1.0`).
