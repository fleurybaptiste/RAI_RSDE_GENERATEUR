# Changelog — RAI RSDE Générateur IANESCO

Toutes les modifications notables sont documentées ici.
Format : `MAJEUR.MINEUR.PATCH` — voir [Semantic Versioning](https://semver.org/lang/fr/).

---

## [4.1.0] — 2026-10-23

### Corrections
- **A4.81 auto** : si le résultat < LQ et CdRemAnalyse = 1, correction automatique
  en code 10 (valeur < seuil, non certifiée) pour éviter le rejet SANDRE.
- **NumeroPointMesure** : valeurs par défaut corrigées — entrée = 3 (A3),
  sortie = 4 (A4) ; les valeurs 1/2 étaient incompatibles avec Neptune V4.8.x.

### Divers
- Ajout de `__version__` dans le source (`sandre_etape1_gui.py`).
- Correction manuelle du XML D24-01-1493 (EDV Châtellerault, campagne 2024-01-30) :
  NumeroPointMesure, HeureReceptionEchantillon (chronologie ENTREE), 2 × CdRemAnalyse.
- Documentation `Guide_Utilisateur_RSDE_IANESCOV4.docx` mise à jour :
  noms de fichiers référentiels, mapping complet Tetraed Builder → SANDRE.

---

## [4.0.0] — avant 2026-10-23

Version initiale mise en production (IANESCO V4).
Fonctionnalités : génération XML FCT_ASSAIN RSDE v4 depuis extraction Tetraed Builder,
gestion profils station, tables de correspondance Parametres/Methodes/Fractions/LQ.
