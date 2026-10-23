# Changelog — RAI RSDE Générateur IANESCO

Toutes les modifications notables sont documentées ici.
Format : `MAJEUR.MINEUR.PATCH` — voir [Semantic Versioning](https://semver.org/lang/fr/).

---

## [4.2.0] — 2026-10-23

### Nouveautés
- **DateDepot (colonne AC)** : nouvelle colonne Builder `DateDepot` (format `30/01/2024 14:30`)
  utilisée en priorité pour `DateReceptionEchantillon`, `HeureReceptionEchantillon`,
  `DateAnalyse` et `HeureAnalyse` — résout les problèmes de chronologie (réception
  avant prélèvement) et les 2 jours d'écart entre date prélèvement et date analyse.
- `HeureAnalyse` désormais renseignée (était toujours vide).
- Fallback conservé : si `DateDepot` absent → `Date_rcpt`/`heure_rcpt` puis convention 11h00.

---

## [4.1.1] — 2026-10-23

### Corrections (build)
- Recompilation avec le bon point d'entrée `rsde_app.py` (icône goutte d'eau,
  customtkinter, support xlsm — la compilation précédente utilisait
  `sandre_etape1_gui.py` par erreur).
- Ajout de `SANDRE_RSDE_App.spec`, `rsde_app.py`, `xlsx_to_sandre_xml.py`,
  `ianesco_rsde.ico` dans le repo git.

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
