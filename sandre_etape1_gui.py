# -*- coding: utf-8 -*-
"""
Etape 1 : transforme une extraction brute de labo (+ tables de correspondance
Ianesco) en fichier "a plat" (onglet Resultats) pret pour l'etape 2
(SANDRE_Xlsx_to_XML.exe).

Gere des "profils de station" (SIRET, ouvrage, dates par defaut...) que l'on
sauvegarde/recharge d'une campagne a l'autre, pour n'importe quel client.
"""

__version__ = "4.2.0"
import os
import re
import sys
import json
import datetime as dt
import traceback
import threading

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Erreur", "Le module openpyxl est introuvable.")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Emplacements (a cote de l'exe/du script)
# ---------------------------------------------------------------------------
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


PROFILES_DIR = os.path.join(app_dir(), "profils_station")
os.makedirs(PROFILES_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Regles metier generales (independantes du client, propres au scenario RSDE
# et aux conventions d'export du LIMS Ianesco - voir Notice de l'outil)
# ---------------------------------------------------------------------------
CODE_PRODUIT_TO_FRACTION_LABEL = {
    "ER": "Eau brute",
    "ER_MOYEN_24H": "Eau brute",
    "ER_ICPE": "Eau brute",
    "ER_FILTRAT": "Phase aqueuse de l'eau (filtrÃ©e, centrifugÃ©e...)",
    "ER_MES": "Phase particulaire de l'eau",
}

POINT_MESURE_RULES_DEFAUT = [
    {"motcle": "entr", "numero": "3", "libelle": "ENTREE STATION", "code": "A3"},
    {"motcle": "sorti", "numero": "4", "libelle": "SORTIE STATION", "code": "A4"},
]

PARAM_METADATA_EXCLUS = {
    "DATE_PRELEVEMENT", "INTERVALLE_DE_VOLUME", "VOLUME_ECHANTILLON_PONCTUEL",
    "INTERVALLE_DE_TEMPS", "MINERALISATION_IMPRIMABLE",
}

METAUX_METHODE_715 = {
    "CADMIUM", "PLOMB", "NICKEL", "ARSENIC", "ZINC", "CUIVRE", "CHROME",
    "COBALT", "TITANE", "ARGENT", "THALLIUM",
}
PARAMS_AGREMENT_OBLIGATOIRE = {"1305", "1313", "1314", "6396", "1841"}
PARAM_DEBIT = "1552"
CD_UNITE_DEBIT = "120"  # m3/j, regle E4.23

NUM_RE = re.compile(r"^(<)?\s*([0-9]+(?:[.,][0-9]+)?)")


def to_float(s):
    if s is None:
        return None
    s = str(s).strip().replace(",", ".")
    m = re.match(r"^(-?[0-9]+(?:\.[0-9]+)?)", s)
    if not m:
        return None
    try:
        return float(m.group(1))
    except ValueError:
        return None


def _norm_code(s):
    if s is None:
        return None
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def load_referentiel(path, log):
    """Lit le referentiel unique LQ/Incertitude (fichier 'Referentiel_LQ_Incertitude...xlsx',
    feuille 'Referentiel', colonnes CODE_PARAMETRE/CODE_METHODE/LQ/Incertitude), qui remplace
    les anciennes sources eparpillees (RDD labo, Structure_Excel_RSDE.xlsm, retours labo par
    campagne) desormais fusionnees dans un seul fichier a maintenir dans le temps.
    Construit un dictionnaire par (CODE_PARAMETRE, CODE_METHODE) puis un par CODE_PARAMETRE
    seul (le plus complet/fiable en cas de doublon) pour le repli."""
    by_pair, by_param = {}, {}
    if not path:
        return by_pair, by_param
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet_name = next((n for n in wb.sheetnames if "referentiel" in n.lower()), wb.sheetnames[0])
    ws = wb[sheet_name]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    needed = ["CODE_PARAMETRE", "LQ", "Incertitude"]
    missing = [n for n in needed if n not in idx]
    if missing:
        log(f"/!\\ Colonnes {', '.join(missing)} absentes de '{sheet_name}', referentiel ignore.")
        return by_pair, by_param
    col_meth = idx.get("CODE_METHODE")
    best_score = {}
    for r in range(2, ws.max_row + 1):
        cp = ws.cell(row=r, column=idx["CODE_PARAMETRE"]).value
        if not cp:
            continue
        meth = ws.cell(row=r, column=col_meth).value if col_meth else None
        lq = ws.cell(row=r, column=idx["LQ"]).value
        incert = ws.cell(row=r, column=idx["Incertitude"]).value
        if lq in (None, "") and incert in (None, ""):
            continue
        cp_norm, meth_norm = _norm_code(cp), _norm_code(meth)
        by_pair[(cp_norm, meth_norm)] = (lq, incert)
        score = (lq not in (None, "")) + (incert not in (None, ""))
        if score >= best_score.get(cp_norm, -1):
            by_param[cp_norm] = (lq, incert)
            best_score[cp_norm] = score
    log(f"Referentiel LQ/Incertitude charge ({len(by_pair)} couple(s) parametre/methode, {len(by_param)} parametre(s) distinct(s)).")
    return by_pair, by_param


def parse_resultat(raw):
    if raw is None or str(raw).strip() == "":
        return "", "0", None
    txt = str(raw).strip()
    if " " in txt and ("Â°" in txt or " a " in txt.lower() or " Ã  " in txt.lower()):
        m = re.match(r"^([0-9]+(?:[.,][0-9]+)?)", txt)
        if m:
            return m.group(1).replace(",", "."), "1", f"Valeur brute labo : {txt}"
    m = NUM_RE.match(txt)
    if not m:
        return txt, "1", "Format de resultat non reconnu automatiquement, a verifier"
    is_lt = m.group(1) == "<"
    value = m.group(2).replace(",", ".")
    return value, ("10" if is_lt else "1"), None


def fmt_date(x):
    if x in (None, ""):
        return ""
    if isinstance(x, dt.time):
        # Cellule Excel vide/0 relue par openpyxl comme une heure pure (ex: 00:00:00) :
        # ca ne represente aucune date, la traiter comme absente plutot que stringifier "00:00:00".
        return ""
    if isinstance(x, (dt.date, dt.datetime)):
        return x.strftime("%Y-%m-%d")
    return str(x)


def fmt_pct(x):
    """Formate un pourcentage sans decimale inutile (100.0 -> '100'), requis par la
    regle A4.78 (l'incertitude DOIT etre exprimee en pourcentage) qui rejette le
    format flottant avec point decimal."""
    if x is None:
        return ""
    if float(x) == int(x):
        return str(int(x))
    return str(x)


def fmt_heure(x):
    if x in (None, ""):
        return ""
    if isinstance(x, dt.date) and not isinstance(x, dt.datetime):
        # Symetrique de fmt_date : une cellule qui ne contient qu'une date n'a pas d'heure.
        return ""
    if isinstance(x, (dt.time, dt.datetime)):
        return x.strftime("%H:%M:%S")
    return str(x)


def parse_fr_datetime(txt):
    m = re.match(r"(\d{2})/(\d{2})/(\d{4}).*?(\d{1,2})h(\d{2})", txt)
    if not m:
        return None
    j, mo, a, h, mi = (int(x) for x in m.groups())
    return dt.datetime(a, mo, j, h, mi)


# ---------------------------------------------------------------------------
# Champs d'un profil de station
# ---------------------------------------------------------------------------
PROFILE_FIELDS = [
    ("nom_profil", "Nom du profil (ex: Chatellerault)"),
    ("siret_emetteur", "SIRET Emetteur"),
    ("nom_emetteur", "Nom Emetteur"),
    ("siret_destinataire", "SIRET Destinataire (Agence de l'Eau)"),
    ("nom_destinataire", "Nom Destinataire"),
    ("cd_ouvrage", "Code ouvrage (CdOuvrageDepollution)"),
    ("nom_ouvrage", "Nom de l'ouvrage (STEP)"),
    ("type_ouvrage", "Type d'ouvrage (defaut 4 = STEP)"),
    ("cd_commune", "Code commune (INSEE)"),
    ("lb_commune", "Nom commune"),
    ("nature_syst", "Nature systeme traitement (defaut 1 = urbain)"),
    ("siret_mo", "SIRET Maitre d'ouvrage"),
    ("nom_mo", "Nom Maitre d'ouvrage"),
    ("siret_exploitant", "SIRET Exploitant"),
    ("nom_exploitant", "Nom Exploitant"),
    ("siret_labo", "SIRET Laboratoire (Ianesco)"),
    ("nom_labo", "Nom Laboratoire"),
    ("siret_preleveur", "SIRET Preleveur"),
    ("nom_preleveur", "Nom Preleveur"),
    ("siret_producteur", "SIRET Producteur"),
    ("nom_producteur", "Nom Producteur"),
    ("finalite_analyse", "FinaliteAnalyse (11 = RSDE, 1 = autosurv. standard)"),
    ("statut_rs", "StatutRsAnalyse par defaut (A = donnees brutes)"),
    ("qual_rs", "QualRsAnalyse par defaut (4 = non qualifie)"),
    ("conformite_prlvt", "ConformitePrlvt par defaut (1 = conforme)"),
    ("agre_ana_defaut", "AgreAna par defaut (0 = non agree)"),
    ("agre_ana_agrement", "AgreAna si agrement obligatoire (1 = agree)"),
]

PROFILE_DEFAULTS = {
    "type_ouvrage": "4",
    "nature_syst": "1",
    "finalite_analyse": "11",
    "statut_rs": "A",
    "qual_rs": "4",
    "conformite_prlvt": "1",
    "agre_ana_defaut": "0",
    "agre_ana_agrement": "1",
}


def list_profiles():
    out = []
    for fn in sorted(os.listdir(PROFILES_DIR)):
        if fn.lower().endswith(".json"):
            out.append(fn[:-5])
    return out


def load_profile(name):
    path = os.path.join(PROFILES_DIR, name + ".json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_profile(data):
    name = data["nom_profil"].strip()
    if not name:
        raise ValueError("Le profil doit avoir un nom.")
    safe = re.sub(r"[^A-Za-z0-9_\-Ã€-Å¿ ]", "_", name)
    path = os.path.join(PROFILES_DIR, safe + ".json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return safe


# ---------------------------------------------------------------------------
# Colonnes de sortie (doivent correspondre a ce qu'attend SANDRE_Xlsx_to_XML)
# ---------------------------------------------------------------------------
OUTPUT_COLUMNS = [
    ("SIRET_Emetteur", "ENVOI", True), ("Nom_Emetteur", "ENVOI", True),
    ("SIRET_Destinataire", "ENVOI", True), ("Nom_Destinataire", "ENVOI", True),
    ("DateDebutReference", "ENVOI", True), ("DateFinReference", "ENVOI", True),
    ("CdOuvrageDepollution", "OUVRAGE", True), ("TypeOuvrageDepollution", "OUVRAGE", True),
    ("NomOuvrageDepollution", "OUVRAGE", True), ("CdCommune", "OUVRAGE", False),
    ("LbCommune", "OUVRAGE", False), ("NatureSystTraitementEauxUsees", "OUVRAGE", True),
    ("SIRET_MaitreOuvrage", "OUVRAGE", False), ("Nom_MaitreOuvrage", "OUVRAGE", False),
    ("SIRET_Exploitant", "OUVRAGE", False), ("Nom_Exploitant", "OUVRAGE", False),
    ("NumeroPointMesure", "POINT DE MESURE", True), ("LbPointMesure", "POINT DE MESURE", True),
    ("LocGlobalePointMesure", "POINT DE MESURE", True), ("ComPointMesure", "POINT DE MESURE", False),
    ("DatePrlvt", "PRELEVEMENT", True), ("HeurePrlvt", "PRELEVEMENT", True),
    ("DureePrlvt", "PRELEVEMENT", True), ("SIRET_Preleveur", "PRELEVEMENT", True),
    ("Nom_Preleveur", "PRELEVEMENT", False), ("ConformitePrlvt", "PRELEVEMENT", True),
    ("AccrePrlvt", "PRELEVEMENT", True), ("CdSupport", "PRELEVEMENT", True),
    ("DateReceptionEchantillon", "ANALYSE / RESULTAT", True),
    ("HeureReceptionEchantillon", "ANALYSE / RESULTAT", True),
    ("DateAnalyse", "ANALYSE / RESULTAT", True), ("HeureAnalyse", "ANALYSE / RESULTAT", False),
    ("CdParametre", "ANALYSE / RESULTAT", True), ("RsAnalyse", "ANALYSE / RESULTAT", True),
    ("CdUniteMesure", "ANALYSE / RESULTAT", True), ("CdRemAnalyse", "ANALYSE / RESULTAT", True),
    ("InSituAnalyse", "ANALYSE / RESULTAT", True), ("StatutRsAnalyse", "ANALYSE / RESULTAT", True),
    ("QualRsAnalyse", "ANALYSE / RESULTAT", True), ("CdFractionAnalysee", "ANALYSE / RESULTAT", True),
    ("CdMethode", "ANALYSE / RESULTAT", True), ("SIRET_Laboratoire", "ANALYSE / RESULTAT", True),
    ("Nom_Laboratoire", "ANALYSE / RESULTAT", False), ("SIRET_Producteur", "ANALYSE / RESULTAT", False),
    ("Nom_Producteur", "ANALYSE / RESULTAT", False), ("FinaliteAnalyse", "ANALYSE / RESULTAT", True),
    ("LDAna", "ANALYSE / RESULTAT", False), ("LQAna", "ANALYSE / RESULTAT", True),
    ("LSAna", "ANALYSE / RESULTAT", False), ("AccreAna", "ANALYSE / RESULTAT", True),
    ("AgreAna", "ANALYSE / RESULTAT", True), ("IncertAna", "ANALYSE / RESULTAT", True),
    ("ComAna", "ANALYSE / RESULTAT", False),
]

GROUP_COLORS = {
    "ENVOI": "1F4E78", "OUVRAGE": "2E75B6", "POINT DE MESURE": "548235",
    "PRELEVEMENT": "BF8F00", "ANALYSE / RESULTAT": "833C00",
}


# ---------------------------------------------------------------------------
# Coeur du traitement
# ---------------------------------------------------------------------------
def find_point_mesure(denomination, rules):
    d = (denomination or "").lower()
    for rule in rules:
        if rule["motcle"].lower() in d:
            return rule["numero"], rule["libelle"], rule["code"]
    return "", "", ""


# ---------------------------------------------------------------------------
# Reference multi-stations/points (optionnelle) : permet de traiter en un seul
# passage une extraction qui melange plusieurs ouvrages (ex: Grand Poitiers,
# qui regroupe La Folie et Chasseneuil dans le meme Builder), chacun avec ses
# propres intervenants. Deux onglets :
#   - "Stations"  : un ouvrage par ligne, memes champs que OUVRAGE dans le profil.
#   - "Points"    : un prefixe de Denomination par ligne (match "commence par",
#                    insensible a la casse) -> code ouvrage + intervenants du
#                    point. Le prefixe le plus long qui correspond l'emporte.
# Sans ce fichier, le comportement est inchange (tout vient du profil, comme
# pour Chatellerault).
# ---------------------------------------------------------------------------
STATIONS_COLUMNS = [
    "CdOuvrageDepollution", "NomOuvrageDepollution", "TypeOuvrageDepollution",
    "CdCommune", "LbCommune", "NatureSystTraitementEauxUsees",
    "SIRET_MaitreOuvrage", "Nom_MaitreOuvrage", "SIRET_Exploitant", "Nom_Exploitant",
]
POINTS_COLUMNS = [
    "POINT_EAU", "CdOuvrageDepollution", "CdSupport",
    "SIRET_Producteur", "SIRET_Preleveur", "SIRET_Laboratoire",
]


def load_points_stations(path, log):
    """Lit le classeur de reference multi-stations (onglets 'Stations' et
    'Points'). Retourne (stations: {cd_ouvrage: {champ: valeur}},
    points: [(prefixe_normalise, cd_ouvrage, support, producteur, preleveur,
    laboratoire), ...] trie par longueur de prefixe decroissante)."""
    stations, points = {}, []
    if not path:
        return stations, points
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Stations" not in wb.sheetnames or "Points" not in wb.sheetnames:
        log("/!\\ Le fichier de reference multi-stations doit contenir les onglets 'Stations' et 'Points' ; ignore.")
        return stations, points

    ws_st = wb["Stations"]
    headers = [ws_st.cell(row=1, column=c).value for c in range(1, ws_st.max_column + 1)]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    missing = [c for c in STATIONS_COLUMNS if c not in idx]
    if missing:
        log(f"/!\\ Colonnes {', '.join(missing)} absentes de l'onglet 'Stations' ; reference multi-stations ignoree.")
        return {}, []
    for r in range(2, ws_st.max_row + 1):
        cd = ws_st.cell(row=r, column=idx["CdOuvrageDepollution"]).value
        if not cd:
            continue
        stations[str(cd).strip()] = {
            col: (ws_st.cell(row=r, column=idx[col]).value or "") for col in STATIONS_COLUMNS
        }

    ws_pt = wb["Points"]
    headers = [ws_pt.cell(row=1, column=c).value for c in range(1, ws_pt.max_column + 1)]
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    missing = [c for c in POINTS_COLUMNS if c not in idx]
    if missing:
        log(f"/!\\ Colonnes {', '.join(missing)} absentes de l'onglet 'Points' ; reference multi-stations ignoree.")
        return {}, []
    for r in range(2, ws_pt.max_row + 1):
        prefix = ws_pt.cell(row=r, column=idx["POINT_EAU"]).value
        cd = ws_pt.cell(row=r, column=idx["CdOuvrageDepollution"]).value
        if not prefix or not cd:
            continue
        points.append((
            str(prefix).strip().lower(),
            str(cd).strip(),
            str(ws_pt.cell(row=r, column=idx["CdSupport"]).value or "").strip(),
            str(ws_pt.cell(row=r, column=idx["SIRET_Producteur"]).value or "").strip(),
            str(ws_pt.cell(row=r, column=idx["SIRET_Preleveur"]).value or "").strip(),
            str(ws_pt.cell(row=r, column=idx["SIRET_Laboratoire"]).value or "").strip(),
        ))
    points.sort(key=lambda t: len(t[0]), reverse=True)
    log(f"Reference multi-stations chargee : {len(stations)} station(s), {len(points)} point(s).")
    return stations, points


def resolve_station_point(denomination, stations, points, profile):
    """Resout les champs ouvrage/intervenants pour une ligne, a partir de la
    reference multi-stations si elle matche, sinon par repli sur le profil
    (comportement d'origine, mono-station)."""
    d = (denomination or "").strip().lower()
    for prefix, cd_ouvrage, support, producteur, preleveur, laboratoire in points:
        if d.startswith(prefix):
            station = stations.get(cd_ouvrage, {})
            return {
                "CdOuvrageDepollution": cd_ouvrage,
                "NomOuvrageDepollution": station.get("NomOuvrageDepollution", ""),
                "TypeOuvrageDepollution": station.get("TypeOuvrageDepollution") or profile.get("type_ouvrage", "4"),
                "CdCommune": station.get("CdCommune", ""),
                "LbCommune": station.get("LbCommune", ""),
                "NatureSystTraitementEauxUsees": station.get("NatureSystTraitementEauxUsees") or profile.get("nature_syst", "1"),
                "SIRET_MaitreOuvrage": station.get("SIRET_MaitreOuvrage", ""),
                "Nom_MaitreOuvrage": station.get("Nom_MaitreOuvrage", ""),
                "SIRET_Exploitant": station.get("SIRET_Exploitant", ""),
                "Nom_Exploitant": station.get("Nom_Exploitant", ""),
                "CdSupport": support or "3",
                "SIRET_Producteur": producteur or profile.get("siret_producteur", ""),
                "SIRET_Preleveur": preleveur or profile.get("siret_preleveur", ""),
                "SIRET_Laboratoire": laboratoire or profile.get("siret_labo", ""),
                "_matched": True,
            }
    return {
        "CdOuvrageDepollution": profile.get("cd_ouvrage", ""),
        "NomOuvrageDepollution": profile.get("nom_ouvrage", ""),
        "TypeOuvrageDepollution": profile.get("type_ouvrage", "4"),
        "CdCommune": profile.get("cd_commune", ""),
        "LbCommune": profile.get("lb_commune", ""),
        "NatureSystTraitementEauxUsees": profile.get("nature_syst", "1"),
        "SIRET_MaitreOuvrage": profile.get("siret_mo", ""),
        "Nom_MaitreOuvrage": profile.get("nom_mo", ""),
        "SIRET_Exploitant": profile.get("siret_exploitant", ""),
        "Nom_Exploitant": profile.get("nom_exploitant", ""),
        "CdSupport": "3",
        "SIRET_Producteur": profile.get("siret_producteur", ""),
        "SIRET_Preleveur": profile.get("siret_preleveur", ""),
        "SIRET_Laboratoire": profile.get("siret_labo", ""),
        "_matched": False,
    }


def run_generation(extraction_path, parametres_path, unites_path, fractions_path,
                    methodes_path, profile, date_debut, date_fin, log, referentiel_path=None,
                    points_stations_path=None):
    log(f"Lecture du profil '{profile.get('nom_profil')}'...")

    ref_by_pair, ref_by_param = load_referentiel(referentiel_path, log)
    ref_stations, ref_points = load_points_stations(points_stations_path, log)

    log("Chargement des tables de correspondance...")
    wb_param = openpyxl.load_workbook(parametres_path, data_only=True)
    ws_param = wb_param["Param"]
    param_map = {}
    for r in range(2, ws_param.max_row + 1):
        key = ws_param.cell(row=r, column=1).value
        code = ws_param.cell(row=r, column=3).value
        if key is not None:
            param_map[key] = code

    wb_unit = openpyxl.load_workbook(unites_path, data_only=True)
    ws_unit = wb_unit["Unite"]
    unit_map = {}
    for r in range(2, ws_unit.max_row + 1):
        key = ws_unit.cell(row=r, column=1).value
        code = ws_unit.cell(row=r, column=2).value
        if key is not None:
            unit_map[key] = code

    wb_frac = openpyxl.load_workbook(fractions_path, data_only=True)
    ws_frac = wb_frac["FRACTIONS"]
    frac_label_to_code = {}
    for r in range(2, ws_frac.max_row + 1):
        frac_label = ws_frac.cell(row=r, column=3).value
        frac_code = ws_frac.cell(row=r, column=4).value
        if frac_label is not None:
            frac_label_to_code[frac_label] = frac_code

    wb_meth = openpyxl.load_workbook(methodes_path, data_only=True)
    ws_meth = wb_meth[wb_meth.sheetnames[0]]
    methode_map = {}
    for r in range(2, ws_meth.max_row + 1):
        key = ws_meth.cell(row=r, column=1).value
        code = ws_meth.cell(row=r, column=2).value
        if key is not None:
            nk = _norm_code(key)
            if nk:
                methode_map[nk] = code

    default_support = "3"  # EAU - le seul rencontre a ce jour dans les extractions Ianesco

    log("Lecture de l'extraction brute...")
    wb_ext = openpyxl.load_workbook(extraction_path, data_only=True)
    ws_ext = wb_ext["Builder"]
    headers = [ws_ext.cell(row=7, column=c).value for c in range(1, 31)]
    idx = {h: i + 1 for i, h in enumerate(headers)}

    # Colonnes indispensables pour identifier une ligne (sans elles, impossible
    # de savoir de quel resultat il s'agit).
    required_cols = ["CODE_PARAMETRE", "Denomination", "Resultat_Analyse", "DatePrelv"]
    missing_required = [c for c in required_cols if c not in idx]
    if missing_required:
        raise ValueError(
            "Colonnes indispensables absentes de l'extraction (structure inattendue) : "
            + ", ".join(missing_required)
        )

    # Colonnes utiles mais facultatives : si absentes, on continue avec des
    # valeurs vides (signalees dans le rapport "a verifier" au cas par cas),
    # plutot que de bloquer toute la generation.
    optional_cols = ["Analyse_Denomination", "Unite", "Date_Fin_Examen", "LimiteQ",
                      "Incertitude", "Cofrac_analyse", "CODE_Produit", "heure_rcpt",
                      "Date_rcpt", "CODE_METHODE", "ID_ECHANT", "DateDepot"]
    missing_optional = [c for c in optional_cols if c not in idx]
    if missing_optional:
        log(f"/!\\ Colonnes absentes de l'extraction (poursuite avec champs vides) : {', '.join(missing_optional)}")

    def g(row, name):
        if name not in idx:
            return None
        return ws_ext.cell(row=row, column=idx[name]).value

    point_rules = profile.get("point_mesure_rules") or POINT_MESURE_RULES_DEFAUT

    # Fenetre de prelevement (Heure/Duree) par echantillon, si la metadonnee
    # DATE_PRELEVEMENT est presente (convention export Ianesco)
    prlvt_window = {}
    for r in range(8, ws_ext.max_row + 1):
        if g(r, "CODE_PARAMETRE") == "DATE_PRELEVEMENT":
            ech = g(r, "ID_ECHANT")
            d = parse_fr_datetime(str(g(r, "Resultat_Analyse")))
            if d:
                prlvt_window.setdefault(ech, []).append(d)
    prlvt_info = {}
    for ech, dts in prlvt_window.items():
        dts.sort()
        start, end = dts[0], dts[-1]
        total_sec = int((end - start).total_seconds())
        h, rem = divmod(total_sec, 3600)
        m, s = divmod(rem, 60)
        prlvt_info[ech] = (end.strftime("%H:%M:%S"), f"{h:02d}:{m:02d}:{s:02d}")

    output_rows = []
    issues = []

    for r in range(8, ws_ext.max_row + 1):
        code_param = g(r, "CODE_PARAMETRE")
        if not code_param or code_param in PARAM_METADATA_EXCLUS:
            continue

        denomination = g(r, "Denomination")
        num_pm, lb_pm, loc_pm = find_point_mesure(denomination, point_rules)
        if not loc_pm:
            issues.append((r, "Point de mesure", f"Non reconnu depuis Denomination='{denomination}'"))

        station = resolve_station_point(denomination, ref_stations, ref_points, profile)
        if ref_points and not station["_matched"]:
            issues.append((r, "Station", f"Denomination='{denomination}' ne correspond a aucun prefixe de la reference multi-stations ; repli sur le profil"))

        cd_parametre = param_map.get(code_param)
        if cd_parametre in (None, ""):
            issues.append((r, "Parametre", f"CODE_PARAMETRE '{code_param}' absent de Parametres.xlsx"))
            cd_parametre = ""

        unite_txt = g(r, "Unite")
        cd_unite = unit_map.get(unite_txt, "")
        if unite_txt not in (None, "", "/") and cd_unite in (None, ""):
            issues.append((r, "Unite", f"Unite '{unite_txt}' absente de Unites.xlsx"))

        rs_analyse, cd_rem, note_resultat = parse_resultat(g(r, "Resultat_Analyse"))

        cofrac = g(r, "Cofrac_analyse")
        accre = {"Vrai": "1", "Faux": "2"}.get(cofrac, "0")

        ech = g(r, "ID_ECHANT")
        heure_prlvt, duree_prlvt = prlvt_info.get(ech, ("", ""))

        incertitude_pct = to_float(g(r, "Incertitude"))
        limiteq_val = to_float(g(r, "LimiteQ"))

        if incertitude_pct is None or limiteq_val is None:
            cp_norm, meth_norm = _norm_code(code_param), _norm_code(g(r, "CODE_METHODE"))
            ref_pair = ref_by_pair.get((cp_norm, meth_norm))
            ref_param = ref_by_param.get(cp_norm)
            approx = False
            if ref_pair is None and ref_param is None and cp_norm and cp_norm.endswith("ST"):
                # Parametre sous-traite (suffixe _ST) sans entree dediee : a defaut, on
                # retombe sur la version "classique" (methode interne) du meme parametre,
                # moins fiable car methode/sous-traitant potentiellement differents ->
                # toujours signalee pour verification par le labo.
                ref_param = ref_by_param.get(cp_norm[:-2])
                approx = True
            # Combine les 2 niveaux : la valeur specifique au couple (parametre,
            # methode) est prioritaire champ par champ, mais si elle ne renseigne
            # que la LQ (ou que l'Incertitude), on complete avec le meilleur
            # candidat au niveau parametre plutot que de laisser le champ vide.
            ref_lq = ref_pair[0] if ref_pair and ref_pair[0] not in (None, "") else None
            ref_incert = ref_pair[1] if ref_pair and ref_pair[1] not in (None, "") else None
            if ref_param:
                if ref_lq is None and ref_param[0] not in (None, ""):
                    ref_lq = ref_param[0]
                if ref_incert is None and ref_param[1] not in (None, ""):
                    ref_incert = ref_param[1]
            if ref_lq is not None or ref_incert is not None:
                suffix = " (approx.)" if approx else ""
                detail_suffix = (" - version standard (non-ST) du meme parametre, methode "
                                  "potentiellement differente - A VERIFIER avec le labo/sous-traitant"
                                  if approx else " (a verifier)")
                if incertitude_pct is None and to_float(ref_incert) is not None:
                    incertitude_pct = to_float(ref_incert)
                    issues.append((r, f"Incertitude auto{suffix}", f"Completee depuis le referentiel LQ/Incertitude pour parametre {cd_parametre}{detail_suffix}"))
                if limiteq_val is None and to_float(ref_lq) is not None:
                    limiteq_val = to_float(ref_lq)
                    issues.append((r, f"LQ auto{suffix}", f"Completee depuis le referentiel LQ/Incertitude pour parametre {cd_parametre}{detail_suffix}"))
        if incertitude_pct is not None and incertitude_pct > 100:
            issues.append((r, "Incertitude", f"{incertitude_pct}% plafonnee a 99.6% (decision labo/responsable, cf regle E4.63)"))
            incertitude_pct = 99.6

        if incertitude_pct is None and profile.get("finalite_analyse", "11") == "11":
            incertitude_pct = 50.0
            issues.append((r, "Incertitude defaut 50%", f"Aucune incertitude connue pour parametre {cd_parametre} (ni extraction, ni referentiel) ; valeur par defaut 50% appliquee sur directive N+1"))

        param_upper = (g(r, "Analyse_Denomination") or "").upper()
        agre_ana = profile.get("agre_ana_defaut", "0")
        if str(cd_parametre) in PARAMS_AGREMENT_OBLIGATOIRE:
            agre_ana = profile.get("agre_ana_agrement", "1")
            issues.append((r, "Agrement", f"Parametre {cd_parametre} soumis a agrement obligatoire (A4.75)"))

        code_methode_labo = g(r, "CODE_METHODE")
        cd_methode = methode_map.get(_norm_code(code_methode_labo) or "", "")
        if any(m in param_upper for m in METAUX_METHODE_715):
            cd_methode = "715"
        if cd_methode == "" and code_methode_labo not in (None, "", "NA"):
            issues.append((r, "Methode", f"Methode labo '{code_methode_labo}' (parametre {cd_parametre}) sans correspondance connue"))
        if any(m in param_upper for m in METAUX_METHODE_715):
            issues.append((r, "Methode metal", f"Verifier que la methode Sandre 715 convient pour '{param_upper}' (regle A4.79)"))

        code_produit = g(r, "CODE_Produit")
        frac_label = CODE_PRODUIT_TO_FRACTION_LABEL.get(code_produit)
        cd_fraction = frac_label_to_code.get(frac_label, "")
        if cd_fraction == "":
            issues.append((r, "Fraction", f"Fraction non reconnue pour CODE_Produit={code_produit} (parametre {cd_parametre})"))

        cd_support = station["CdSupport"] or default_support
        cd_unite_final = cd_unite
        in_situ = "2"
        if str(cd_parametre) == PARAM_DEBIT:
            cd_unite_final = CD_UNITE_DEBIT
            in_situ = "1"

        if limiteq_val is None:
            issues.append((r, "LQAna", f"Limite de quantification manquante pour parametre {cd_parametre}"))
        if incertitude_pct is None:
            issues.append((r, "IncertAna", f"Incertitude manquante pour parametre {cd_parametre}"))

        # DateDepot (colonne AC) : date + heure de depot au labo = reception ET analyse.
        # Prioritaire sur les anciennes colonnes Date_rcpt/heure_rcpt/Date_Fin_Examen.
        date_depot_val = g(r, "DateDepot")
        date_rcpt     = fmt_date(date_depot_val)
        heure_rcpt    = fmt_heure(date_depot_val)
        date_analyse  = fmt_date(date_depot_val)
        heure_analyse = fmt_heure(date_depot_val)

        # Fallback si DateDepot absent : anciennes colonnes puis convention 11h00
        if date_rcpt == "" or heure_rcpt == "":
            if date_rcpt == "":
                date_rcpt = fmt_date(g(r, "Date_rcpt"))
            if heure_rcpt == "":
                heure_rcpt = fmt_heure(g(r, "heure_rcpt"))
            if (date_rcpt == "" or heure_rcpt == "") and profile.get("finalite_analyse", "11") == "11":
                if date_rcpt == "":
                    date_rcpt = fmt_date(g(r, "DatePrelv"))
                if heure_rcpt == "":
                    heure_rcpt = "11:00:00"
                issues.append((r, "Reception auto", "DateDepot absent, completee par convention RSDE (jour du prelevement, 11h00)"))

        if date_analyse == "":
            date_analyse = fmt_date(g(r, "Date_Fin_Examen"))
            heure_analyse = ""
            if date_analyse == "":
                date_analyse = fmt_date(g(r, "DatePrelv"))
                heure_analyse = ""
                if date_analyse != "":
                    issues.append((r, "DateAnalyse auto", "DateDepot et Date_Fin_Examen absents, completee avec DatePrelv"))

        # A4.81 : si resultat < LQ et code remarque = 1 (valeur certifiee), corriger
        # automatiquement en 10 (valeur < LQ, non certifiee) pour eviter le rejet SANDRE
        if cd_rem == "1" and limiteq_val is not None and rs_analyse not in (None, ""):
            try:
                if float(rs_analyse) < float(limiteq_val):
                    cd_rem = "10"
                    issues.append((r, "A4.81 auto", f"Resultat {rs_analyse} < LQ {limiteq_val} pour parametre {cd_parametre} : CdRemAnalyse 1 -> 10"))
            except (ValueError, TypeError):
                pass

        row = {
            "SIRET_Emetteur": profile.get("siret_emetteur", ""),
            "Nom_Emetteur": profile.get("nom_emetteur", ""),
            "SIRET_Destinataire": profile.get("siret_destinataire", ""),
            "Nom_Destinataire": profile.get("nom_destinataire", ""),
            "DateDebutReference": date_debut,
            "DateFinReference": date_fin,
            "CdOuvrageDepollution": station["CdOuvrageDepollution"],
            "TypeOuvrageDepollution": station["TypeOuvrageDepollution"],
            "NomOuvrageDepollution": station["NomOuvrageDepollution"],
            "CdCommune": station["CdCommune"],
            "LbCommune": station["LbCommune"],
            "NatureSystTraitementEauxUsees": station["NatureSystTraitementEauxUsees"],
            "SIRET_MaitreOuvrage": station["SIRET_MaitreOuvrage"],
            "Nom_MaitreOuvrage": station["Nom_MaitreOuvrage"],
            "SIRET_Exploitant": station["SIRET_Exploitant"],
            "Nom_Exploitant": station["Nom_Exploitant"],
            "NumeroPointMesure": num_pm, "LbPointMesure": lb_pm,
            "LocGlobalePointMesure": loc_pm, "ComPointMesure": "",
            "DatePrlvt": fmt_date(g(r, "DatePrelv")),
            "HeurePrlvt": heure_prlvt, "DureePrlvt": duree_prlvt,
            "SIRET_Preleveur": station["SIRET_Preleveur"],
            "Nom_Preleveur": profile.get("nom_preleveur", ""),
            "ConformitePrlvt": profile.get("conformite_prlvt", "1"),
            "AccrePrlvt": accre, "CdSupport": cd_support,
            "DateReceptionEchantillon": date_rcpt,
            "HeureReceptionEchantillon": heure_rcpt,
            "DateAnalyse": date_analyse, "HeureAnalyse": heure_analyse,
            "CdParametre": cd_parametre, "RsAnalyse": rs_analyse,
            "CdUniteMesure": cd_unite_final, "CdRemAnalyse": cd_rem,
            "InSituAnalyse": in_situ,
            "StatutRsAnalyse": profile.get("statut_rs", "A"),
            "QualRsAnalyse": profile.get("qual_rs", "4"),
            "CdFractionAnalysee": cd_fraction, "CdMethode": cd_methode,
            "SIRET_Laboratoire": station["SIRET_Laboratoire"],
            "Nom_Laboratoire": profile.get("nom_labo", ""),
            "SIRET_Producteur": station["SIRET_Producteur"],
            "Nom_Producteur": profile.get("nom_producteur", ""),
            "FinaliteAnalyse": profile.get("finalite_analyse", "11"),
            "LDAna": "", "LQAna": "" if limiteq_val is None else str(limiteq_val),
            "LSAna": "", "AccreAna": accre, "AgreAna": agre_ana,
            "IncertAna": fmt_pct(incertitude_pct),
            "ComAna": note_resultat or "",
            "_src_row": r,
            "_src_designation": g(r, "Designation"),
            "_src_code_produit": code_produit,
            "_src_analyse_denomination": g(r, "Analyse_Denomination"),
            "_src_code_parametre": code_param,
            "_src_code_methode": code_methode_labo,
            "_src_sous_traite": bool(code_param) and str(code_param).upper().endswith("_ST"),
        }
        output_rows.append(row)

    log(f"{len(output_rows)} ligne(s) d'analyse lues.")

    # Deduplication Sandre (E4.43) : un seul releve par (point, date, parametre,
    # fraction, finalite, in-situ/labo)
    seen = {}
    deduped = []
    for row in output_rows:
        # Si CdParametre n'a pas ete resolu (signale en "Parametre" dans les
        # issues), on ne peut pas savoir si 2 lignes sont de vrais doublons :
        # on les garde toutes les deux (cle rendue unique via la ligne source)
        # plutot que de risquer de faire disparaitre une mesure a tort.
        cd_param_key = row["CdParametre"] or f"_non_resolu_{row['_src_row']}"
        key = (row["CdOuvrageDepollution"], row["NumeroPointMesure"], row["DatePrlvt"], cd_param_key,
               row["CdFractionAnalysee"], row["FinaliteAnalyse"], row["InSituAnalyse"])
        if key in seen:
            issues.append((row["_src_row"], "Doublon",
                            f"Ecartee, doublon de la ligne {seen[key]} (meme point/date/parametre/fraction)"))
            continue
        seen[key] = row["_src_row"]
        deduped.append(row)
    output_rows = deduped
    log(f"{len(output_rows)} ligne(s) apres suppression des doublons.")
    log(f"{len(issues)} point(s) signale(s) - voir le rapport joint.")

    return output_rows, issues


# ---------------------------------------------------------------------------
# Ecriture des classeurs de sortie
# ---------------------------------------------------------------------------
def write_resultats_xlsx(output_rows, out_path):
    FONT_NAME = "Arial"
    normal_font = Font(name=FONT_NAME, size=10)
    bold_font = Font(name=FONT_NAME, size=10, bold=True)
    group_font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    fill_oblig = PatternFill("solid", fgColor="FFC7CE")
    fill_fac = PatternFill("solid", fgColor="DDEBF7")
    flag_fill = PatternFill("solid", fgColor="FFEB9C")
    thin = Side(style="thin", color="BFBFBF")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resultats"
    ws.sheet_view.showGridLines = False

    prev_group, group_start = None, 1
    for i, (name, group, oblig) in enumerate(OUTPUT_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = 16
        if group != prev_group:
            if prev_group is not None:
                ws.merge_cells(start_row=1, start_column=group_start, end_row=1, end_column=i - 1)
            group_start, prev_group = i, group
    ws.merge_cells(start_row=1, start_column=group_start, end_row=1, end_column=len(OUTPUT_COLUMNS))

    prev_group = None
    for i, (name, group, oblig) in enumerate(OUTPUT_COLUMNS, start=1):
        cell = ws.cell(row=1, column=i)
        if group != prev_group:
            cell.value = group
            prev_group = group
        cell.fill = PatternFill("solid", fgColor=GROUP_COLORS[group])
        cell.font = group_font
        cell.alignment = wrap_center
        cell.border = border_all

    for i, (name, group, oblig) in enumerate(OUTPUT_COLUMNS, start=1):
        label = f"{name}\n({'Obligatoire' if oblig else 'Facultatif'})"
        cell = ws.cell(row=2, column=i, value=label)
        cell.fill = fill_oblig if oblig else fill_fac
        cell.font = bold_font
        cell.alignment = wrap_center
        cell.border = border_all
    ws.row_dimensions[2].height = 42

    for i, row in enumerate(output_rows):
        rr = 3 + i
        for ci, (col_name, group, oblig) in enumerate(OUTPUT_COLUMNS, start=1):
            val = row.get(col_name, "")
            cell = ws.cell(row=rr, column=ci, value=val)
            cell.border = border_all
            cell.font = normal_font
            if val == "" and oblig:
                cell.fill = flag_fill

    ws.freeze_panes = "A3"
    wb.save(out_path)


def write_issues_xlsx(issues, out_path):
    FONT_NAME = "Arial"
    header_font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    normal_font = Font(name=FONT_NAME, size=10)
    thin = Side(style="thin", color="BFBFBF")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
    wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "A_verifier"
    ws.sheet_view.showGridLines = False
    headers = ["Ligne extraction", "Categorie", "Detail"]
    widths = [16, 20, 90]
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = header_font, header_fill, wrap_center, border_all
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    for i, (r, cat, detail) in enumerate(issues, start=2):
        ws.cell(row=i, column=1, value=r).font = normal_font
        ws.cell(row=i, column=2, value=cat).font = normal_font
        ws.cell(row=i, column=3, value=detail).font = normal_font
        for c in range(1, 4):
            ws.cell(row=i, column=c).border = border_all
            ws.cell(row=i, column=c).alignment = wrap if c == 3 else wrap_center
    wb.save(out_path)


def write_completion_request_xlsx(output_rows, out_path):
    """Genere un classeur 'a completer' avec un onglet recapitulatif et le
    detail des lignes ou LQAna et/ou IncertAna manquent (colonnes CODE_PARAMETRE/
    CODE_METHODE internes + LQ/Incertitude a remplir), pour renvoi au laboratoire
    puis fusion dans le referentiel unique (memes cles que load_referentiel)."""
    FONT_NAME = "Arial"
    header_font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    normal_font = Font(name=FONT_NAME, size=10)
    bold_font = Font(name=FONT_NAME, size=11, bold=True)
    fill_a_saisir = PatternFill("solid", fgColor="FFEB9C")
    thin = Side(style="thin", color="BFBFBF")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap_left = Alignment(horizontal="left", vertical="center")

    missing_rows = []
    for row in output_rows:
        lq_missing = row["LQAna"] == ""
        incert_missing = row["IncertAna"] == ""
        if lq_missing or incert_missing:
            missing_rows.append((row, lq_missing, incert_missing))

    nb_total = len(output_rows)
    nb_lq = sum(1 for _, lq, _ in missing_rows if lq)
    nb_incert = sum(1 for _, _, inc in missing_rows if inc)
    nb_both = sum(1 for _, lq, inc in missing_rows if lq and inc)
    nb_lq_st = sum(1 for r, lq, _ in missing_rows if lq and r.get("_src_sous_traite"))
    nb_lq_int = nb_lq - nb_lq_st
    nb_incert_st = sum(1 for r, _, inc in missing_rows if inc and r.get("_src_sous_traite"))
    nb_incert_int = nb_incert - nb_incert_st
    from collections import Counter
    par_point = Counter(r["LbPointMesure"] for r, _, _ in missing_rows)

    wb = openpyxl.Workbook()
    ws_recap = wb.active
    ws_recap.title = "Recap"
    ws_recap.sheet_view.showGridLines = False
    ws_recap.column_dimensions["A"].width = 45
    ws_recap.column_dimensions["B"].width = 14

    def put(r, label, value, bold=False):
        c1 = ws_recap.cell(row=r, column=1, value=label)
        c2 = ws_recap.cell(row=r, column=2, value=value)
        c1.font = bold_font if bold else normal_font
        c2.font = bold_font if bold else normal_font
        c1.alignment, c2.alignment = wrap_left, wrap_center

    put(1, "Recapitulatif LQ / Incertitude manquantes", "", bold=True)
    put(2, "Lignes d'analyse au total", nb_total)
    put(3, "", "")
    put(4, "Lignes avec LQ manquante", nb_lq, bold=True)
    put(5, "   dont parametres sous-traites (_ST)", nb_lq_st)
    put(6, "   dont parametres analyses en interne", nb_lq_int)
    put(7, "", "")
    put(8, "Lignes avec Incertitude manquante", nb_incert, bold=True)
    put(9, "   dont parametres sous-traites (_ST)", nb_incert_st)
    put(10, "   dont parametres analyses en interne", nb_incert_int)
    put(11, "", "")
    put(12, "Lignes avec LQ et Incertitude manquantes", nb_both, bold=True)
    put(13, "", "")
    r = 14
    for point, n in sorted(par_point.items()):
        put(r, f"   dont point de mesure : {point}", n)
        r += 1
    r += 1
    put(r, "NB : les parametres sous-traites (suffixe _ST) dependent du",  "")
    put(r + 1, "sous-traitant et ne peuvent pas etre completes depuis un",  "")
    put(r + 2, "referentiel labo fixe (LQ/Incertitude variables par lot).", "")

    ws = wb.create_sheet("A_completer")
    ws.sheet_view.showGridLines = False
    columns = [
        ("NumeroPointMesure", 16), ("LbPointMesure", 18), ("DatePrlvt", 14),
        ("CODE_PARAMETRE", 40), ("Parametre (libelle)", 40), ("CODE_METHODE", 30),
        ("Sous-traite ?", 12),
        ("RsAnalyse", 12), ("CdRemAnalyse", 12), ("CdUniteMesure", 14),
        ("LQAna manquante ?", 16), ("IncertAna manquante ?", 18),
        ("LQ", 16), ("Incertitude", 14),
    ]
    for i, (h, w) in enumerate(columns, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = header_font, header_fill, wrap_center, border_all
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    row_i = 2
    for row, lq_missing, incert_missing in missing_rows:
        values = [
            row["NumeroPointMesure"], row["LbPointMesure"], row["DatePrlvt"],
            row.get("_src_code_parametre") or "", row.get("_src_analyse_denomination") or "",
            row.get("_src_code_methode") or "",
            "OUI" if row.get("_src_sous_traite") else "",
            row["RsAnalyse"], row["CdRemAnalyse"], row["CdUniteMesure"],
            "OUI" if lq_missing else "", "OUI" if incert_missing else "",
            "", "",
        ]
        for ci, val in enumerate(values, start=1):
            cell = ws.cell(row=row_i, column=ci, value=val)
            cell.font = normal_font
            cell.border = border_all
            cell.alignment = wrap_center
            if ci in (13, 14) and ((ci == 13 and lq_missing) or (ci == 14 and incert_missing)):
                cell.fill = fill_a_saisir
        row_i += 1

    wb.save(out_path)
    return len(missing_rows)


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------
class ProfileDialog(tk.Toplevel):
    def __init__(self, parent, initial=None, on_saved=None):
        super().__init__(parent)
        self.title("Profil de station")
        self.geometry("560x640")
        self.on_saved = on_saved
        self.vars = {}

        canvas = tk.Canvas(self, borderwidth=0)
        frame = ttk.Frame(canvas)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        initial = initial or {}
        for i, (key, label) in enumerate(PROFILE_FIELDS):
            ttk.Label(frame, text=label, wraplength=260).grid(row=i, column=0, sticky="w", padx=8, pady=4)
            var = tk.StringVar(value=initial.get(key, PROFILE_DEFAULTS.get(key, "")))
            ttk.Entry(frame, textvariable=var, width=32).grid(row=i, column=1, padx=8, pady=4)
            self.vars[key] = var

        btns = ttk.Frame(self)
        btns.pack(side="bottom", fill="x", pady=8)
        ttk.Button(btns, text="Enregistrer", command=self.on_save).pack(side="right", padx=8)
        ttk.Button(btns, text="Annuler", command=self.destroy).pack(side="right")

    def on_save(self):
        data = {k: v.get().strip() for k, v in self.vars.items()}
        if not data["nom_profil"]:
            messagebox.showwarning("Nom manquant", "Donnez un nom au profil.", parent=self)
            return
        try:
            save_profile(data)
        except Exception as e:
            messagebox.showerror("Erreur", str(e), parent=self)
            return
        if self.on_saved:
            self.on_saved()
        self.destroy()


class App:
    def __init__(self, root):
        self.root = root
        root.title("SANDRE - Etape 1 : extraction brute -> fichier Ã  plat")
        root.geometry("820x640")
        root.minsize(700, 500)

        self.paths = {
            "extraction": tk.StringVar(),
            "parametres": tk.StringVar(value=os.path.join(app_dir(), "Parametres.xlsx")),
            "unites": tk.StringVar(value=os.path.join(app_dir(), "UnitÃ©s.xlsx")),
            "fractions": tk.StringVar(value=os.path.join(app_dir(), "Fractions.xlsx")),
            "methodes": tk.StringVar(value=os.path.join(app_dir(), "Methodes.xlsx")),
            "referentiel": tk.StringVar(value=os.path.join(app_dir(), "LQ_Incertitude.xlsx")),
            "points_stations": tk.StringVar(),
        }
        self.date_debut = tk.StringVar()
        self.date_fin = tk.StringVar()
        self.profile_var = tk.StringVar()

        files_frame = ttk.LabelFrame(root, text="Fichiers")
        files_frame.pack(fill="x", padx=10, pady=8)
        self._file_row(files_frame, 0, "Extraction brute du labo (.xlsx)", "extraction")
        self._file_row(files_frame, 1, "Parametres.xlsx", "parametres")
        self._file_row(files_frame, 2, "Unites.xlsx", "unites")
        self._file_row(files_frame, 3, "Fractions.xlsx", "fractions")
        self._file_row(files_frame, 4, "Methodes.xlsx", "methodes")
        self._file_row(files_frame, 5, "RÃ©fÃ©rentiel LQ/Incertitude (optionnel)", "referentiel")
        self._file_row(files_frame, 6, "RÃ©fÃ©rence multi-stations (optionnel)", "points_stations")

        prof_frame = ttk.LabelFrame(root, text="Station")
        prof_frame.pack(fill="x", padx=10, pady=8)
        ttk.Label(prof_frame, text="Profil :").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        self.profile_combo = ttk.Combobox(prof_frame, textvariable=self.profile_var, state="readonly", width=30)
        self.profile_combo.grid(row=0, column=1, padx=8, pady=6, sticky="w")
        ttk.Button(prof_frame, text="Nouveau...", command=self.new_profile).grid(row=0, column=2, padx=4)
        ttk.Button(prof_frame, text="Modifier...", command=self.edit_profile).grid(row=0, column=3, padx=4)
        ttk.Button(prof_frame, text="Actualiser", command=self.refresh_profiles).grid(row=0, column=4, padx=4)

        ttk.Label(prof_frame, text="Date dÃ©but pÃ©riode (AAAA-MM-JJ) :").grid(row=1, column=0, padx=8, pady=6, sticky="w")
        ttk.Entry(prof_frame, textvariable=self.date_debut, width=16).grid(row=1, column=1, padx=8, pady=6, sticky="w")
        ttk.Label(prof_frame, text="Date fin pÃ©riode (AAAA-MM-JJ) :").grid(row=1, column=2, padx=8, pady=6, sticky="w")
        ttk.Entry(prof_frame, textvariable=self.date_fin, width=16).grid(row=1, column=3, padx=8, pady=6, sticky="w")

        action = ttk.Frame(root)
        action.pack(fill="x", padx=10)
        self.gen_btn = ttk.Button(action, text="GÃ©nÃ©rer le fichier Ã  plat", command=self.on_generate)
        self.gen_btn.pack(side="left")
        self.open_folder_btn = ttk.Button(action, text="Ouvrir le dossier du rÃ©sultat",
                                           command=self.open_folder, state="disabled")
        self.open_folder_btn.pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="En attente...")
        ttk.Label(root, textvariable=self.status_var).pack(fill="x", padx=10, pady=(4, 0))

        log_frame = ttk.Frame(root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled", bg="#f5f5f5")
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.last_output_dir = None
        self.refresh_profiles()

    def _file_row(self, parent, row, label, key):
        ttk.Label(parent, text=label, width=28).grid(row=row, column=0, padx=8, pady=4, sticky="w")
        ttk.Entry(parent, textvariable=self.paths[key], width=55).grid(row=row, column=1, padx=8, pady=4)
        ttk.Button(parent, text="Parcourir...", command=lambda k=key: self.browse(k)).grid(row=row, column=2, padx=4)

    def browse(self, key):
        path = filedialog.askopenfilename(title="Choisir un fichier", filetypes=[("Classeur Excel", "*.xlsx")],
                                           parent=self.root)
        if path:
            self.paths[key].set(path)

    def refresh_profiles(self):
        names = list_profiles()
        self.profile_combo["values"] = names
        if names and not self.profile_var.get():
            self.profile_var.set(names[0])

    def new_profile(self):
        ProfileDialog(self.root, on_saved=self.refresh_profiles)

    def edit_profile(self):
        name = self.profile_var.get()
        if not name:
            messagebox.showwarning("Aucun profil", "Choisissez d'abord un profil.", parent=self.root)
            return
        data = load_profile(name)
        ProfileDialog(self.root, initial=data, on_saved=self.refresh_profiles)

    def open_folder(self):
        if self.last_output_dir:
            os.startfile(self.last_output_dir)

    def log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", str(msg) + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.root.update_idletasks()

    def on_generate(self):
        extraction = self.paths["extraction"].get().strip()
        parametres = self.paths["parametres"].get().strip()
        unites = self.paths["unites"].get().strip()
        fractions = self.paths["fractions"].get().strip()
        methodes = self.paths["methodes"].get().strip()
        referentiel = self.paths["referentiel"].get().strip()
        points_stations = self.paths["points_stations"].get().strip()
        for label, p in [("extraction", extraction), ("Parametres.xlsx", parametres),
                          ("Unites.xlsx", unites), ("Fractions.xlsx", fractions),
                          ("Methodes_Labo_Sandre.xlsx", methodes)]:
            if not p or not os.path.isfile(p):
                messagebox.showerror("Fichier manquant", f"Fichier introuvable : {label}\n{p}", parent=self.root)
                return
        if referentiel and not os.path.isfile(referentiel):
            messagebox.showerror("Fichier manquant", f"Fichier introuvable : RÃ©fÃ©rentiel LQ/Incertitude\n{referentiel}", parent=self.root)
            return
        if points_stations and not os.path.isfile(points_stations):
            messagebox.showerror("Fichier manquant", f"Fichier introuvable : RÃ©fÃ©rence multi-stations\n{points_stations}", parent=self.root)
            return
        if not self.profile_var.get():
            messagebox.showwarning("Profil manquant", "Choisissez ou crÃ©ez un profil de station.", parent=self.root)
            return
        if not self.date_debut.get().strip() or not self.date_fin.get().strip():
            messagebox.showwarning("Dates manquantes", "Renseignez les dates de dÃ©but et de fin de pÃ©riode.", parent=self.root)
            return

        self.gen_btn.configure(state="disabled")
        self.open_folder_btn.configure(state="disabled")
        self.status_var.set("GÃ©nÃ©ration en cours...")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        profile = load_profile(self.profile_var.get())
        thread = threading.Thread(
            target=self._run,
            args=(extraction, parametres, unites, fractions, methodes, profile,
                  self.date_debut.get().strip(), self.date_fin.get().strip(), referentiel or None,
                  points_stations or None),
            daemon=True,
        )
        thread.start()

    def _run(self, extraction, parametres, unites, fractions, methodes, profile, date_debut, date_fin,
              referentiel=None, points_stations=None):
        try:
            output_rows, issues = run_generation(
                extraction, parametres, unites, fractions, methodes, profile, date_debut, date_fin, self.log,
                referentiel_path=referentiel, points_stations_path=points_stations
            )
            base_dir = os.path.dirname(extraction)
            safe_name = re.sub(r"[^A-Za-z0-9_\-]", "_", profile.get("nom_profil", "Station"))
            out_xlsx = os.path.join(base_dir, f"Resultats_{safe_name}_{date_debut}_{date_fin}.xlsx")
            out_issues = os.path.join(base_dir, f"Resultats_{safe_name}_{date_debut}_{date_fin}_a_verifier.xlsx")
            out_completer = os.path.join(base_dir, f"Resultats_{safe_name}_{date_debut}_{date_fin}_LQ_Incertitude_a_completer.xlsx")
            write_resultats_xlsx(output_rows, out_xlsx)
            if issues:
                write_issues_xlsx(issues, out_issues)
            else:
                out_issues = None
            nb_a_completer = write_completion_request_xlsx(output_rows, out_completer)
            if nb_a_completer == 0:
                out_completer = None
            else:
                self.log(f"{nb_a_completer} ligne(s) sans LQ/Incertitude : voir {os.path.basename(out_completer)}")
        except Exception as e:
            self.log(f"\nERREUR : {e}")
            self.log(traceback.format_exc())
            self.root.after(0, lambda: self._on_done(False, None))
            return
        self.root.after(0, lambda: self._on_done(True, (out_xlsx, out_issues, out_completer, len(issues))))

    def _on_done(self, success, result):
        self.gen_btn.configure(state="normal")
        if success:
            out_xlsx, out_issues, out_completer, nb_issues = result
            self.last_output_dir = os.path.dirname(out_xlsx)
            self.open_folder_btn.configure(state="normal")
            self.status_var.set(f"TerminÃ© : {out_xlsx} ({nb_issues} point(s) Ã  vÃ©rifier)")
            msg = f"Fichier gÃ©nÃ©rÃ© :\n{out_xlsx}"
            if out_issues:
                msg += f"\n\n{nb_issues} point(s) Ã  vÃ©rifier :\n{out_issues}"
            if out_completer:
                msg += f"\n\nLQ/Incertitude manquantes Ã  faire complÃ©ter par le labo :\n{out_completer}"
            messagebox.showinfo("TerminÃ©", msg, parent=self.root)
        else:
            self.status_var.set("Ã‰chec, voir le journal.")
            messagebox.showerror("Erreur", "La gÃ©nÃ©ration a Ã©chouÃ©, voir le dÃ©tail dans la fenÃªtre.", parent=self.root)


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
