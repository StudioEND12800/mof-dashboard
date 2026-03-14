"""
normalize_mof.py
Normalise Département (+ numéro) et Métier pour tri facile.
"""
import pandas as pd
import unicodedata
import re
from pathlib import Path

XLSX = Path("/Users/nicolas/projets_mof/output/MOF_France_FUSIONNÉ.xlsx")

# ─────────────────────────────────────────
# 1. TABLE DÉPARTEMENT → NUMÉRO + NOM OFFICIEL
# ─────────────────────────────────────────
DEPT_MAP = {
    # num : (nom_officiel, [alias])
    "01": ("Ain",                      ["ain"]),
    "02": ("Aisne",                    ["aisne"]),
    "03": ("Allier",                   ["allier"]),
    "04": ("Alpes-de-Haute-Provence",  ["alpes-de-haute-provence","alpes de haute provence"]),
    "05": ("Hautes-Alpes",             ["hautes-alpes","hautes alpes"]),
    "06": ("Alpes-Maritimes",          ["alpes-maritimes","alpes maritimes"]),
    "07": ("Ardèche",                  ["ardeche","ardèche"]),
    "08": ("Ardennes",                 ["ardennes"]),
    "09": ("Ariège",                   ["ariege","ariège"]),
    "10": ("Aube",                     ["aube"]),
    "11": ("Aude",                     ["aude"]),
    "12": ("Aveyron",                  ["aveyron"]),
    "13": ("Bouches-du-Rhône",         ["bouches-du-rhone","bouches-du-rhône","bouches du rhone"]),
    "14": ("Calvados",                 ["calvados"]),
    "15": ("Cantal",                   ["cantal"]),
    "16": ("Charente",                 ["charente"]),
    "17": ("Charente-Maritime",        ["charente-maritime","charente maritime"]),
    "18": ("Cher",                     ["cher"]),
    "19": ("Corrèze",                  ["correze","corrèze"]),
    "2A": ("Corse-du-Sud",             ["corse-du-sud","corse du sud"]),
    "2B": ("Haute-Corse",              ["haute-corse","haute corse"]),
    "21": ("Côte-d'Or",                ["cote-d'or","côte-d'or","cote d or","côte d'or"]),
    "22": ("Côtes-d'Armor",            ["cotes-d'armor","côtes-d'armor","cotes d armor","côtes d'armor"]),
    "23": ("Creuse",                   ["creuse"]),
    "24": ("Dordogne",                 ["dordogne"]),
    "25": ("Doubs",                    ["doubs"]),
    "26": ("Drôme",                    ["drome","drôme"]),
    "27": ("Eure",                     ["eure"]),
    "28": ("Eure-et-Loir",             ["eure-et-loir","eure et loir"]),
    "29": ("Finistère",                ["finistere","finistère"]),
    "30": ("Gard",                     ["gard"]),
    "31": ("Haute-Garonne",            ["haute-garonne","haute garonne"]),
    "32": ("Gers",                     ["gers"]),
    "33": ("Gironde",                  ["gironde"]),
    "34": ("Hérault",                  ["herault","hérault"]),
    "35": ("Ille-et-Vilaine",          ["ille-et-vilaine","ille et vilaine"]),
    "36": ("Indre",                    ["indre"]),
    "37": ("Indre-et-Loire",           ["indre-et-loire","indre et loire"]),
    "38": ("Isère",                    ["isere","isère"]),
    "39": ("Jura",                     ["jura"]),
    "40": ("Landes",                   ["landes"]),
    "41": ("Loir-et-Cher",             ["loir-et-cher","loir et cher"]),
    "42": ("Loire",                    ["loire"]),
    "43": ("Haute-Loire",              ["haute-loire","haute loire"]),
    "44": ("Loire-Atlantique",         ["loire-atlantique","loire atlantique"]),
    "45": ("Loiret",                   ["loiret"]),
    "46": ("Lot",                      ["lot"]),
    "47": ("Lot-et-Garonne",           ["lot-et-garonne","lot et garonne"]),
    "48": ("Lozère",                   ["lozere","lozère"]),
    "49": ("Maine-et-Loire",           ["maine-et-loire","maine et loire"]),
    "50": ("Manche",                   ["manche"]),
    "51": ("Marne",                    ["marne"]),
    "52": ("Haute-Marne",              ["haute-marne","haute marne"]),
    "53": ("Mayenne",                  ["mayenne"]),
    "54": ("Meurthe-et-Moselle",       ["meurthe-et-moselle","meurthe et moselle"]),
    "55": ("Meuse",                    ["meuse"]),
    "56": ("Morbihan",                 ["morbihan"]),
    "57": ("Moselle",                  ["moselle"]),
    "58": ("Nièvre",                   ["nievre","nièvre"]),
    "59": ("Nord",                     ["nord"]),
    "60": ("Oise",                     ["oise"]),
    "61": ("Orne",                     ["orne"]),
    "62": ("Pas-de-Calais",            ["pas-de-calais","pas de calais"]),
    "63": ("Puy-de-Dôme",              ["puy-de-dome","puy-de-dôme","puy de dome"]),
    "64": ("Pyrénées-Atlantiques",     ["pyrenees-atlantiques","pyrénées-atlantiques","pyrenees atlantiques"]),
    "65": ("Hautes-Pyrénées",          ["hautes-pyrenees","hautes-pyrénées","hautes pyrenees"]),
    "66": ("Pyrénées-Orientales",      ["pyrenees-orientales","pyrénées-orientales"]),
    "67": ("Bas-Rhin",                 ["bas-rhin","bas rhin"]),
    "68": ("Haut-Rhin",                ["haut-rhin","haut rhin"]),
    "69": ("Rhône",                    ["rhone","rhône"]),
    "70": ("Haute-Saône",              ["haute-saone","haute-saône","haute saone"]),
    "71": ("Saône-et-Loire",           ["saone-et-loire","saône-et-loire","saone et loire"]),
    "72": ("Sarthe",                   ["sarthe"]),
    "73": ("Savoie",                   ["savoie"]),
    "74": ("Haute-Savoie",             ["haute-savoie","haute savoie"]),
    "75": ("Paris",                    ["paris"]),
    "76": ("Seine-Maritime",           ["seine-maritime","seine maritime"]),
    "77": ("Seine-et-Marne",           ["seine-et-marne","seine et marne"]),
    "78": ("Yvelines",                 ["yvelines"]),
    "79": ("Deux-Sèvres",              ["deux-sevres","deux-sèvres","deux sevres"]),
    "80": ("Somme",                    ["somme"]),
    "81": ("Tarn",                     ["tarn"]),
    "82": ("Tarn-et-Garonne",          ["tarn-et-garonne","tarn et garonne"]),
    "83": ("Var",                      ["var"]),
    "84": ("Vaucluse",                 ["vaucluse"]),
    "85": ("Vendée",                   ["vendee","vendée"]),
    "86": ("Vienne",                   ["vienne"]),
    "87": ("Haute-Vienne",             ["haute-vienne","haute vienne"]),
    "88": ("Vosges",                   ["vosges"]),
    "89": ("Yonne",                    ["yonne"]),
    "90": ("Territoire de Belfort",    ["territoire de belfort","belfort"]),
    "91": ("Essonne",                  ["essonne"]),
    "92": ("Hauts-de-Seine",           ["hauts-de-seine","hauts de seine"]),
    "93": ("Seine-Saint-Denis",        ["seine-saint-denis","seine-st-denis","seine st denis","seine saint denis"]),
    "94": ("Val-de-Marne",             ["val-de-marne","val de marne"]),
    "95": ("Val-d'Oise",               ["val-d'oise","val-d'oise","val d'oise","val-d oise","val d oise"]),
    "971": ("Guadeloupe",              ["guadeloupe"]),
    "972": ("Martinique",              ["martinique"]),
    "973": ("Guyane",                  ["guyane"]),
    "974": ("La Réunion",              ["reunion","réunion","la reunion","la réunion"]),
    "976": ("Mayotte",                 ["mayotte"]),
}

# Construire l'index de lookup alias→(num, nom_officiel)
dept_lookup = {}  # alias_normalisé → (num, nom)
for num, (nom, aliases) in DEPT_MAP.items():
    dept_lookup[nom.lower()] = (num, nom)
    for alias in aliases:
        dept_lookup[alias.lower()] = (num, nom)

def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def lookup_dept(raw):
    if not raw or not str(raw).strip():
        return ("", "")
    s = str(raw).strip()
    # Supprimer ".0" que pandas peut ajouter (ex: "28.0" → "28")
    s = re.sub(r'\.0$', '', s)
    # Déjà un numéro de département (1 ou 2 chiffres, ou "2A"/"2B")
    if re.match(r'^\d+$', s):
        num = s.zfill(2)  # "5"→"05", "28"→"28", "971"→"971"
        if len(s) >= 3:
            num = s  # DOM-TOM 971-976
        if num in DEPT_MAP:
            return (num, DEPT_MAP[num][0])
    if s.upper() in ("2A", "2B"):
        num = s.upper()
        return (num, DEPT_MAP[num][0])
    sl = s.lower()
    if sl in dept_lookup:
        return dept_lookup[sl]
    # Sans accents
    s_na = strip_accents(sl)
    for key, val in dept_lookup.items():
        if strip_accents(key) == s_na:
            return val
    return ("", s)  # inconnu → garder tel quel

# ─────────────────────────────────────────
# 2. NORMALISATION MÉTIER
# ─────────────────────────────────────────

# Mapping de variantes connues → forme canonique
METIER_CANON = {
    # Coiffure
    "coiffure": "Coiffure",
    "coiffure mixte": "Coiffure",
    # Pâtisserie
    "patisserie confiserie": "Pâtisserie-confiserie",
    "pâtisserie confiserie": "Pâtisserie-confiserie",
    "pâtisserie, confiserie": "Pâtisserie-confiserie",
    "patisserie, confiserie": "Pâtisserie-confiserie",
    "pâtissier-confiseur": "Pâtisserie-confiserie",
    "patissier-confiseur": "Pâtisserie-confiserie",
    "pâtissier confiseur": "Pâtisserie-confiserie",
    "confiserie pâtisserie": "Pâtisserie-confiserie",
    "confiserie patisserie": "Pâtisserie-confiserie",
    # Boulangerie
    "boulanger": "Boulangerie",
    "boulangerie": "Boulangerie",
    "boulangerie viennoiserie": "Boulangerie-viennoiserie",
    "boulangerie-viennoiserie": "Boulangerie-viennoiserie",
    # Cuisine
    "cuisine gastronomie": "Cuisine, gastronomie",
    "cuisine, gastronomie": "Cuisine, gastronomie",
    "cuisinier gastronomie": "Cuisine, gastronomie",
    "cuisine froide": "Cuisine froide",
    # Fabrication mécanique
    "fabrication mecanique": "Fabrication mécanique",
    "fabrication mécanique": "Fabrication mécanique",
    # Chaudronnerie
    "chaudronnerie": "Chaudronnerie",
    # Pierre
    "metiers de la pierre": "Métiers de la pierre",
    "métiers de la pierre": "Métiers de la pierre",
    # Menuiserie
    "menuiserie": "Menuiserie",
    "menuiserie bois": "Menuiserie",
    # Broderie
    "broderie main": "Broderie main",
    "broderie a la main": "Broderie main",
    # Photographie
    "photographie": "Photographie",
    # Verrerie
    "verrerie cristallerie": "Verrerie-cristallerie",
    "verrerie-cristallerie": "Verrerie-cristallerie",
    # Art floral
    "art floral": "Art floral",
    # Ébénisterie
    "ebenisterie": "Ébénisterie",
    "ébénisterie": "Ébénisterie",
    # Carrelage
    "carrelage mosaique": "Carrelage-mosaïque",
    "carrelage - mosaïque": "Carrelage-mosaïque",
    "carrelage mosaïque": "Carrelage-mosaïque",
    # Charcuterie
    "charcutier traiteur": "Charcutier-traiteur",
    "charcutier-traiteur": "Charcutier-traiteur",
    "charcuterie": "Charcutier-traiteur",
    # Esthétique
    "esthetique art du maquillage": "Esthétique, art du maquillage",
    "esthétique, art du maquillage": "Esthétique, art du maquillage",
    "esthetique cosmétique": "Esthétique, art du maquillage",
    # Sommellerie
    "sommellerie": "Sommellerie",
    # Maîtrise hôtellerie / service
    "maitre d hotel du service et des arts de la table": "Maître d'hôtel, service et arts de la table",
    "maître d'hôtel, du service et des arts de la table": "Maître d'hôtel, service et arts de la table",
    "service en salle": "Maître d'hôtel, service et arts de la table",
    # Optique
    "actions commerciales en optique lunetterie": "Optique-lunetterie",
    "opticien lunettier": "Optique-lunetterie",
    "optique lunetterie": "Optique-lunetterie",
    # Génie climatique
    "genie climatique chauffage": "Génie climatique, chauffage",
    "génie climatique, chauffage": "Génie climatique, chauffage",
    "froid et climatisation": "Génie climatique, chauffage",
    # Couverture
    "couverture ornemaniste metallique": "Couverture-ornemaniste métallique",
    "couverture-ornemaniste métallique": "Couverture-ornemaniste métallique",
    # Chocolaterie
    "chocolatier confiseur": "Chocolaterie-confiserie",
    "chocolaterie confiserie": "Chocolaterie-confiserie",
    "chocolaterie": "Chocolaterie-confiserie",
}

def normalize_metier(raw):
    """Normalise un métier vers la forme canonique."""
    if not raw or not raw.strip():
        return ""
    s = raw.strip()
    # Lookup direct (insensible à la casse et aux accents)
    key = strip_accents(s.lower()).replace("  ", " ")
    # Chercher dans le mapping
    for k, v in METIER_CANON.items():
        if strip_accents(k) == key:
            return v
    # Sinon : Title Case + préservation des accents
    # Convertir de MAJUSCULES vers Title Case si entièrement en majuscules
    if s == s.upper() and len(s) > 2:
        # Reconstituer en title case
        words = s.split()
        titled = []
        for w in words:
            if len(w) <= 2 and w in ("DE", "DU", "DES", "LA", "LE", "LES", "ET", "EN", "AU", "/"):
                titled.append(w.lower())
            else:
                titled.append(w.capitalize())
        return " ".join(titled)
    return s

def sort_key_num(num):
    """Clé de tri (string zéro-paddée) pour les numéros de département.
    01→001, 2A→020A, 2B→020B, 95→095, 971→971, ''→999
    """
    if not num:
        return "999"
    if num == "2A":
        return "020A"
    if num == "2B":
        return "020B"
    try:
        return str(int(num)).zfill(3)
    except:
        return "998"

# ─────────────────────────────────────────
# 3. CHARGEMENT ET NORMALISATION
# ─────────────────────────────────────────
print("Chargement Excel...")
df = pd.read_excel(XLSX, dtype=str)
df.fillna("", inplace=True)
print(f"  {len(df)} lignes, {len(df.columns)} colonnes")

# Normaliser département
print("Normalisation département...")
dept_results = df["Département"].apply(lookup_dept)
df["Num_Dept"] = dept_results.apply(lambda x: x[0])
df["Département"] = dept_results.apply(lambda x: x[1] if x[1] else df.loc[dept_results.index == dept_results.index[dept_results.tolist().index(x)], "Département"].values[0] if x[1] == "" else x[1])

# Recalculer proprement
nums = []
noms = []
for raw in df["Département"].values:
    num, nom = lookup_dept(raw)
    nums.append(num)
    noms.append(nom if nom else raw.strip())
df["Num_Dept"] = nums
df["Département"] = noms

# Corriger la Région "70" (Haute-Saône mal classée)
df.loc[df["Région"] == "70", "Région"] = ""

# Normaliser métier
print("Normalisation métier...")
df["Métier"] = df["Métier"].apply(normalize_metier)

# Stats après normalisation
print(f"\nDépartements résolus: {(df['Num_Dept'].ne('')).sum()} / {len(df)}")
print(f"Métiers uniques avant: 356 → après: {df['Métier'].str.strip().ne('').sum()} non-vides, {df['Métier'].nunique()} uniques")

# Réorganiser les colonnes : mettre Num_Dept juste avant Département
cols = df.columns.tolist()
cols.remove("Num_Dept")
dept_idx = cols.index("Département")
cols.insert(dept_idx, "Num_Dept")
df = df[cols]

# Trier : par Num_Dept puis Nom
df["_sort_dept"] = df["Num_Dept"].apply(sort_key_num)
df = df.sort_values(["_sort_dept", "Nom", "Promotion"])
df = df.drop(columns=["_sort_dept"])
df = df.reset_index(drop=True)

# Aperçu
print(f"\nTop métiers après normalisation:")
print(df["Métier"].str.strip().value_counts().head(20).to_string())
print(f"\nDépartements inconnus (sans numéro):")
unknown = df[df["Num_Dept"] == ""]["Département"].value_counts()
if len(unknown):
    print(unknown.head(20).to_string())
else:
    print("  Aucun !")

# Sauvegarder
print(f"\nSauvegarde...")
with pd.ExcelWriter(XLSX, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="MOFs", index=False)
    # Feuille stats métiers
    metier_stats = df["Métier"].str.strip().value_counts().reset_index()
    metier_stats.columns = ["Métier", "Nb MOFs"]
    metier_stats.to_excel(writer, sheet_name="Stats Métiers", index=False)
    # Feuille stats départements
    dept_stats = df.groupby(["Num_Dept", "Département"]).size().reset_index(name="Nb MOFs")
    dept_stats = dept_stats[dept_stats["Département"].ne("")].sort_values("Num_Dept")
    dept_stats.to_excel(writer, sheet_name="Stats Départements", index=False)

    # Mise en forme feuille principale
    ws = writer.sheets["MOFs"]
    ws.freeze_panes = "A2"           # Ligne d'en-têtes figée
    ws.auto_filter.ref = ws.dimensions  # Filtres automatiques sur toutes les colonnes
    # Largeur colonnes
    col_widths = {
        "A": 22, "B": 18, "C": 30, "D": 28, "E": 12, "F": 10,
        "G": 22, "H": 10, "I": 22, "J": 18, "K": 12, "L": 30,
        "M": 18, "N": 12,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

print(f"✅ Sauvegardé: {XLSX}")
print(f"   {len(df)} MOFs, colonnes: {df.columns.tolist()}")
