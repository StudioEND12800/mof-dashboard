"""
dashboard.py  –  Tableau de bord MOF France
Lancer : streamlit run dashboard.py
"""

import re
import unicodedata
import base64
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from pathlib import Path
from PIL import Image

# ──────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────
_MEDAILLE_PATH = Path(__file__).parent / "assets" / "medaille-mof.png"
_medaille_icon = Image.open(_MEDAILLE_PATH) if _MEDAILLE_PATH.exists() else "🏆"
_medaille_b64 = base64.b64encode(_MEDAILLE_PATH.read_bytes()).decode() if _MEDAILLE_PATH.exists() else ""

st.set_page_config(
    page_title="MOF France – Tableau de bord",
    page_icon=_medaille_icon,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────
# ACCÈS PRIVÉ – MOT DE PASSE
# ──────────────────────────────────────────
import hmac

def _check_password() -> bool:
    """Affiche un champ mot de passe. Retourne True si correct."""
    if st.session_state.get("_auth_ok"):
        return True

    # Pas de secret configuré → accès libre (dev local sans secrets.toml)
    try:
        expected = st.secrets["password"]
    except (FileNotFoundError, KeyError):
        return True

    with st.container():
        st.markdown(
            "<div style='max-width:360px; margin:80px auto 0 auto'>",
            unsafe_allow_html=True,
        )
        if _medaille_b64:
            st.markdown(
                f'<div style="text-align:center; margin-bottom:8px;">'  
                f'<img src="data:image/png;base64,{_medaille_b64}" width="60">'  
                f'</div>',
                unsafe_allow_html=True,
            )
        st.markdown("### MOF France – Accès privé")
        pwd = st.text_input("Mot de passe", type="password", key="_pwd_input")
        if st.button("Connexion", use_container_width=True):
            if hmac.compare_digest(pwd, expected):
                st.session_state["_auth_ok"] = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")
        st.markdown("</div>", unsafe_allow_html=True)
    return False

if not _check_password():
    st.stop()

EXCEL = Path(__file__).parent / "output" / "MOF_France_FUSIONNÉ.xlsx"

# ──────────────────────────────────────────
# CHARGEMENT DES DONNÉES
# ──────────────────────────────────────────
@st.cache_data
def load_data():
    df = pd.read_excel(EXCEL, dtype=str).fillna("")
    df["Promotion"] = pd.to_numeric(df["Promotion"], errors="coerce")
    df["_nom_lower"]    = df["Nom complet"].str.lower().fillna("")
    df["_metier_lower"] = df["Métier"].str.lower().fillna("")
    df["_dept_lower"]   = df["Département"].str.lower().fillna("")
    df["_region_lower"] = df["Région"].str.lower().fillna("")
    df["_ville_lower"]  = df["Ville"].str.lower().fillna("")
    return df

df_full = load_data()

# ──────────────────────────────────────────
# MOTEUR DE QUESTIONS EN FRANÇAIS
# ──────────────────────────────────────────

def strip_acc(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    ).lower()

# Valeurs possibles (sans accents pour matching)
ALL_REGIONS  = {strip_acc(r): r for r in df_full["Région"].unique() if r}
ALL_DEPTS    = {strip_acc(d): d for d in df_full["Département"].unique() if d}
ALL_METIERS  = {strip_acc(m): m for m in df_full["Métier"].unique() if m}
ALL_VILLES   = {strip_acc(v): v for v in df_full["Ville"].unique() if v}

# Alias fréquents région / département
ALIASES = {
    # Régions
    "occitanie": "Occitanie",
    "idf": "Île-de-France",
    "ile-de-france": "Île-de-France",
    "ile de france": "Île-de-France",
    "paca": "Provence-Alpes-Côte d'Azur",
    "aura": "Auvergne-Rhône-Alpes",
    "auvergne rhone alpes": "Auvergne-Rhône-Alpes",
    "bretagne": "Bretagne",
    "normandie": "Normandie",
    "hauts de france": "Hauts-de-France",
    "grand est": "Grand Est",
    "nouvelle aquitaine": "Nouvelle-Aquitaine",
    "pays de la loire": "Pays de la Loire",
    "bourgogne franche comte": "Bourgogne-Franche-Comté",
    "centre val de loire": "Centre-Val de Loire",
    # Métiers alias
    "coiffeur": "Coiffure",
    "coiffeurs": "Coiffure",
    "boulanger": "Boulangerie",
    "boulangers": "Boulangerie",
    "patissier": "Pâtisserie-confiserie",
    "patissiers": "Pâtisserie-confiserie",
    "cuisinier": "Cuisine, gastronomie",
    "cuisiniers": "Cuisine, gastronomie",
    "chef": "Cuisine, gastronomie",
    "chefs": "Cuisine, gastronomie",
    "charcutier": "Charcutier-traiteur",
    "menuisier": "Menuiserie",
    "menuisiers": "Menuiserie",
    "ebeniste": "Ébénisterie",
    "photographe": "Photographie",
    "sommelier": "Sommellerie",
    "sommeliers": "Sommellerie",
    "chocolatier": "Chocolaterie-confiserie",
    "fleuriste": "Art floral",
    "optique": "Optique-lunetterie",
    "opticien": "Optique-lunetterie",
    "brodeur": "Broderie main",
    "verrier": "Verrerie-cristallerie",
    "tonnelier": "Tonnellerie",
    "chaudronnier": "Chaudronnerie",
}

# Mapping département → région (pour répondre à "combien dans le Rhône")
DEPT_TO_REGION = {
    row["Département"]: row["Région"]
    for _, row in df_full[["Département", "Région"]].drop_duplicates().iterrows()
    if row["Département"] and row["Région"]
}


def parse_question(question: str) -> dict:
    """Extrait filtre(s) d'une question en français. Retourne dict de filtres."""
    q_raw  = question.strip()
    q      = strip_acc(q_raw)
    result = {"region": None, "dept": None, "metier": None, "annee": None,
              "ville": None, "raw": q_raw}

    # ── Années / promotions ──────────────────────────────────────────────
    annees = re.findall(r"\b(19[0-9]{2}|20[0-2][0-9])\b", q)
    if len(annees) == 1:
        result["annee"] = int(annees[0])
    elif len(annees) == 2:
        result["annee_range"] = (int(annees[0]), int(annees[1]))

    # ── Alias directs ────────────────────────────────────────────────────
    for alias, canon in ALIASES.items():
        if alias in q:
            # Détecter si c'est une région, métier…
            if canon in ALL_REGIONS.values():
                result["region"] = canon
            elif canon in ALL_METIERS.values():
                result["metier"] = canon

    # ── Régions ──────────────────────────────────────────────────────────
    if not result["region"]:
        for key, val in ALL_REGIONS.items():
            if key in q:
                result["region"] = val
                break

    # ── Départements ─────────────────────────────────────────────────────
    for key, val in ALL_DEPTS.items():
        if key and key in q:
            result["dept"] = val
            break
    # Numéro de département (ex: "departement 69", "dept. 75")
    m = re.search(r"\bdept?(?:artement)?\s*\.?\s*(\d{2,3})\b", q)
    if m:
        num = m.group(1).zfill(2)
        # chercher le dept avec ce numéro
        matches = df_full[df_full["Num_Dept"] == num]["Département"].dropna().unique()
        if len(matches):
            result["dept"] = matches[0]

    # ── Métiers ───────────────────────────────────────────────────────────
    if not result["metier"]:
        for key, val in ALL_METIERS.items():
            if key and key in q:
                result["metier"] = val
                break

    # ── Villes ────────────────────────────────────────────────────────────
    for key, val in ALL_VILLES.items():
        if key and len(key) > 4 and key in q:
            result["ville"] = val
            break

    return result


def apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    out = df.copy()
    if filters.get("region"):
        out = out[out["_region_lower"] == strip_acc(filters["region"])]
    if filters.get("dept"):
        out = out[out["_dept_lower"] == strip_acc(filters["dept"])]
    if filters.get("metier"):
        out = out[out["_metier_lower"] == strip_acc(filters["metier"])]
    if filters.get("ville"):
        out = out[out["_ville_lower"] == strip_acc(filters["ville"])]
    if filters.get("annee"):
        out = out[out["Promotion"] == filters["annee"]]
    if filters.get("annee_range"):
        a, b = filters["annee_range"]
        out = out[out["Promotion"].between(a, b)]
    return out


def answer_question(question: str, df: pd.DataFrame) -> tuple[str, pd.DataFrame]:
    """Retourne (réponse textuelle, DataFrame filtré)."""
    filters = parse_question(question)
    sub = apply_filters(df, filters)

    # Construction de la réponse
    parts = []
    if filters["region"]:
        parts.append(f"la région **{filters['region']}**")
    if filters["dept"]:
        parts.append(f"le département **{filters['dept']}**")
    if filters["metier"]:
        parts.append(f"le métier **{filters['metier']}**")
    if filters["ville"]:
        parts.append(f"la ville de **{filters['ville']}**")
    if filters.get("annee"):
        parts.append(f"la promotion **{filters['annee']}**")
    if filters.get("annee_range"):
        parts.append(f"les promotions **{filters['annee_range'][0]}–{filters['annee_range'][1]}**")

    ctx = " · ".join(parts) if parts else "l'ensemble de la base"
    n = len(sub)
    pct = n / len(df) * 100

    if n == 0:
        msg = f"Aucun MOF trouvé pour {ctx}."
    elif n == 1:
        row = sub.iloc[0]
        msg = f"**1 MOF** pour {ctx} : {row['Nom complet']} ({row['Métier']}, promo {int(row['Promotion']) if pd.notna(row['Promotion']) else '?'})"
    else:
        msg = f"**{n} MOFs** pour {ctx} ({pct:.1f}% de la base totale)"
        # Métier dominant
        if not filters["metier"] and sub["Métier"].ne("").any():
            top_m = sub[sub["Métier"].ne("")]["Métier"].value_counts().idxmax()
            top_n = sub["Métier"].value_counts().iloc[0]
            msg += f"\n\nMétier dominant : **{top_m}** ({top_n} MOFs)"
        # Dept dominant
        if not filters["dept"] and not filters["region"] and sub["Département"].ne("").any():
            top_d = sub[sub["Département"].ne("")]["Département"].value_counts().idxmax()
            top_dn = sub["Département"].value_counts().iloc[0]
            msg += f"\nDépartement le plus représenté : **{top_d}** ({top_dn} MOFs)"

    return msg, sub, filters


# ──────────────────────────────────────────
# SIDEBAR – FILTRES MANUELS
# ──────────────────────────────────────────
with st.sidebar:
    logo_path = Path(__file__).parent / "assets" / "logo_snmof_color.png"
    if logo_path.exists():
        logo_b64 = base64.b64encode(logo_path.read_bytes()).decode()
        st.markdown(
            f'<div style="text-align:center; padding: 8px 0 4px 0;">'
            f'<img src="data:image/png;base64,{logo_b64}" width="150">'
            f'</div>',
            unsafe_allow_html=True,
        )
    else:
        if _medaille_b64:
            st.markdown(
                f'<div style="text-align:center; padding: 8px 0 4px 0;">'  
                f'<img src="data:image/png;base64,{_medaille_b64}" width="60">'  
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown("### MOF")
    st.markdown("---")
    st.title("Filtres")

    # Recherche par nom
    search_nom = st.text_input("🔍 Rechercher un nom", "", placeholder="ex : Dupont, Marie…")

    regions_opts = sorted([r for r in df_full["Région"].unique() if r])
    sel_region = st.multiselect("Région", regions_opts)

    if sel_region:
        depts_opts = sorted([d for d in df_full[df_full["Région"].isin(sel_region)]["Département"].unique() if d])
    else:
        depts_opts = sorted([d for d in df_full["Département"].unique() if d])
    sel_dept = st.multiselect("Département", depts_opts)

    metiers_opts = sorted([m for m in df_full["Métier"].unique() if m])
    sel_metier = st.multiselect("Métier", metiers_opts)

    promo_min = int(df_full["Promotion"].dropna().min())
    promo_max = int(df_full["Promotion"].dropna().max())
    sel_promo = st.slider("Promotion", promo_min, promo_max, (promo_min, promo_max))

    st.divider()
    st.caption(f"Base : **{len(df_full):,} MOFs**  •  {promo_min}–{promo_max}")

    # ── Logo Astérion + crédits ──────────────────────────
    asterion_path = Path(__file__).parent / "assets" / "logo_asterion.png"
    if asterion_path.exists():
        _ast_b64 = base64.b64encode(asterion_path.read_bytes()).decode()
        st.markdown(
            f'<div style="text-align:center; padding: 18px 0 4px 0;">'
            f'<a href="https://asterion.studio-end.com" target="_blank">'
            f'<img src="data:image/png;base64,{_ast_b64}" width="80" style="opacity:0.85; transition:opacity 0.2s;" '
            f'onmouseover="this.style.opacity=\'1\'" onmouseout="this.style.opacity=\'0.85\'">'
            f'</a>'
            f'</div>',
            unsafe_allow_html=True,
        )
    st.markdown(
        '<div style="text-align:center; font-size:0.68rem; color:#8899aa; line-height:1.6; padding-bottom:8px;">'
        'Dev par <strong>Nicolas DELSAUT</strong><br>'
        '©&nbsp;Studio END WEBDESIGN&nbsp;— Rodez'
        '</div>',
        unsafe_allow_html=True,
    )

# Filtre sidebar
df_view = df_full.copy()
if search_nom:
    _sn = search_nom.lower()
    df_view = df_view[df_view["_nom_lower"].str.contains(_sn, na=False)]
if sel_region:
    df_view = df_view[df_view["Région"].isin(sel_region)]
if sel_dept:
    df_view = df_view[df_view["Département"].isin(sel_dept)]
if sel_metier:
    df_view = df_view[df_view["Métier"].isin(sel_metier)]
df_view = df_view[
    df_view["Promotion"].isna() |
    df_view["Promotion"].between(sel_promo[0], sel_promo[1])
]

# ──────────────────────────────────────────
# HEADER PRINCIPAL
# ──────────────────────────────────────────
_hcol1, _hcol2 = st.columns([1, 9])
with _hcol1:
    if _MEDAILLE_PATH.exists():
        st.image(str(_MEDAILLE_PATH), width=72)
with _hcol2:
    st.markdown("# MOF France — Tableau de bord")
st.info(
    "**Meilleurs Ouvriers de France · Base consolidée multi-sources**  \n"
    "⚠️ Cette base est une extraction et une recompilation des informations trouvées sur le web "
    "par des robots et des scripts. Elle peut contenir des erreurs ! "
    "N'hésitez pas à nous en faire part : "
    "[contact@studio-end.fr](mailto:contact@studio-end.fr).  \n"
    "Utilisez la même adresse si vous souhaitez être effacé de cette liste (**RGPD**).",
)

# ──────────────────────────────────────────
# ZONE DE QUESTION EN LANGAGE NATUREL
# ──────────────────────────────────────────
st.markdown("### 💬 Posez une question")

EXEMPLES = [
    "Combien y a-t-il de MOF en Occitanie ?",
    "Combien de coiffeurs ?",
    "Combien de MOF dans le Rhône ?",
    "Combien de MOF boulangers en Île-de-France ?",
    "Combien de MOF en promotion 2004 ?",
    "Combien de MOF entre 1980 et 2000 ?",
    "Combien de MOF à Lyon ?",
    "Combien de MOF pâtissiers à Paris ?",
    "Combien de charcutiers en PACA ?",
    "Combien de MOF tonneliers ?",
]

col_q, col_ex = st.columns([3, 1])
with col_q:
    question = st.text_input(
        "Question",
        placeholder="Ex : combien y a-t-il de coiffeurs en Occitanie ?",
        label_visibility="collapsed",
    )
with col_ex:
    exemple_choisi = st.selectbox("Exemples", ["— choisir —"] + EXEMPLES,
                                   label_visibility="collapsed")
    if exemple_choisi != "— choisir —":
        question = exemple_choisi

if question:
    answer, df_q, flt = answer_question(question, df_view)
    st.info(answer, icon="🔎")

    if len(df_q) > 0 and len(df_q) < len(df_full):
        # Petites stats rapides sur le résultat
        sub_cols = st.columns(4)
        sub_cols[0].metric("MOFs trouvés", len(df_q))
        if df_q["Métier"].ne("").any():
            sub_cols[1].metric("Métiers distincts", df_q[df_q["Métier"].ne("")]["Métier"].nunique())
        if df_q["Département"].ne("").any():
            sub_cols[2].metric("Départements", df_q[df_q["Département"].ne("")]["Département"].nunique())
        if df_q["Promotion"].notna().any():
            sub_cols[3].metric("Avec promo connue", int(df_q["Promotion"].notna().sum()))

        # Graphiques contextuels
        c1, c2 = st.columns(2)
        if not flt["metier"] and df_q["Métier"].ne("").any():
            top_m = df_q[df_q["Métier"].ne("")]["Métier"].value_counts().head(12).reset_index()
            top_m.columns = ["Métier", "Nb"]
            fig = px.bar(top_m, x="Nb", y="Métier", orientation="h",
                         color="Nb", color_continuous_scale="Blues",
                         title="Top métiers dans la sélection")
            fig.update_layout(showlegend=False, coloraxis_showscale=False,
                              yaxis={"categoryorder": "total ascending"},
                              height=360, margin=dict(l=0, r=0, t=40, b=0))
            c1.plotly_chart(fig, width="stretch")

        if not flt["region"] and not flt["dept"] and df_q["Département"].ne("").any():
            top_d = df_q[df_q["Département"].ne("")]["Département"].value_counts().head(12).reset_index()
            top_d.columns = ["Département", "Nb"]
            fig2 = px.bar(top_d, x="Nb", y="Département", orientation="h",
                          color="Nb", color_continuous_scale="Greens",
                          title="Top départements dans la sélection")
            fig2.update_layout(showlegend=False, coloraxis_showscale=False,
                               yaxis={"categoryorder": "total ascending"},
                               height=360, margin=dict(l=0, r=0, t=40, b=0))
            c2.plotly_chart(fig2, width="stretch")

        # Table des résultats
        with st.expander(f"📋 Voir les {len(df_q)} MOFs ({min(len(df_q), 200)} affichés)", expanded=len(df_q) <= 50):
            disp_cols = ["Nom complet", "Métier", "Promotion", "Département", "Région", "Ville", "Classe"]
            st.dataframe(
                df_q[disp_cols].sort_values(["Métier", "Nom complet"]).head(200),
                width="stretch", hide_index=True
            )

st.divider()

# ──────────────────────────────────────────
# KPI GLOBAUX (vue filtrée par sidebar)
# ──────────────────────────────────────────
st.markdown(f"### 📊 Vue d'ensemble — {len(df_view):,} MOFs sélectionnés")

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total MOFs", f"{len(df_view):,}")
k2.metric("Métiers distincts", df_view[df_view["Métier"].ne("")]["Métier"].nunique())
k3.metric("Départements", df_view[df_view["Département"].ne("")]["Département"].nunique())
k4.metric("Régions", df_view[df_view["Région"].ne("")]["Région"].nunique())
k5.metric("Années de promo", int(df_view["Promotion"].nunique()))

st.markdown("---")

# ──────────────────────────────────────────
# LISTE FILTRÉE (affichée quand filtres actifs)
# ──────────────────────────────────────────
_filtered = len(df_view) < len(df_full)
if _filtered:
    with st.expander(f"📋 Liste des {len(df_view):,} MOFs sélectionnés", expanded=True):
        disp_cols = ["Num_Dept", "Département", "Région", "Nom complet", "Métier",
                     "Promotion", "Classe", "Ville", "Entreprise", "Téléphone", "Email"]
        st.dataframe(
            df_view[disp_cols].sort_values(["Num_Dept", "Nom complet"]),
            hide_index=True,
            height=400,
            use_container_width=True,
        )
        csv_bytes = df_view[disp_cols].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Télécharger la sélection (CSV)",
            data=csv_bytes,
            file_name="mof_selection.csv",
            mime="text/csv",
        )
    st.markdown("---")

# ──────────────────────────────────────────
# GRAPHIQUES PRINCIPAUX
# ──────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["🗺️ Géographie", "🔨 Métiers", "📅 Promotions", "📋 Données"])

# ── TAB 1 : GÉOGRAPHIE ───────────────────
with tab1:
    col_r, col_d = st.columns(2)

    # ── Barre Régions (cliquable) ──────────────────────────────
    reg_data = (
        df_view[df_view["Région"].ne("")]
        .groupby("Région").size().reset_index(name="Nb")
        .sort_values("Nb", ascending=True)
    )
    fig_r = px.bar(reg_data, x="Nb", y="Région", orientation="h",
                   color="Nb", color_continuous_scale="Teal",
                   title="📍 Cliquez sur une région", text="Nb")
    fig_r.update_traces(textposition="outside")
    fig_r.update_layout(coloraxis_showscale=False, showlegend=False,
                        height=420, margin=dict(l=0, r=0, t=40, b=0))
    sel_r = col_r.plotly_chart(fig_r, width="stretch",
                                on_select="rerun", key="geo_region_bar")
    clicked_region = None
    if sel_r and sel_r.selection and sel_r.selection.get("points"):
        clicked_region = sel_r.selection["points"][0].get("label") or \
                         sel_r.selection["points"][0].get("y")

    # ── Barre Top 20 Départements (cliquable) ─────────────────
    dept_data = (
        df_view[df_view["Département"].ne("")]
        .groupby(["Num_Dept", "Département"]).size().reset_index(name="Nb")
        .sort_values("Nb", ascending=False).head(20)
        .sort_values("Nb", ascending=True)
    )
    fig_d = px.bar(dept_data, x="Nb", y="Département", orientation="h",
                   color="Nb", color_continuous_scale="Blues",
                   title="📍 Cliquez sur un département", text="Nb")
    fig_d.update_traces(textposition="outside")
    fig_d.update_layout(coloraxis_showscale=False, showlegend=False,
                        height=560, margin=dict(l=0, r=0, t=40, b=0))
    sel_d = col_d.plotly_chart(fig_d, width="stretch",
                                on_select="rerun", key="geo_dept_bar")
    clicked_dept = None
    if sel_d and sel_d.selection and sel_d.selection.get("points"):
        clicked_dept = sel_d.selection["points"][0].get("label") or \
                       sel_d.selection["points"][0].get("y")

    # ── Carte choroplèthe (cliquable) ─────────────────────────
    from_all = (
        df_full[df_full["Num_Dept"].ne("") & df_full["Département"].ne("")]
        .groupby(["Num_Dept", "Département"]).size().reset_index(name="Nb")
    )
    clicked_map_dept = None
    if len(from_all) > 10:
        try:
            import json, urllib.request
            geo_cache = Path("/tmp/depts_geojson.json")
            if not geo_cache.exists():
                url = "https://raw.githubusercontent.com/gregoiredavid/france-geojson/master/departements.geojson"
                urllib.request.urlretrieve(url, geo_cache)
            with open(geo_cache) as f:
                geojson = json.load(f)

            fig_map = px.choropleth(
                from_all,
                geojson=geojson,
                locations="Num_Dept",
                featureidkey="properties.code",
                color="Nb",
                color_continuous_scale="YlOrRd",
                hover_data={"Département": True, "Nb": True},
                title="🖱️ Cliquez sur un département pour voir ses MOFs",
            )
            fig_map.update_geos(
                visible=False,
                projection_type="mercator",
                lataxis_range=[41.0, 51.5],
                lonaxis_range=[-5.5, 10.0],
            )
            fig_map.update_layout(height=720, margin=dict(l=0, r=0, t=40, b=0),
                                  clickmode="event+select")
            sel_map = st.plotly_chart(fig_map, width="stretch",
                                      on_select="rerun", key="geo_map")

            if sel_map and sel_map.selection and sel_map.selection.get("points"):
                pt = sel_map.selection["points"][0]
                num_dept = pt.get("location")  # Num_Dept
                if num_dept:
                    match = from_all[from_all["Num_Dept"] == num_dept]
                    if not match.empty:
                        clicked_map_dept = match.iloc[0]["Département"]

        except Exception as e:
            st.caption(f"Carte non disponible : {e}")

    # ── Panneau de résultats pour le clic géo ─────────────────
    active_dept   = clicked_map_dept or clicked_dept
    active_region = clicked_region

    if active_dept or active_region:
        if active_dept:
            subset = df_view[df_view["Département"] == active_dept]
            label  = f"département **{active_dept}**"
        else:
            subset = df_view[df_view["Région"] == active_region]
            label  = f"région **{active_region}**"

        st.markdown(f"---\n### 📋 {len(subset)} MOFs — {label}")

        # Métriques rapides
        mc1, mc2, mc3 = st.columns(3)
        mc1.metric("MOFs", len(subset))
        mc2.metric("Métiers distincts", subset[subset["Métier"].ne("")]["Métier"].nunique())
        if subset["Promotion"].notna().any():
            mc3.metric("Dernière promo", int(subset["Promotion"].dropna().max()))

        # Top métiers du département
        if subset["Métier"].ne("").any():
            top_local = (subset[subset["Métier"].ne("")]["Métier"]
                         .value_counts().head(8).reset_index())
            top_local.columns = ["Métier", "Nb"]
            fig_local = px.bar(top_local, x="Nb", y="Métier", orientation="h",
                               color="Nb", color_continuous_scale="Oranges",
                               title=f"Top métiers – {active_dept or active_region}",
                               text="Nb")
            fig_local.update_layout(coloraxis_showscale=False, showlegend=False,
                                    height=280, margin=dict(l=0, r=0, t=40, b=0),
                                    yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig_local, width="stretch")

        # Liste des MOFs
        disp_cols = ["Nom complet", "Métier", "Promotion", "Ville", "Classe"]
        st.dataframe(
            subset[disp_cols].sort_values(["Métier", "Nom complet"]),
            use_container_width=True, hide_index=True, height=400,
        )

# ── TAB 2 : MÉTIERS ──────────────────────
with tab2:
    n_show = st.slider("Nombre de métiers à afficher", 10, 50, 25, key="n_met")
    met_data = (
        df_view[df_view["Métier"].ne("")]
        .groupby("Métier").size().reset_index(name="Nb")
        .sort_values("Nb", ascending=False).head(n_show)
        .sort_values("Nb", ascending=True)
    )
    fig_m = px.bar(met_data, x="Nb", y="Métier", orientation="h",
                   color="Nb", color_continuous_scale="Viridis",
                   title=f"Top {n_show} métiers", text="Nb")
    fig_m.update_traces(textposition="outside")
    fig_m.update_layout(coloraxis_showscale=False, showlegend=False,
                        height=max(400, n_show * 22),
                        margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig_m, width="stretch")

    # Pie top 10
    pie_data = met_data.tail(10)
    fig_pie = px.pie(pie_data, values="Nb", names="Métier",
                     title="Répartition top 10 métiers",
                     color_discrete_sequence=px.colors.qualitative.Set3)
    fig_pie.update_layout(height=400)
    st.plotly_chart(fig_pie, width="stretch")

    # Métier × Région heatmap
    if df_view["Région"].ne("").any() and df_view["Métier"].ne("").any():
        top_met = df_view["Métier"].value_counts().head(15).index.tolist()
        top_reg = df_view["Région"].value_counts().head(10).index.tolist()
        heat = (
            df_view[df_view["Métier"].isin(top_met) & df_view["Région"].isin(top_reg)]
            .groupby(["Métier", "Région"]).size().unstack(fill_value=0)
        )
        fig_h = px.imshow(heat, text_auto=True, aspect="auto",
                          color_continuous_scale="Blues",
                          title="Heatmap : Top 15 métiers × Top 10 régions")
        fig_h.update_layout(height=500, margin=dict(l=0, r=0, t=50, b=0))
        st.plotly_chart(fig_h, width="stretch")

# ── TAB 3 : PROMOTIONS ───────────────────
with tab3:
    promo_data = (
        df_view[df_view["Promotion"].notna()]
        .groupby("Promotion").size().reset_index(name="Nb")
        .sort_values("Promotion")
    )
    promo_data["Promotion"] = promo_data["Promotion"].astype(int)

    fig_p = px.bar(promo_data, x="Promotion", y="Nb",
                   color="Nb", color_continuous_scale="Sunset",
                   title="Nombre de MOFs par promotion (année)")
    fig_p.update_layout(coloraxis_showscale=False, showlegend=False,
                        height=400, margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig_p, width="stretch")

    # Par décennie
    promo_data["Décennie"] = (promo_data["Promotion"] // 10 * 10).astype(str) + "s"
    dec_data = promo_data.groupby("Décennie")["Nb"].sum().reset_index()
    fig_dec = px.pie(dec_data, values="Nb", names="Décennie",
                     title="Répartition par décennie",
                     color_discrete_sequence=px.colors.qualitative.Pastel)
    st.plotly_chart(fig_dec, width="stretch")

    # Évolution cumulative
    promo_data_sorted = promo_data.sort_values("Promotion")
    promo_data_sorted["Cumulé"] = promo_data_sorted["Nb"].cumsum()
    fig_cum = px.area(promo_data_sorted, x="Promotion", y="Cumulé",
                      title="Croissance cumulative de la base de données MOF",
                      color_discrete_sequence=["#1f77b4"])
    fig_cum.update_layout(height=350, margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig_cum, width="stretch")

# ── TAB 4 : DONNÉES BRUTES ───────────────
with tab4:
    st.markdown(f"**{len(df_view):,} lignes** dans la sélection courante")

    search_txt = st.text_input("🔍 Recherche texte libre (nom, métier, ville…)", "")
    df_show = df_view.copy()
    if search_txt:
        mask = (
            df_show["_nom_lower"].str.contains(search_txt.lower(), na=False) |
            df_show["_metier_lower"].str.contains(search_txt.lower(), na=False) |
            df_show["_ville_lower"].str.contains(search_txt.lower(), na=False) |
            df_show["_dept_lower"].str.contains(search_txt.lower(), na=False)
        )
        df_show = df_show[mask]
        st.caption(f"{len(df_show)} résultats pour « {search_txt} »")

    disp = ["Num_Dept", "Département", "Région", "Nom complet", "Métier",
            "Promotion", "Classe", "Ville", "Entreprise", "Source"]
    st.dataframe(
        df_show[disp].sort_values(["Num_Dept", "Nom complet"]),
        width="stretch", hide_index=True,
        height=600,
    )

    # Export CSV
    csv_bytes = df_show[disp].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Télécharger la sélection (CSV)",
        data=csv_bytes,
        file_name="mof_selection.csv",
        mime="text/csv",
    )
