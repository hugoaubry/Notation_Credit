import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF

# CONFIG DE BASE
st.set_page_config(
    page_title="Moteur de notation de crédit",
    page_icon="💹",
    layout="wide"
)
#chemins vers les CSV de sortie de chaque secteur
SECTOR_FILES = {
    "Automobile":        "data/output/autos_output_scorecard.csv",
    "Retail & Apparel":  "data/output/retail_output_scorecard.csv",
    "Mining":            "data/output/mining_output_scorecard.csv",
    "Pharmaceuticals":   "data/output/pharma_output_scorecard.csv",
    "Consumer Goods":    "data/output/cpg_output.csv",
    "Technology":        "data/output/tech_output_scorecard.csv",
    "Industrials":       "data/output/industrial_output_scorecard.csv",
    "Telecommunications": "data/output/telecom_output_scorecard.csv",
    "Oil & Gas":         "data/output/oil_output_scorecard.csv"
}
SECTOR_COLORS = {
    "Automobile": "blue",
    "Retail & Apparel": "red",
    "Mining": "orange",
    "Oil & Gas": "brown",
    "Pharmaceuticals": "green",
    "Technology": "grey",        
    "Consumer Goods": "turquoise",
    "Telecommunications": "pink",
    "Industrials": "gold",
}
SECTOR_LABELS_FR = {
    "Automobile": "Automobile",
    "Retail & Apparel": "Distribution / habillement",
    "Mining": "Mines",
    "Oil & Gas": "Pétrole et gaz intégrés",
    "Pharmaceuticals": "Pharmaceutique",
    "Technology": "Technologie",
    "Consumer Goods": "Biens de consommation",
    "Telecommunications": "Télécommunications",
    "Industrials": "Industriel",
}

def sector_label_fr(sector: str) -> str:
    """Retourne le libellé français d'un secteur."""
    return SECTOR_LABELS_FR.get(str(sector), str(sector))

# Au préalable

@st.cache_data
def load_sector_data():
    """Charge tous les CSV secteurs en un seul DataFrame, avec une colonne 'sector'."""
    frames = []
    for sector, path in SECTOR_FILES.items():
        try:
            df = pd.read_csv(path)
            df["sector"] = sector
            frames.append(df)
        except FileNotFoundError:
            st.warning(f"⚠️ Fichier introuvable pour le secteur {sector} : {path}")
    if not frames:
        return pd.DataFrame()
    all_df = pd.concat(frames, ignore_index=True)
    return all_df

@st.cache_data
def load_metadata():
    """Charge les métadonnées des sociétés (pays, secteur, note agence, coordonnées, etc.)."""
    paths = [
        "data/metadata/universe_metadata.xlsx",
        "data/metadata/universe_metadata.csv",
    ]
    for path in paths:
        try:
            if path.endswith(".csv"):
                # Lecture CSV classique
                meta = pd.read_csv(path)
                # Si tout est dans une seule colonne avec des ';', on retente avec sep=';'
                if len(meta.columns) == 1 and ";" in str(meta.columns[0]):
                    meta = pd.read_csv(path, sep=";")
            else:
                # Lecture Excel
                meta = pd.read_excel(path)
            # Nettoyage des noms de colonnes : suppression d'espaces parasites
            meta.columns = [str(c).strip() for c in meta.columns]
            return meta
        except FileNotFoundError:
            continue
        except Exception as e:
            st.warning(f"Impossible de lire {path} : {e}")
            continue

    st.warning(
        "Aucun fichier universe_metadata (XLSX ou CSV) n'a été trouvé dans data/metadata/. "
    )
    return pd.DataFrame()

def rating_to_notch(rating: str) -> int:
    """
    Convertir une note de type 'A1','Baa2', etc. en un indice numérique.
    Aaa = 1, Aa1 = 2, ..., C = 22 par exemple.
    On s'en sert pour calculer des écarts
    """
    if pd.isna(rating):
        return None
    s = str(rating).strip()

    scale = [
        "Aaa",
        "Aa1","Aa2","Aa3",
        "A1","A2","A3",
        "Baa1","Baa2","Baa3",
        "Ba1","Ba2","Ba3",
        "B1","B2","B3",
        "Caa1","Caa2","Caa3",
        "Ca","C"
    ]
    if s not in scale:
        return None
    return scale.index(s) + 1

def notch_diff(r1: str, r2: str):
    n1 = rating_to_notch(r1)
    n2 = rating_to_notch(r2)
    if n1 is None or n2 is None:
        return None
    return n2 - n1   # si positif = modèle plus généreux que Agences

#afficher l'échelle des notes dans la sidebar
def render_sidebar_rating_scale():
    """Affiche dans la sidebar une table d'équivalence des notes (3 agences + score interne)."""
    st.sidebar.markdown("### Échelle des notes")

    # Définition des lignes (Moody's, Fitch, S&P, catégorie de risque)
    rows = [
        ("Aaa", "AAA", "AAA", "Sécurité maximale"),
        ("Aa1", "AA+", "AA+", "Haute qualité"),
        ("Aa2", "AA", "AA", "Haute qualité"),
        ("Aa3", "AA-", "AA-", "Haute qualité"),
        ("A1", "A+", "A+", "Qualité moyenne"),
        ("A2", "A", "A", "Qualité moyenne"),
        ("A3", "A-", "A-", "Qualité moyenne"),
        ("Baa1", "BBB+", "BBB+", "Qualité moyenne inférieure"),
        ("Baa2", "BBB", "BBB", "Qualité moyenne inférieure"),
        ("Baa3", "BBB-", "BBB-", "Qualité moyenne inférieure"),
        ("Ba1", "BB+", "BB+", "Spéculatif"),
        ("Ba2", "BB", "BB", "Spéculatif"),
        ("Ba3", "BB-", "BB-", "Spéculatif"),
        ("B1", "B+", "B+", "Hautement spéculatif"),
        ("B2", "B", "B", "Hautement spéculatif"),
        ("B3", "B-", "B-", "Hautement spéculatif"),
        ("Caa1", "CCC+", "CCC+", "Mauvaise condition"),
        ("Caa2", "CCC", "CCC", "Mauvaise condition"),
        ("Caa3", "CCC-", "CCC-", "Mauvaise condition"),
        ("Ca", "CC", "CC", "Extrêmement spéculatif"),
        ("C", "C", "C", "En défaut / quasi défaut"),
    ]

    bucket_colors = {
        "Sécurité maximale": "#1f77b4",        # bleu
        "Haute qualité": "#2ca02c",           # vert
        "Qualité moyenne": "#8dd35f",         # vert clair
        "Qualité moyenne inférieure": "#c7e9b4",
        "Spéculatif": "#ffdd57",             # jaune
        "Hautement spéculatif": "#ffb347",    # orange
        "Mauvaise condition": "#ff7f7f",      # rouge clair
        "Extrêmement spéculatif": "#ff4c4c",  # rouge vif
        "En défaut / quasi défaut": "#b10026", # rouge sombre
    }

    # Style minimal pour que le tableau rentre bien dans la sidebar
    html = """
    <style>
    .rating-table {font-size: 11px; border-collapse: collapse; width: 100%;}
    .rating-table th, .rating-table td {
        border: 1px solid #444;
        padding: 2px 4px;
        text-align: center;
    }
    .rating-table th {background-color: #222; color: #f5f5f5;}
    </style>
    <table class="rating-table">
      <tr>
        <th>Moody's</th>
        <th>Fitch</th>
        <th>S&P</th>
        <th>Score interne</th>
      </tr>
    """

    for moody, fitch, sp, bucket in rows:
        notch = rating_to_notch(moody)
        bg = bucket_colors.get(bucket, "#ffffff")
        html += f"<tr style='background-color:{bg};'>"
        html += f"<td>{moody}</td><td>{fitch}</td><td>{sp}</td><td>{notch}</td>"
        html += "</tr>"

    html += "</table>"

    st.sidebar.markdown(html, unsafe_allow_html=True)

    # Légende des couleurs
    st.sidebar.markdown("#### Légende des couleurs")
    legend_items = [
        ("#1f77b4", "Sécurité maximale"),
        ("#2ca02c", "Haute qualité"),
        ("#8dd35f", "Qualité moyenne"),
        ("#c7e9b4", "Qualité moyenne inférieure"),
        ("#ffdd57", "Spéculatif"),
        ("#ffb347", "Hautement spéculatif"),
        ("#ff7f7f", "Mauvaise condition"),
        ("#ff4c4c", "Extrêmement spéculatif"),
        ("#b10026", "En défaut / quasi défaut"),
    ]
    legend_html = "<div style='font-size:11px;'>"
    for color, label in legend_items:
        legend_html += f"<div style='display:flex;align-items:center;margin-bottom:2px;'>"
        legend_html += f"<div style='width:10px;height:10px;background-color:{color};margin-right:6px;border:1px solid #444;'></div>"
        legend_html += f"<span>{label}</span></div>"
    legend_html += "</div>"
    st.sidebar.markdown(legend_html, unsafe_allow_html=True)


#conversion score numérique 1–20 en note qualitative
def numeric_score_to_rating(x: float) -> str:
    """
    Convertit un score numérique (1 = meilleur, 20 = plus faible)
    en note qualitative
    """
    try:
        v = float(x)
    except (TypeError, ValueError):
        return "N/A"

    bins = [
        (1.5, "Aaa"),
        (2.5, "Aa1"), (3.5, "Aa2"), (4.5, "Aa3"),
        (5.5, "A1"),  (6.5, "A2"),  (7.5, "A3"),
        (8.5, "Baa1"),(9.5, "Baa2"),(10.5,"Baa3"),
        (11.5,"Ba1"), (12.5,"Ba2"), (13.5,"Ba3"),
        (14.5,"B1"),  (15.5,"B2"),  (16.5,"B3"),
        (17.5,"Caa1"),(18.5,"Caa2"),(19.5,"Caa3"),
        (20.5,"Ca")
    ]
    for thr, lab in bins:
        if v <= thr:
            return lab
    return "C"

def factor_columns(df: pd.DataFrame):
    """Liste des colonnes factor_..."""
    return [c for c in df.columns if c.startswith("factor_")]

# Aide pour les sous-facteurs spécifiques à chaque méthodologie de secteur

def subfactor_columns(df: pd.DataFrame):
    """Liste des colonnes sf_... (sous-facteurs spécifiques à chaque méthodologie de secteur)."""
    return [c for c in df.columns if c.startswith("sf_")]

# Libellés français pour les facteurs principaux
FACTOR_LABELS_FR = {
    "factor_scale": "Taille / échelle",
    "factor_business_profile": "Profil économique",
    "factor_profitability_efficiency": "Rentabilité et efficacité",
    "factor_leverage_coverage": "Levier et couverture",
    "factor_financial_policy": "Politique financière",
}

def factor_label_fr(col_name: str) -> str:
    if col_name in FACTOR_LABELS_FR:
        return FACTOR_LABELS_FR[col_name]
    base = col_name.replace("factor_", "")
    base = base.replace("_", " ")
    return base[:1].upper() + base[1:]

# Libellés français pour les sous-facteurs fréquents 
SUBFACTOR_LABELS_FR = {
    "sf_trend_global_share": "Tendance de part de marché mondiale",
    "sf_market_position": "Position de marché",
    "sf_ebit_margin": "Marge EBIT",
    "sf_debt_ebitda": "Dette / EBITDA",
    "sf_ebit_interest": "Couverture des intérêts (EBIT / intérêts)",
    "sf_rcf_debt": "RCF / Dette",
    "sf_fcf_debt": "FCF / Dette",
    "sf_financial_policy": "Politique financière",
}


def subfactor_label_fr(col_name: str) -> str:
    """
    Libellé français lisible pour un nom de colonne sf_...
    On utilise un mapping explicite quand il existe, sinon on nettoie le nom.
    """
    if col_name in SUBFACTOR_LABELS_FR:
        return SUBFACTOR_LABELS_FR[col_name]
    base = col_name.replace("sf_", "")
    base = base.replace("_", " ")
    return base[:1].upper() + base[1:]

# Facteurs et sous-facteurs "importants" par secteur pour le comparateur
IMPORTANT_METRICS_BY_SECTOR = {
    # Secteur Automobile :
    "Automobile": {
        "factors": [
            "factor_business_profile",
            "factor_profitability_efficiency",
            "factor_leverage_coverage",
            "factor_financial_policy",
        ],
        "subfactors": [
            "sf_trend_global_share",
            "sf_market_position",
            "sf_ebit_margin",
            "sf_debt_ebitda",
            "sf_ebit_interest",
            "sf_rcf_debt",
            "sf_fcf_debt",
            "sf_financial_policy",
        ],
    },

    # Retail & Apparel : importance du profil business et de la structure financière
    "Retail & Apparel": {
        "factors": [
            "factor_business_profile",
            "factor_profitability_efficiency",
            "factor_leverage_coverage",
            "factor_financial_policy",
        ],
        "subfactors": [
            # Profil économique
            "sf_scale",
            "sf_market_position",
            "sf_brand_strength",
            "sf_format_diversification",
            "sf_geographic_diversification",
            # Rentabilité / cash-flow
            "sf_ebit_margin",
            "sf_ebitda_margin",
            "sf_rcf_debt",
            "sf_fcf_debt",
            # Structure financière / politique
            "sf_debt_ebitda",
            "sf_financial_policy",
        ],
    },

    # Semiconductors : scale, techno, concentration clients, levier et cash-flow
    "Semiconductors": {
        "factors": [
            "factor_business_profile",
            "factor_profitability_efficiency",
            "factor_leverage_coverage",
            "factor_financial_policy",
        ],
        "subfactors": [
            # Business / techno
            "sf_scale",
            "sf_market_position",
            "sf_geographic_diversification",
            "sf_customer_concentration",
            "sf_technological_leadership",
            # Rentabilité / cash-flow
            "sf_ebit_margin",
            "sf_ebitda_margin",
            "sf_rcf_debt",
            "sf_fcf_debt",
            # Levier
            "sf_debt_ebitda",
            "sf_financial_policy",
        ],
    },
     # Mining : échelle, coûts, diversification, levier, couverture
    "Mining": {
        "factors": [
            "factor_business_profile",
            "factor_profitability_efficiency",
            "factor_leverage_coverage",
            "factor_financial_policy",
        ],
        "subfactors": [
            # Profil opérationnel
            "sf_scale",
            "sf_reserve_life",
            "sf_cost_position",
            "sf_diversification",
            # Structure financière
            "sf_leverage",
            "sf_debt_ebitda",
            "sf_ffo_debt",
            "sf_coverage",
        ],
    },

    # Oil & Gas intégrés : scale, intégration, réserves, levier, couverture
    "Oil & Gas": {
        "factors": [
            "factor_business_profile",
            "factor_profitability_efficiency",
            "factor_leverage_coverage",
            "factor_financial_policy",
        ],
        "subfactors": [
            # Profil business
            "sf_scale",
            "sf_integration",
            "sf_reserve_life",
            "sf_upstream_downstream_balance",
            "sf_geographic_diversification",
            # Structure financière
            "sf_leverage",
            "sf_debt_ebitda",
            "sf_ffo_debt",
            "sf_coverage",
        ],
    },

    # Pharmaceutique : taille, diversification produit, pipeline, levier & couverture
    "Pharmaceuticals": {
        "factors": [
            "factor_business_profile",
            "factor_profitability_efficiency",
            "factor_leverage_coverage",
            "factor_financial_policy",
        ],
        "subfactors": [
            # Profil business / R&D
            "sf_scale",
            "sf_product_diversification",
            "sf_geographic_diversification",
            "sf_patent_expiry_risk",
            "sf_pipeline_quality",
            "sf_rnd",
            # Structure financière
            "sf_leverage",
            "sf_debt_ebitda",
            "sf_rcf_debt",
            "sf_coverage",
        ],
    },
}

def go_to_scorecard(company: str, sector: str):
    """
    Bouton pour passer de la carte à la page Scorecard détaillée
    avec la bonne société et le bon secteur.
    """
    st.session_state["jump_to_company"] = company
    st.session_state["jump_to_sector"] = sector
    st.session_state["page"] = "Scorecard détaillée"

def build_pdf_report(row: pd.Series, factors: dict) -> bytes:
    """
    Génère PDF de synthèse pour une entreprise.
    row = ligne du DataFrame principale
    factors = dict {nom_factor: valeur}
    Retourne les bytes du PDF.
    """
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Titre
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Rapport de notation de crédit", ln=True)

    pdf.ln(5)
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 8, f"Société : {row['name']}", ln=True)
    if "sector" in row:
        pdf.cell(0, 8, f"Secteur : {row['sector']}", ln=True)
    if "final_assigned_rating" in row:
        pdf.cell(0, 8, f"Note du modèle : {row['final_assigned_rating']}", ln=True)
    if "scorecard_rating" in row:
        pdf.cell(0, 8, f"Note de la scorecard : {row['scorecard_rating']}", ln=True)
    if "moody_rating" in row:
        pdf.cell(0, 8, f"Note Moody's (en entrée) : {row['moody_rating']}", ln=True)

    pdf.ln(5)
    pdf.set_font("Arial","B",12)
    pdf.cell(0, 8, "Facteurs :", ln=True)

    pdf.set_font("Arial","",11)
    for fname, val in factors.items():
        pdf.cell(0, 7, f"- {fname}: {val}", ln=True)

    # Ajouter un petit commentaire automatique
    pdf.ln(5)
    pdf.set_font("Arial","B",12)
    pdf.cell(0,8, "Commentaire :", ln=True)
    pdf.set_font("Arial","",11)

    agg = row.get("final_adjusted_score", row.get("scorecard_aggregate", None))
    txt = "Cette notation reflète un profil équilibré entre le risque économique, le levier financier et la politique financière."
    if agg is not None:
        try:
            agg = float(agg)
            if agg <= 7:
                txt = "L'entité présente un profil de crédit solide, avec des fondamentaux robustes et un levier conservateur."
            elif agg <= 11:
                txt = "L'entité présente un profil de crédit intermédiaire, avec des forces et faiblesses globalement équilibrées."
            else:
                txt = "L'entité présente un risque de crédit plus élevé, lié à un levier important et/ou à des fondamentaux plus fragiles."
        except:
            pass

    pdf.multi_cell(0, 7, txt)

    # Export : gérer les cas où FPDF retourne une str ou déjà des bytes/bytearray
    out = pdf.output(dest="S")
    if isinstance(out, str):
        pdf_bytes = out.encode("latin1")
    else:
        # fpdf2 peut déjà renvoyer un bytearray / bytes
        pdf_bytes = bytes(out)
    return pdf_bytes

# CHARGEMENT DES DONNÉES

all_data = load_sector_data()
meta = load_metadata()

# Merge metadata  dans all_data si possible
if not all_data.empty and not meta.empty:
    if "name" in meta.columns:
        # On enlève 'sector' de meta pour éviter sector_x / sector_y
        meta_cols = [c for c in meta.columns if c != "sector"]
        all_data = all_data.merge(meta[meta_cols], on="name", how="left")
    else:
        st.warning(
            "Le fichier universe_metadata.csv ne contient pas de colonne 'name'. "
        )

# SIDE BAR

st.sidebar.title("📊 Moteur de notation de crédit")
page = st.sidebar.radio(
    "Navigation",
    [
        "Accueil – Carte monde",
        "Fiche de Notation détaillée",
        "Mode analyste",
        "Comparateur d'entreprises",
        "Ecarts Vs Agences",
    ],
    key="page"
)
# Tableau d'échelle des notes (agences + score interne)
render_sidebar_rating_scale()

# PAGE 1 – CARTE MONDE

if page == "Accueil – Carte monde":
    st.title("🌍 Couverture mondiale des sociétés analysées")

    # Vérifier que le metadata est disponible
    if meta.empty:
        st.info("Ajoute un fichier data/metadata/universe_metadata.csv avec au moins : name, sector, country, lat, lon, moody_rating.")
    else:
        # S'assurer des colonnes minimales
        required_cols = {"name", "sector", "country", "lat", "lon"}
        missing = required_cols.difference(meta.columns)
        if missing:
            st.error(
                "Le fichier universe_metadata.csv doit contenir au minimum les colonnes : "
                "name, sector, country, lat, lon.\n"
                f"Colonnes manquantes : {', '.join(missing)}"
            )
        else:
            st.markdown(
                "Chaque point représente une société analysée. "
                "Les pays où nous avons au moins une société sont ombrés en gris pour visualiser l'exposition mondiale."
            )

            # EXPOSTION PAR PAYS (CHOROPLETH)
            expo = (
                meta.groupby("country")
                    .size()
                    .reset_index(name="count")
            )

            choropleth = go.Choropleth(
                locations=expo["country"],
                locationmode="country names",
                z=[1] * len(expo),  # on code juste présence/absence
                colorscale=[[0, "rgb(255,255,255)"], [1, "rgb(90,140,220)"]],  # bleu plus foncé
                showscale=False,
                marker_line_color="black",
                marker_line_width=0.5,
                hovertemplate="<b>%{location}</b><extra></extra>",
            )
            #  🔹 FILTRES SECTEUR + ENTREPRISE
            sectors = sorted(meta["sector"].dropna().unique())
            sector_options_fr = ["Tous les secteurs"] + [sector_label_fr(s) for s in sectors]
            sector_map_fr_to_en = {sector_label_fr(s): s for s in sectors}
            col1, col2 = st.columns(2)
            with col1:
                sector_choice_label = st.selectbox(
                    "🎯 Filtrer par secteur",
                    sector_options_fr,
                    index=0
                )

            if sector_choice_label == "Tous les secteurs":
                subset = meta.copy()
            else:
                sector_choice_en = sector_map_fr_to_en[sector_choice_label]
                subset = meta[meta["sector"] == sector_choice_en].copy()

            companies = sorted(subset["name"].dropna().unique())
            with col2:
                company_choice = st.selectbox(
                    "🏢 Filtrer par entreprise",
                    ["Aucune"] + companies,
                    index=0
                )
            #  🔹 CONSTRUCTION DES TRACES POUR LES SOCIÉTÉS
            traces = [choropleth]

            # Cas 1 : entreprise spécifique sélectionnée: mettre un seul gros point
            if company_choice != "Aucune":
                row = subset[subset["name"] == company_choice].iloc[0]

                traces.append(
                    go.Scattergeo(
                        lon=[row["lon"]],
                        lat=[row["lat"]],
                        mode="markers",
                        name=sector_label_fr(row["sector"]),
                        marker=dict(
                            size=18,
                            color=SECTOR_COLORS.get(row["sector"], "black"),
                            line=dict(width=2, color="white"),
                        ),
                        text=[f"{row['name']} ({sector_label_fr(row['sector'])})"],
                        hovertemplate="<b>%{text}</b><extra></extra>",
                    )
                )

                info_text = f"📍 **{row['name']} — {sector_label_fr(row['sector'])} — {row['country']}**"

                # Bouton pour accéder directement à la scorecard détaillée
                st.markdown("### 🔎 Analyse détaillée")
                st.button(
                    "📄 Voir la scorecard complète",
                    on_click=go_to_scorecard,
                    args=(row["name"], row["sector"])
                )

            # Cas 2 : pas d'entreprise spécifique: on affiche toutes les sociétés filtrées
            else:
                for secteur in sorted(subset["sector"].dropna().unique()):
                    ens = subset[subset["sector"] == secteur]
                    if ens.empty:
                        continue

                    traces.append(
                        go.Scattergeo(
                            lon=ens["lon"],
                            lat=ens["lat"],
                            mode="markers",
                            name=sector_label_fr(secteur),
                            marker=dict(
                                size=10,
                                color=SECTOR_COLORS.get(secteur, "gray"),
                                line=dict(width=1, color="black"),
                            ),
                            text=[f"{n} ({sector_label_fr(secteur)})" for n in ens["name"]],
                            hovertemplate="<b>%{text}</b><extra></extra>",
                        )
                    )

                info_text = (
                    f"🌐 Exposition actuelle : **{subset['name'].nunique()} sociétés** "
                    f"dans **{subset['country'].nunique()} pays**."
                )

            fig = go.Figure(data=traces)
            fig.update_layout(
                geo=dict(
                    projection_type="natural earth",
                    showland=True,
                    landcolor="rgb(255,255,255)",       # pays sans société : blanc
                    showcountries=True,
                    countrycolor="black",               # frontières noires
                    countrywidth=0.5,
                    showocean=True,
                    oceancolor="rgb(255,255,255)",      # eau comme le fond
                ),
                legend=dict(title="Secteurs"),
                margin=dict(l=0, r=0, t=0, b=0),
            )

            st.plotly_chart(fig, use_container_width=True)
            st.info(info_text)
            
# PAGE 2 – SCORECARD DÉTAILLÉE

elif page == "Fiche de Notation détaillée":
    st.title("📋 Fiche détaillée par société")

    if all_data.empty:
        st.error("Aucune donnée secteur n'a été chargée. Vérifie les CSV dans data/outputs/.")
    else:
        # Pré-sélection éventuelle si redirection depuis la carte
        pre_selected_company = st.session_state.pop("jump_to_company", None) if "jump_to_company" in st.session_state else None
        pre_selected_sector = st.session_state.pop("jump_to_sector", None) if "jump_to_sector" in st.session_state else None

        sectors = sorted(all_data["sector"].dropna().unique())

        sector_labels_fr = [sector_label_fr(s) for s in sectors]
        sector_map_fr_to_en = {sector_label_fr(s): s for s in sectors}

        # Secteur pré-sélectionné si fourni, sinon premier secteur
        default_sector_index = sectors.index(pre_selected_sector) if pre_selected_sector in sectors else 0

        col1, col2 = st.columns(2)
        with col1:
            sector_choice_label = st.selectbox("Choisissez un secteur", sector_labels_fr, index=default_sector_index)
        sector_choice = sector_map_fr_to_en[sector_choice_label]

        subset = all_data[all_data["sector"] == sector_choice]
        companies = sorted(subset["name"].dropna().unique())

        # Société pré-sélectionnée si fournie et présente dans la liste du secteur
        default_company_index = companies.index(pre_selected_company) if pre_selected_company in companies else 0

        with col2:
            company_choice = st.selectbox("Choisissez une société", companies, index=default_company_index)

        row = subset[subset["name"] == company_choice].iloc[0]

        st.subheader(f"{company_choice} – {sector_label_fr(sector_choice)}")
        colA, colB, colC = st.columns(3)
        with colA:
            st.metric("Score agrégé de la scorecard", f"{row.get('scorecard_aggregate', 'N/A')}")
            st.metric("Note de la scorecard", row.get("scorecard_rating", "N/A"))
        with colB:
            st.metric("Score final ajusté", f"{row.get('final_adjusted_score', 'N/A')}")
            st.metric("Note finale", row.get("final_assigned_rating", "N/A"))
        with colC:
            st.metric("Note Agences", row.get("Note_Agences", "N/A"))
            diff = notch_diff(row.get("Note_Agences", None), row.get("final_assigned_rating", None))
            if diff is not None:
                st.metric("Écart vs Agences (crans)", diff)

        st.markdown("### ⚙️ Détail par facteur")

        fcols = factor_columns(subset)
        if fcols:
            # Récupérer les scores numériques 1–20 pour chaque factor_...
            factor_data = {}
            for c in fcols:
                val = row.get(c, None)
                if pd.notna(val):
                    try:
                        v = float(val)
                    except (TypeError, ValueError):
                        continue
                    # borne sécurité 1..20
                    v = max(1.0, min(20.0, v))
                    factor_data[c] = v

            if factor_data:
                # Construire un tableau lisible : nom de facteur, score numérique, note qualitative
                display_rows = []
                for col_name, score in factor_data.items():
                    label = factor_label_fr(col_name)
                    display_rows.append({
                        "Facteur": label,
                        "Score numérique": round(score, 2),
                        "Note qualitative": numeric_score_to_rating(score),
                    })
                df_factors = pd.DataFrame(display_rows).set_index("Facteur")
                st.dataframe(df_factors)

                #Synthèse qualitative automatique
                sorted_factors = sorted(factor_data.items(), key=lambda kv: kv[1])  # 1 = meilleur
                best = sorted_factors[:2]
                worst = sorted_factors[-2:] if len(sorted_factors) >= 2 else []

                synth_lines = []
                if best:
                    synth_lines.append("**Forces principales :**")
                    for col_name, score in best:
                        label = factor_label_fr(col_name)
                        synth_lines.append(f"- {label} ({numeric_score_to_rating(score)})")
                if worst:
                    synth_lines.append("")
                    synth_lines.append("**Points de vigilance :**")
                    for col_name, score in worst:
                        label = factor_label_fr(col_name)
                        synth_lines.append(f"- {label} ({numeric_score_to_rating(score)})")

                if synth_lines:
                    st.markdown("### 🧾 Synthèse qualitative")
                    st.info("\n".join(synth_lines))

                # 🔷 Radar : société vs médiane de secteur
                if len(factor_data) >= 3:
                    labels = [factor_label_fr(col) for col in factor_data.keys()]
                    numeric_vals = list(factor_data.values())
                    # On inverse l'échelle pour que 1 (meilleur) donne une valeur élevée sur le radar
                    core_vals = [21.0 - v for v in numeric_vals]

                    # Médiane du secteur sur les mêmes colonnes factor_...
                    sector_medians = []
                    for col in factor_data.keys():
                        try:
                            med = float(subset[col].median())
                        except Exception:
                            med = None
                        if pd.isna(med):
                            med = None
                        sector_medians.append(med)
                    median_vals = [
                        (21.0 - v) if v is not None else None
                        for v in sector_medians
                    ]

                    # Fermeture des polygones
                    labels_closed = labels + [labels[0]]
                    core_closed = core_vals + [core_vals[0]]

                    # On remplace les None éventuels par 0 pour le tracé (peu fréquent)
                    median_clean = [
                        (v if v is not None else 0.0) for v in median_vals
                    ]
                    median_closed = median_clean + [median_clean[0]]

                    fig_radar = go.Figure()
                    # Société : polygone bleu foncé rempli
                    fig_radar.add_trace(
                        go.Scatterpolar(
                            r=core_closed,
                            theta=labels_closed,
                            fill="toself",
                            name=company_choice,
                            line=dict(color="rgb(0,76,153)")
                        )
                    )
                    # Médiane secteur : trait orange sans remplissage
                    fig_radar.add_trace(
                        go.Scatterpolar(
                            r=median_closed,
                            theta=labels_closed,
                            fill=None,
                            name="Médiane secteur",
                            line=dict(color="orange", width=3),
                            mode="lines"
                        )
                    )
                    fig_radar.update_layout(
                        height=500,
                        polar=dict(radialaxis=dict(showticklabels=False)),
                        legend=dict(title="Profil")
                    )
                    st.plotly_chart(fig_radar, use_container_width=True)


            else:
                st.info("Aucune valeur de facteur disponible pour cette société.")
        else:
            st.info("Aucune colonne factor_... trouvée dans ce fichier pour construire un radar.")

        st.markdown("### 📝 Export PDF de cette société")

        if st.button("Générer un PDF de synthèse"):
            factor_data = {}
            for c in factor_columns(subset):
                val = row.get(c, None)
                if pd.notna(val):
                    fname = factor_label_fr(c)
                    factor_data[fname] = round(float(val), 2)
            pdf_bytes = build_pdf_report(row, factor_data)
            st.download_button(
                label="📄 Télécharger le rapport PDF",
                data=pdf_bytes,
                file_name=f"{company_choice}_rating_report.pdf",
                mime="application/pdf"
            )

# PAGE 3 – MODE ANALYSTE

elif page == "Mode analyste":
    st.title("🧠 Mode analyste – jouer avec les facteurs")

    if all_data.empty:
        st.error("Aucune donnée secteur n'a été chargée.")
    else:
        sectors = sorted(all_data["sector"].dropna().unique())
        sector_labels_fr = [sector_label_fr(s) for s in sectors]
        sector_map_fr_to_en = {sector_label_fr(s): s for s in sectors}
        col1, col2 = st.columns(2)
        with col1:
            sector_choice_label = st.selectbox("Secteur", sector_labels_fr)
        sector_choice = sector_map_fr_to_en[sector_choice_label]
        subset = all_data[all_data["sector"] == sector_choice]
        companies = sorted(subset["name"].dropna().unique())
        with col2:
            company_choice = st.selectbox("Société", companies)

        row = subset[subset["name"] == company_choice].iloc[0]

        st.write(f"**Base :** note de scorecard = {row.get('scorecard_rating','N/A')}, "
                 f"note finale = {row.get('final_assigned_rating','N/A')}")

        st.markdown("#### Ajustez les sous-facteurs clés du secteur")

        # Sous-facteurs spécifiques à la méthodologie de chaque secteur
        sfcols = subfactor_columns(subset)
        # Si aucun sf_... n'est disponible (secteur pas encore détaillé), on retombe sur les factor_...
        if sfcols:
            editable_cols = sfcols
            slider_label_prefix = "sf_"
        else:
            st.info("Aucun sous-facteur sf_... détecté pour ce secteur, utilisation des facteurs agrégés.")
            editable_cols = factor_columns(subset)
            slider_label_prefix = "factor_"

        # On ne garde que les principaux : on enlève tout ce qui est exactement neutre (score = 10)
        filtered_cols = []
        for c in editable_cols:
            val = row.get(c, None)
            if pd.notna(val):
                try:
                    v = float(val)
                except (TypeError, ValueError):
                    continue
                # 10 = neutre dans notre échelle, on ne le propose pas dans les sliders
                if abs(v - 10.0) > 1e-6:
                    filtered_cols.append(c)
        editable_cols = filtered_cols

        if not editable_cols:
            st.info("Aucun sous-facteur/facteur significatif à ajuster (tous étaient neutres à 10).")
        else:
            sliders = {}
            cols = st.columns(2)
            for i, c in enumerate(editable_cols):
                base_val = float(row[c]) if pd.notna(row[c]) else 10.0
                # Utilisation de libellés français pour les facteurs / sous-facteurs
                if c.startswith("factor_"):
                    label = factor_label_fr(c)
                elif c.startswith("sf_"):
                    label = subfactor_label_fr(c)
                else:
                    base_name = c.replace(slider_label_prefix, "").replace("_", " ")
                    label = base_name[:1].upper() + base_name[1:]
                with cols[i % 2]:
                    # moyenne du secteur pour ce facteur / sous-facteur
                    try:
                        sector_mean = float(subset[c].mean())
                    except Exception:
                        sector_mean = None

                    # Slider principal pour la société
                    sliders[c] = st.slider(
                        label,
                        min_value=1.0,
                        max_value=20.0,
                        value=float(round(base_val, 1)),
                        step=0.1,
                        key=f"slider_{c}"
                    )

                    # Ajout d'un point orange sur la "ligne" du slider pour représenter la moyenne du secteur
                    if sector_mean is not None and not pd.isna(sector_mean):
                        # conversion de la moyenne en position (%) entre 1 et 20
                        try:
                            pos = (sector_mean - 1.0) / (20.0 - 1.0) * 100.0
                        except Exception:
                            pos = None
                        if pos is not None:
                            # on borne la position entre 0% et 100% par sécurité
                            pos = max(0.0, min(100.0, pos))
                            st.markdown(
                                f"""
                                <div style="
                                    position: relative;
                                    height: 0px;
                                    margin-top: -18px;
                                    margin-bottom: 10px;">
                                    <div style="
                                        position: absolute;
                                        left: {pos}%;
                                        transform: translateX(-50%);
                                        width: 12px;
                                        height: 12px;
                                        border-radius: 50%;
                                        background-color: orange;
                                        border: 2px solid #000;">
                                    </div>
                                </div>
                                """,
                                unsafe_allow_html=True
                            )

            # Slider spécifique pour les 'other considerations soft' si la colonne existe
            soft_slider = None
            if "delta_other_considerations_soft" in subset.columns:
                base_soft = row.get("delta_other_considerations_soft", 0.0)
                try:
                    base_soft = float(base_soft)
                except (TypeError, ValueError):
                    base_soft = 0.0
                st.markdown("#### Ajustez les autres considérations (delta, peut améliorer ou détériorer la note)")
                soft_slider = st.slider(
                    "Autres considérations (delta, de -1,0 à +1,0, par pas de 0,25)",
                    min_value=-1.0,
                    max_value=1.0,
                    value=base_soft,
                    step=0.25
                )

            if st.button("Recalculer la note simulée"):
                # --- 1. baseline à partir du scorecard_aggregate ---
                try:
                    core_baseline = float(row.get("scorecard_aggregate", new_score if 'new_score' in locals() else 10))
                except:
                    core_baseline = 10.0

                # --- 2. moyenne des valeurs originales ---
                orig_vals = []
                for c in sliders.keys():
                    try:
                        ov = float(row.get(c, 10))
                        orig_vals.append(ov)
                    except:
                        pass
                mean_orig = sum(orig_vals) / len(orig_vals) if orig_vals else 10.0

                # --- 3. moyenne des nouveaux sliders ---
                new_vals = list(sliders.values())
                mean_new = sum(new_vals) / len(new_vals)

                # --- 4. delta entre nouvelles valeurs et originales ---
                delta_mean = mean_new - mean_orig

                # --- 5. nouveau score "core" cohérent avec baseline réelle ---
                new_core = core_baseline + delta_mean

                # --- 6. gestion du delta other considerations ---
                if soft_slider is not None:
                    try:
                        new_soft = float(soft_slider)
                    except:
                        new_soft = 0.0
                else:
                    new_soft = 0.0

                # --- 7. score final ---
                new_final = new_core - new_soft

                st.write(f"**Score core de départ (réel)** : {core_baseline:.2f}")
                st.write(f"**Nouveau score core** : {new_core:.2f}")
                st.write(f"**Score final après autres considérations** : {new_final:.2f}")

                # --- 8. conversion rating ---
                def score_to_rating(x):
                    bins = [
                        (1.5, "Aaa"),
                        (2.5, "Aa1"), (3.5, "Aa2"), (4.5, "Aa3"),
                        (5.5, "A1"),  (6.5, "A2"),  (7.5, "A3"),
                        (8.5, "Baa1"),(9.5, "Baa2"),(10.5,"Baa3"),
                        (11.5,"Ba1"), (12.5,"Ba2"), (13.5,"Ba3"),
                        (14.5,"B1"),  (15.5,"B2"),  (16.5,"B3"),
                        (17.5,"Caa1"),(18.5,"Caa2"),(19.5,"Caa3"),
                        (20.5,"Ca")
                    ]
                    for thr, lab in bins:
                        if x <= thr:
                            return lab
                    return "C"

                new_rating = score_to_rating(new_final)
                st.success(f"Nouvelle note indicative : **{new_rating}**")

                old_rating = row.get("final_assigned_rating", "N/A")
                st.write(f"Comparaison : ancienne note = {old_rating}, nouvelle note = {new_rating}")
                diff = notch_diff(old_rating, new_rating)
                if diff is not None:
                    st.write(f"Écart (crans) : {diff}")

# PAGE 4 – COMPARATEUR D'ENTREPRISES

elif page == "Comparateur d'entreprises":
    st.title("🆚 Comparateur de sociétés")

    if all_data.empty:
        st.error("Aucune donnée secteur n'a été chargée.")
    else:
        # --- Sélection du secteur au centre sous le titre ---
        sectors = sorted(all_data["sector"].dropna().unique())
        sector_labels_fr = [sector_label_fr(s) for s in sectors]
        sector_map_fr_to_en = {sector_label_fr(s): s for s in sectors}

        st.markdown("### Choisissez un secteur à comparer")
        sector_choice_label = st.selectbox(
            "Secteur",
            sector_labels_fr,
            index=0,
            key="comp_sector_select",
        )
        sector_choice = sector_map_fr_to_en[sector_choice_label]

        subset_sector = all_data[all_data["sector"] == sector_choice]
        companies_sector = sorted(subset_sector["name"].dropna().unique())

        if len(companies_sector) < 2:
            st.info("Il faut au moins deux sociétés dans ce secteur pour effectuer une comparaison.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                company1 = st.selectbox("Société 1", companies_sector, index=0, key="comp_company1")
            with col2:
                company2 = st.selectbox("Société 2", companies_sector, index=min(1, len(companies_sector)-1), key="comp_company2")

            df1 = subset_sector[subset_sector["name"] == company1].iloc[0]
            df2 = subset_sector[subset_sector["name"] == company2].iloc[0]

            st.subheader("Notes globales")
            colA, colB = st.columns(2)
            with colA:
                st.markdown(f"**{company1}** ({sector_label_fr(df1.get('sector',''))})")
                st.write(f"Note finale : {df1.get('final_assigned_rating','N/A')}")
                st.write(f"Note de la scorecard : {df1.get('scorecard_rating','N/A')}")
                st.write(f"Note Agences : {df1.get('Note_Agences','N/A')}")
            with colB:
                st.markdown(f"**{company2}** ({sector_label_fr(df2.get('sector',''))})")
                st.write(f"Note finale : {df2.get('final_assigned_rating','N/A')}")
                st.write(f"Note de la scorecard : {df2.get('scorecard_rating','N/A')}")
                st.write(f"Note Agences : {df2.get('Note_Agences','N/A')}")

            st.markdown("### Comparaison des facteurs et sous-facteurs importants")

            # --- Récupération des facteurs et sous-facteurs pertinents pour ce secteur ---
            sector_importance = IMPORTANT_METRICS_BY_SECTOR.get(sector_choice, None)

            all_factor_cols = factor_columns(subset_sector)
            all_sf_cols = subfactor_columns(subset_sector)

            if sector_importance is not None:
                chosen_factors = [c for c in sector_importance.get("factors", []) if c in all_factor_cols]
                chosen_subfactors = [c for c in sector_importance.get("subfactors", []) if c in all_sf_cols]
            else:
                # Si pas de configuration spécifique : on prend tous les factors et sf_ disponibles
                chosen_factors = all_factor_cols
                chosen_subfactors = all_sf_cols

            metrics_cols = chosen_factors + chosen_subfactors

            if not metrics_cols:
                st.info("Aucun facteur ou sous-facteur pertinent trouvé pour ce secteur.")
            else:
                data_comp = []
                for col in metrics_cols:
                    v1 = df1.get(col, None)
                    v2 = df2.get(col, None)
                    try:
                        v1 = float(v1) if pd.notna(v1) else None
                    except (TypeError, ValueError):
                        v1 = None
                    try:
                        v2 = float(v2) if pd.notna(v2) else None
                    except (TypeError, ValueError):
                        v2 = None

                    # Choix du label en français
                    if col.startswith("factor_"):
                        label = factor_label_fr(col)
                    elif col.startswith("sf_"):
                        label = subfactor_label_fr(col)
                    else:
                        base = col.replace("_", " ")
                        label = base[:1].upper() + base[1:]

                    data_comp.append({
                        "Indicateur": label,
                        company1: v1,
                        company2: v2,
                    })

                df_comp = pd.DataFrame(data_comp)
                st.dataframe(df_comp.set_index("Indicateur"))

                # Graphique comparatif
                df_melt = df_comp.melt(id_vars="Indicateur", var_name="Société", value_name="Score")
                # On garde seulement les lignes avec des scores valides
                df_melt = df_melt.dropna(subset=["Score"])

                if df_melt.empty:
                    st.info("Pas de scores numériques comparables à tracer pour ces deux sociétés.")
                else:
                    fig_bar = px.bar(
                        df_melt,
                        x="Indicateur",
                        y="Score",
                        color="Société",
                        barmode="group",
                    )
                    fig_bar.update_layout(
                        xaxis_title="Facteurs / Sous-facteurs",
                        yaxis_title="Score (1 = meilleur, 20 = plus faible)",
                        height=500
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)

# PAGE 5 – HEATMAP ÉCART MOODY'S
elif page == "Ecarts Vs Agences":
    st.title("🧯 Heatmap des écarts vs Agences")

    if all_data.empty:
        st.error("Aucune donnée secteur n'a été chargée.")
    else:
        df = all_data.copy()
        if "Note_Agences" not in df.columns:
            st.info("Ajoute une colonne Note_Agences dans data/metadata/universe_metadata.csv pour pouvoir comparer.")
        else:
            df["notch_model"] = df["final_assigned_rating"].apply(rating_to_notch)
            df["notch_moody"] = df["Note_Agences"].apply(rating_to_notch)
            df["notch_diff"] = df.apply(
                lambda r: r["notch_model"] - r["notch_moody"]
                if (pd.notna(r["notch_model"]) and pd.notna(r["notch_moody"]))
                else None,
                axis=1
            )

            st.markdown("### Tableau des écarts")
            df_display = df[["name","sector","Note_Agences","final_assigned_rating","notch_diff"]].copy()
            df_display["sector"] = df_display["sector"].apply(sector_label_fr)
            df_display = df_display.rename(columns={
                "name": "Nom",
                "sector": "Secteur",
                "Note_Agences": "Note agences",
                "final_assigned_rating": "Note finale",
                "notch_diff": "Écart (crans)"
            })
            st.dataframe(
                df_display.sort_values("Écart (crans)", ascending=True)
            )

            st.markdown("### Distribution des écarts")

            df_valid = df.dropna(subset=["notch_diff"])
            if not df_valid.empty:
                # Histogramme avec classes entières (1 cran par 1 cran)
                fig_hist = px.histogram(df_valid, x="notch_diff")

                # Bins centrés sur les valeurs entières de notch_diff
                min_diff = df_valid["notch_diff"].min()
                max_diff = df_valid["notch_diff"].max()
                fig_hist.update_traces(
                    xbins=dict(
                        start=min_diff - 0.5,
                        end=max_diff + 0.5,
                        size=1
                    )
                )

                fig_hist.update_layout(
                    xaxis_title="Écart (modèle - Agences) en crans",
                    yaxis_title="Nombre de sociétés",
                    height=500,
                )
                fig_hist.update_layout(bargap=0.25)
                # Graduation de l'axe X uniquement sur les entiers
                fig_hist.update_xaxes(dtick=1)

                st.plotly_chart(fig_hist, use_container_width=True)

                # Écart moyen absolu global
                st.markdown("#### Écart moyen absolu (global)")
                mae = df_valid["notch_diff"].abs().mean()
                st.write(f"Écart moyen absolu : **{mae:.2f} crans**")

                # Décomposition par secteur
                st.markdown("#### Écart moyen par secteur")
                sector_stats = (
                    df_valid
                    .groupby("sector")["notch_diff"]
                    .agg(["mean", "median", "count"])
                    .reset_index()
                )

                # Ajout de l'écart absolu moyen par secteur
                mean_abs_by_sector = (
                    df_valid
                    .groupby("sector")["notch_diff"]
                    .apply(lambda s: s.abs().mean())
                    .reset_index(name="mean_abs")
                )
                sector_stats = sector_stats.merge(mean_abs_by_sector, on="sector", how="left")

                # Libellés français pour les secteurs
                sector_stats["sector"] = sector_stats["sector"].apply(sector_label_fr)

                # Renommage des colonnes pour affichage
                sector_stats = sector_stats.rename(columns={
                    "sector": "Secteur",
                    "mean": "Écart moyen (crans)",
                    "median": "Médiane (crans)",
                    "count": "Nombre de sociétés",
                    "mean_abs": "Écart absolu moyen (crans)",
                })

                st.dataframe(sector_stats)
            else:
                st.info("Aucun notch_diff valide calculable. Vérifie que moody_rating et final_assigned_rating sont bien renseignées.")