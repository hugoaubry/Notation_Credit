### Pharmaceuticals – Scorecard ###
from typing import Optional, Dict, Tuple
import pandas as pd

# =========================
#   CONSTANTES I/O
# =========================

INPUT_XLSX  = "data/intput/pharma_input_template.xlsx"
INPUT_SHEET = "pharma_input_template"
OUTPUT_CSV  = "data/output/pharma_output_scorecard.csv"

# =========================
#   COLONNES ATTENDUES
# =========================
# y1 = plus ancien ; y3 = plus récent

COL_NAME    = "name"
COL_COUNTRY = "country"   # nouvelle colonne pays

# --------- FACTEUR SCALE (25%) ----------
# Revenue (valeur en devise locale dans l’xlsx -> convertie en USD dans le code)
COL_REV_Y1 = "revenue_y1"
COL_REV_Y2 = "revenue_y2"
COL_REV_Y3 = "revenue_y3"

# --------- BUSINESS PROFILE (40%) ----------
# 4 sous-facteurs qualitatifs (10% chacun)
COL_PROD_DIVERSITY   = "product_therapeutic_diversity"   # Aaa..Ca
COL_GEO_DIVERSITY    = "geographic_diversity"            # Aaa..Ca
COL_PATENT_EXPOSURE  = "patent_exposures"                # Aaa..Ca
COL_PIPELINE_QUALITY = "pipeline_quality"                # Aaa..Ca

# --------- LEVERAGE & COVERAGE (20%) ----------
# Debt / EBITDA (x)
COL_DEBT_EBITDA_Y1 = "debt_ebitda_y1"
COL_DEBT_EBITDA_Y2 = "debt_ebitda_y2"
COL_DEBT_EBITDA_Y3 = "debt_ebitda_y3"

# RCF / Net Debt (%)
COL_RCF_NETDEBT_Y1 = "rcf_net_debt_y1"
COL_RCF_NETDEBT_Y2 = "rcf_net_debt_y2"
COL_RCF_NETDEBT_Y3 = "rcf_net_debt_y3"

# EBIT / Interest Expense (x)
COL_EBIT_INT_Y1 = "ebit_interest_y1"
COL_EBIT_INT_Y2 = "ebit_interest_y2"
COL_EBIT_INT_Y3 = "ebit_interest_y3"

# --------- FINANCIAL POLICY (15%) ----------
COL_FINANCIAL_POLICY = "financial_policy"  # Aaa..Ca

# --------- LIQUIDITÉ (pour other considerations) ----------
COL_LIQ_Y1 = "liq_y1"
COL_LIQ_Y2 = "liq_y2"
COL_LIQ_Y3 = "liq_y3"

# --------- OTHER CONSIDERATIONS (gabarit générique) ----------
COL_ESG_SCORE        = "esg_score"          # 1..5
COL_CAPTIVE_RATIO    = "captive_ratio"      # 0..1
COL_REGULATION_SCORE = "regulation_score"   # 1..5
COL_MANAGEMENT_SCORE = "management_score"   # 1..5
COL_NONWHOLLY_SALES  = "nonwholly_sales"    # 0..1
COL_EVENT_RISK_SCORE = "event_risk_score"   # 0/1/2
COL_PARENTAL_SUPPORT = "parental_support"   # -3..+3 (on cappe à ±1 cran)

# (facultatif) cash loggé
COL_CASH_Y1 = "cash_y1"
COL_CASH_Y2 = "cash_y2"
COL_CASH_Y3 = "cash_y3"


# =========================
#   FX PAR PAYS (UNIQUEMENT PAYS PRÉSENTS DANS LE XLSX)
# =========================

# Pays présents dans pharma_input_template.xlsx :
#   - USA
#   - Switzerland
#   - Japan
#
# On associe à chacun une devise "native" et un facteur de conversion
# vers USD (valeur approx, mais fixe et cohérente pour tout le fichier).

FX_BY_COUNTRY: Dict[str, Dict[str, float]] = {
    "USA": {
        "currency": "USD",
        "fx_to_usd": 1.0,      # 1 USD = 1 USD
    },
    "Switzerland": {
        "currency": "CHF",
        "fx_to_usd": 1.10,     # 1 CHF ≈ 1.10 USD (approx, fixe)
    },
    "Japan": {
        "currency": "JPY",
        "fx_to_usd": 0.0065,   # 1 JPY ≈ 0.0065 USD (approx, fixe)
    },
}


def get_fx_for_country(country_val) -> Tuple[str, float]:
    """
    Retourne (currency, fx_to_usd) à partir du pays.
    Si pays inconnu ou vide -> on suppose USD par défaut.
    """
    if country_val is None:
        return "USD", 1.0
    s = str(country_val).strip()
    info = FX_BY_COUNTRY.get(s)
    if info is None:
        # Fallback défensif : USD si jamais un pays non mappé apparaît.
        return "USD", 1.0
    return info["currency"], info["fx_to_usd"]


# Colonnes monétaires à convertir en USD (valeurs absolues)
MONETARY_COLS_REVENUE = [COL_REV_Y1, COL_REV_Y2, COL_REV_Y3]
MONETARY_COLS_CASH    = [COL_CASH_Y1, COL_CASH_Y2, COL_CASH_Y3]


def convert_row_monetary_to_usd(row: pd.Series) -> pd.Series:
    """
    Convertit les montants monétaires (revenues, cash) de la devise locale
    vers USD en fonction du pays de l'entreprise.
    Les ratios (Debt/EBITDA, RCF/Net Debt, EBIT/Interest, etc.)
    NE SONT PAS modifiés.
    """
    country = row.get(COL_COUNTRY)
    cur, fx = get_fx_for_country(country)

    # On stocke éventuellement l'info pour debug (non utilisé dans les scores)
    row["_currency"] = cur
    row["_fx_to_usd"] = fx

    for col in MONETARY_COLS_REVENUE + MONETARY_COLS_CASH:
        if col in row:
            from_val = row[col]
            if from_val is not None and not pd.isna(from_val):
                v = None
                if isinstance(from_val, (int, float)):
                    v = float(from_val)
                else:
                    v = to_float(from_val)
                if v is not None:
                    row[col] = v * fx

    return row


# =========================
#   QUALI -> NUM
# =========================

QUALI_NUM = {
    "Aaa": 1.0, "Aa": 3.0, "A": 6.0, "Baa": 9.0,
    "Ba": 12.0, "B": 15.0, "Caa": 18.0, "Ca": 20.0
}
NOTCH_TO_ALPHA = {
    "Aaa": "Aaa",
    "Aa1": "Aa", "Aa2": "Aa", "Aa3": "Aa",
    "A1": "A", "A2": "A", "A3": "A",
    "Baa1": "Baa", "Baa2": "Baa", "Baa3": "Baa",
    "Ba1": "Ba", "Ba2": "Ba", "Ba3": "Ba",
    "B1": "B", "B2": "B", "B3": "B",
    "Caa1": "Caa", "Caa2": "Caa", "Caa3": "Caa",
    "Ca": "Ca", "C": "Ca", "D": "Ca"
}

def _alpha_from_label(lbl: Optional[str]) -> Optional[str]:
    if lbl is None or (isinstance(lbl, float) and pd.isna(lbl)):
        return None
    s = str(lbl).strip()
    if s in QUALI_NUM:
        return s
    if s in NOTCH_TO_ALPHA:
        return NOTCH_TO_ALPHA[s]
    sU = s.upper()
    if sU in NOTCH_TO_ALPHA:
        return NOTCH_TO_ALPHA[sU]
    return None

def score_quali(label: Optional[str]) -> float:
    a = _alpha_from_label(label)
    return QUALI_NUM[a] if a else 12.0   # défaut Ba

# =========================
#   PARSING NUM
# =========================

def to_float(x) -> Optional[float]:
    if x is None:
        return None
    if isinstance(x, (int, float)):
        if pd.isna(x):
            return None
        return float(x)
    s = str(x).strip()
    if s == "":
        return None
    s = s.replace(",", "")
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    if s.endswith("%"):
        s = s[:-1].strip()
    try:
        val = float(s)
    except:
        return None
    if neg:
        val = -val
    return val

# =========================
#   NUMERIC RANGES (Exhibit Pharma)
# =========================

NUM_RANGES = {
    "Aaa": (0.5, 1.5),
    "Aa":  (1.5, 4.5),
    "A":   (4.5, 7.5),
    "Baa": (7.5, 10.5),
    "Ba":  (10.5, 13.5),
    "B":   (13.5, 16.5),
    "Caa": (16.5, 19.5),
    "Ca":  (19.5, 20.5)
}

def _interp_linear(x: float, lo: float, hi: float, ylo: float, yhi: float) -> float:
    if hi == lo:
        return (ylo + yhi) / 2.0
    t = (x - lo) / (hi - lo)
    if t < 0:
        t = 0.0
    if t > 1:
        t = 1.0
    return ylo + t * (yhi - ylo)

def _score_quant_from_bounds(x: float, bounds: list, higher_is_better: bool) -> float:
    for alpha, lo, hi in bounds:
        in_band = (lo < hi and (x >= lo and x < hi)) or (lo > hi and (x <= lo and x > hi))
        if in_band:
            num_lo, num_hi = NUM_RANGES[alpha]
            if higher_is_better:
                return _interp_linear(x, lo, hi, num_hi, num_lo)
            else:
                return _interp_linear(x, lo, hi, num_lo, num_hi)
    # Hors bornes
    if higher_is_better:
        return 0.5 if x >= bounds[0][2] else 20.5
    else:
        return 0.5 if x <= bounds[0][1] else 20.5

# ====== Bornes spécifiques PHARMA (Exhibit 2) ======

# 1) Revenue (USD) – Scale (25 %)
# On applique directement les bornes en USD absolus (x ≈ chiffre d'affaires en USD)
REVENUE_BOUNDS = [
    ("Aaa", 60e9,   1e18),   # ≥ $60bn
    ("Aa",  30e9,   60e9),
    ("A",   15e9,   30e9),
    ("Baa",  8e9,   15e9),
    ("Ba",   3e9,    8e9),
    ("B",    1e9,    3e9),
    ("Caa", 0.25e9,  1e9),
    ("Ca",  -5e9,  0.25e9),
]

# 2) Debt / EBITDA (x)
DEBT_EBITDA_BOUNDS = [
    ("Aaa", 0.0,  0.5),
    ("Aa",  0.5,  1.5),
    ("A",   1.5,  2.5),
    ("Baa", 2.5,  3.5),
    ("Ba",  3.5,  4.5),
    ("B",   4.5,  6.0),
    ("Caa", 6.0,  9.0),
    ("Ca",  -20.0,  20.0),
]

# 3) RCF / Net Debt (%)
RCF_NETDEBT_BOUNDS = [
    ("Aaa", 70.0, 120.0),
    ("Aa",  50.0, 70.0),
    ("A",   35.0, 50.0),
    ("Baa", 20.0, 35.0),
    ("Ba",  12.5, 20.0),
    ("B",    5.0, 12.5),
    ("Caa",  0.0, 5.0),
    ("Ca",  -20.0, 0.0),
]

# 4) EBIT / Interest Expense (x)
EBIT_INT_BOUNDS = [
    ("Aaa", 18.0, 100.0),
    ("Aa",  12.0, 18.0),
    ("A",    7.0, 12.0),
    ("Baa",  4.0, 7.0),
    ("Ba",   2.25, 4.0),
    ("B",    1.0, 2.25),
    ("Caa",  0.5, 1.0),
    ("Ca",   -20.0, 0.5),
]

# =========================
#   SCORE QUANTI
# =========================

def score_revenue(usd_amount: Optional[float]) -> float:
    # usd_amount = chiffre d'affaires en USD (après conversion FX)
    v = to_float(usd_amount)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, REVENUE_BOUNDS, True)

def score_debt_ebitda(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    if v < 0:
        return 0.5
    return _score_quant_from_bounds(v, DEBT_EBITDA_BOUNDS, False)

def score_rcf_net_debt(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, RCF_NETDEBT_BOUNDS, True)

def score_ebit_interest(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBIT_INT_BOUNDS, True)

# =========================
#   PONDÉRATIONS PHARMA
# =========================

W = {
    "scale_revenue":          25.0,
    "prod_therapeutic_div":   10.0,
    "geo_diversity":          10.0,
    "patent_exposure":        10.0,
    "pipeline_quality":       10.0,
    "debt_ebitda":            10.0,
    "rcf_net_debt":            5.0,
    "ebit_interest":           5.0,
    "financial_policy":       15.0,
}
assert abs(sum(W.values()) - 100.0) < 1e-8

# =========================
#   EXHIBIT 5 : SCORE -> RATING
# =========================
EX5_BINS = [
    (1.5, "Aaa"),
    (2.5, "Aa1"), (3.5, "Aa2"), (4.5, "Aa3"),
    (5.5, "A1"),  (6.5, "A2"),  (7.5, "A3"),
    (8.5, "Baa1"),(9.5, "Baa2"),(10.5,"Baa3"),
    (11.5,"Ba1"), (12.5,"Ba2"), (13.5,"Ba3"),
    (14.5,"B1"),  (15.5,"B2"),  (16.5,"B3"),
    (17.5,"Caa1"),(18.5,"Caa2"),(19.5,"Caa3"),
    (20.5,"Ca")
]
def score_to_rating(x: float) -> str:
    for thr, lab in EX5_BINS:
        if x <= thr:
            return lab
    return "C"

# =========================
#   OUTILS : Moyennes 3 ans sur SCORES
# =========================

def weighted_average_scores_3y(s1=None, s2=None, s3=None):
    """0.2*y1 + 0.3*y2 + 0.5*y3 sur des scores déjà numérisés."""
    vals, weights = [], []
    for v, w in [(s1,0.2),(s2,0.3),(s3,0.5)]:
        fv = to_float(v)
        if fv is None:
            continue
        vals.append(fv); weights.append(w)
    if not vals:
        return None
    sw = sum(weights)
    weights = [w/sw for w in weights]
    return sum(v*w for v,w in zip(vals,weights))

def scored_revenue_3y(y1=None, y2=None, y3=None):
    s1 = score_revenue(y1)
    s2 = score_revenue(y2)
    s3 = score_revenue(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_debt_ebitda_3y(y1=None, y2=None, y3=None):
    s1 = score_debt_ebitda(y1)
    s2 = score_debt_ebitda(y2)
    s3 = score_debt_ebitda(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_rcf_net_debt_3y(y1=None, y2=None, y3=None):
    s1 = score_rcf_net_debt(y1)
    s2 = score_rcf_net_debt(y2)
    s3 = score_rcf_net_debt(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_ebit_interest_3y(y1=None, y2=None, y3=None):
    s1 = score_ebit_interest(y1)
    s2 = score_ebit_interest(y2)
    s3 = score_ebit_interest(y3)
    return weighted_average_scores_3y(s1,s2,s3)

# =========================
#   OTHER CONSIDERATIONS (±1 cran)
# =========================
def other_considerations_soft_delta(
    *,
    esg_score=None,
    liquidity_ratio=None,
    captive_ratio=None,
    regulation_score=None,
    management_score=None,
    nonwholly_sales=None,
    event_risk_score=None,
    parental_support=None,
    financial_policy_label=None
) -> float:
    """
    +delta = amélioration (score numérique diminue),
    CAP global de cette partie : ±1 cran.
    """
    delta = 0.0

    # ESG
    esg = to_float(esg_score)
    if esg is not None:
        if esg <= 2: delta += 0.25
        elif esg >= 4: delta -= 0.25

    # Liquidité
    lr = to_float(liquidity_ratio)
    if lr is not None:
        if lr > 2.0: delta += 0.25
        elif lr < 1.0: delta -= 0.5

    # Captive / JV / minoritaires
    cr = to_float(captive_ratio)
    if cr is not None:
        if cr > 0.40: delta -= 0.5
        elif cr > 0.20: delta -= 0.25

    # Régulation / risque pays
    rs = to_float(regulation_score)
    if rs is not None:
        if rs <= 2: delta += 0.25
        elif rs >= 4: delta -= 0.25

    # Management (éviter double compte si FP déjà Aa/Aaa)
    fp_alpha = _alpha_from_label(financial_policy_label)
    mgmt_contrib_allowed = (fp_alpha not in ("Aaa","Aa"))
    ms = to_float(management_score)
    if ms is not None and mgmt_contrib_allowed:
        if ms <= 2: delta += 0.25
        elif ms >= 4: delta -= 0.5

    # Non-wholly owned sales
    nws = to_float(nonwholly_sales)
    if nws is not None:
        if nws > 0.25: delta -= 0.5
        elif nws > 0.10: delta -= 0.25

    # Event risk
    er = to_float(event_risk_score)
    if er is not None:
        er = int(er)
        if er == 1: delta -= 0.5
        elif er >= 2: delta -= 1.0

    # Parental / Gov support (cap ±1)
    ps = to_float(parental_support)
    if ps is not None:
        ps = int(ps)
        if ps > 0: delta += min(ps, 1) * 0.5
        elif ps < 0: delta += max(ps, -1) * 0.5

    delta = max(-1.0, min(1.0, delta))
    return round(delta, 3)

def apply_adjustment(scorecard_aggregate: float, delta_crans: float) -> Tuple[float,str]:
    # +delta améliore la note => score numérique diminue
    adjusted = scorecard_aggregate - delta_crans
    return round(adjusted,3), score_to_rating(adjusted)

# =========================
#   SCORECARD – PHARMA
# =========================

def moodys_pharma_score_from_scores(
    *,
    name,
    s_rev,
    s_prod_div,
    s_geo_div,
    s_patent,
    s_pipeline,
    s_debt_ebitda,
    s_rcf_nd,
    s_ebit_int,
    s_finpol
) -> Dict[str, float]:

    # Scale (25%) : revenue
    score_scale = s_rev

    # Business Profile (40%) : 4 sous-facteurs à 10% chacun
    score_business = (
        W["prod_therapeutic_div"] * s_prod_div +
        W["geo_diversity"]        * s_geo_div   +
        W["patent_exposure"]      * s_patent    +
        W["pipeline_quality"]     * s_pipeline
    ) / 40.0

    # Leverage & Coverage (20%)
    score_levcov = (
        W["debt_ebitda"]   * s_debt_ebitda +
        W["rcf_net_debt"]  * s_rcf_nd      +
        W["ebit_interest"] * s_ebit_int
    ) / 20.0

    # Financial Policy (15%)
    score_policy = s_finpol

    agg = (
        0.25 * score_scale    +
        0.40 * score_business +
        0.20 * score_levcov   +
        0.15 * score_policy
    )

    rating = score_to_rating(agg)

    return {
        "name": name,
        "scorecard_aggregate": round(agg,3),
        "scorecard_rating": rating,
        "factor_scale": round(score_scale,3),
        "factor_business_profile": round(score_business,3),
        "factor_leverage_coverage": round(score_levcov,3),
        "factor_financial_policy": round(score_policy,3),
        "sf_scale_revenue": round(s_rev,3),
        "sf_product_therapeutic_diversity": round(s_prod_div,3),
        "sf_geographic_diversity": round(s_geo_div,3),
        "sf_patent_exposures": round(s_patent,3),
        "sf_pipeline_quality": round(s_pipeline,3),
        "sf_debt_ebitda": round(s_debt_ebitda,3),
        "sf_rcf_net_debt": round(s_rcf_nd,3),
        "sf_ebit_interest": round(s_ebit_int,3),
        "sf_financial_policy": round(s_finpol,3),
    }

# =========================
#   PIPELINE PRINCIPAL
# =========================
if __name__ == "__main__":
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET)

    results = []
    for _, r in df.iterrows():
        name = r[COL_NAME]

        # --------- CONVERSION FX -> USD (revenues + cash) ----------
        r = convert_row_monetary_to_usd(r)

        # --------- SCALE (Revenue 3Y, maintenant en USD) ----------
        s_rev = scored_revenue_3y(
            r.get(COL_REV_Y1), r.get(COL_REV_Y2), r.get(COL_REV_Y3)
        )

        # --------- BUSINESS PROFILE (4 quali) ----------
        s_prod_div = score_quali(r.get(COL_PROD_DIVERSITY))
        s_geo_div  = score_quali(r.get(COL_GEO_DIVERSITY))
        s_patent   = score_quali(r.get(COL_PATENT_EXPOSURE))
        s_pipe     = score_quali(r.get(COL_PIPELINE_QUALITY))

        # --------- LEVERAGE & COVERAGE 3Y (ratios, pas d'FX à appliquer) ----------
        s_debt_ebitda = scored_debt_ebitda_3y(
            r.get(COL_DEBT_EBITDA_Y1), r.get(COL_DEBT_EBITDA_Y2), r.get(COL_DEBT_EBITDA_Y3)
        )
        s_rcf_nd = scored_rcf_net_debt_3y(
            r.get(COL_RCF_NETDEBT_Y1), r.get(COL_RCF_NETDEBT_Y2), r.get(COL_RCF_NETDEBT_Y3)
        )
        s_ebit_int = scored_ebit_interest_3y(
            r.get(COL_EBIT_INT_Y1), r.get(COL_EBIT_INT_Y2), r.get(COL_EBIT_INT_Y3)
        )

        # --------- FINANCIAL POLICY ----------
        s_finpol = score_quali(r.get(COL_FINANCIAL_POLICY))

        # --------- Scorecard de base ----------
        base = moodys_pharma_score_from_scores(
            name=name,
            s_rev=s_rev,
            s_prod_div=s_prod_div,
            s_geo_div=s_geo_div,
            s_patent=s_patent,
            s_pipeline=s_pipe,
            s_debt_ebitda=s_debt_ebitda,
            s_rcf_nd=s_rcf_nd,
            s_ebit_int=s_ebit_int,
            s_finpol=s_finpol
        )

        # --------- Liquidité 3Y moyenne ----------
        liq_vals, weights = [], []
        for v, w in [(r.get(COL_LIQ_Y1),0.2),
                     (r.get(COL_LIQ_Y2),0.3),
                     (r.get(COL_LIQ_Y3),0.5)]:
            fv = to_float(v)
            if fv is not None:
                liq_vals.append(fv); weights.append(w)
        liq3 = None
        if liq_vals:
            sw = sum(weights)
            weights = [w/sw for w in weights]
            liq3 = sum(v*w for v,w in zip(liq_vals,weights))

        # --------- Other considerations (±1 cran) ----------
        delta_soft = other_considerations_soft_delta(
            esg_score=r.get(COL_ESG_SCORE),
            liquidity_ratio=liq3,
            captive_ratio=r.get(COL_CAPTIVE_RATIO),
            regulation_score=r.get(COL_REGULATION_SCORE),
            management_score=r.get(COL_MANAGEMENT_SCORE),
            nonwholly_sales=r.get(COL_NONWHOLLY_SALES),
            event_risk_score=r.get(COL_EVENT_RISK_SCORE),
            parental_support=r.get(COL_PARENTAL_SUPPORT),
            financial_policy_label=r.get(COL_FINANCIAL_POLICY)
        )

        delta_total = delta_soft

        adj_score, final_rating = apply_adjustment(base["scorecard_aggregate"], delta_total)

        results.append({
            **base,
            "inputs_liquidity_ratio_3y": liq3,
            "delta_other_considerations_soft": delta_soft,
            "delta_total_adjustment": delta_total,
            "final_adjusted_score": adj_score,
            "final_assigned_rating": final_rating
        })

    out = pd.DataFrame(results)
    out.to_csv(OUTPUT_CSV, index=False)

    print("\n✅ Pharmaceuticals – calcul terminé (3Y + FX -> USD + soft adjustments ±1). Résumé :\n")
    print(out[[ "name",
                "scorecard_aggregate","scorecard_rating",
                "delta_total_adjustment",
                "final_adjusted_score","final_assigned_rating"]].to_string(index=False))
    print(f"\n➡️ Détail sauvegardé dans {OUTPUT_CSV}")