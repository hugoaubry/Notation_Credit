### Integrated Oil & Gas – Scorecard ###
from typing import Optional, Dict, Tuple
import pandas as pd

# =========================
#   CONSTANTES I/O
# =========================

INPUT_XLSX  = "data/intput/oil_input_template.xlsx"
INPUT_SHEET = "oil_input_template"
OUTPUT_CSV  = "data/output/oil_output_scorecard.csv"

# =========================
#   COLONNES ATTENDUES
# =========================
# y1 = plus ancien ; y3 = plus récent

COL_NAME             = "name"
COL_COUNTRY          = "country"   # <- pays dans le XLSX

# Business Profile (qualitatif, 25%)
COL_BUSINESS_PROFILE = "business_profile"      # Aaa..Ca

# Financial Policy (qualitatif, 20%)
COL_FINANCIAL_POLICY = "financial_policy"      # Aaa..Ca

# --------- SCALE (20%) : 3 sous-facteurs quanti ----------
# Average Daily Production (Mboe/d)
COL_PROD_Y1          = "avg_production_y1"
COL_PROD_Y2          = "avg_production_y2"
COL_PROD_Y3          = "avg_production_y3"

# Proved Reserves (MMboe)
COL_RES_Y1           = "proved_reserves_y1"
COL_RES_Y2           = "proved_reserves_y2"
COL_RES_Y3           = "proved_reserves_y3"

# Crude Distillation Capacity (Mbbls/d)
COL_CRUDE_Y1         = "crude_capacity_y1"
COL_CRUDE_Y2         = "crude_capacity_y2"
COL_CRUDE_Y3         = "crude_capacity_y3"

# --------- PROFITABILITY & EFFICIENCY (10%) ----------
# EBIT / Average Book Capitalization (%)
COL_EBIT_BOOK_Y1     = "ebit_bookcap_y1"
COL_EBIT_BOOK_Y2     = "ebit_bookcap_y2"
COL_EBIT_BOOK_Y3     = "ebit_bookcap_y3"

# Downstream EBIT / Total Throughput Barrels ($/bbl)
COL_DOWN_EBIT_BBL_Y1 = "downstream_ebit_bbl_y1"
COL_DOWN_EBIT_BBL_Y2 = "downstream_ebit_bbl_y2"
COL_DOWN_EBIT_BBL_Y3 = "downstream_ebit_bbl_y3"

# --------- LEVERAGE & COVERAGE (25%) ----------
# EBIT / Interest Expense (x)
COL_EBIT_INT_Y1      = "ebit_interest_y1"
COL_EBIT_INT_Y2      = "ebit_interest_y2"
COL_EBIT_INT_Y3      = "ebit_interest_y3"

# RCF / Net Debt (%)
COL_RCF_NETDEBT_Y1   = "rcf_net_debt_y1"
COL_RCF_NETDEBT_Y2   = "rcf_net_debt_y2"
COL_RCF_NETDEBT_Y3   = "rcf_net_debt_y3"

# Debt / Book Capitalization (%)
COL_DEBT_BOOK_Y1     = "debt_bookcap_y1"
COL_DEBT_BOOK_Y2     = "debt_bookcap_y2"
COL_DEBT_BOOK_Y3     = "debt_bookcap_y3"

# --------- LIQUIDITÉ pour Other Considerations ----------
COL_LIQ_Y1           = "liq_y1"
COL_LIQ_Y2           = "liq_y2"
COL_LIQ_Y3           = "liq_y3"

# --------- OTHER CONSIDERATIONS (gabarit générique) ----------
COL_ESG_SCORE        = "esg_score"          # 1..5 (1 = meilleur)
COL_CAPTIVE_RATIO    = "captive_ratio"      # 0..1 (JV, minoritaires, etc.)
COL_REGULATION_SCORE = "regulation_score"   # 1..5 (1 = faible risque)
COL_MANAGEMENT_SCORE = "management_score"   # 1..5
COL_NONWHOLLY_SALES  = "nonwholly_sales"    # 0..1
COL_EVENT_RISK_SCORE = "event_risk_score"   # 0/1/2
COL_PARENTAL_SUPPORT = "parental_support"   # -3..+3 (on cape à ±1 cran)
# Notching spécifique Oil : Government Policy Framework (0..10 downward notches)
COL_GOV_POLICY       = "gov_policy_notches"

# (facultatif) cash loggé, pas utilisé directement dans la scorecard
COL_CASH_Y1          = "cash_y1"
COL_CASH_Y2          = "cash_y2"
COL_CASH_Y3          = "cash_y3"

# =========================
#   PAYS -> MONNAIES (uniquement ceux du fichier)
# =========================

COUNTRY_TO_CURRENCY = {
    "USA": "USD",
    "China": "CNY",
    "Saudi Arabia": "SAR",
}

def get_currency_for_country(country: Optional[str]) -> Optional[str]:
    if country is None or (isinstance(country, float) and pd.isna(country)):
        return None
    return COUNTRY_TO_CURRENCY.get(str(country).strip())

# =========================
#   QUALI -> NUM
# =========================

QUALI_NUM = {
    "Aaa":1.0, "Aa":3.0, "A":6.0, "Baa":9.0,
    "Ba":12.0, "B":15.0, "Caa":18.0, "Ca":20.0
}
NOTCH_TO_ALPHA = {
    "Aaa":"Aaa",
    "Aa1":"Aa","Aa2":"Aa","Aa3":"Aa",
    "A1":"A","A2":"A","A3":"A",
    "Baa1":"Baa","Baa2":"Baa","Baa3":"Baa",
    "Ba1":"Ba","Ba2":"Ba","Ba3":"Ba",
    "B1":"B","B2":"B","B3":"B",
    "Caa1":"Caa","Caa2":"Caa","Caa3":"Caa",
    "Ca":"Ca","C":"Ca","D":"Ca"
}

def _alpha_from_label(lbl: Optional[str]) -> Optional[str]:
    if lbl is None or (isinstance(lbl,float) and pd.isna(lbl)):
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
    if isinstance(x, (int,float)):
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
#   NUMERIC RANGES (Exhibit 2 Oil)
# =========================

NUM_RANGES = {
    "Aaa":(0.5,1.5),
    "Aa": (1.5,4.5),
    "A":  (4.5,7.5),
    "Baa":(7.5,10.5),
    "Ba": (10.5,13.5),
    "B":  (13.5,16.5),
    "Caa":(16.5,19.5),
    "Ca": (19.5,20.5)
}

def _interp_linear(x: float, lo: float, hi: float, ylo: float, yhi: float) -> float:
    if hi == lo:
        return (ylo + yhi)/2.0
    t = (x - lo)/(hi - lo)
    if t < 0: t = 0.0
    if t > 1: t = 1.0
    return ylo + t*(yhi - ylo)

def _score_quant_from_bounds(x: float, bounds: list, higher_is_better: bool) -> float:
    for alpha, lo, hi in bounds:
        in_band = (lo < hi and (x >= lo and x < hi)) or (lo > hi and (x <= lo and x > hi))
        if in_band:
            num_lo, num_hi = NUM_RANGES[alpha]
            if higher_is_better:
                return _interp_linear(x, lo, hi, num_hi, num_lo)
            else:
                return _interp_linear(x, lo, hi, num_lo, num_hi)
    # Hors bornes : on envoie au mieux ou au pire
    if higher_is_better:
        return 0.5 if x >= bounds[0][2] else 20.5
    return 0.5 if x <= bounds[0][1] else 20.5

# ====== Bornes spécifiques Integrated Oil (Exhibit 2) ======

AVG_PROD_BOUNDS = [
    ("Aaa", 2750.0, 1e9),
    ("Aa",  1100.0, 2750.0),
    ("A",   550.0, 1100.0),
    ("Baa", 140.0, 550.0),
    ("Ba",  55.0, 140.0),
    ("B",   20.0, 55.0),
    ("Caa", 10.0, 20.0),
    ("Ca",  0.0, 10.0),
]

RESERVES_BOUNDS = [
    ("Aaa", 10000.0, 1e9),
    ("Aa",   5000.0, 10000.0),
    ("A",    2000.0, 5000.0),
    ("Baa",   500.0, 2000.0),
    ("Ba",    100.0, 500.0),
    ("B",      30.0, 100.0),
    ("Caa",    10.0, 30.0),
    ("Ca",      0.0, 10.0),
]

CRUDE_CAP_BOUNDS = [
    ("Aaa", 3000.0, 1e9),
    ("Aa",  2000.0, 3000.0),
    ("A",   1000.0, 2000.0),
    ("Baa",  500.0, 1000.0),
    ("Ba",   250.0, 500.0),
    ("B",     50.0, 250.0),
    ("Caa",   25.0, 50.0),
    ("Ca",     0.0, 25.0),
]

EBIT_BOOKCAP_BOUNDS = [
    ("Aaa", 25.0, 100.0),
    ("Aa",  20.0, 25.0),
    ("A",   15.0, 20.0),
    ("Baa", 10.0, 15.0),
    ("Ba",   5.0, 10.0),
    ("B",    3.0, 5.0),
    ("Caa",  0.0, 3.0),
    ("Ca",  -20.0, 0.0),
]

DOWN_EBIT_BBL_BOUNDS = [
    ("Aaa", 15.0, 1e9),
    ("Aa",  10.0, 15.0),
    ("A",    7.0, 10.0),
    ("Baa",  4.0, 7.0),
    ("Ba",   2.0, 4.0),
    ("B",    1.0, 2.0),
    ("Caa",  0.0, 1.0),
    ("Ca",  -20.0, 0.0),
]

EBIT_INT_BOUNDS = [
    ("Aaa", 25.0, 1e9),
    ("Aa",  15.0, 25.0),
    ("A",    7.0, 15.0),
    ("Baa",  4.0, 7.0),
    ("Ba",   2.0, 4.0),
    ("B",    1.0, 2.0),
    ("Caa",  0.5, 1.0),
    ("Ca",   0.0, 0.5),
]

RCF_NETDEBT_BOUNDS = [
    ("Aaa", 60.0, 120.0),
    ("Aa",  40.0, 60.0),
    ("A",   30.0, 40.0),
    ("Baa", 20.0, 30.0),
    ("Ba",  10.0, 20.0),
    ("B",    5.0, 10.0),
    ("Caa",  2.0, 5.0),
    ("Ca",  -20.0, 2.0),
]

DEBT_BOOK_BOUNDS = [
    ("Aaa", -20.0,  20.0),
    ("Aa",  20.0, 30.0),
    ("A",   30.0, 40.0),
    ("Baa", 40.0, 50.0),
    ("Ba",  50.0, 60.0),
    ("B",   60.0, 70.0),
    ("Caa", 70.0, 80.0),
    ("Ca",  80.0, 120.0),
]

# =========================
#   SCORE QUANTI (par sous-facteur)
# =========================

def score_avg_production(mboed: Optional[float]) -> float:
    v = to_float(mboed)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, AVG_PROD_BOUNDS, True)

def score_reserves(mmboe: Optional[float]) -> float:
    v = to_float(mmboe)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, RESERVES_BOUNDS, True)

def score_crude_cap(mbblsd: Optional[float]) -> float:
    v = to_float(mbblsd)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, CRUDE_CAP_BOUNDS, True)

def score_ebit_bookcap(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, EBIT_BOOKCAP_BOUNDS, True)

def score_downstream_ebit_bbl(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, DOWN_EBIT_BBL_BOUNDS, True)

def score_ebit_interest(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, EBIT_INT_BOUNDS, True)

def score_rcf_net_debt(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, RCF_NETDEBT_BOUNDS, True)

def score_debt_bookcap(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, DEBT_BOOK_BOUNDS, False)

# =========================
#   PONDÉRATIONS OIL
# =========================

W = {
    "scale_prod":        10.0,
    "scale_reserves":     5.0,
    "scale_crude":        5.0,
    "business_profile":  25.0,
    "ebit_bookcap":       5.0,
    "down_ebit_bbl":      5.0,
    "ebit_interest":      7.5,
    "rcf_net_debt":      10.0,
    "debt_bookcap":       7.5,
    "financial_policy":  20.0,
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
#   OUTILS : Moyennes 3 ans (sur les SCORES)
# =========================

def weighted_average_scores_3y(s1=None, s2=None, s3=None):
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

def scored_avg_production_3y(y1=None, y2=None, y3=None):
    s1 = score_avg_production(y1)
    s2 = score_avg_production(y2)
    s3 = score_avg_production(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_reserves_3y(y1=None, y2=None, y3=None):
    s1 = score_reserves(y1)
    s2 = score_reserves(y2)
    s3 = score_reserves(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_crude_cap_3y(y1=None, y2=None, y3=None):
    s1 = score_crude_cap(y1)
    s2 = score_crude_cap(y2)
    s3 = score_crude_cap(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_ebit_bookcap_3y(y1=None, y2=None, y3=None):
    s1 = score_ebit_bookcap(y1)
    s2 = score_ebit_bookcap(y2)
    s3 = score_ebit_bookcap(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_down_ebit_bbl_3y(y1=None, y2=None, y3=None):
    s1 = score_downstream_ebit_bbl(y1)
    s2 = score_downstream_ebit_bbl(y2)
    s3 = score_downstream_ebit_bbl(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_ebit_interest_3y(y1=None, y2=None, y3=None):
    s1 = score_ebit_interest(y1)
    s2 = score_ebit_interest(y2)
    s3 = score_ebit_interest(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_rcf_net_debt_3y(y1=None, y2=None, y3=None):
    s1 = score_rcf_net_debt(y1)
    s2 = score_rcf_net_debt(y2)
    s3 = score_rcf_net_debt(y3)
    return weighted_average_scores_3y(s1,s2,s3)

def scored_debt_bookcap_3y(y1=None, y2=None, y3=None):
    s1 = score_debt_bookcap(y1)
    s2 = score_debt_bookcap(y2)
    s3 = score_debt_bookcap(y3)
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
    delta = 0.0

    esg = to_float(esg_score)
    if esg is not None:
        if esg <= 2: delta += 0.25
        elif esg >= 4: delta -= 0.25

    lr = to_float(liquidity_ratio)
    if lr is not None:
        if lr > 2.0: delta += 0.25
        elif lr < 1.0: delta -= 0.5

    cr = to_float(captive_ratio)
    if cr is not None:
        if cr > 0.40: delta -= 0.5
        elif cr > 0.20: delta -= 0.25

    rs = to_float(regulation_score)
    if rs is not None:
        if rs <= 2: delta += 0.25
        elif rs >= 4: delta -= 0.25

    fp_alpha = _alpha_from_label(financial_policy_label)
    mgmt_contrib_allowed = (fp_alpha not in ("Aaa","Aa"))
    ms = to_float(management_score)
    if ms is not None and mgmt_contrib_allowed:
        if ms <= 2: delta += 0.25
        elif ms >= 4: delta -= 0.5

    nws = to_float(nonwholly_sales)
    if nws is not None:
        if nws > 0.25: delta -= 0.5
        elif nws > 0.10: delta -= 0.25

    er = to_float(event_risk_score)
    if er is not None:
        er = int(er)
        if er == 1: delta -= 0.5
        elif er >= 2: delta -= 1.0

    ps = to_float(parental_support)
    if ps is not None:
        ps = int(ps)
        if ps > 0: delta += min(ps, 1) * 0.5
        elif ps < 0: delta += max(ps, -1) * 0.5

    delta = max(-1.0, min(1.0, delta))
    return round(delta, 3)

def apply_adjustment(scorecard_aggregate: float, delta_crans: float) -> Tuple[float,str]:
    adjusted = scorecard_aggregate - delta_crans
    return round(adjusted,3), score_to_rating(adjusted)

# =========================
#   SCORECARD – OIL
# =========================

def moodys_oil_score_from_scores(
    *,
    name,
    s_prod,
    s_reserves,
    s_crude,
    s_business_profile,
    s_ebit_book,
    s_down_bbl,
    s_ebit_int,
    s_rcf_nd,
    s_debt_book,
    s_pol
) -> Dict[str, float]:

    score_scale = (W["scale_prod"]*s_prod +
                   W["scale_reserves"]*s_reserves +
                   W["scale_crude"]*s_crude) / 20.0

    score_business = s_business_profile

    score_profit = (W["ebit_bookcap"]*s_ebit_book +
                    W["down_ebit_bbl"]*s_down_bbl) / 10.0

    score_levcov = (W["ebit_interest"]*s_ebit_int +
                    W["rcf_net_debt"]*s_rcf_nd +
                    W["debt_bookcap"]*s_debt_book) / 25.0

    score_policy = s_pol

    agg = (0.20*score_scale +
           0.25*score_business +
           0.10*score_profit +
           0.25*score_levcov +
           0.20*score_policy)

    rating = score_to_rating(agg)

    return {
        "name": name,
        "scorecard_aggregate": round(agg,3),
        "scorecard_rating": rating,
        "factor_scale": round(score_scale,3),
        "factor_business_profile": round(score_business,3),
        "factor_profitability_efficiency": round(score_profit,3),
        "factor_leverage_coverage": round(score_levcov,3),
        "factor_financial_policy": round(score_policy,3),
        "sf_scale_avg_production": round(s_prod,3),
        "sf_scale_reserves": round(s_reserves,3),
        "sf_scale_crude_capacity": round(s_crude,3),
        "sf_business_profile": round(score_business,3),
        "sf_ebit_bookcap": round(s_ebit_book,3),
        "sf_downstream_ebit_bbl": round(s_down_bbl,3),
        "sf_ebit_interest": round(s_ebit_int,3),
        "sf_rcf_net_debt": round(s_rcf_nd,3),
        "sf_debt_bookcap": round(s_debt_book,3),
        "sf_financial_policy": round(s_pol,3),
    }

# =========================
#   PIPELINE PRINCIPAL
# =========================
if __name__ == "__main__":
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET)

    results = []
    for _, r in df.iterrows():
        name    = r[COL_NAME]
        country = r.get(COL_COUNTRY)
        cc      = get_currency_for_country(country)

        s_prod = scored_avg_production_3y(
            r.get(COL_PROD_Y1), r.get(COL_PROD_Y2), r.get(COL_PROD_Y3)
        )
        s_res  = scored_reserves_3y(
            r.get(COL_RES_Y1), r.get(COL_RES_Y2), r.get(COL_RES_Y3)
        )
        s_cru  = scored_crude_cap_3y(
            r.get(COL_CRUDE_Y1), r.get(COL_CRUDE_Y2), r.get(COL_CRUDE_Y3)
        )

        s_ebit_book = scored_ebit_bookcap_3y(
            r.get(COL_EBIT_BOOK_Y1), r.get(COL_EBIT_BOOK_Y2), r.get(COL_EBIT_BOOK_Y3)
        )
        s_down_bbl  = scored_down_ebit_bbl_3y(
            r.get(COL_DOWN_EBIT_BBL_Y1), r.get(COL_DOWN_EBIT_BBL_Y2), r.get(COL_DOWN_EBIT_BBL_Y3)
        )

        s_ebit_int  = scored_ebit_interest_3y(
            r.get(COL_EBIT_INT_Y1), r.get(COL_EBIT_INT_Y2), r.get(COL_EBIT_INT_Y3)
        )
        s_rcf_nd    = scored_rcf_net_debt_3y(
            r.get(COL_RCF_NETDEBT_Y1), r.get(COL_RCF_NETDEBT_Y2), r.get(COL_RCF_NETDEBT_Y3)
        )
        s_debt_book = scored_debt_bookcap_3y(
            r.get(COL_DEBT_BOOK_Y1), r.get(COL_DEBT_BOOK_Y2), r.get(COL_DEBT_BOOK_Y3)
        )

        s_bp  = score_quali(r.get(COL_BUSINESS_PROFILE))
        s_pol = score_quali(r.get(COL_FINANCIAL_POLICY))

        base = moodys_oil_score_from_scores(
            name=name,
            s_prod=s_prod,
            s_reserves=s_res,
            s_crude=s_cru,
            s_business_profile=s_bp,
            s_ebit_book=s_ebit_book,
            s_down_bbl=s_down_bbl,
            s_ebit_int=s_ebit_int,
            s_rcf_nd=s_rcf_nd,
            s_debt_book=s_debt_book,
            s_pol=s_pol
        )

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

        gov_delta = 0.0
        gov = to_float(r.get(COL_GOV_POLICY))
        if gov is not None and gov > 0:
            gov_delta = -min(max(gov, 0), 10)

        delta_total = delta_soft + gov_delta

        adj_score, final_rating = apply_adjustment(base["scorecard_aggregate"], delta_total)

        results.append({
            "country": country,
            "currency": cc,
            **base,
            "inputs_liquidity_ratio_3y": liq3,
            "delta_other_considerations_soft": delta_soft,
            "delta_gov_policy_notching": gov_delta,
            "delta_total_adjustment": delta_total,
            "final_adjusted_score": adj_score,
            "final_assigned_rating": final_rating
        })

    out = pd.DataFrame(results)
    out.to_csv(OUTPUT_CSV, index=False)

    print("\n✅ Integrated Oil & Gas – calcul terminé (3Y + soft adjustments ±1 + gov policy). Résumé :\n")
    print(out[[ "name",
                "scorecard_aggregate","scorecard_rating",
                "delta_total_adjustment",
                "final_adjusted_score","final_assigned_rating"]].to_string(index=False))
    print(f"\n➡️ Détail sauvegardé dans {OUTPUT_CSV}")