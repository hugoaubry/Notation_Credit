### LVMH / Retail & Apparel – Scorecard Moody’s-like ###
from typing import Optional, Dict, Tuple
import pandas as pd

# =========================
#   CONSTANTES I/O
# =========================

INPUT_XLSX  = "data/intput/retail_input_template.xlsx"
OUTPUT_CSV  = "data/output/retail_output_scorecard.csv"
INPUT_SHEET = "retail_input_template"


# =========================
#   COLONNES ATTENDUES
# =========================
# y1 = plus ancien ; y3 = plus récent

COL_NAME                = "name"
COL_COUNTRY             = "country"  # colonne pays dans l'Excel (adapter si besoin)

# Business Profile (qualitatif)
COL_MARKET_CHAR         = "market_characteristics"   # Market Characteristics (Aaa..Ca)
COL_MARKET_POSITION     = "market_position"          # Market Position (Aaa..Ca)

# Profitability & Efficiency (qualitatif)
COL_REV_EARN_STAB       = "revenue_earnings_stability"  # Revenue and Earnings Stability (Aaa..Ca)

# Ratios Leverage & Coverage (quanti, mais on les recalcule si bruts dispo)
COL_DEBT_EBITDA_Y1      = "debt_ebitda_y1"
COL_DEBT_EBITDA_Y2      = "debt_ebitda_y2"
COL_DEBT_EBITDA_Y3      = "debt_ebitda_y3"

COL_RCF_NETDEBT_Y1      = "rcf_netdebt_y1"
COL_RCF_NETDEBT_Y2      = "rcf_netdebt_y2"
COL_RCF_NETDEBT_Y3      = "rcf_netdebt_y3"

COL_EBITDA_CAPEX_INT_Y1 = "ebitda_capex_int_y1"
COL_EBITDA_CAPEX_INT_Y2 = "ebitda_capex_int_y2"
COL_EBITDA_CAPEX_INT_Y3 = "ebitda_capex_int_y3"

# Liquidity ratio (Sources / Uses) – pour Other Considerations
COL_LIQ_Y1              = "liq_y1"
COL_LIQ_Y2              = "liq_y2"
COL_LIQ_Y3              = "liq_y3"

# Financial Policy (qualitatif)
COL_FINANCIAL_POLICY    = "financial_policy"

# Other considerations inputs (same esprit que l'automobile)
COL_ESG_SCORE           = "esg_score"          # 1..5 (1 = meilleur)
COL_CAPTIVE_RATIO       = "captive_ratio"      # 0..1 (souvent 0 pour le luxe)
COL_REGULATION_SCORE    = "regulation_score"   # 1..5 (1 = meilleur)
COL_MANAGEMENT_SCORE    = "management_score"   # 1..5 (1 = meilleur)
COL_NONWHOLLY_SALES     = "nonwholly_sales"    # 0..1
COL_EVENT_RISK_SCORE    = "event_risk_score"   # 0/1/2 (2 = risque élevé)
COL_PARENTAL_SUPPORT    = "parental_support"   # -3..+3 (on cape à ±1 cran d'effet)

# (facultatif) cash loggé, pas utilisé dans le score
COL_CASH_Y1             = "cash_y1"
COL_CASH_Y2             = "cash_y2"
COL_CASH_Y3             = "cash_y3"

# =========================
#   MAPPING PAYS / DEVISE / FX
# =========================

# Mapping pays -> devise locale
COUNTRY_TO_CCY = {
    "United States": "USD",
    "USA": "USD",
    "US": "USD",
    "France": "EUR",
    "Germany": "EUR",
    "Italy": "EUR",
    "Spain": "EUR",
    "Netherlands": "EUR",
    "Belgium": "EUR",
    "Luxembourg": "EUR",
    "Portugal": "EUR",
    "Ireland": "EUR",
    "United Kingdom": "GBP",
    "UK": "GBP",
    "Great Britain": "GBP",
    "Japan": "JPY",
    "Switzerland": "CHF",
    # à compléter si besoin
}

# Taux de change (1 unité de devise locale -> USD)
# À mettre à jour manuellement si besoin.
FX_TO_USD = {
    "USD": 1.0,
    "EUR": 1.10,   # exemple : 1 EUR ≈ 1.10 USD
    "GBP": 1.25,   # exemple : 1 GBP ≈ 1.25 USD
    "JPY": 0.007,  # exemple : 1 JPY ≈ 0.007 USD
    "CHF": 1.10,   # exemple : 1 CHF ≈ 1.10 USD
}

def get_fx_from_country(country: Optional[str]) -> float:
    if country is None or (isinstance(country, float) and pd.isna(country)):
        return 1.0
    c = str(country).strip()
    ccy = COUNTRY_TO_CCY.get(c, "USD")
    return FX_TO_USD.get(ccy, 1.0)

# =========================
#   COLONNES BRUTES
# =========================
# revenue_y* en UNITÉ de devise locale (idéalement milliards), qui sera convertie en USD
RAW = dict(
    revenue = ["revenue_y1","revenue_y2","revenue_y3"],      # scale factor utilisera revenue_y3
    ebit = ["ebit_y1","ebit_y2","ebit_y3"],
    ebitda = ["ebitda_y1","ebitda_y2","ebitda_y3"],
    interest_exp = ["interest_exp_y1","interest_exp_y2","interest_exp_y3"],   # valeur absolue positive
    ocf = ["ocf_y1","ocf_y2","ocf_y3"],                      # operating cash flow
    capex = ["capex_y1","capex_y2","capex_y3"],              # POSITIF (on prendra abs sinon)
    dividends = ["dividends_y1","dividends_y2","dividends_y3"],
    delta_wcap = ["delta_wcap_y1","delta_wcap_y2","delta_wcap_y3"],  # + = hausse de BFR
    st_debt = ["st_debt_y1","st_debt_y2","st_debt_y3"],
    cash_sti = ["cash_sti_y1","cash_sti_y2","cash_sti_y3"],
    total_debt = ["total_debt_y1","total_debt_y2","total_debt_y3"],
    lease_liab_current = ["lease_liab_current_y1","lease_liab_current_y2","lease_liab_current_y3"],
    lease_liab_noncurrent = ["lease_liab_noncurrent_y1","lease_liab_noncurrent_y2","lease_liab_noncurrent_y3"],
    lease_payments = ["lease_payments_y1","lease_payments_y2","lease_payments_y3"],
)

# =========================
#   QUALI -> NUM
# =========================

QUALI_NUM = {"Aaa":1.0, "Aa":3.0, "A":6.0, "Baa":9.0, "Ba":12.0, "B":15.0, "Caa":18.0, "Ca":20.0}
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
    return QUALI_NUM[a] if a else 12.0  # défaut Ba si inconnu

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

def safe_abs(x):
    v = to_float(x)
    if v is None:
        return None
    return abs(v)

# =========================
#   NUMERIC RANGES (génériques Moody’s)
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
    # Hors bornes -> on envoie au mieux ou au pire
    if higher_is_better:
        return 0.5 if x >= bounds[0][2] else 20.5
    return 0.5 if x <= bounds[0][1] else 20.5

# =========================
#   BORNES SPÉCIFIQUES RETAIL & APPAREL
# =========================
# 1) Scale : Revenue (USD bn), factor 15%
REVENUE_BOUNDS = [
    ("Aaa",100e9,100e10),
    ("Aa", 50e9,100e9),
    ("A",  25e9,50e9),
    ("Baa",10e9,25e9),
    ("Ba",  4e9,10e9),
    ("B",   1e9,4e9),
    ("Caa", 0.5e9,1.5e9),
    ("Ca",  -1e9,0.5e9),
]

# 2) Debt/EBITDA (Retail table : ≤0.5 ; 0.5–1 ; 1–2 ; 2–3 ; 3–4 ; 4–6 ; 6–8 ; >8 ; Ca endpoint ~12+)
DEBT_EBITDA_BOUNDS = [
    ("Aaa", -1e9,  0.5),
    ("Aa",  0.5,  1.0),
    ("A",   1.0,  2.0),
    ("Baa", 2.0,  3.0),
    ("Ba",  3.0,  4.0),
    ("B",   4.0,  6.0),
    ("Caa", 6.0,  8.0),
    ("Ca",  8.0,  1e9),
]

# 3) RCF/Net Debt (%) (Retail table : 60–80 ; 40–60 ; 25–40 ; 15–25 ; 10–15 ; 5–10 ; <5 ; Ca endpoint 0)
RCF_NETDEBT_BOUNDS = [
    ("Aaa", 80.0,1e9),
    ("Aa",  60.0,80.0),
    ("A",   40.0,60.0),
    ("Baa", 25.0,40.0),
    ("Ba",  15.0,25.0),
    ("B",   10.0,15.0),
    ("Caa", 5.0,10.0),
    ("Ca",  -1e9,5.0),
]

# 4) (EBITDA - Capex) / Interest Expense (x)
EBITDA_CAPEX_INT_BOUNDS = [
    ("Aaa",20.0,1e9),
    ("Aa", 16.0,20.0),
    ("A",  9.0,16.0),
    ("Baa",4.5,9),
    ("Ba", 2.5,4.5),
    ("B",  1.25,2.5),
    ("Caa",0.25,1.25),
    ("Ca", -1e9,0.25),
]

# =========================
#   SCORE QUANTI
# =========================

def score_revenue_scale(rev_usd_bn: Optional[float]) -> float:
    v = to_float(rev_usd_bn)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, REVENUE_BOUNDS, True)

def score_debt_ebitda(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, DEBT_EBITDA_BOUNDS, False)

def score_rcf_netdebt(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, RCF_NETDEBT_BOUNDS, True)

def score_ebitda_capex_int(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBITDA_CAPEX_INT_BOUNDS, True)

# =========================
#   PONDÉRATIONS RETAIL & APPAREL
# =========================
# Sub-factors, total = 100
W = {
    "revenue_scale":        15.0,  # Scale
    "market_characteristics":10.0, # Business Profile
    "market_position":      10.0,  # Business Profile
    "rev_earn_stability":   10.0,  # Profitability & Efficiency
    "debt_ebitda":          15.0,  # Leverage & Coverage
    "rcf_net_debt":         10.0,  # Leverage & Coverage
    "ebitda_capex_int":     15.0,  # Leverage & Coverage
    "financial_policy":     15.0,  # Financial Policy
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
#   OUTILS 3 ANS
# =========================

def weighted_average_scores_3y(s1=None, s2=None, s3=None):
    """ Moyenne pondérée de scores : 0.2*y1 + 0.3*y2 + 0.5*y3. """
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

def scored_debt_ebitda_3y(y1=None, y2=None, y3=None):
    s1 = score_debt_ebitda(y1)
    s2 = score_debt_ebitda(y2)
    s3 = score_debt_ebitda(y3)
    return weighted_average_scores_3y(s1, s2, s3)

def scored_rcf_netdebt_3y(y1=None, y2=None, y3=None):
    s1 = score_rcf_netdebt(y1)
    s2 = score_rcf_netdebt(y2)
    s3 = score_rcf_netdebt(y3)
    return weighted_average_scores_3y(s1, s2, s3)

def scored_ebitda_capex_int_3y(y1=None, y2=None, y3=None):
    s1 = score_ebitda_capex_int(y1)
    s2 = score_ebitda_capex_int(y2)
    s3 = score_ebitda_capex_int(y3)
    return weighted_average_scores_3y(s1, s2, s3)

# =========================
#   AJUSTEMENTS "OTHER CONSIDERATIONS" (CAP ±1 cran)
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
    +delta = amélioration (score numérique diminue), CAP total ±1.
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

    # Captive finance (surtout pour retail non-luxe)
    cr = to_float(captive_ratio)
    if cr is not None:
        if cr > 0.40: delta -= 0.5
        elif cr > 0.20: delta -= 0.25

    # Régulation
    rs = to_float(regulation_score)
    if rs is not None:
        if rs <= 2: delta += 0.25
        elif rs >= 4: delta -= 0.25

    # Management (évite double-compte si FP déjà très élevée)
    fp_alpha = _alpha_from_label(financial_policy_label)
    mgmt_contrib_allowed = (fp_alpha not in ("Aaa","Aa"))
    ms = to_float(management_score)
    if ms is not None and mgmt_contrib_allowed:
        if ms <= 2: delta += 0.25
        elif ms >= 4: delta -= 0.5

    # Non-wholly owned (subs, IPO partiels, etc.)
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
    adjusted = scorecard_aggregate - delta_crans
    return round(adjusted,3), score_to_rating(adjusted)

# =========================
#   RE-CALCUL MOODY’S DES RATIOS (DEPUIS BRUTS)
# =========================

def compute_adjusted_debt(total_debt, lease_cur=None, lease_non=None):
    td = to_float(total_debt) or 0.0
    lc = to_float(lease_cur) or 0.0
    ln = to_float(lease_non) or 0.0
    return td + lc + ln

def safe_interest(interest, revenue):
    ival = abs(to_float(interest) or 0.0)
    rev  = abs(to_float(revenue) or 0.0)
    floor = max(ival, 1e-6 * rev)
    return floor

def safe_ebitda(ebitda, revenue):
    eval_ = to_float(ebitda) or 0.0
    rev   = abs(to_float(revenue) or 0.0)
    return max(eval_, 1e-6 * rev)

def compute_rcf_fcf(ocf, capex_pos, dividends, delta_wcap=None):
    ocf_v = to_float(ocf) or 0.0
    cap_v = abs(to_float(capex_pos) or 0.0)
    div_v = to_float(dividends) or 0.0
    if delta_wcap is not None:
        dwc = to_float(delta_wcap) or 0.0
        rcf = ocf_v - dwc - div_v   # RCF ~ OCF - ΔBFR - Dividendes
    else:
        rcf = ocf_v - div_v
    fcf = ocf_v - cap_v - div_v
    return rcf, fcf

def liquidity_ratio_moodys(cash_sti, ocf, st_debt, capex_pos, dividends=0.0, lease_payments=0.0):
    sources = (to_float(cash_sti) or 0.0) + (to_float(ocf) or 0.0)
    uses    = (to_float(st_debt) or 0.0) + abs(to_float(capex_pos) or 0.0) \
              + (to_float(dividends) or 0.0) + (to_float(lease_payments) or 0.0)
    return sources / max(uses, 1e-6)

def derive_ratios_from_raw(row, suffix: str) -> dict:
    """
    Recalcule les ratios Moody's Retail & Apparel pour une année suffix (y1/y2/y3).
    Remplit :
      - debt_ebitda_suffix
      - rcf_netdebt_suffix
      - ebitda_capex_int_suffix
      - liq_suffix
    """
    R = lambda col: row.get(f"{col}_{suffix}")
    revenue    = R("revenue")
    ebitda     = R("ebitda")
    interest   = R("interest_exp")
    ocf        = R("ocf")
    capex_pos  = R("capex")
    dividends  = R("dividends")
    delta_wc   = R("delta_wcap")
    st_debt    = R("st_debt")
    cash_sti   = R("cash_sti")
    tot_debt   = R("total_debt")
    lease_cur  = R("lease_liab_current")
    lease_non  = R("lease_liab_noncurrent")
    lease_pay  = R("lease_payments")

    out = {}

    adj_debt = None
    if tot_debt is not None:
        adj_debt = compute_adjusted_debt(tot_debt, lease_cur, lease_non)

    # Debt/EBITDA
    if adj_debt is not None and ebitda is not None:
        ebitda_val = safe_ebitda(ebitda, revenue)
        out[f"debt_ebitda_{suffix}"] = adj_debt / ebitda_val

    # (EBITDA - Capex) / Interest Expense
    if ebitda is not None and capex_pos is not None and interest is not None:
        ebitda_val = safe_ebitda(ebitda, revenue)
        cap_val    = abs(to_float(capex_pos) or 0.0)
        numer      = ebitda_val - cap_val
        numer      = max(numer, 0.0)  # coverage <0 => on floor à 0
        cov        = numer / safe_interest(interest, revenue)
        out[f"ebitda_capex_int_{suffix}"] = cov

    # RCF / Net Debt
    if adj_debt is not None and cash_sti is not None and ocf is not None:
        net_debt = adj_debt - (to_float(cash_sti) or 0.0)
        rcf, _   = compute_rcf_fcf(ocf, capex_pos, dividends, delta_wc)
        if net_debt > 0:
            out[f"rcf_netdebt_{suffix}"] = (rcf / max(net_debt,1e-6)) * 100.0
        else:
            # Net debt <=0 : approximation simple
            out[f"rcf_netdebt_{suffix}"] = 100.0 if rcf > 0 else 0.0

    # Liquidity ratio S/U
    if st_debt is not None and cash_sti is not None and ocf is not None and capex_pos is not None:
        liq = liquidity_ratio_moodys(cash_sti, ocf, st_debt, capex_pos, dividends, lease_pay)
        out[f"liq_{suffix}"] = liq

    return out

def convert_row_monetary_to_usd(row: pd.Series, fx: float) -> pd.Series:
    """
    Multiplie toutes les colonnes monétaires RAW de la ligne par fx
    pour les exprimer en USD (même unité qu'à l'origine, ex : bn).
    """
    if fx == 1.0:
        return row
    for cols in RAW.values():
        for col in cols:
            if col in row:
                val = to_float(row[col])
                if val is not None:
                    row[col] = val * fx
    return row

# =========================
#   SCORECARD – RETAIL & APPAREL
# =========================

def moodys_retail_score_from_scores(
    *,
    name,
    s_scale,
    s_market_char,
    s_market_pos,
    s_rev_earn,
    s_debt,
    s_rcfnd,
    s_cov,
    s_pol
) -> Dict[str, float]:

    # Facteurs
    score_scale    = s_scale
    score_business = (W["market_characteristics"]*s_market_char + W["market_position"]*s_market_pos) / 20.0
    score_profit   = s_rev_earn
    score_levcov   = (W["debt_ebitda"]*s_debt + W["rcf_net_debt"]*s_rcfnd + W["ebitda_capex_int"]*s_cov) / 40.0
    score_policy   = s_pol

    agg = (0.15*score_scale +
           0.20*score_business +
           0.10*score_profit +
           0.40*score_levcov +
           0.15*score_policy)

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
        "sf_revenue_scale": round(s_scale,3),
        "sf_market_characteristics": round(s_market_char,3),
        "sf_market_position": round(s_market_pos,3),
        "sf_revenue_earnings_stability": round(s_rev_earn,3),
        "sf_debt_ebitda": round(s_debt,3),
        "sf_rcf_net_debt": round(s_rcfnd,3),
        "sf_ebitda_capex_int": round(s_cov,3),
        "sf_financial_policy": round(s_pol,3),
    }

# =========================
#   PIPELINE PRINCIPAL
# =========================
if __name__ == "__main__":
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET)

    results = []
    for _, r in df.iterrows():
        name = r[COL_NAME]

        # --- Conversion devise locale -> USD en fonction du pays
        country = r.get(COL_COUNTRY, None)
        fx = get_fx_from_country(country)
        r = convert_row_monetary_to_usd(r, fx)

        # --- Recalcule Moody's depuis BRUTS si dispo (désormais en USD)
        for suf in ("y1","y2","y3"):
            upd = derive_ratios_from_raw(r, suf)
            for k,v in upd.items():
                if v is not None:
                    r[k] = v

        # --- SCALE : Revenue (USD bn), on prend l'année la plus récente (y3)
        rev_raw = r.get(RAW["revenue"][2]) if RAW["revenue"][2] in r else None
        s_scale = score_revenue_scale(rev_raw)

        # --- LEVERAGE & COVERAGE : scores 3Y pondérés
        s_debt  = scored_debt_ebitda_3y(
            r.get(COL_DEBT_EBITDA_Y1),
            r.get(COL_DEBT_EBITDA_Y2),
            r.get(COL_DEBT_EBITDA_Y3)
        )
        s_rcfnd = scored_rcf_netdebt_3y(
            r.get(COL_RCF_NETDEBT_Y1),
            r.get(COL_RCF_NETDEBT_Y2),
            r.get(COL_RCF_NETDEBT_Y3)
        )
        s_cov   = scored_ebitda_capex_int_3y(
            r.get(COL_EBITDA_CAPEX_INT_Y1),
            r.get(COL_EBITDA_CAPEX_INT_Y2),
            r.get(COL_EBITDA_CAPEX_INT_Y3)
        )

        # --- QUALITATIFS
        s_market_char = score_quali(r.get(COL_MARKET_CHAR))
        s_market_pos  = score_quali(r.get(COL_MARKET_POSITION))
        s_rev_earn    = score_quali(r.get(COL_REV_EARN_STAB))
        s_pol         = score_quali(r.get(COL_FINANCIAL_POLICY))

        # --- Scorecard outcome
        base = moodys_retail_score_from_scores(
            name=name,
            s_scale=s_scale,
            s_market_char=s_market_char,
            s_market_pos=s_market_pos,
            s_rev_earn=s_rev_earn,
            s_debt=s_debt,
            s_rcfnd=s_rcfnd,
            s_cov=s_cov,
            s_pol=s_pol
        )

        # --- Liquidity 3Y moyenne pour Other Considerations
        liq_vals, weights = [], []
        for v, w in [(r.get(COL_LIQ_Y1),0.2),(r.get(COL_LIQ_Y2),0.3),(r.get(COL_LIQ_Y3),0.5)]:
            fv = to_float(v)
            if fv is not None:
                liq_vals.append(fv); weights.append(w)
        liq3 = None
        if liq_vals:
            sw = sum(weights)
            weights = [w/sw for w in weights]
            liq3 = sum(v*w for v,w in zip(liq_vals,weights))

        # --- Ajustements hors scorecard (soft, cap ±1 cran)
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

        adj_score, final_rating = apply_adjustment(base["scorecard_aggregate"], delta_soft)

        results.append({
            **base,
            "inputs_liquidity_ratio_3y": liq3,
            "delta_other_considerations_soft": delta_soft,
            "final_adjusted_score": adj_score,
            "final_assigned_rating": final_rating
        })

    out = pd.DataFrame(results)
    out.to_csv(OUTPUT_CSV, index=False)

    print("\n✅ Retail & Apparel – calcul terminé (3Y + soft adjustments ±1). Résumé :\n")
    print(out[[ "name",
                "scorecard_aggregate","scorecard_rating",
                "delta_other_considerations_soft",
                "final_adjusted_score","final_assigned_rating"]].to_string(index=False))
    print(f"\n➡️ Détail sauvegardé dans {OUTPUT_CSV}")