# =========================
#   Mining – Scorecard Moody's (version avec FX)
# =========================

from typing import Optional, Dict, Tuple
import pandas as pd

# ====== I/O ======
INPUT_XLSX  = "data/intput/mining_input_template.xlsx"
INPUT_SHEET = "mining_input_template"
OUTPUT_CSV  = "data/output/mining_output_scorecard.csv"

# ====== Colonnes principales ======
COL_NAME             = "name"
COL_COUNTRY          = "country"  
COL_BUSINESS_PROFILE = "business_profile"
COL_FINANCIAL_POLICY = "financial_policy"

# Ratios 3Y (déjà présents dans le xlsx, mais écrasés si on peut les recalculer)
COL_EBIT_MARGIN_Y1      = "ebit_margin_y1"
COL_EBIT_MARGIN_Y2      = "ebit_margin_y2"
COL_EBIT_MARGIN_Y3      = "ebit_margin_y3"
COL_DEBT_EBITDA_Y1      = "debt_ebitda_y1"
COL_DEBT_EBITDA_Y2      = "debt_ebitda_y2"
COL_DEBT_EBITDA_Y3      = "debt_ebitda_y3"
COL_EBITDA_CAPEX_INT_Y1 = "ebitda_capex_int_y1"
COL_EBITDA_CAPEX_INT_Y2 = "ebitda_capex_int_y2"
COL_EBITDA_CAPEX_INT_Y3 = "ebitda_capex_int_y3"
COL_RCF_DEBT_Y1         = "rcf_debt_y1"
COL_RCF_DEBT_Y2         = "rcf_debt_y2"
COL_RCF_DEBT_Y3         = "rcf_debt_y3"
COL_LIQ_Y1              = "liq_y1"
COL_LIQ_Y2              = "liq_y2"
COL_LIQ_Y3              = "liq_y3"

# Other Considerations
COL_ESG_SCORE        = "esg_score"
COL_CAPTIVE_RATIO    = "captive_ratio"
COL_REGULATION_SCORE = "regulation_score"
COL_MANAGEMENT_SCORE = "management_score"
COL_NONWHOLLY_SALES  = "nonwholly_sales"
COL_EVENT_RISK_SCORE = "event_risk_score"
COL_PARENTAL_SUPPORT = "parental_support"

# Colonnes brutes pour recalculer les ratios Moody's
RAW = dict(
    revenue    = ["revenue_y1","revenue_y2","revenue_y3"],
    ebit       = ["ebit_y1","ebit_y2","ebit_y3"],
    ebitda     = ["ebitda_y1","ebitda_y2","ebitda_y3"],
    interest   = ["interest_exp_y1","interest_exp_y2","interest_exp_y3"],
    ocf        = ["ocf_y1","ocf_y2","ocf_y3"],
    capex      = ["capex_y1","capex_y2","capex_y3"],
    dividends  = ["dividends_y1","dividends_y2","dividends_y3"],
    delta_wc   = ["delta_wcap_y1","delta_wcap_y2","delta_wcap_y3"],
    st_debt    = ["st_debt_y1","st_debt_y2","st_debt_y3"],
    cash_sti   = ["cash_sti_y1","cash_sti_y2","cash_sti_y3"],
    total_debt = ["total_debt_y1","total_debt_y2","total_debt_y3"],
    lease_cur  = ["lease_liab_current_y1","lease_liab_current_y2","lease_liab_current_y3"],
    lease_non  = ["lease_liab_noncurrent_y1","lease_liab_noncurrent_y2","lease_liab_noncurrent_y3"],
    lease_pay  = ["lease_payments_y1","lease_payments_y2","lease_payments_y3"],
)

# =========================
#   FX PAR PAYS (SEULEMENT CEUX DU XLSX)
# =========================
# USA      -> USD
# Australia -> AUD converti en USD
FX_BY_COUNTRY = {
    "USA":       ("USD", 1.00),  # déjà en USD
    "Australia": ("AUD", 0.65),  # 1 AUD ≈ 0.65 USD (approx, à affiner si besoin)
}

def get_fx_factor(country: Optional[str]) -> Tuple[str, float]:
    if country is None:
        return ("USD", 1.0)
    c = str(country).strip()
    if c in FX_BY_COUNTRY:
        return FX_BY_COUNTRY[c]
    # fallback : pas de conversion si pays non mappé
    return ("USD", 1.0)

def convert_row_monetary_to_usd(row: pd.Series) -> pd.Series:
    """
    Convertit toutes les colonnes MONÉTAIRES brutes (revenue, ebit, ebitda, interest,
    ocf, capex, dividends, delta_wcap, st_debt, cash_sti, total_debt, leases, lease_payments)
    en USD en fonction du pays, pour y1 / y2 / y3.
    """
    country = row.get(COL_COUNTRY)
    _, fx = get_fx_factor(country)

    if fx == 1.0:
        # USA -> déjà en USD, on ne touche rien
        return row

    for key, cols in RAW.items():
        for col in cols:
            if col in row:
                v = row[col]
                if v is not None and not (isinstance(v, float) and pd.isna(v)):
                    from math import isnan
                    try:
                        val = to_float(v)
                    except Exception:
                        val = None
                    if val is not None:
                        row[col] = val * fx
    return row

# =========================
#   QUALITATIF -> NUMÉRIQUE
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
    if lbl is None:
        return None
    if isinstance(lbl,float) and pd.isna(lbl):
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
    return QUALI_NUM[a] if a else 12.0   # défaut = Ba

# =========================
#   PARSING NUMÉRIQUE
# =========================

def to_float(x) -> Optional[float]:
    if x is None:
        return None
    if isinstance(x,(int,float)):
        if pd.isna(x):
            return None
        return float(x)
    s = str(x).strip()
    if s == "":
        return None
    s = s.replace(",","")
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    if s.endswith("%"):
        s = s[:-1].strip()
    try:
        v = float(s)
    except Exception:
        return None
    return -v if neg else v

# =========================
#   GRILLES MINING (Exhibit 2)
# =========================

NUM_RANGES = {
    "Aaa":(0.5,1.5),
    "Aa": (1.5,4.5),
    "A":  (4.5,7.5),
    "Baa":(7.5,10.5),
    "Ba": (10.5,13.5),
    "B":  (13.5,16.5),
    "Caa":(16.5,19.5),
    "Ca": (19.5,20.5),
    "C": (20.5,21.5)
}

def _interp_linear(x: float, lo: float, hi: float, ylo: float, yhi: float) -> float:
    if hi == lo:
        return (ylo + yhi)/2.0
    t = (x - lo)/(hi - lo)
    t = max(0.0, min(1.0, t))
    return ylo + t*(yhi - ylo)

def _score_quant_from_bounds(x: float, bounds: list, higher_is_better: bool) -> float:
    for alpha, lo, hi in bounds:
        in_band = (lo < hi and lo <= x < hi) or (lo > hi and hi < x <= lo)
        if in_band:
            num_lo, num_hi = NUM_RANGES[alpha]
            if higher_is_better:
                return _interp_linear(x, lo, hi, num_hi, num_lo)
            else:
                return _interp_linear(x, lo, hi, num_lo, num_hi)
    if higher_is_better:
        return 0.5 if x >= bounds[0][2] else 20.5
    return 0.5 if x <= bounds[0][1] else 20.5

# 1) Scale – Revenue (USD)
REVENUE_BOUNDS = [
    ("Aaa", 100e9, 1e18),
    ("Aa",   50e9, 100e9),
    ("A",    25e9, 50e9),
    ("Baa",  10e9, 25e9),
    ("Ba",    4e9, 10e9),
    ("B",    1.5e9, 4e9),
    ("Caa",   1e9, 1.5e9),
    ("Ca",   0.5e9, 1e9),
    ("C",   -1e9, 0.5e9),
]
EBIT_MARGIN_BOUNDS = [
    ("Aaa", 60.0, 100.0),
    ("Aa",  35.0, 60.0),
    ("A",   25.0, 35.0),
    ("Baa", 20.0, 25.0),
    ("Ba",  15.0, 20.0),
    ("B",   10.0, 15.0),
    ("Caa", 5.0, 10.0),
    ("Ca",  0.0, 5.0),
    ("C",   -20.0, 0.0),
]
DEBT_EBITDA_BOUNDS = [
    ("Aaa", 0.0, 0.5),
    ("Aa",  0.5, 1.0),
    ("A",   1.0, 2.0),
    ("Baa", 2.0, 3.0),
    ("Ba",  3.0, 4.0),
    ("B",   4.0, 6.0),
    ("Caa", 6.0, 7.5),
    ("Ca",  7.5, 9.0),
    ("C",   9.0, 50.0),
]
EBITDA_CAPEX_INT_BOUNDS = [
    ("Aaa", 20.0, 50.0),
    ("Aa",  16.0, 20.0),
    ("A",    9.0, 16.0),
    ("Baa",  4.5, 9.0),
    ("Ba",   2.5, 4.5),
    ("B",    1.25, 2.5),
    ("Caa",  0.25, 1.25),
    ("Ca",   0.0, 0.25),
    ("C",   -50.0, 0.0),
]
RCF_DEBT_BOUNDS = [
    ("Aaa", 100.0, 200.0),
    ("Aa",   75.0, 100.0),
    ("A",    50.0, 75.0),
    ("Baa",  30.0, 50.0),
    ("Ba",   20.0, 30.0),
    ("B",    10.0, 20.0),
    ("Caa",   5.0, 10.0),
    ("Ca",    2.5, 5.0),
    ("C",    -50.0, 2.5),
]

def score_revenue_scale(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, REVENUE_BOUNDS, True)

def score_ebit_margin(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBIT_MARGIN_BOUNDS, True)

def score_debt_ebitda(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, DEBT_EBITDA_BOUNDS, False)

def score_ebitda_capex_int(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBITDA_CAPEX_INT_BOUNDS, True)

def score_rcf_debt(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, RCF_DEBT_BOUNDS, True)

# =========================
#   PONDÉRATIONS MINING
# =========================
# Scale 15 / Business 30 / Profit 10 / Lev+Cov 25 / FP 20
W = {
    "revenue_scale":    15.0,
    "business_profile": 30.0,
    "ebit_margin":      10.0,
    "debt_ebitda":      10.0,
    "ebitda_capex_int": 10.0,
    "rcf_debt":         5.0,
    "financial_policy": 20.0,
}
assert abs(sum(W.values()) - 100.0) < 1e-8

# =========================
#   SCORE -> RATING (Exhibit 5)
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

def score_3y(score_fn, y1=None, y2=None, y3=None) -> Optional[float]:
    vals, weights = [], []
    for v, w in [(y1,0.2),(y2,0.3),(y3,0.5)]:
        fv = to_float(v)
        if fv is None:
            continue
        vals.append(score_fn(fv)); weights.append(w)
    if not vals:
        return None
    sw = sum(weights)
    weights = [w/sw for w in weights]
    return sum(v*w for v,w in zip(vals, weights))

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
    d = 0.0

    esg = to_float(esg_score)
    if esg is not None:
        if esg <= 2: d += 0.25
        elif esg >= 4: d -= 0.25

    lr = to_float(liquidity_ratio)
    if lr is not None:
        if lr > 2.0: d += 0.25
        elif lr < 1.0: d -= 0.5

    cr = to_float(captive_ratio)
    if cr is not None:
        if cr > 0.40: d -= 0.5
        elif cr > 0.20: d -= 0.25

    rs = to_float(regulation_score)
    if rs is not None:
        if rs <= 2: d += 0.25
        elif rs >= 4: d -= 0.25

    fp_alpha = _alpha_from_label(financial_policy_label)
    ms = to_float(management_score)
    if ms is not None and fp_alpha not in ("Aaa","Aa"):
        if ms <= 2: d += 0.25
        elif ms >= 4: d -= 0.5

    nws = to_float(nonwholly_sales)
    if nws is not None:
        if nws > 0.25: d -= 0.5
        elif nws > 0.10: d -= 0.25

    er = to_float(event_risk_score)
    if er is not None:
        er = int(er)
        if er == 1: d -= 0.5
        elif er >= 2: d -= 1.0

    ps = to_float(parental_support)
    if ps is not None:
        ps = int(ps)
        if ps > 0: d += min(ps,1)*0.5
        elif ps < 0: d += max(ps,-1)*0.5

    d = max(-1.0, min(1.0, d))
    return round(d, 3)

def apply_adjustment(scorecard_aggregate: float, delta_crans: float) -> Tuple[float,str]:
    adj = scorecard_aggregate - delta_crans
    return round(adj,3), score_to_rating(adj)

# =========================
#   RE-CALCUL DES RATIOS DEPUIS BRUTS
# =========================

def compute_adjusted_debt(total_debt, lease_cur=None, lease_non=None):
    td = to_float(total_debt) or 0.0
    lc = to_float(lease_cur) or 0.0
    ln = to_float(lease_non) or 0.0
    return td + lc + ln

def safe_interest(interest, revenue):
    ival = abs(to_float(interest) or 0.0)
    rev  = abs(to_float(revenue) or 0.0)
    return max(ival, 1e-6 * rev)

def safe_ebitda(ebitda, revenue):
    val = to_float(ebitda) or 0.0
    rev = abs(to_float(revenue) or 0.0)
    return max(val, 1e-6 * rev)

def compute_rcf(ocf, dividends, delta_wc=None):
    ocf_v = to_float(ocf) or 0.0
    div_v = to_float(dividends) or 0.0
    if delta_wc is not None:
        dwc = to_float(delta_wc) or 0.0
        return ocf_v - dwc - div_v
    return ocf_v - div_v

def liquidity_ratio_moodys(cash_sti, ocf, st_debt, capex_pos,
                           dividends=0.0, lease_payments=0.0):
    sources = (to_float(cash_sti) or 0.0) + (to_float(ocf) or 0.0)
    uses    = (to_float(st_debt) or 0.0) + abs(to_float(capex_pos) or 0.0) \
              + (to_float(dividends) or 0.0) + (to_float(lease_payments) or 0.0)
    return sources / max(uses, 1e-6)

def derive_ratios_from_raw(row: pd.Series, suffix: str) -> dict:
    def R(key): return row.get(f"{key}_{suffix}")
    revenue   = R("revenue")
    ebit      = R("ebit")
    ebitda    = R("ebitda")
    interest  = R("interest_exp")
    ocf       = R("ocf")
    capex     = R("capex")
    dividends = R("dividends")
    delta_wc  = R("delta_wcap")
    st_debt   = R("st_debt")
    cash_sti  = R("cash_sti")
    tot_debt  = R("total_debt")
    lease_cur = R("lease_liab_current")
    lease_non = R("lease_liab_noncurrent")
    lease_pay = R("lease_payments")

    out = {}

    # EBIT margin
    rev_v  = to_float(revenue)
    ebit_v = to_float(ebit)
    if rev_v not in (None, 0) and ebit_v is not None:
        out[f"ebit_margin_{suffix}"] = (ebit_v / rev_v) * 100.0

    # dette ajustée
    adj_debt = compute_adjusted_debt(tot_debt, lease_cur, lease_non) if tot_debt is not None else None

    # Debt / EBITDA
    ebitda_v = to_float(ebitda)
    if adj_debt is not None and ebitda_v is not None:
        out[f"debt_ebitda_{suffix}"] = adj_debt / safe_ebitda(ebitda_v, revenue)

    # (EBITDA - capex)/interest
    if ebitda_v is not None and capex is not None and interest is not None and revenue is not None:
        cap_v = abs(to_float(capex) or 0.0)
        num   = ebitda_v - cap_v
        denom = safe_interest(interest, revenue)
        out[f"ebitda_capex_int_{suffix}"] = num / denom

    # RCF / Debt
    if adj_debt is not None and ocf is not None:
        rcf = compute_rcf(ocf, dividends, delta_wc)
        if adj_debt > 0:
            out[f"rcf_debt_{suffix}"] = (rcf / max(adj_debt,1e-6)) * 100.0
        else:
            out[f"rcf_debt_{suffix}"] = 100.0 if rcf > 0 else 0.0

    # Liquidity Sources / Uses
    if st_debt is not None and cash_sti is not None and ocf is not None and capex is not None:
        out[f"liq_{suffix}"] = liquidity_ratio_moodys(cash_sti, ocf, st_debt, capex, dividends, lease_pay)

    return out

# =========================
#   SCORECARD MINING
# =========================

def moodys_mining_score_from_scores(
    *,
    name,
    s_scale,
    s_bp,
    s_ebit,
    s_debt,
    s_ecint,
    s_rcf,
    s_pol
) -> Dict[str,float]:

    factor_scale    = s_scale
    factor_business = s_bp
    factor_profit   = s_ebit
    factor_levcov   = (W["debt_ebitda"]*s_debt +
                       W["ebitda_capex_int"]*s_ecint +
                       W["rcf_debt"]*s_rcf) / 25.0
    factor_policy   = s_pol

    agg = (0.15*factor_scale +
           0.30*factor_business +
           0.10*factor_profit +
           0.25*factor_levcov +
           0.20*factor_policy)

    rating = score_to_rating(agg)

    return {
        "name": name,
        "scorecard_aggregate": round(agg,3),
        "scorecard_rating": rating,
        "factor_scale": round(factor_scale,3),
        "factor_business_profile": round(factor_business,3),
        "factor_profitability_efficiency": round(factor_profit,3),
        "factor_leverage_coverage": round(factor_levcov,3),
        "factor_financial_policy": round(factor_policy,3),
        "sf_revenue_scale": round(s_scale,3),
        "sf_business_profile": round(s_bp,3),
        "sf_ebit_margin": round(s_ebit,3),
        "sf_debt_ebitda": round(s_debt,3),
        "sf_ebitda_capex_int": round(s_ecint,3),
        "sf_rcf_debt": round(s_rcf,3),
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

        # 0) Conversion en USD de TOUTES les données monétaires brutes selon le pays
        r = convert_row_monetary_to_usd(r)

        # 1) Recalcule tous les ratios Moody’s possibles à partir des bruts (en USD)
        for suf in ("y1","y2","y3"):
            upd = derive_ratios_from_raw(r, suf)
            for k,v in upd.items():
                if v is not None:
                    r[k] = v

        # 2) Scale : revenue y3 (en USD)
        rev_y3 = r.get(RAW["revenue"][2])
        s_scale = score_revenue_scale(rev_y3)

        # 3) Scores 3Y pondérés
        s_ebit = score_3y(
            score_ebit_margin,
            r.get(COL_EBIT_MARGIN_Y1),
            r.get(COL_EBIT_MARGIN_Y2),
            r.get(COL_EBIT_MARGIN_Y3)
        )
        s_debt = score_3y(
            score_debt_ebitda,
            r.get(COL_DEBT_EBITDA_Y1),
            r.get(COL_DEBT_EBITDA_Y2),
            r.get(COL_DEBT_EBITDA_Y3)
        )
        s_ecint = score_3y(
            score_ebitda_capex_int,
            r.get(COL_EBITDA_CAPEX_INT_Y1),
            r.get(COL_EBITDA_CAPEX_INT_Y2),
            r.get(COL_EBITDA_CAPEX_INT_Y3)
        )
        s_rcf = score_3y(
            score_rcf_debt,
            r.get(COL_RCF_DEBT_Y1),
            r.get(COL_RCF_DEBT_Y2),
            r.get(COL_RCF_DEBT_Y3)
        )

        # 4) Facteurs qualitatifs
        s_bp  = score_quali(r.get(COL_BUSINESS_PROFILE))
        s_pol = score_quali(r.get(COL_FINANCIAL_POLICY))

        base = moodys_mining_score_from_scores(
            name=name,
            s_scale=s_scale,
            s_bp=s_bp,
            s_ebit=s_ebit,
            s_debt=s_debt,
            s_ecint=s_ecint,
            s_rcf=s_rcf,
            s_pol=s_pol
        )

        # 5) Liquidité 3Y pour Other Considerations
        liq_vals, w = [], []
        for v, wt in [(r.get(COL_LIQ_Y1),0.2),
                      (r.get(COL_LIQ_Y2),0.3),
                      (r.get(COL_LIQ_Y3),0.5)]:
            fv = to_float(v)
            if fv is not None:
                liq_vals.append(fv); w.append(wt)
        liq3 = None
        if liq_vals:
            sw = sum(w)
            w = [wt/sw for wt in w]
            liq3 = sum(v*wt for v,wt in zip(liq_vals, w))

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

    print("\n✅ Mining – calcul terminé (3Y + soft adjustments ±1). Résumé :\n")
    print(out[[ "name",
                "scorecard_aggregate","scorecard_rating",
                "delta_other_considerations_soft",
                "final_adjusted_score","final_assigned_rating"]].to_string(index=False))
    print(f"\n➡️ Détail sauvegardé dans {OUTPUT_CSV}")