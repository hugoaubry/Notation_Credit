### Diversified Technology – Scorecard Moody’s-like ###
from typing import Optional, Dict, Tuple
import pandas as pd

# =========================
#   CONSTANTES I/O
# =========================

INPUT_XLSX  = "data/intput/tech_input_template.xlsx"
OUTPUT_CSV  = "data/output/tech_output_scorecard.csv"
INPUT_SHEET = "tech_input_template"

# =========================
#   COLONNES ATTENDUES
# =========================
# y1 = plus ancien ; y3 = plus récent

COL_NAME                = "name"
COL_COUNTRY             = "country"  # colonne 'country' dans ton XLSX

# Business Profile (qualitatif)
COL_BUSINESS_PROFILE    = "business_profile"        # Business Profile (Aaa..Ca)

# Profitability & Efficiency (quanti)
# -> EBIT Margin (en %)
COL_EBIT_MARGIN_Y1      = "ebit_margin_y1"
COL_EBIT_MARGIN_Y2      = "ebit_margin_y2"
COL_EBIT_MARGIN_Y3      = "ebit_margin_y3"

# Leverage & Coverage (quanti)
# Debt / EBITDA (x)
COL_DEBT_EBITDA_Y1      = "debt_ebitda_y1"
COL_DEBT_EBITDA_Y2      = "debt_ebitda_y2"
COL_DEBT_EBITDA_Y3      = "debt_ebitda_y3"

# FCF / Debt (%)
COL_FCF_DEBT_Y1         = "fcf_debt_y1"
COL_FCF_DEBT_Y2         = "fcf_debt_y2"
COL_FCF_DEBT_Y3         = "fcf_debt_y3"

# EBITDA / Interest Expense (x)
COL_EBITDA_INT_Y1       = "ebitda_int_y1"
COL_EBITDA_INT_Y2       = "ebitda_int_y2"
COL_EBITDA_INT_Y3       = "ebitda_int_y3"

# Liquidity ratio (Sources / Uses) – pour Other Considerations
COL_LIQ_Y1              = "liq_y1"
COL_LIQ_Y2              = "liq_y2"
COL_LIQ_Y3              = "liq_y3"

# Financial Policy (qualitatif)
COL_FINANCIAL_POLICY    = "financial_policy"

# Other considerations inputs (gabarit générique)
COL_ESG_SCORE           = "esg_score"          # 1..5 (1 = meilleur)
COL_CAPTIVE_RATIO       = "captive_ratio"      # 0..1
COL_REGULATION_SCORE    = "regulation_score"   # 1..5 (1 = meilleur)
COL_MANAGEMENT_SCORE    = "management_score"   # 1..5 (1 = meilleur)
COL_NONWHOLLY_SALES     = "nonwholly_sales"    # 0..1
COL_EVENT_RISK_SCORE    = "event_risk_score"   # 0/1/2
COL_PARENTAL_SUPPORT    = "parental_support"   # -3..+3 (on cape à ±1 cran d'effet)

# (facultatif) cash loggé, pas utilisé direct dans la scorecard
COL_CASH_Y1             = "cash_y1"
COL_CASH_Y2             = "cash_y2"
COL_CASH_Y3             = "cash_y3"

# =========================
#   MAPPING PAYS / DEVISE / FX
# =========================
# Seulement les pays réellement présents dans tech_input_template.xlsx

COUNTRY_TO_CCY = {
    "USA": "USD",
    "Netherlands": "EUR",
    "South Korea": "KRW",
    "Taiwan": "TWD",
}

# Taux de change (1 unité de devise locale -> USD)
# À ajuster si tu veux coller à une date précise.
FX_TO_USD = {
    "USD": 1.0,       # Apple, Microsoft, Alphabet...
    "EUR": 1.10,      # Netherlands -> EUR
    "KRW": 0.00075,   # South Korea -> KRW
    "TWD": 0.031,     # Taiwan -> TWD
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
# On recalcule les ratios Moody's à partir des bruts.
# Les valeurs sont saisies dans la devise locale, on les convertit en USD via FX.

RAW = dict(
    revenue = ["revenue_y1","revenue_y2","revenue_y3"],
    ebit    = ["ebit_y1","ebit_y2","ebit_y3"],
    ebitda  = ["ebitda_y1","ebitda_y2","ebitda_y3"],
    interest_exp = ["interest_exp_y1","interest_exp_y2","interest_exp_y3"],
    ocf     = ["ocf_y1","ocf_y2","ocf_y3"],
    capex   = ["capex_y1","capex_y2","capex_y3"],
    dividends = ["dividends_y1","dividends_y2","dividends_y3"],
    delta_wcap = ["delta_wcap_y1","delta_wcap_y2","delta_wcap_y3"],
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
#   NUMERIC RANGES (Diversified Technology – Exhibit 2)
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
    if t < 0:
        t = 0.0
    if t > 1:
        t = 1.0
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
    # Hors bornes -> au mieux ou au pire
    if higher_is_better:
        return 0.5 if x >= bounds[0][2] else 20.5
    return 0.5 if x <= bounds[0][1] else 20.5

# ====== Bornes spécifiques Diversified Technology (Exhibit 2) ======
# (après conversion en USD)

REVENUE_BOUNDS = [
    ("Aaa", 60e9, 1e18),   # ≥ 60 billion USD
    ("Aa",  30e9, 60e9),
    ("A",   15e9, 30e9),
    ("Baa", 5e9,  15e9),
    ("Ba",  2e9,  5e9),
    ("B",   1e9,  2e9),
    ("Caa", 0.25e9, 1e9),
    ("Ca",  -1e9, 0.25e9),
]

EBIT_MARGIN_BOUNDS = [
    ("Aaa", 27.0, 100.0),
    ("Aa",  24.0, 27.0),
    ("A",   21.0, 24.0),
    ("Baa", 18.0, 21.0),
    ("Ba",  15.0, 18.0),
    ("B",   12.0, 15.0),
    ("Caa", 5.0, 12.0),
    ("Ca",  0.0, 5.0),
]

EBITDA_INT_BOUNDS = [
    ("Aaa", 16.0, 100.0),
    ("Aa",  12.0, 16.0),
    ("A",   8.0, 12.0),
    ("Baa", 4.0, 8.0),
    ("Ba",  2.0, 4.0),
    ("B",   1.0, 2.0),
    ("Caa", 0.0, 1.0),
    ("Ca",  -100.0, 0.0),
]

FCF_DEBT_BOUNDS = [
    ("Aaa", 35.0, 100.0),
    ("Aa",  30.0, 35.0),
    ("A",   25.0, 30.0),
    ("Baa", 20.0, 25.0),
    ("Ba",  10.0, 20.0),
    ("B",   5.0, 10.0),
    ("Caa", 0.0, 5.0),
    ("Ca",  -100.0, 0.0),
]

DEBT_EBITDA_BOUNDS = [
    ("Aaa", 0.0, 0.5),
    ("Aa",  0.5, 1.0),
    ("A",   1.0, 1.5),
    ("Baa", 1.5, 2.5),
    ("Ba",  2.5, 4.0),
    ("B",   4.0, 6.0),
    ("Caa", 6.0, 8.0),
    ("Ca",  8.0, 100.0),
]

# =========================
#   SCORE QUANTI
# =========================

def score_revenue_scale(rev_usd: Optional[float]) -> float:
    v = to_float(rev_usd)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, REVENUE_BOUNDS, True)

def score_ebit_margin(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBIT_MARGIN_BOUNDS, True)

def score_debt_ebitda(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, DEBT_EBITDA_BOUNDS, False)

def score_fcf_debt(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, FCF_DEBT_BOUNDS, True)

def score_ebitda_int(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBITDA_INT_BOUNDS, True)

# =========================
#   PONDÉRATIONS – DIVERSIFIED TECHNOLOGY
# =========================

W = {
    "revenue_scale":     20.0,
    "business_profile":  20.0,
    "ebit_margin":       10.0,
    "ebitda_int":        10.0,
    "fcf_debt":          10.0,
    "debt_ebitda":       10.0,
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
#   OUTILS 3 ANS
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

def scored_ebit_margin_3y(y1=None, y2=None, y3=None):
    s1 = score_ebit_margin(y1)
    s2 = score_ebit_margin(y2)
    s3 = score_ebit_margin(y3)
    return weighted_average_scores_3y(s1, s2, s3)

def scored_debt_ebitda_3y(y1=None, y2=None, y3=None):
    s1 = score_debt_ebitda(y1)
    s2 = score_debt_ebitda(y2)
    s3 = score_debt_ebitda(y3)
    return weighted_average_scores_3y(s1, s2, s3)

def scored_fcf_debt_3y(y1=None, y2=None, y3=None):
    s1 = score_fcf_debt(y1)
    s2 = score_fcf_debt(y2)
    s3 = score_fcf_debt(y3)
    return weighted_average_scores_3y(s1, s2, s3)

def scored_ebitda_int_3y(y1=None, y2=None, y3=None):
    s1 = score_ebitda_int(y1)
    s2 = score_ebitda_int(y2)
    s3 = score_ebitda_int(y3)
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
#   RE-CALCUL MOODY’S DES RATIOS (DEPUIS BRUTS, EN USD)
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
        rcf = ocf_v - dwc - div_v
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
    Recalcule les ratios Moody's Diversified Technology pour une année suffix (y1/y2/y3),
    en partant des BRUTS convertis en USD selon le pays de l'entreprise.
    Remplit :
      - ebit_margin_suffix (%)
      - debt_ebitda_suffix
      - fcf_debt_suffix (%)
      - ebitda_int_suffix (x)
      - liq_suffix
    """
    country = row.get(COL_COUNTRY) if COL_COUNTRY in row else None
    fx = get_fx_from_country(country)

    R = lambda col: row.get(f"{col}_{suffix}")

    def conv(v):
        tv = to_float(v)
        return tv * fx if tv is not None else None

    revenue    = conv(R("revenue"))
    ebit       = conv(R("ebit"))
    ebitda     = conv(R("ebitda"))
    interest   = conv(R("interest_exp"))
    ocf        = conv(R("ocf"))
    capex_pos  = conv(R("capex"))
    dividends  = conv(R("dividends"))
    delta_wc   = conv(R("delta_wcap"))
    st_debt    = conv(R("st_debt"))
    cash_sti   = conv(R("cash_sti"))
    tot_debt   = conv(R("total_debt"))
    lease_cur  = conv(R("lease_liab_current"))
    lease_non  = conv(R("lease_liab_noncurrent"))
    lease_pay  = conv(R("lease_payments"))

    out = {}

    # EBIT Margin (%)
    if revenue is not None and revenue != 0 and ebit is not None:
        margin = (ebit / revenue) * 100.0
        out[f"ebit_margin_{suffix}"] = margin

    adj_debt = None
    if tot_debt is not None:
        adj_debt = compute_adjusted_debt(tot_debt, lease_cur, lease_non)

    # Debt/EBITDA
    if adj_debt is not None and ebitda is not None:
        ebitda_safe = safe_ebitda(ebitda, revenue)
        out[f"debt_ebitda_{suffix}"] = adj_debt / ebitda_safe

    # FCF / Debt (%)
    if adj_debt is not None and ocf is not None:
        _, fcf = compute_rcf_fcf(ocf, capex_pos, dividends, delta_wc)
        if adj_debt > 0:
            out[f"fcf_debt_{suffix}"] = (fcf / max(adj_debt,1e-6)) * 100.0
        else:
            out[f"fcf_debt_{suffix}"] = 100.0 if fcf > 0 else 0.0

    # EBITDA / Interest Expense (x)
    if ebitda is not None and interest is not None:
        int_v = safe_interest(interest, revenue)
        ratio = ebitda / max(int_v, 1e-9)
        out[f"ebitda_int_{suffix}"] = ratio

    # Liquidity ratio S/U
    if st_debt is not None and cash_sti is not None and ocf is not None and capex_pos is not None:
        liq = liquidity_ratio_moodys(cash_sti, ocf, st_debt, capex_pos, dividends, lease_pay)
        out[f"liq_{suffix}"] = liq

    return out

# =========================
#   SCORECARD – DIVERSIFIED TECHNOLOGY
# =========================

def moodys_tech_score_from_scores(
    *,
    name,
    s_scale,
    s_business_profile,
    s_ebit_margin,
    s_debt,
    s_fcf,
    s_ebitda_int,
    s_pol
) -> Dict[str, float]:

    score_scale    = s_scale
    score_business = s_business_profile
    score_profit   = s_ebit_margin
    score_levcov   = (W["ebitda_int"]*s_ebitda_int +
                      W["fcf_debt"]*s_fcf +
                      W["debt_ebitda"]*s_debt) / 30.0
    score_policy   = s_pol

    agg = (0.20*score_scale +
           0.20*score_business +
           0.10*score_profit +
           0.30*score_levcov +
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
        "sf_revenue_scale": round(s_scale,3),
        "sf_business_profile": round(s_business_profile,3),
        "sf_ebit_margin": round(s_ebit_margin,3),
        "sf_debt_ebitda": round(s_debt,3),
        "sf_fcf_debt": round(s_fcf,3),
        "sf_ebitda_int": round(s_ebitda_int,3),
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
        fx      = get_fx_from_country(country)

        # --- Recalcule Moody's depuis BRUTS (données d'abord converties en USD dans derive_ratios_from_raw)
        for suf in ("y1","y2","y3"):
            upd = derive_ratios_from_raw(r, suf)
            for k,v in upd.items():
                if v is not None:
                    r[k] = v

        # --- SCALE : Revenue (USD), on prend l'année la plus récente (y3)
        rev_raw = r.get(RAW["revenue"][2]) if RAW["revenue"][2] in r else None
        rev_usd = None
        tv = to_float(rev_raw)
        if tv is not None:
            rev_usd = tv * fx  # conversion en USD
        s_scale = score_revenue_scale(rev_usd)

        # --- LEVERAGE & COVERAGE / PROFITABILITY : scores 3Y pondérés
        s_ebitm = scored_ebit_margin_3y(
            r.get(COL_EBIT_MARGIN_Y1),
            r.get(COL_EBIT_MARGIN_Y2),
            r.get(COL_EBIT_MARGIN_Y3)
        )
        s_debt  = scored_debt_ebitda_3y(
            r.get(COL_DEBT_EBITDA_Y1),
            r.get(COL_DEBT_EBITDA_Y2),
            r.get(COL_DEBT_EBITDA_Y3)
        )
        s_fcf   = scored_fcf_debt_3y(
            r.get(COL_FCF_DEBT_Y1),
            r.get(COL_FCF_DEBT_Y2),
            r.get(COL_FCF_DEBT_Y3)
        )
        s_eint  = scored_ebitda_int_3y(
            r.get(COL_EBITDA_INT_Y1),
            r.get(COL_EBITDA_INT_Y2),
            r.get(COL_EBITDA_INT_Y3)
        )

        # --- QUALITATIFS
        s_bp   = score_quali(r.get(COL_BUSINESS_PROFILE))
        s_pol  = score_quali(r.get(COL_FINANCIAL_POLICY))

        # --- Scorecard outcome
        base = moodys_tech_score_from_scores(
            name=name,
            s_scale=s_scale,
            s_business_profile=s_bp,
            s_ebit_margin=s_ebitm,
            s_debt=s_debt,
            s_fcf=s_fcf,
            s_ebitda_int=s_eint,
            s_pol=s_pol
        )

        # --- Liquidity 3Y moyenne pour Other Considerations (déjà en USD car calculée à partir de données converties)
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

    print("\n✅ Diversified Technology – calcul terminé (3Y + soft adjustments ±1). Résumé :\n")
    print(out[[ "name",
                "scorecard_aggregate","scorecard_rating",
                "delta_other_considerations_soft",
                "final_adjusted_score","final_assigned_rating"]].to_string(index=False))
    print(f"\n➡️ Détail sauvegardé dans {OUTPUT_CSV}")