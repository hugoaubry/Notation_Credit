from typing import Optional, Dict, Tuple
import pandas as pd

INPUT_XLSX  = "data/intput/telecom_input_template.xlsx"
OUTPUT_CSV  = "data/output/telecom_output_scorecard.csv"
INPUT_SHEET = "telecom_input_template"

# =========================
#   COLONNES ATTENDUES
# =========================

COL_NAME             = "name"
COL_COUNTRY          = "country"   # pays pour la devise

# Qualitatifs principaux
COL_BUSINESS_PROFILE = "business_profile"   # Aaa, Aa, A, Baa, ...
COL_FINANCIAL_POLICY = "financial_policy"   # Aaa, Aa, A, Baa, ...

# Ratios (fallback si bruts absents) – même schéma que les autres scripts
COL_EBIT_MARGIN_Y1   = "ebit_margin_y1"
COL_EBIT_MARGIN_Y2   = "ebit_margin_y2"
COL_EBIT_MARGIN_Y3   = "ebit_margin_y3"

# Ici, on stocke le ratio **Debt / EBITDA**
COL_DEBT_EBITDA_Y1   = "debt_ebitda_y1"
COL_DEBT_EBITDA_Y2   = "debt_ebitda_y2"
COL_DEBT_EBITDA_Y3   = "debt_ebitda_y3"

# Ici, on stocke le ratio **EBITDA / Interest Expense**
COL_EBITDA_INT_Y1    = "ebitda_interest_y1"
COL_EBITDA_INT_Y2    = "ebitda_interest_y2"
COL_EBITDA_INT_Y3    = "ebitda_interest_y3"

# FCF / Debt
COL_FCF_DEBT_Y1      = "fcf_debt_y1"
COL_FCF_DEBT_Y2      = "fcf_debt_y2"
COL_FCF_DEBT_Y3      = "fcf_debt_y3"

# Liquidity ratio (Sources / Uses) : optionnel mais utile pour "other considerations"
COL_LIQ_Y1           = "liq_y1"
COL_LIQ_Y2           = "liq_y2"
COL_LIQ_Y3           = "liq_y3"

# Other considerations inputs (mêmes conventions que les autres secteurs)
COL_ESG_SCORE        = "esg_score"          # 1..5 (1 = meilleur)
COL_CAPTIVE_RATIO    = "captive_ratio"      # 0..1
COL_REGULATION_SCORE = "regulation_score"   # 1..5 (1 = meilleur)
COL_MANAGEMENT_SCORE = "management_score"   # 1..5 (1 = meilleur)
COL_NONWHOLLY_SALES  = "nonwholly_sales"    # 0..1
COL_EVENT_RISK_SCORE = "event_risk_score"   # 0/1/2 (2 = risque élevé)
COL_PARENTAL_SUPPORT = "parental_support"   # -3..+3 (on cape à ±1 cran)

# Cash de confort (non utilisé dans le score mais conservé)
COL_CASH_Y1          = "cash_y1"
COL_CASH_Y2          = "cash_y2"
COL_CASH_Y3          = "cash_y3"

# =========================
#   COLONNES BRUTES (recalcule Moody’s)
# =========================
# On garde le même schéma que pour l’automobile.
# Les colonnes sont attendues dans le XLSX, si présentes on recalculera les ratios.

RAW = dict(
    revenue=["revenue_y1","revenue_y2","revenue_y3"],          # chiffre d’affaires (en devise locale, converti en USD)
    ebit=["ebit_y1","ebit_y2","ebit_y3"],
    ebitda=["ebitda_y1","ebitda_y2","ebitda_y3"],
    interest_exp=["interest_exp_y1","interest_exp_y2","interest_exp_y3"],

    ocf=["ocf_y1","ocf_y2","ocf_y3"],
    capex=["capex_y1","capex_y2","capex_y3"],
    dividends=["dividends_y1","dividends_y2","dividends_y3"],
    delta_wcap=["delta_wcap_y1","delta_wcap_y2","delta_wcap_y3"],

    st_debt=["st_debt_y1","st_debt_y2","st_debt_y3"],
    cash_sti=["cash_sti_y1","cash_sti_y2","cash_sti_y3"],
    total_debt=["total_debt_y1","total_debt_y2","total_debt_y3"],

    lease_liab_current=["lease_liab_current_y1","lease_liab_current_y2","lease_liab_current_y3"],
    lease_liab_noncurrent=["lease_liab_noncurrent_y1","lease_liab_noncurrent_y2","lease_liab_noncurrent_y3"],
    lease_payments=["lease_payments_y1","lease_payments_y2","lease_payments_y3"],
)

# =========================
#   PAYS & DEVISES (UNIQUEMENT CEUX DU XLSX TÉLÉCOM)
# =========================
# Pays présents dans telecom_input_template.xlsx :
#   - USA      → USD
#   - Germany  → EUR
#   - Japan    → JPY
#   - China    → CNY
#
# fx_to_usd = montant_en_devise_locale × fx_to_usd = montant converti en USD

COUNTRY_CCY_FX: Dict[str, Tuple[str, float]] = {
    "USA":     ("USD", 1.00),
    "Germany": ("EUR", 1.10),   # ~1 EUR = 1.10 USD (approx)
    "Japan":   ("JPY", 0.007),  # ~1 JPY = 0.007 USD (approx)
    "China":   ("CNY", 0.14),   # ~1 CNY = 0.14 USD (approx)
}

def get_fx_for_country(country: Optional[str]) -> Tuple[str, float]:
    if country is None or (isinstance(country, float) and pd.isna(country)):
        return ("USD", 1.0)
    c = str(country).strip()
    info = COUNTRY_CCY_FX.get(c)
    if info is None:
        # fallback neutre si jamais un pays hors liste apparaît
        return ("USD", 1.0)
    return info

# =========================
#   QUALI -> NUM
# =========================

QUALI_NUM = {
    "Aaa":1.0, "Aa":3.0, "A":6.0, "Baa":9.0,
    "Ba":12.0,"B":15.0,"Caa":18.0,"Ca":20.0
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
    """Convertit un label (Aaa, Aa1, Baa2, …) en score numérique de 1 à 20."""
    a = _alpha_from_label(label)
    return QUALI_NUM[a] if a else 12.0  # défaut Ba si inconnu

# =========================
#   PARSING NUMÉRIQUE
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

def safe_abs(x):
    v = to_float(x)
    if v is None:
        return None
    return abs(v)

# =========================
#   ÉCHELLE NUMÉRIQUE & BORNES
# =========================

NUM_RANGES = {
    "Aaa": (0.5, 1.5),
    "Aa":  (1.5, 4.5),
    "A":   (4.5, 7.5),
    "Baa": (7.5, 10.5),
    "Ba":  (10.5, 13.5),
    "B":   (13.5, 16.5),
    "Caa": (16.5, 19.5),
    "Ca":  (19.5, 20.5),
}

#   MAPPINGS & BORNES – DIVERSIFIED TECHNOLOGY (réutilisés pour Telecom)

REVENUE_BOUNDS = [
    ("Aaa", 60e9, 1e18), 
    ("Aa",  30e9, 60e9),
    ("A",   15e9, 30e9),
    ("Baa",  5e9, 15e9),
    ("Ba",   2e9,  5e9),
    ("B",    1e9,  2e9),
    ("Caa", 0.25e9, 1e9),
    ("Ca",  -1e9, 0.25e9),
]
EBIT_MARGIN_BOUNDS = [
    ("Aaa", 40.0, 100.0),
    ("Aa",  30.0, 40.0),
    ("A",   20.0, 30.0),
    ("Baa", 15.0, 20.0),
    ("Ba",  10.0, 15.0),
    ("B",    5.0, 10.0),
    ("Caa",  2.5, 5.0),
    ("Ca",  -100.0, 2.5),
]
EBITDA_INT_BOUNDS = [
    ("Aaa", 30.0, 100.0),
    ("Aa",  20.0, 30.0),
    ("A",   10.0, 20.0),
    ("Baa", 7.0,  10.0),
    ("Ba",  4.0,  7.0),
    ("B",   2.0,  4.0),
    ("Caa", 1.0,  2.0),
    ("Ca",  -100.0, 1.0),
]
FCF_DEBT_BOUNDS = [
    ("Aaa", 45.0, 100.0),
    ("Aa",  35.0, 45.0),
    ("A",   25.0, 35.0),
    ("Baa", 20.0, 25.0),
    ("Ba",  10.0, 20.0),
    ("B",    5.0, 10.0),
    ("Caa",  0.0, 5.0),
    ("Ca",  -20.0, 0.0),
]
DEBT_EBITDA_BOUNDS = [
    ("Aaa", 0.0,  0.5),
    ("Aa",  0.5,  1.0),
    ("A",   1.0,  1.5),
    ("Baa", 1.5,  2.5),
    ("Ba",  2.5,  4.0),
    ("B",   4.0,  6.0),
    ("Caa", 6.0,  8.5),
    ("Ca",  8.5, 15.0),
]

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
    # hors bornes : on projette
    if higher_is_better:
        return 0.5 if x >= bounds[0][2] else 20.5
    return 0.5 if x <= bounds[0][1] else 20.5

def score_revenue_billion(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, REVENUE_BOUNDS, True)

def score_ebit_margin(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBIT_MARGIN_BOUNDS, True)

def score_ebitda_interest(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBITDA_INT_BOUNDS, True)

def score_fcf_debt(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, FCF_DEBT_BOUNDS, True)

def score_debt_ebitda(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, DEBT_EBITDA_BOUNDS, False)

# =========================
#   EXHIBIT 5 : SCORE -> RATING
# =========================
EX5_BINS = [
    (1.5, "Aaa"),
    (2.5, "Aa1"), (3.5, "Aa2"), (4.5, "Aa3"),
    (5.5, "A1"),  (6.5, "A2"),  (7.5, "A3"),
    (8.5, "Baa1"),(9.5, "Baa2"),(10.5, "Baa3"),
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
#   OUTILS Y1-Y3
# =========================

def weighted_average_scores_3y(s1=None, s2=None, s3=None):
    """Moyenne pondérée des SCORES : 0.2*y1 + 0.3*y2 + 0.5*y3."""
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

def scored_ebitda_interest_3y(y1=None, y2=None, y3=None):
    s1 = score_ebitda_interest(y1)
    s2 = score_ebitda_interest(y2)
    s3 = score_ebitda_interest(y3)
    return weighted_average_scores_3y(s1, s2, s3)

def scored_fcf_debt_3y(y1=None, y2=None, y3=None):
    s1 = score_fcf_debt(y1)
    s2 = score_fcf_debt(y2)
    s3 = score_fcf_debt(y3)
    return weighted_average_scores_3y(s1, s2, s3)

# =========================
#   OTHER CONSIDERATIONS (soft ±1 cran)
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
    Même logique que dans les autres secteurs.
    +delta = amélioration (score numérique diminue), -delta = pénalité.
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

    # Captive / JV
    cr = to_float(captive_ratio)
    if cr is not None:
        if cr > 0.40: delta -= 0.5
        elif cr > 0.20: delta -= 0.25

    # Régulation
    rs = to_float(regulation_score)
    if rs is not None:
        if rs <= 2: delta += 0.25
        elif rs >= 4: delta -= 0.25

    # Management (évite double-compte si FP déjà ultra conservatrice)
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

    # Parental / Gov support
    ps = to_float(parental_support)
    if ps is not None:
        ps = int(ps)
        if ps > 0:   delta += min(ps, 1) * 0.5
        elif ps < 0: delta += max(ps, -1) * 0.5

    # Cap global ±1 cran
    delta = max(-1.0, min(1.0, delta))
    return round(delta, 3)

def apply_adjustment(scorecard_aggregate: float, delta_crans: float) -> Tuple[float,str]:
    adjusted = scorecard_aggregate - delta_crans  # +delta = amélioration → score diminue
    return round(adjusted,3), score_to_rating(adjusted)

# =========================
#   RE-CALCUL DES RATIOS À PARTIR DES BRUTS
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

def convert_row_monetary_to_usd(row: pd.Series) -> pd.Series:
    """
    Convertit tous les montants monétaires de la ligne en USD
    en fonction du pays (USA, Germany, Japan, China uniquement).
    """
    country = row.get(COL_COUNTRY)
    ccy, fx = get_fx_for_country(country)

    # trace de la devise d'origine et du fx utilisé
    row["_currency_original"] = ccy
    row["_fx_to_usd"] = fx

    if fx == 1.0:
        return row  # USA -> déjà en USD

    # conversion de toutes les colonnes brutes monétaires
    for key_list in RAW.values():
        for col in key_list:
            if col in row:
                v = to_float(row[col])
                if v is not None:
                    row[col] = v * fx

    # conversion du cash loggé
    for col in (COL_CASH_Y1, COL_CASH_Y2, COL_CASH_Y3):
        if col in row:
            v = to_float(row[col])
            if v is not None:
                row[col] = v * fx

    return row

def derive_ratios_from_raw(row, suffix: str) -> dict:
    """
    Recalcule les ratios Moody’s pour une année suffix (y1/y2/y3) :
    - marge EBIT
    - Debt / EBITDA
    - EBITDA / Interest Expense
    - FCF / Debt
    - ratio de liquidité
    """
    R = lambda col: row.get(f"{col}_{suffix}")
    revenue    = R("revenue")
    ebit       = R("ebit")
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

    if revenue is not None and ebit is not None and ebitda is not None and tot_debt is not None:
        # Adjusted Debt (en USD après conversion)
        adj_debt   = compute_adjusted_debt(tot_debt, lease_cur, lease_non)

        rev_val    = max(to_float(revenue) or 1e-6, 1e-6)
        ebit_val   = to_float(ebit) or 0.0
        ebitda_val = safe_ebitda(ebitda, revenue)

        # EBIT margin (%)
        out[f"ebit_margin_{suffix}"] = (ebit_val / rev_val) * 100.0

        # Debt / EBITDA
        out[f"debt_ebitda_{suffix}"] = adj_debt / ebitda_val

        # EBITDA / Interest Expense (capé à 40x par prudence)
        cov = ebitda_val / safe_interest(interest, revenue)
        out[f"ebitda_interest_{suffix}"] = min(cov, 40.0)

        # RCF & FCF
        rcf, fcf = compute_rcf_fcf(ocf, capex_pos, dividends, delta_wc)
        out[f"fcf_debt_{suffix}"] = (fcf / max(adj_debt, 1e-6)) * 100.0

        # Liquidity ratio
        if st_debt is not None and cash_sti is not None and ocf is not None and capex_pos is not None:
            out[f"liq_{suffix}"] = liquidity_ratio_moodys(cash_sti, ocf, st_debt, capex_pos, dividends, lease_pay)

    return out

# =========================
#   SCORECARD – CALCUL AGRÉGÉ
# =========================

def moodys_telecom_score_from_scores(
    *,
    name: str,
    s_scale: float,
    s_business: float,
    s_profit: float,
    s_cov: float,        # EBITDA / Interest
    s_fcf_debt: float,
    s_debt_ebitda: float,
    s_policy: float
) -> Dict[str, float]:

    factor_scale      = s_scale
    factor_business   = s_business
    factor_profit     = s_profit
    factor_lev_cov    = (s_cov + s_fcf_debt + s_debt_ebitda) / 3.0
    factor_fin_policy = s_policy

    agg = (
        0.20 * factor_scale +
        0.20 * factor_business +
        0.10 * factor_profit +
        0.30 * factor_lev_cov +
        0.20 * factor_fin_policy
    )

    rating = score_to_rating(agg)

    return {
        "name": name,
        "scorecard_aggregate": round(agg,3),
        "scorecard_rating": rating,
        "factor_scale": round(factor_scale,3),
        "factor_business_profile": round(factor_business,3),
        "factor_profitability_efficiency": round(factor_profit,3),
        "factor_leverage_coverage": round(factor_lev_cov,3),
        "factor_financial_policy": round(factor_fin_policy,3),
        # sous-facteurs
        "sf_revenue": round(s_scale,3),
        "sf_business_profile": round(s_business,3),
        "sf_ebit_margin": round(s_profit,3),
        "sf_ebitda_interest": round(s_cov,3),
        "sf_fcf_debt": round(s_fcf_debt,3),
        "sf_debt_ebitda": round(s_debt_ebitda,3),
        "sf_financial_policy": round(s_policy,3),
    }

# =========================
#   PIPELINE
# =========================

if __name__ == "__main__":
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET)

    results = []
    for _, r in df.iterrows():
        name = r[COL_NAME]

        # 0) Conversion devise → USD en fonction du pays (USA, Germany, Japan, China UNIQUEMENT)
        r = convert_row_monetary_to_usd(r)

        # 1) Recalcule Moody’s depuis les BRUTS si dispo (et écrase les ratios d’entrée)
        for suf in ("y1","y2","y3"):
            upd = derive_ratios_from_raw(r, suf)
            for k,v in upd.items():
                if v is not None:
                    r[k] = v

        # 2) Scale : on prend la moyenne 3 ans du CA (en USD)
        rev_cols = RAW["revenue"]
        rev_vals = [to_float(r.get(c)) for c in rev_cols]
        rev_vals = [v for v in rev_vals if v is not None]
        if rev_vals:
            rev_avg = sum(rev_vals) / len(rev_vals)
        else:
            rev_avg = None
        s_scale = score_revenue_billion(rev_avg)

        # 3) Profitabilité : marge EBIT 3 ans
        s_profit = scored_ebit_margin_3y(
            r.get(COL_EBIT_MARGIN_Y1),
            r.get(COL_EBIT_MARGIN_Y2),
            r.get(COL_EBIT_MARGIN_Y3)
        )

        # 4) Leverage & coverage : trois sous-facteurs 3 ans
        s_debt_ebitda = scored_debt_ebitda_3y(
            r.get(COL_DEBT_EBITDA_Y1),
            r.get(COL_DEBT_EBITDA_Y2),
            r.get(COL_DEBT_EBITDA_Y3)
        )
        s_cov = scored_ebitda_interest_3y(
            r.get(COL_EBITDA_INT_Y1),
            r.get(COL_EBITDA_INT_Y2),
            r.get(COL_EBITDA_INT_Y3)
        )
        s_fcf_debt = scored_fcf_debt_3y(
            r.get(COL_FCF_DEBT_Y1),
            r.get(COL_FCF_DEBT_Y2),
            r.get(COL_FCF_DEBT_Y3)
        )

        # 5) Qualitatifs → scores numériques
        s_business = score_quali(r.get(COL_BUSINESS_PROFILE))
        s_policy   = score_quali(r.get(COL_FINANCIAL_POLICY))

        # 6) Scorecard "de base"
        base = moodys_telecom_score_from_scores(
            name=name,
            s_scale=s_scale,
            s_business=s_business,
            s_profit=s_profit,
            s_cov=s_cov,
            s_fcf_debt=s_fcf_debt,
            s_debt_ebitda=s_debt_ebitda,
            s_policy=s_policy
        )

        # 7) Liquidity 3Y moyenne pour les "other considerations"
        liq3 = weighted_average_scores_3y(
            to_float(r.get(COL_LIQ_Y1)),
            to_float(r.get(COL_LIQ_Y2)),
            to_float(r.get(COL_LIQ_Y3))
        )

        # 8) Ajustements hors scorecard (±1 cran)
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

        adj_score, final_rating = apply_adjustment(
            base["scorecard_aggregate"], delta_soft
        )

        results.append({
            **base,
            "inputs_liquidity_ratio_3y": liq3,
            "delta_other_considerations_soft": delta_soft,
            "final_adjusted_score": adj_score,
            "final_assigned_rating": final_rating
        })

    out = pd.DataFrame(results)
    out.to_csv(OUTPUT_CSV, index=False)

    print("\n✅ Scorecard Télécom / Diversified Technology calculée.\n")
    print(out[[
        "name",
        "scorecard_aggregate","scorecard_rating",
        "delta_other_considerations_soft",
        "final_adjusted_score","final_assigned_rating"
    ]].to_string(index=False))
    print(f"\n➡️ Détail sauvegardé dans {OUTPUT_CSV}")