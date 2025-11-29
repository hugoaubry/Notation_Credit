### Industrials / Manufacturing – Scorecard Moody’s-like ###
from typing import Optional, Dict, Tuple
import pandas as pd

# =========================
#   CONSTANTES I/O
# =========================

# garde la même arbo que pour les autres secteurs (typo "intput" incluse si tu l'as partout)
INPUT_XLSX  = "data/intput/industrial_input_template.xlsx"
OUTPUT_CSV  = "data/output/industrial_output_scorecard.csv"
INPUT_SHEET = "industrial_input_template"


# =========================
#   COLONNES ATTENDUES
# =========================
# y1 = plus ancien ; y3 = plus récent

COL_NAME                = "name"
COL_COUNTRY             = "country"

# Factor 2 – Business Profile (qualitatif)
COL_BUSINESS_PROFILE    = "business_profile"        
# Factor 3 – Profitability & Efficiency : EBITA Margin (%)
COL_EBITA_MARGIN_Y1     = "ebita_margin_y1"
COL_EBITA_MARGIN_Y2     = "ebita_margin_y2"
COL_EBITA_MARGIN_Y3     = "ebita_margin_y3"

# Factor 4 – Leverage & Coverage
# a) Debt / EBITDA (x)
COL_DEBT_EBITDA_Y1      = "debt_ebitda_y1"
COL_DEBT_EBITDA_Y2      = "debt_ebitda_y2"
COL_DEBT_EBITDA_Y3      = "debt_ebitda_y3"

# b) RCF / Net Debt (%)
COL_RCF_NETDEBT_Y1      = "rcf_netdebt_y1"
COL_RCF_NETDEBT_Y2      = "rcf_netdebt_y2"
COL_RCF_NETDEBT_Y3      = "rcf_netdebt_y3"

# c) FCF / Debt (%)
COL_FCF_DEBT_Y1         = "fcf_debt_y1"
COL_FCF_DEBT_Y2         = "fcf_debt_y2"
COL_FCF_DEBT_Y3         = "fcf_debt_y3"

# d) EBITA / Interest Expense (x)
COL_EBITA_INT_Y1        = "ebita_interest_y1"
COL_EBITA_INT_Y2        = "ebita_interest_y2"
COL_EBITA_INT_Y3        = "ebita_interest_y3"

# Liquidity ratio (Sources / Uses) – pour Other Considerations
COL_LIQ_Y1              = "liq_y1"
COL_LIQ_Y2              = "liq_y2"
COL_LIQ_Y3              = "liq_y3"

# Factor 5 – Financial Policy (qualitatif)
COL_FINANCIAL_POLICY    = "financial_policy"

# Other considerations inputs (gabarit générique)
COL_ESG_SCORE           = "esg_score"          # 1..5 (1 = meilleur)
COL_CAPTIVE_RATIO       = "captive_ratio"      # 0..1 (si relevant, sinon laisser vide)
COL_REGULATION_SCORE    = "regulation_score"   # 1..5 (1 = meilleur)
COL_MANAGEMENT_SCORE    = "management_score"   # 1..5 (1 = meilleur)
COL_NONWHOLLY_SALES     = "nonwholly_sales"    # 0..1
COL_EVENT_RISK_SCORE    = "event_risk_score"   # 0/1/2
COL_PARENTAL_SUPPORT    = "parental_support"   # -3..+3 (on cape à ±1 cran d'effet)

# (facultatif) cash loggé, pas utilisé directement
COL_CASH_Y1             = "cash_y1"
COL_CASH_Y2             = "cash_y2"
COL_CASH_Y3             = "cash_y3"

# =========================
#   COLONNES BRUTES
# =========================
# Même logique que pour les autres secteurs : on recalcule les ratios Moody's à partir des bruts si dispo.
RAW = dict(
    revenue = ["revenue_y1","revenue_y2","revenue_y3"],      # montant nominal (on convertit en USD)
    ebita = ["ebita_y1","ebita_y2","ebita_y3"],              # EBITA (ou EBIT si tu n'as que ça)
    ebitda = ["ebitda_y1","ebitda_y2","ebitda_y3"],
    interest_exp = ["interest_exp_y1","interest_exp_y2","interest_exp_y3"],   # valeur absolue positive
    ocf = ["ocf_y1","ocf_y2","ocf_y3"],                      # operating cash flow (approx FFO)
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
    # enlève séparateurs de milliers
    s = s.replace(",", "")
    neg = False
    # parenthèses = négatif
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    # pourcentages
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
#   FX PAR PAYS (d’après l’XLSX : Germany / USA / Japan / France / Switzerland)
# =========================

FX_BY_COUNTRY = {
    "USA": 1.0,
    "United States": 1.0,
    "Germany": 1.08,       # EUR → USD (approx)
    "France": 1.08,        # EUR → USD (approx)     
    "Switzerland": 1.10,   # CHF → USD (approx)
}

# Colonnes nominales à convertir (mêmes que RAW + éventuel cash_y*)
NOMINAL_COLS = []
for _cols in RAW.values():
    NOMINAL_COLS.extend(_cols)
NOMINAL_COLS.extend([COL_CASH_Y1, COL_CASH_Y2, COL_CASH_Y3])
# enlever doublons en gardant l'ordre
NOMINAL_COLS = list(dict.fromkeys(NOMINAL_COLS))

def _fx_for_country(country_val) -> float:
    if country_val is None or (isinstance(country_val,float) and pd.isna(country_val)):
        return 1.0
    s = str(country_val).strip()
    return FX_BY_COUNTRY.get(s, 1.0)

def convert_row_nominals_to_usd(row: pd.Series) -> pd.Series:
    """
    Convertit toutes les colonnes nominales (revenue, debt, cash, OCF, capex, etc.)
    en USD pour le pays concerné, AVANT tout recalcul de ratios / score.
    """
    fx = _fx_for_country(row.get(COL_COUNTRY))
    if fx == 1.0:
        return row
    for col in NOMINAL_COLS:
        if col in row:
            v = to_float(row[col])
            if v is not None:
                row[col] = v * fx
    return row

# =========================
#   NUMERIC RANGES (Manufacturing / Industrials)
# =========================

# Même mapping score numérique qu’ailleurs
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

# ====== Bornes spécifiques Manufacturing / Industrials (approximées à partir de plusieurs scorecards) ======

# 1) Factor 1 – Scale : Revenue (USD), poids 20%
REVENUE_BOUNDS = [
    ("Aaa", 50e9, 1e18),   # ≥ $50bn
    ("Aa",  30e9, 50e9),
    ("A",   15e9, 30e9),
    ("Baa", 5e9,  15e9),
    ("Ba",  1.5e9, 5e9),
    ("B",   0.5e9, 1.5e9),
    ("Caa", 0.25e9, 0.5e9),
    ("Ca",  -1e9, 0.25e9),   # < $0.25bn
]
EBITA_MARGIN_BOUNDS = [
    ("Aaa", 35.0, 100.0),
    ("Aa",  25.0, 35.0),
    ("A",   17.0, 25.0),
    ("Baa", 12.0, 17.0),
    ("Ba",  7.0, 12.0),
    ("B",   2.5, 7.0),
    ("Caa", 0.0, 2.5),
    ("Ca",  -100.0, 0.0),
]
DEBT_EBITDA_BOUNDS = [
    ("Aaa", 0.0,   0.5),
    ("Aa",  0.5,   1.0),
    ("A",   1.0,   1.75),
    ("Baa", 1.75,  3.25),
    ("Ba",  3.25,  4.75),
    ("B",   4.75,  6.25),
    ("Caa", 6.25,  7.75),
    ("Ca",  7.75,  20.0),
]
RCF_NETDEBT_BOUNDS = [
    ("Aaa", 60.0, 100.0),
    ("Aa",  45.0, 60.0),
    ("A",   35.0, 45.0),
    ("Baa", 25.0, 35.0),
    ("Ba",  15.0, 25.0),
    ("B",   7.5, 15.0),
    ("Caa", 0.0, 7.5),
    ("Ca",  -100.0, 0.0),
]
FCF_DEBT_BOUNDS = [
    ("Aaa", 25.0, 100.0),
    ("Aa",  20.0, 25.0),
    ("A",   15.0, 20.0),
    ("Baa", 10.0, 15.0),
    ("Ba",  5.0, 10.0),
    ("B",   0.0, 5.0),
    ("Caa", -5.0, 0.0),
    ("Ca",  -20.0, -5.0),
]
EBITA_INT_BOUNDS = [
    ("Aaa", 20.0, 100.0),
    ("Aa",  15.0, 20.0),
    ("A",   10.0, 15.0),
    ("Baa", 7.0, 10.0),
    ("Ba",  4.0, 7.0),
    ("B",   1.5, 4.0),
    ("Caa", 0.75, 1.5),
    ("Ca",  0.0, 0.75),
]

# =========================
#   SCORE QUANTI
# =========================

def score_revenue_scale(rev_usd: Optional[float]) -> float:
    v = to_float(rev_usd)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, REVENUE_BOUNDS, True)

def score_ebita_margin(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBITA_MARGIN_BOUNDS, True)

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

def score_fcf_debt(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, FCF_DEBT_BOUNDS, True)

def score_ebita_int(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None:
        return 12.0
    return _score_quant_from_bounds(v, EBITA_INT_BOUNDS, True)

# =========================
#   PONDÉRATIONS MANUFACTURING / INDUSTRIALS
# =========================
# Factor 1 – Scale (20%): Revenue 20
# Factor 2 – Business Profile (25%): Business Profile 25
# Factor 3 – Profitability & Efficiency (5%): EBITA Margin 5
# Factor 4 – Leverage & Coverage (35%): Debt/EBITDA 10, RCF/Net Debt 10, FCF/Debt 5, EBITA/Interest 10
# Factor 5 – Financial Policy (15%): Financial Policy 15

W = {
    "revenue_scale":     20.0,
    "business_profile":  25.0,
    "ebita_margin":      5.0,
    "debt_ebitda":       10.0,
    "rcf_netdebt":       10.0,
    "fcf_debt":          5.0,
    "ebita_int":         10.0,
    "financial_policy":  15.0,
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
    """ Moyenne pondérée de scores : 0.2*y1 + 0.3*y2 + 0.5*y3 (y3 = plus récent). """
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

def scored_ebita_margin_3y(y1=None, y2=None, y3=None):
    s1 = score_ebita_margin(y1)
    s2 = score_ebita_margin(y2)
    s3 = score_ebita_margin(y3)
    return weighted_average_scores_3y(s1, s2, s3)

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

def scored_fcf_debt_3y(y1=None, y2=None, y3=None):
    s1 = score_fcf_debt(y1)
    s2 = score_fcf_debt(y2)
    s3 = score_fcf_debt(y3)
    return weighted_average_scores_3y(s1, s2, s3)

def scored_ebita_int_3y(y1=None, y2=None, y3=None):
    s1 = score_ebita_int(y1)
    s2 = score_ebita_int(y2)
    s3 = score_ebita_int(y3)
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

    # Liquidité (Sources / Uses)
    lr = to_float(liquidity_ratio)
    if lr is not None:
        if lr > 2.0: delta += 0.25
        elif lr < 1.0: delta -= 0.5

    # Captive finance (si pertinent pour l’industriel, sinon laisser vide)
    cr = to_float(captive_ratio)
    if cr is not None:
        if cr > 0.40: delta -= 0.5
        elif cr > 0.20: delta -= 0.25

    # Régulation
    rs = to_float(regulation_score)
    if rs is not None:
        if rs <= 2: delta += 0.25
        elif rs >= 4: delta -= 0.25

    # Management (éviter double-compte si FP déjà très élevée)
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

    # Cap total
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
    Recalcule les ratios Moody's Manufacturing pour une année suffix (y1/y2/y3).
    Remplit :
      - ebita_margin_suffix (%)
      - debt_ebitda_suffix
      - rcf_netdebt_suffix (%)
      - fcf_debt_suffix (%)
      - ebita_interest_suffix (x)
      - liq_suffix
    """
    R = lambda col: row.get(f"{col}_{suffix}")
    revenue    = R("revenue")
    ebita      = R("ebita")
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

    # EBITA Margin (%)
    rev_v   = to_float(revenue)
    ebita_v = to_float(ebita)
    if rev_v is not None and rev_v != 0 and ebita_v is not None:
        margin = (ebita_v / rev_v) * 100.0
        out[f"ebita_margin_{suffix}"] = margin

    adj_debt = None
    if tot_debt is not None:
        adj_debt = compute_adjusted_debt(tot_debt, lease_cur, lease_non)

    # Debt / EBITDA
    if adj_debt is not None and ebitda is not None:
        ebitda_safe = safe_ebitda(ebitda, revenue)
        out[f"debt_ebitda_{suffix}"] = adj_debt / ebitda_safe

    # RCF / Net Debt & FCF / Debt
    rcf, fcf = compute_rcf_fcf(ocf, capex_pos, dividends, delta_wc)
    if adj_debt is not None:
        # Net debt = Debt - cash
        cash_v = to_float(cash_sti) or 0.0
        net_debt = max(adj_debt - cash_v, 0.0)
        if net_debt > 0:
            out[f"rcf_netdebt_{suffix}"] = (rcf / max(net_debt,1e-6)) * 100.0
        else:
            out[f"rcf_netdebt_{suffix}"] = 100.0 if rcf > 0 else 0.0

        if adj_debt > 0:
            out[f"fcf_debt_{suffix}"] = (fcf / max(adj_debt,1e-6)) * 100.0
        else:
            out[f"fcf_debt_{suffix}"] = 100.0 if fcf > 0 else 0.0

    # EBITA / Interest Expense
    if ebita_v is not None and interest is not None:
        int_v = safe_interest(interest, revenue)
        ratio = ebita_v / max(int_v, 1e-9)
        out[f"ebita_interest_{suffix}"] = ratio

    # Liquidity ratio S/U
    if st_debt is not None and cash_sti is not None and ocf is not None and capex_pos is not None:
        liq = liquidity_ratio_moodys(cash_sti, ocf, st_debt, capex_pos, dividends, lease_pay)
        out[f"liq_{suffix}"] = liq

    return out

# =========================
#   SCORECARD – INDUSTRIALS
# =========================

def moodys_industrial_score_from_scores(
    *,
    name,
    s_scale,
    s_business_profile,
    s_ebita_margin,
    s_debt,
    s_rcf_nd,
    s_fcf,
    s_ebita_int,
    s_pol
) -> Dict[str, float]:

    # Facteurs
    score_scale    = s_scale
    score_business = s_business_profile
    score_profit   = s_ebita_margin
    score_levcov   = (
        W["debt_ebitda"]  * s_debt +
        W["rcf_netdebt"]  * s_rcf_nd +
        W["fcf_debt"]     * s_fcf +
        W["ebita_int"]    * s_ebita_int
    ) / 35.0
    score_policy   = s_pol

    agg = (0.20*score_scale +
           0.25*score_business +
           0.05*score_profit +
           0.35*score_levcov +
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
        "sf_business_profile": round(s_business_profile,3),
        "sf_ebita_margin": round(s_ebita_margin,3),
        "sf_debt_ebitda": round(s_debt,3),
        "sf_rcf_netdebt": round(s_rcf_nd,3),
        "sf_fcf_debt": round(s_fcf,3),
        "sf_ebita_interest": round(s_ebita_int,3),
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

        # --- Conversion nominals en USD en fonction du pays (USA / Germany / France / Japan / Switzerland)
        r = convert_row_nominals_to_usd(r)

        # --- Recalcule Moody's depuis BRUTS si dispo (à partir des valeurs DÉJÀ converties en USD)
        for suf in ("y1","y2","y3"):
            upd = derive_ratios_from_raw(r, suf)
            for k,v in upd.items():
                if v is not None:
                    r[k] = v

        # --- SCALE : Revenue (USD), on prend l'année la plus récente (y3)
        rev_raw = r.get(RAW["revenue"][2]) if RAW["revenue"][2] in r else None
        s_scale = score_revenue_scale(rev_raw)

        # --- LEVERAGE & COVERAGE / PROFITABILITY : scores 3Y pondérés
        s_ebita  = scored_ebita_margin_3y(
            r.get(COL_EBITA_MARGIN_Y1),
            r.get(COL_EBITA_MARGIN_Y2),
            r.get(COL_EBITA_MARGIN_Y3)
        )
        s_debt   = scored_debt_ebitda_3y(
            r.get(COL_DEBT_EBITDA_Y1),
            r.get(COL_DEBT_EBITDA_Y2),
            r.get(COL_DEBT_EBITDA_Y3)
        )
        s_rcf_nd = scored_rcf_netdebt_3y(
            r.get(COL_RCF_NETDEBT_Y1),
            r.get(COL_RCF_NETDEBT_Y2),
            r.get(COL_RCF_NETDEBT_Y3)
        )
        s_fcf    = scored_fcf_debt_3y(
            r.get(COL_FCF_DEBT_Y1),
            r.get(COL_FCF_DEBT_Y2),
            r.get(COL_FCF_DEBT_Y3)
        )
        s_eint   = scored_ebita_int_3y(
            r.get(COL_EBITA_INT_Y1),
            r.get(COL_EBITA_INT_Y2),
            r.get(COL_EBITA_INT_Y3)
        )

        # --- QUALITATIFS
        s_bp   = score_quali(r.get(COL_BUSINESS_PROFILE))
        s_pol  = score_quali(r.get(COL_FINANCIAL_POLICY))

        # --- Scorecard outcome
        base = moodys_industrial_score_from_scores(
            name=name,
            s_scale=s_scale,
            s_business_profile=s_bp,
            s_ebita_margin=s_ebita,
            s_debt=s_debt,
            s_rcf_nd=s_rcf_nd,
            s_fcf=s_fcf,
            s_ebita_int=s_eint,
            s_pol=s_pol
        )

        # --- Ajustements hors scorecard (soft, cap ±1 cran)
        delta_soft = other_considerations_soft_delta(
            esg_score=r.get(COL_ESG_SCORE),
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
            "delta_other_considerations_soft": delta_soft,
            "final_adjusted_score": adj_score,
            "final_assigned_rating": final_rating
        })

    out = pd.DataFrame(results)
    out.to_csv(OUTPUT_CSV, index=False)

    print("\n✅ Industrials – calcul terminé (3Y + soft adjustments ±1). Résumé :\n")
    print(out[[ "name",
                "scorecard_aggregate","scorecard_rating",
                "delta_other_considerations_soft",
                "final_adjusted_score","final_assigned_rating"]].to_string(index=False))
    print(f"\n➡️ Détail sauvegardé dans {OUTPUT_CSV}")