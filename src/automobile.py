from typing import Optional, Dict, Tuple
import pandas as pd

INPUT_XLSX  = "data/intput/autos_input_template.xlsx"
OUTPUT_CSV  = "data/output/autos_output_scorecard.csv" 
INPUT_SHEET = "autos_input_template"

COL_NAME = "name"
COL_TREND_QUALI = "share_y3" # label qualitatif (si pas d'auto-trend)
COL_MARKET_POSITION = "market_position"
COL_FINANCIAL_POLICY = "financial_policy"

# calcul de ratios sur 3 ans (Y1 = -3ans, Y2 = -2ans, Y3 = -1an)
COL_EBIT_MARGIN_Y1 = "ebit_margin_y1"
COL_EBIT_MARGIN_Y2 = "ebit_margin_y2"
COL_EBIT_MARGIN_Y3 = "ebit_margin_y3"

COL_DEBT_EBITDA_Y1 = "debt_ebitda_y1"
COL_DEBT_EBITDA_Y2 = "debt_ebitda_y2"
COL_DEBT_EBITDA_Y3 = "debt_ebitda_y3"

COL_EBIT_INT_Y1 = "ebit_interest_y1"
COL_EBIT_INT_Y2 = "ebit_interest_y2"
COL_EBIT_INT_Y3 = "ebit_interest_y3"

COL_RCF_DEBT_Y1 = "rcf_debt_y1"
COL_RCF_DEBT_Y2 = "rcf_debt_y2"
COL_RCF_DEBT_Y3 = "rcf_debt_y3"

COL_FCF_DEBT_Y1 = "fcf_debt_y1"
COL_FCF_DEBT_Y2 = "fcf_debt_y2"
COL_FCF_DEBT_Y3 = "fcf_debt_y3"

# Liquidity ratio (S/U)
COL_LIQ_Y1 = "liq_y1"
COL_LIQ_Y2 = "liq_y2"
COL_LIQ_Y3 = "liq_y3"

# inscription des "other considerations"
COL_ESG_SCORE = "esg_score"          # 1..5 (1 = meilleur)
COL_CAPTIVE_RATIO = "captive_ratio"      # 0..1
COL_REGULATION_SCORE = "regulation_score"   # 1..5 (1 = meilleur)
COL_MANAGEMENT_SCORE = "management_score"   # 1..5 (1 = meilleur)
COL_NONWHOLLY_SALES = "nonwholly_sales"    # 0..1
COL_EVENT_RISK_SCORE = "event_risk_score"   # 0/1/2 (2 = risque élevé)
COL_PARENTAL_SUPPORT = "parental_support"   # -3..+3 (on cape à ±1 cran d'effet)

# Part de marché unitaire (% unités) pour l’auto-trend
RAW = dict(share_units = ["share_units_y1","share_units_y2","share_units_y3"],)

# Optionnel : marge “forward”
COL_EBIT_MARGIN_FWD = "ebit_margin_fwd"  # % si fourni

# passade de quali à numérique
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
    if lbl is None or (isinstance(lbl,float) and pd.isna(lbl)): return None
    s = str(lbl).strip()
    if s in QUALI_NUM: return s
    if s in NOTCH_TO_ALPHA: return NOTCH_TO_ALPHA[s]
    sU = s.upper()
    if sU in NOTCH_TO_ALPHA: return NOTCH_TO_ALPHA[sU]
    return None

def score_quali(label: Optional[str]) -> float:
    a = _alpha_from_label(label)
    return QUALI_NUM[a] if a else 12.0  # défaut Ba si inconnu

#   PARSING NUMÉRIQUE ROBUSTE

def to_float(x) -> Optional[float]:
    if x is None: return None
    if isinstance(x, (int, float)):
        if pd.isna(x): return None
        return float(x)
    s = str(x).strip()
    if s == "": return None
    s = s.replace(",", "")
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    if s.endswith("%"):
        s = s[:-1].strip()  # On suppose ici que 25.0 = 25% (pas 0.25)
    try:
        val = float(s)
    except:
        return None
    if neg: val = -val
    return val

#   MAPPINGS & BORNES MOODY'S

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

EBIT_MARGIN_BOUNDS = [
    ("Aaa", 15.0, 40.0),
    ("Aa",  10.0, 15.0),
    ("A",   7.0,  10.0),
    ("Baa", 4.0,  7.0),
    ("Ba",  1.0,  4.0),
    ("B",   0.5,  1.0),
    ("Caa", 0.0,  0.5),
    ("Ca",  -1e9, 0.0)
]
DEBT_EBITDA_BOUNDS = [
    ("Aaa", 0.0,  1.5),
    ("Aa",  1.5,  2.5),
    ("A",   2.5,  3.5),
    ("Baa", 3.5,  4.5),
    ("Ba",  4.5,  6.0),
    ("B",   6.0,  9.0),
    ("Caa", 9.0,  12.0),
    ("Ca",  12.0, 1e9)
]
EBIT_INT_BOUNDS = [
    ("Aaa", 15.0, 40.0),
    ("Aa",  10.0, 15.0),
    ("A",   5.0,  10.0),
    ("Baa", 2.0,  5.0),
    ("Ba",  1.0,  2.0),
    ("B",   0.5,  1.0),
    ("Caa", 0.0,  0.5),
    ("Ca",  -1e9, 0.0)
]
RCF_DEBT_BOUNDS = [
    ("Aaa", 75.0, 140.0),
    ("Aa",  50.0, 75.0),
    ("A",   30.0, 50.0),
    ("Baa", 20.0, 30.0),
    ("Ba",  10.0, 20.0),
    ("B",   5.0,  10.0),
    ("Caa", 2.5,  5.0),
    ("Ca",  -1e9, 2.5)
]
FCF_DEBT_BOUNDS = [
    ("Aaa", 30.0, 100.0),
    ("Aa",  20.0, 30.0),
    ("A",   10.0, 20.0),
    ("Baa", 5.0,  10.0),
    ("Ba",  0.0,  5.0),
    ("B",   -5.0, 0.0),
    ("Caa", -10.0, -5.0),
    ("Ca",  -1e9, -10.0)
]

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
    if higher_is_better:
        return 0.5 if x >= bounds[0][2] else 20.5
    return 0.5 if x <= bounds[0][1] else 20.5

def score_ebit_margin(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, EBIT_MARGIN_BOUNDS, True)

def score_debt_ebitda(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, DEBT_EBITDA_BOUNDS, False)

def score_ebit_interest(x: Optional[float]) -> float:
    v = to_float(x)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, EBIT_INT_BOUNDS, True)

def score_rcf_debt(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, RCF_DEBT_BOUNDS, True)

def score_fcf_debt(pct: Optional[float]) -> float:
    v = to_float(pct)
    if v is None: return 12.0
    return _score_quant_from_bounds(v, FCF_DEBT_BOUNDS, True)

#   PONDÉRATIONS MOODY’S
W = {
    "trend_global_share": 10.0,
    "market_position":    30.0,
    "ebit_margin":        20.0,
    "debt_ebitda":        10.0,
    "ebit_interest":      5.0,
    "rcf_debt":           5.0,
    "fcf_debt":           5.0,
    "financial_policy":   15.0
}
assert abs(sum(W.values()) - 100.0) < 1e-8
 
#   EXHIBIT 5 : SCORE -> RATING

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
        if x <= thr: return lab
    return "C"

#   OUTILS Y1-Y3

def weighted_average_scores_3y(s1=None, s2=None, s3=None):
    """ Moyenne pondérée des SCORES : 0.2*s1 + 0.3*s2 + 0.5*s3 (y1 ancien → y3 récent). """
    vals, weights = [], []
    for v, w in [(s1,0.2),(s2,0.3),(s3,0.5)]:
        fv = to_float(v)
        if fv is None: 
            continue
        vals.append(fv); weights.append(w)
    if not vals: return None
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

def scored_ebit_interest_3y(y1=None, y2=None, y3=None):
    s1 = score_ebit_interest(y1)
    s2 = score_ebit_interest(y2)
    s3 = score_ebit_interest(y3)
    return weighted_average_scores_3y(s1, s2, s3)

def scored_rcf_debt_3y(y1=None, y2=None, y3=None):
    s1 = score_rcf_debt(y1)
    s2 = score_rcf_debt(y2)
    s3 = score_rcf_debt(y3)
    return weighted_average_scores_3y(s1, s2, s3)

def scored_fcf_debt_3y(y1=None, y2=None, y3=None):
    s1 = score_fcf_debt(y1)
    s2 = score_fcf_debt(y2)
    s3 = score_fcf_debt(y3)
    return weighted_average_scores_3y(s1, s2, s3)

#   AJUSTEMENTS MOODY’S-LIKE (SOFT, ±1 cran)

def other_considerations_soft_delta(
    *,
    esg_score=None,              # 1..5 (1 = meilleur)
    liquidity_ratio=None,        # Sources/Uses (moyenne 3y)
    captive_ratio=None,          # 0..1
    regulation_score=None,       # 1..5 (1 = meilleur)
    management_score=None,       # 1..5 (1 = meilleur)
    nonwholly_sales=None,        # 0..1
    event_risk_score=None,       # 0,1,2
    parental_support=None,       # -3..+3  (effect cap +1)
    financial_policy_label=None
) -> float:
    """
    Delta symétrique, modéré, CAP ±1 cran au total.
    +delta = amélioration (score numérique diminue), -delta = pénalité.
    """
    delta = 0.0

    # ESG (modéré)
    esg = to_float(esg_score)
    if esg is not None:
        if esg <= 2: delta += 0.25
        elif esg >= 4: delta -= 0.25

    # Liquidité (modérée ; pas +0.5 automatique)
    lr = to_float(liquidity_ratio)
    if lr is not None:
        if lr > 2.0: delta += 0.25
        elif lr < 1.0: delta -= 0.5

    # Captive finance (plus de captive = pénalité)
    cr = to_float(captive_ratio)
    if cr is not None:
        if cr > 0.40: delta -= 0.5
        elif cr > 0.20: delta -= 0.25

    # Régulation
    rs = to_float(regulation_score)
    if rs is not None:
        if rs <= 2: delta += 0.25
        elif rs >= 4: delta -= 0.25

    # Management (éviter double-compte si FP élevée)
    fp_alpha = _alpha_from_label(financial_policy_label)
    mgmt_contrib_allowed = (fp_alpha not in ("Aaa","Aa"))
    ms = to_float(management_score)
    if ms is not None and mgmt_contrib_allowed:
        if ms <= 2: delta += 0.25
        elif ms >= 4: delta -= 0.5

    # Non-wholly owned sales (élevé = pénalité)
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

    # Parental / Gov support (limité à +1 cran d'effet total)
    ps = to_float(parental_support)
    if ps is not None:
        ps = int(ps)
        if ps > 0: delta += min(ps, 1) * 0.5
        elif ps < 0: delta += max(ps, -1) * 0.5

    # CAP global ±1 cran
    delta = max(-1.0, min(1.0, delta))
    return round(delta, 3)

def apply_adjustment(scorecard_aggregate: float, delta_crans: float) -> Tuple[float,str]:
    adjusted = scorecard_aggregate - delta_crans  # +delta = amélioration -> score diminue
    return round(adjusted,3), score_to_rating(adjusted)

#   AUTRES OUTILS

def classify_share_trend(s1, s2, s3, thr_up=+0.5, thr_down=-0.5):
    """ Auto-classement de la tendance de part de marché (en points de % sur 2 ans). """
    v1, v3 = to_float(s1), to_float(s3)
    if v1 is None or v3 is None:
        return None
    delta = v3 - v1
    if delta >= thr_up:   return "A"    # hausse
    if delta <= thr_down: return "Ba"   # baisse
    return "Baa"          # stable

#   SCORECARD FROM SCORES

def moodys_auto_score_from_scores(*, name,
                                  s_trend, s_market,
                                  s_ebitm, s_debt, s_cov, s_rcf, s_fcf,
                                  s_pol) -> Dict[str, float]:
    score_business = (W["trend_global_share"]*s_trend + W["market_position"]*s_market) / 40.0
    score_profit   = s_ebitm
    score_levcov   = (W["debt_ebitda"]*s_debt + W["ebit_interest"]*s_cov +
                      W["rcf_debt"]*s_rcf + W["fcf_debt"]*s_fcf) / 25.0
    score_policy   = s_pol

    agg = (0.40*score_business + 0.20*score_profit + 0.25*score_levcov + 0.15*score_policy)
    rating = score_to_rating(agg)

    return {
        "name": name,
        "scorecard_aggregate": round(agg,3),
        "scorecard_rating": rating,
        "factor_business_profile": round(score_business,3),
        "factor_profitability_efficiency": round(score_profit,3),
        "factor_leverage_coverage": round(score_levcov,3),
        "factor_financial_policy": round(score_policy,3),
        "sf_trend_global_share": round(s_trend,3),
        "sf_market_position": round(s_market,3),
        "sf_ebit_margin": round(s_ebitm,3),
        "sf_debt_ebitda": round(s_debt,3),
        "sf_ebit_interest": round(s_cov,3),
        "sf_rcf_debt": round(s_rcf,3),
        "sf_fcf_debt": round(s_fcf,3),
        "sf_financial_policy": round(s_pol,3),
    }

#   PIPELINE

if __name__ == "__main__":
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET)

    results = []
    for _, r in df.iterrows():
        name = r[COL_NAME]

        # PROFITABILITY : option "forward" si tu fournis ebit_margin_fwd
        ebit_margin_fwd = r.get(COL_EBIT_MARGIN_FWD)
        if ebit_margin_fwd not in (None, ""):
            s_y3 = score_ebit_margin(r.get(COL_EBIT_MARGIN_Y3))
            s_f  = score_ebit_margin(ebit_margin_fwd)
            s_ebitm = (s_y3 + s_f)/2.0
        else:
            s_ebitm = scored_ebit_margin_3y(
                r.get(COL_EBIT_MARGIN_Y1),
                r.get(COL_EBIT_MARGIN_Y2),
                r.get(COL_EBIT_MARGIN_Y3)
            )

        # SCORES quanti 3Y (moyenne pondérée des scores)
        s_debt  = scored_debt_ebitda_3y(
            r.get(COL_DEBT_EBITDA_Y1),
            r.get(COL_DEBT_EBITDA_Y2),
            r.get(COL_DEBT_EBITDA_Y3)
        )
        s_cov   = scored_ebit_interest_3y(
            r.get(COL_EBIT_INT_Y1),
            r.get(COL_EBIT_INT_Y2),
            r.get(COL_EBIT_INT_Y3)
        )
        s_rcf   = scored_rcf_debt_3y(
            r.get(COL_RCF_DEBT_Y1),
            r.get(COL_RCF_DEBT_Y2),
            r.get(COL_RCF_DEBT_Y3)
        )
        s_fcf   = scored_fcf_debt_3y(
            r.get(COL_FCF_DEBT_Y1),
            r.get(COL_FCF_DEBT_Y2),
            r.get(COL_FCF_DEBT_Y3)
        )

        # Qualitatifs -> scores numériques
        # Auto-trend si parts de marché (% unités) fournies :
        auto_trend = None
        su = RAW["share_units"]
        if su[0] in r and su[1] in r and su[2] in r:
            auto_trend = classify_share_trend(r.get(su[0]), r.get(su[1]), r.get(su[2]))
        trend_label = auto_trend if auto_trend else r.get(COL_TREND_QUALI)

        s_trend  = score_quali(trend_label)
        s_market = score_quali(r.get(COL_MARKET_POSITION))
        s_pol    = score_quali(r.get(COL_FINANCIAL_POLICY))

        # ---- Scorecard outcome (scores agrégés)
        base = moodys_auto_score_from_scores(
            name=name,
            s_trend=s_trend, s_market=s_market,
            s_ebitm=s_ebitm, s_debt=s_debt, s_cov=s_cov, s_rcf=s_rcf, s_fcf=s_fcf,
            s_pol=s_pol
        )

        # ---- Liquidity 3Y moyenne (à partir des ratios Excel)
        liq3 = weighted_average_scores_3y(
            to_float(r.get(COL_LIQ_Y1)),
            to_float(r.get(COL_LIQ_Y2)),
            to_float(r.get(COL_LIQ_Y3))
        )

        # ---- Ajustements hors scorecard (modérés, cap ±1 cran)
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

        # ---- LOG / SORTIE
        results.append({
            **base,
            "inputs_liquidity_ratio_3y": liq3,
            "delta_other_considerations_soft": delta_soft,
            "final_adjusted_score": adj_score,
            "final_assigned_rating": final_rating
        })

    out = pd.DataFrame(results)
    out.to_csv(OUTPUT_CSV, index=False)

    print("\n✅ Calcul 3Y terminé (scores à partir des ratios Excel + soft adjustments cap ±1). Résumé :\n")
    print(out[[ "name",
                "scorecard_aggregate","scorecard_rating",
                "delta_other_considerations_soft",
                "final_adjusted_score","final_assigned_rating"]].to_string(index=False))
    print(f"\n➡️ Détail sauvegardé dans {OUTPUT_CSV}")