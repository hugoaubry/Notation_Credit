"""
Économétrie – Validation du modèle de notation crédit

Objectifs :
- Fusionner toutes les sorties de scorecards sectorielles
- Les enrichir avec la vraie note d’agence (depuis universe_metadata.csv)
- Convertir les notes qualitatives (Aaa, Aa1, Baa2, …) en scores numériques
- Calculer les écarts (notre modèle vs agence)
- Estimer une régression OLS : rating_agence_num ~ facteurs du modèle
"""

import os
from typing import Optional, Dict

import numpy as np
import pandas as pd
import statsmodels.api as sm


OUTPUT_DIR = "data/output"
META_PATH = "data/metadata/universe_metadata.csv"
AGENCY_COL = "moody_rating"

# =========================
#   MAPPING NOTES -> NUM
# =========================

# On garde la même logique que dans tes modèles :
# Aaa = 1, Aa1 = 2, ..., C = 21

RATING_ORDER = [
    "Aaa",
    "Aa1","Aa2","Aa3",
    "A1","A2","A3",
    "Baa1","Baa2","Baa3",
    "Ba1","Ba2","Ba3",
    "B1","B2","B3",
    "Caa1","Caa2","Caa3",
    "Ca","C"
]

RATING_TO_NUM: Dict[str, int] = {lab: i+1 for i, lab in enumerate(RATING_ORDER)}

def normalize_rating_label(x: Optional[str]) -> Optional[str]:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).strip()
    if s == "":
        return None
    # Harmonisation simple (ex: "baa2" -> "Baa2")
    s = s.replace(" ", "")
    s_up = s.upper()

    # Quelques alias fréquents possibles
    aliases = {
        "AAA": "Aaa",
        "AA+": "Aa1", "AA": "Aa2", "AA-": "Aa3",
        "A+": "A1",   "A": "A2",   "A-": "A3",
        "BBB+": "Baa1", "BBB": "Baa2", "BBB-": "Baa3",
        "BB+": "Ba1",   "BB": "Ba2",   "BB-": "Ba3",
        "B+": "B1",     "B": "B2",     "B-": "B3",
        "CCC+": "Caa1", "CCC": "Caa2", "CCC-": "Caa3",
        "CC": "Ca", "C": "C"
    }
    if s_up in aliases:
        return aliases[s_up]

    # Si déjà au bon format (Aaa, Baa2, etc.)
    if s in RATING_TO_NUM:
        return s

    # Tentative de formatage type "BAA2" -> "Baa2"
    if len(s_up) >= 3 and s_up[0].isalpha():
        base = s_up[0] + s_up[1:].lower()   # BAA2 -> Baa2
        if base in RATING_TO_NUM:
            return base

    return None

def rating_to_num(x: Optional[str]) -> Optional[float]:
    lab = normalize_rating_label(x)
    if lab is None:
        return None
    return float(RATING_TO_NUM.get(lab, np.nan))

# =========================
#   CHARGEMENT DES OUTPUTS DE SECTEURS
# =========================

def load_all_sector_outputs(output_dir: str) -> pd.DataFrame:
    """
    Charge tous les fichiers *_output_scorecard.csv dans data/output
    et les concatène dans un seul DataFrame.
    Ajoute une colonne 'sector' si possible.
    Ignore les fichiers vides (comme oil_output_scorecard.csv pour l'instant).
    """
    frames = []
    for fname in os.listdir(output_dir):
        if not fname.endswith("_output_scorecard.csv"):
            continue
        fpath = os.path.join(output_dir, fname)
        try:
            df = pd.read_csv(fpath)
        except pd.errors.EmptyDataError:
            print(f"⚠️ Fichier vide ignoré : {fpath}")
            continue
        except Exception as e:
            print(f"⚠️ Impossible de lire {fpath} : {e}")
            continue

        # Si le CSV est lisible mais vide / sans colonnes, on ignore aussi
        if df is None or df.empty or len(df.columns) == 0:
            print(f"⚠️ Fichier sans données utiles ignoré : {fpath}")
            continue

        if "sector" not in df.columns:
            # On essaie de déduire le secteur du nom de fichier : ex 'automobile_output_scorecard.csv'
            sector_name = fname.replace("_output_scorecard.csv", "")
        else:
            sector_name = None

        if sector_name is not None:
            df["sector"] = sector_name

        frames.append(df)

    if not frames:
        raise RuntimeError("Aucun fichier *_output_scorecard.csv exploitable trouvé dans data/output")

    all_data = pd.concat(frames, ignore_index=True)
    return all_data

# =========================
#   FONCTION PRINCIPALE
# =========================

if __name__ == "__main__":
    # 1) Chargement des outputs de scorecards
    print("📥 Chargement des fichiers de scorecards sectorielles...")
    all_scores = load_all_sector_outputs(OUTPUT_DIR)
    print(f"✅ {len(all_scores)} lignes chargées depuis {OUTPUT_DIR}")

    # 2) Chargement des métadonnées (dont la note agence)
    print("📥 Chargement de universe_metadata...")
    meta = pd.read_csv(META_PATH, sep=";")

    if "name" not in meta.columns:
        raise KeyError("Le fichier universe_metadata.csv doit contenir une colonne 'name'.")

    if AGENCY_COL not in meta.columns:
        raise KeyError(
            f"Le fichier universe_metadata.csv doit contenir une colonne '{AGENCY_COL}'. "
            f"Tu peux soit la créer, soit changer la variable AGENCY_COL dans econometrie.py."
        )

    # 3) Merge sur le nom de la société
    if "sector" in meta.columns:
        meta_cols = ["name", AGENCY_COL, "sector"]
    else:
        meta_cols = ["name", AGENCY_COL]

    merged = all_scores.merge(meta[meta_cols], on="name", how="left")

    # 3bis) Normalisation de la colonne secteur après le merge
    # Selon les fichiers, on peut avoir 'sector', 'sector_x' ou 'sector_y'.
    sector_col = None
    for cand in ["sector", "sector_x", "sector_y"]:
        if cand in merged.columns:
            sector_col = cand
            break

    if sector_col is not None and sector_col != "sector":
        merged["sector"] = merged[sector_col]

    # 4) Conversion des notes qualitatives en numériques
    # - Note agence (référence)
    merged["agency_rating_norm"] = merged[AGENCY_COL].apply(normalize_rating_label)
    merged["agency_rating_num"]  = merged["agency_rating_norm"].apply(rating_to_num)

    # - Notre note finale (final_assigned_rating)
    if "final_assigned_rating" not in merged.columns:
        raise KeyError("Les outputs de scorecard doivent contenir la colonne 'final_assigned_rating'.")

    merged["model_rating_norm"] = merged["final_assigned_rating"].apply(normalize_rating_label)
    merged["model_rating_num"]  = merged["model_rating_norm"].apply(rating_to_num)

    # 5) Filtrage lignes valides (où on a bien une note agence ET une note modèle)
    valid = merged.dropna(subset=["agency_rating_num", "model_rating_num"]).copy()
    print(f"✅ {len(valid)} lignes utilisées pour l'analyse économétrique (avec note agence dispo).")

    if valid.empty:
        print("⚠️ Aucune ligne avec notes agence + modèle disponibles. Vérifie universe_metadata.moody_rating.")
        exit(0)

    # 6) Calcul des écarts (en 'crans')
    valid["diff_model_minus_agency"] = valid["model_rating_num"] - valid["agency_rating_num"]
    valid["abs_diff"]               = valid["diff_model_minus_agency"].abs()

    # 7) Statistiques globales
    print("\n📊 Statistiques globales des écarts (notre modèle vs agence) :\n")
    print(valid[["diff_model_minus_agency", "abs_diff"]].describe())

    # 8) Statistiques par secteur (si la colonne sector existe)
    if "sector" in valid.columns:
        print("\n📊 Écarts moyens par secteur :\n")
        grp = valid.groupby("sector").agg(
            mean_diff=("diff_model_minus_agency", "mean"),
            mean_abs_diff=("abs_diff", "mean"),
            n=("name", "count")
        ).reset_index()
        print(grp.to_string(index=False))

    # 9) RÉGRESSION OLS : note agence expliquée par les facteurs de notre modèle
    # On prend quelques facteurs communs aux secteurs (si dispo)
    candidate_X_cols = [
        "factor_scale",
        "factor_business_profile",
        "factor_profitability_efficiency",
        "factor_leverage_coverage",
        "factor_financial_policy",
    ]
    X_cols = [c for c in candidate_X_cols if c in valid.columns]

    if not X_cols:
        print("\n⚠️ Aucun des facteurs 'factor_*' n'a été trouvé dans les données. "
              "Vérifie les noms de colonnes dans les CSV de sortie.")
    else:
        reg_data = valid.dropna(subset=["agency_rating_num"] + X_cols).copy()

        if reg_data.empty:
            print("\n⚠️ Pas assez de données complètes pour lancer la régression OLS.")
        else:
            Y = reg_data["agency_rating_num"]
            X = reg_data[X_cols]
            X = sm.add_constant(X)

            print("\n📈 Régression OLS : agency_rating_num ~ facteurs du modèle")
            print(f"Variables explicatives utilisées : {X_cols}")

            model = sm.OLS(Y, X).fit()
            print(model.summary())

    # 9bis) RÉGRESSIONS PAR SECTEUR
    if "sector" in valid.columns and X_cols:
        print("\n📊 RÉGRESSIONS PAR SECTEUR\n")

        for sector_name, sub in valid.groupby("sector"):
            sub_reg = sub.dropna(subset=["agency_rating_num"] + X_cols).copy()
            if len(sub_reg) < 6:
                print(f"➡️ Secteur '{sector_name}': échantillon trop petit ({len(sub_reg)} obs), régression ignorée.")
                continue

            Y_s = sub_reg["agency_rating_num"]
            X_s = sm.add_constant(sub_reg[X_cols])

            print(f"\n=== Secteur : {sector_name} (n={len(sub_reg)}) ===")
            try:
                model_s = sm.OLS(Y_s, X_s).fit()
                print(f"R² = {model_s.rsquared:.3f}  |  R² ajusté = {model_s.rsquared_adj:.3f}")
                coef_df = pd.DataFrame({
                    "coef": model_s.params,
                    "t_stat": model_s.tvalues,
                    "p_value": model_s.pvalues,
                })
                print(coef_df.to_string())
            except Exception as e:
                print(f"⚠️ Régression impossible pour le secteur '{sector_name}': {e}")

    # 10) Sauvegarde optionnelle des données fusionnées pour d'autres analyses
    out_path = os.path.join(OUTPUT_DIR, "econometrics_merged_dataset.csv")
    merged.to_csv(out_path, index=False)
    print(f"\n💾 Dataset fusionné sauvegardé dans : {out_path}")

    print("\n✅ Partie économétrique terminée.")
    print("   → Utilise les stats globales, par secteur et le résumé OLS pour ta présentation.")