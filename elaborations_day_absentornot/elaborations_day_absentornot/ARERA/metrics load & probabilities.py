from pathlib import Path

import numpy as np
import pandas as pd
from scipy.stats import pearsonr, spearmanr


# Modifica la cartella se i file si trovano altrove.
BASE_DIR = Path(r"")

FILES = {
    2021: BASE_DIR / "Database_grafico_subactivities_2021.xlsx",
    2022: BASE_DIR / "Database_grafico_subactivities_2022.xlsx",
    2024: BASE_DIR / "Database_grafico_subactivities_2024.xlsx",
    2025: BASE_DIR / "Database_grafico_subactivities_2025.xlsx",
}

OUTPUT_PATH = BASE_DIR / "profile_similarity_metrics.xlsx"

KEYS = ["rip", "day_type", "meseri_y", "activity_class"]
PROB_COLS = [f"ora_{h}" for h in range(24)]
LOAD_COLS = [f"load_ora_{h}" for h in range(24)]


def safe_corr(x, y, method="pearson"):
    """Restituisce il coefficiente di correlazione, o NaN se non calcolabile."""
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    valid = np.isfinite(x) & np.isfinite(y)
    x, y = x[valid], y[valid]

    if len(x) != 24 or np.ptp(x) == 0 or np.ptp(y) == 0:
        return np.nan

    if method == "pearson":
        return float(pearsonr(x, y).statistic)

    return float(spearmanr(x, y).statistic)


def profile_errors(prob, load):
    """
    Normalizza entrambi i profili affinché ciascuno sommi a 1,
    quindi calcola MAE e RMSE tra le quote orarie.
    """
    prob = np.asarray(prob, dtype=float)
    load = np.asarray(load, dtype=float)

    if not (np.isfinite(prob).all() and np.isfinite(load).all()):
        return np.nan, np.nan

    if np.any(prob < 0) or np.any(load < 0):
        return np.nan, np.nan

    sum_prob = prob.sum()
    sum_load = load.sum()

    if sum_prob <= 0 or sum_load <= 0:
        return np.nan, np.nan

    prob_share = prob / sum_prob
    load_share = load / sum_load
    diff = prob_share - load_share

    mae = float(np.mean(np.abs(diff)))
    rmse = float(np.sqrt(np.mean(diff ** 2)))

    return mae, rmse


# Legge i file e calcola le metriche per ogni gruppo-anno.
annual_results = []

for year, path in FILES.items():
    if not path.exists():
        raise FileNotFoundError(f"File non trovato per {year}: {path}")

    df = pd.read_excel(path)
    df.columns = df.columns.astype(str).str.strip()

    required = KEYS + PROB_COLS + LOAD_COLS
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise KeyError(
            f"Nel file {path.name} mancano queste colonne:\n{missing}"
        )

    # Controlla che le colonne orarie siano numeriche.
    for col in PROB_COLS + LOAD_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Un record per chiave e anno è necessario per il confronto annuale.
    duplicate = df.duplicated(KEYS, keep=False)
    if duplicate.any():
        examples = df.loc[duplicate, KEYS].head(10)
        raise ValueError(
            f"{path.name}: trovate righe duplicate per le chiavi {KEYS}.\n"
            f"Esempi:\n{examples.to_string(index=False)}"
        )

    for _, row in df.iterrows():
        prob = row[PROB_COLS].to_numpy(dtype=float)
        load = row[LOAD_COLS].to_numpy(dtype=float)

        mae, rmse = profile_errors(prob, load)

        annual_results.append({
            **{key: row[key] for key in KEYS},
            "year": year,
            "pearson_r": safe_corr(prob, load, method="pearson"),
            "spearman_rho": safe_corr(prob, load, method="spearman"),
            "MAE_normalized": mae,
            "RMSE_normalized": rmse,
        })

annual = pd.DataFrame(annual_results)

# Riassume la stabilità delle metriche tra gli anni disponibili.
summary = (
    annual.groupby(KEYS, dropna=False)
    .agg(
        n_years_pearson=("pearson_r", "count"),
        pearson_r_mean=("pearson_r", "mean"),
        pearson_r_median=("pearson_r", "median"),
        pearson_r_sd=("pearson_r", "std"),
        pearson_r_min=("pearson_r", "min"),
        pearson_r_max=("pearson_r", "max"),

        n_years_spearman=("spearman_rho", "count"),
        spearman_rho_mean=("spearman_rho", "mean"),
        spearman_rho_median=("spearman_rho", "median"),
        spearman_rho_sd=("spearman_rho", "std"),
        spearman_rho_min=("spearman_rho", "min"),
        spearman_rho_max=("spearman_rho", "max"),

        n_years_MAE=("MAE_normalized", "count"),
        MAE_normalized_mean=("MAE_normalized", "mean"),
        MAE_normalized_sd=("MAE_normalized", "std"),
        MAE_normalized_min=("MAE_normalized", "min"),
        MAE_normalized_max=("MAE_normalized", "max"),

        RMSE_normalized_mean=("RMSE_normalized", "mean"),
        RMSE_normalized_sd=("RMSE_normalized", "std"),
        RMSE_normalized_min=("RMSE_normalized", "min"),
        RMSE_normalized_max=("RMSE_normalized", "max"),
    )
    .reset_index()
)

# Esporta:
# 1) una riga per gruppo e anno;
# 2) il riassunto interannuale per gruppo.
with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    annual.to_excel(writer, sheet_name="annual_metrics", index=False)
    summary.to_excel(writer, sheet_name="between_year_summary", index=False)

print(f"Salvato: {OUTPUT_PATH}")
print(f"Righe metriche annuali: {len(annual)}")
print(f"Combinazioni nel riepilogo: {len(summary)}")