"""Non-regression of PCF results against a validated baseline file.

The baseline is a validated copy of the per-product PCF results, stored in
`baselines/pcf_results_baseline.csv` (not the customer Excel files: only the
computed results are versioned). The comparison is tolerant to floating-point
noise and reports every field that drifted.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

DEFAULT_BASELINE = Path(__file__).resolve().parents[2] / "baselines" / "pcf_results_baseline.csv"
KEY_COLUMN = "Product SKU"
TOLERANCE = 1e-9


def save_baseline(product_results: pd.DataFrame, baseline_path: str | Path | None = None) -> Path:
    """Store the current validated results as the new regression baseline."""
    path = Path(baseline_path or DEFAULT_BASELINE)
    path.parent.mkdir(parents=True, exist_ok=True)
    product_results.to_csv(path, index=False)
    return path


def load_baseline(baseline_path: str | Path | None = None) -> pd.DataFrame:
    path = Path(baseline_path or DEFAULT_BASELINE)
    if not path.is_file():
        raise FileNotFoundError(f"Baseline not found: {path}")
    return pd.read_csv(path)


def compare_with_baseline(product_results: pd.DataFrame,
                          baseline: pd.DataFrame,
                          key: str = KEY_COLUMN) -> pd.DataFrame:
    """Compare current results with the baseline.

    Returns a report with one row per difference:
    - 'new product' (in results, not in baseline)
    - 'missing product' (in baseline, not in results)
    - 'value drift' (numeric field beyond TOLERANCE, with both values)
    - 'text change' (non-numeric field that differs)
    Empty report = no regression.
    """
    results = product_results.copy()
    base = baseline.copy()
    results[key] = results[key].astype(str).str.strip()
    base[key] = base[key].astype(str).str.strip()

    diffs = []
    for sku in set(results[key]) - set(base[key]):
        diffs.append({key: sku, "Field": "*", "Type": "new product", "Baseline": None, "Current": None})
    for sku in set(base[key]) - set(results[key]):
        diffs.append({key: sku, "Field": "*", "Type": "missing product", "Baseline": None, "Current": None})

    merged = results.merge(base, on=key, how="inner", suffixes=("_current", "_baseline"))
    common_columns = [
        c for c in results.columns
        if c != key and c in base.columns
        and c + "_current" in merged.columns and c + "_baseline" in merged.columns
    ]
    for column in common_columns:
        current_values = merged[column + "_current"]
        baseline_values = merged[column + "_baseline"]
        numeric = pd.to_numeric(current_values, errors="coerce").notna() & pd.to_numeric(
            baseline_values, errors="coerce"
        ).notna()
        for idx in merged.index:
            cur, base_val = merged.at[idx, column + "_current"], merged.at[idx, column + "_baseline"]
            if pd.isna(cur) and pd.isna(base_val):
                continue
            if pd.isna(cur) != pd.isna(base_val):
                diffs.append({key: merged.at[idx, key], "Field": column, "Type": "value drift",
                              "Baseline": base_val, "Current": cur})
            elif numeric.loc[idx] and abs(float(cur) - float(base_val)) > TOLERANCE:
                diffs.append({key: merged.at[idx, key], "Field": column, "Type": "value drift",
                              "Baseline": base_val, "Current": cur})
            elif not numeric.loc[idx] and str(cur) != str(base_val):
                diffs.append({key: merged.at[idx, key], "Field": column, "Type": "text change",
                              "Baseline": base_val, "Current": cur})
    return pd.DataFrame(diffs, columns=[key, "Field", "Type", "Baseline", "Current"])


def check_regression(product_results: pd.DataFrame,
                     baseline_path: str | Path | None = None) -> pd.DataFrame:
    """Load the baseline and return the regression report (empty = OK)."""
    return compare_with_baseline(product_results, load_baseline(baseline_path))
