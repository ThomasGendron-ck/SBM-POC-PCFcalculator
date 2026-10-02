"""Missing activity data assessment on the filled data collection file.

Reads back the collection workbook once SBM / suppliers have filled it and
produces a prioritized list of missing or suspicious activity data.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

MANDATORY_COMPONENT_FIELDS = [
    "Supplier Code",
    "Supplier Name",
    "Raw Material (MB Product)",
    "Net Weight",
]

ACTIVITY_FIELDS = [
    "Supplier PCF value",
    "Supplier PCF framework",
    "Supplier PCF scope",
    "Transformation Process Name",
    "Transformation Energy Type",
    "Transformation Energy Consumption",
    "Scrap Rate",
]


def import_filled_collection(collection_path: str | Path) -> pd.DataFrame:
    """Read the Product and Component sheets of the filled collection file
    into a single DataFrame tagged by level."""
    path = Path(collection_path)
    frames = []
    for sheet, level, sku_col in [("Product", "Product", "Product SKU"), ("Component", "Component", "Component SKU")]:
        df = pd.read_excel(path, sheet_name=sheet, header=1)
        df = df.loc[:, ~pd.Index(df.columns).duplicated()]
        df["Level"] = level
        df = df.rename(columns={sku_col: "SKU"})
        frames.append(df)
    combined = pd.concat(frames, ignore_index=True)
    combined["SKU"] = combined["SKU"].astype(str).str.strip()
    return combined


def assess_missing_activity_data(filled_collection: pd.DataFrame) -> pd.DataFrame:
    """One row per missing mandatory/activity field, sorted by number of
    impacted SKUs (priority list)."""
    rows = []
    fields_ = MANDATORY_COMPONENT_FIELDS + ACTIVITY_FIELDS
    for field_name in fields_:
        if field_name not in filled_collection.columns:
            continue
        missing = filled_collection[
            filled_collection[field_name].isna()
            | (filled_collection[field_name].astype(str).str.strip() == "")
        ]
        if missing.empty:
            continue
        sku_col = next(
            (c for c in ("SKU", "Product SKU", "Component SKU") if c in missing.columns),
            None,
        )
        rows.append(
            {
                "Field": field_name,
                "Missing count": len(missing),
                "Products impacted": missing[sku_col].nunique() if sku_col else len(missing),
                "Impact level": "HIGH" if field_name in MANDATORY_COMPONENT_FIELDS else "MEDIUM",
            }
        )
    report = pd.DataFrame(rows)
    if report.empty:
        return pd.DataFrame(
            columns=["Field", "Missing count", "Products impacted", "Impact level"]
        )
    return report.sort_values(["Products impacted", "Missing count"], ascending=False).reset_index(drop=True)


def activity_data_report(filled_collection: pd.DataFrame) -> pd.DataFrame:
    """'To complete' worksheet: one row per SKU with its list of missing fields."""
    fields_ = MANDATORY_COMPONENT_FIELDS + ACTIVITY_FIELDS

    def _missing_fields(row) -> str | None:
        missing = [
            f
            for f in fields_
            if f in filled_collection.columns
            and (pd.isna(row[f]) or str(row[f]).strip() == "")
        ]
        return " | ".join(missing) if missing else None

    report = filled_collection.copy()
    report["Missing fields"] = report.apply(_missing_fields, axis=1)
    report = report[report["Missing fields"].notna()]
    sku_col = next(
        (c for c in ("SKU", "Product SKU", "Component SKU") if c in report.columns),
        None,
    )
    level_col = "Level" if "Level" in report.columns else None
    keep = [c for c in [sku_col, level_col, "Missing fields"] if c] + [
        c for c in ["Product Designation", "Component Designation", "Supplier Name"] if c in report.columns
    ]
    return report[keep].reset_index(drop=True)
