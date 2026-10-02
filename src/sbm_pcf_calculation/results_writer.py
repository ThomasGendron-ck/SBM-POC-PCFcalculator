"""Writer of the pcf_results.xlsx deliverable (spec v0.97).

Builds the two output sheets from the session tables:
- PCF_Calculation: one line per (product, component) with the product
  PCF/PDS/DQR block, the component details and the phase blocks
  (RM, Transformation, Freight, Use, EoL) plus the validation flags.
- MissingEF_Matching: one line per component without an emission factor,
  with the ecoinvent automatch proposal and the empty user validation
  block (mandatory fields to fill).

Both sheets follow the v0.97 specification: column order, number formats,
column grouping, a mandatory-fields row above the header and frozen panes
at C3.
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from .results_spec import MANDATORY_FIELDS, MISSING_EF_COLUMNS, PCF_COLUMNS

GHOUPING_PATTERN = re.compile(r"Group between '?(.*?)'? and '?(.*?)'?\s*$")

COLUMN_GROUPS_PCF: list[tuple[str, str]] = [
    ("Product Category Code", "Product Stock unit"),
    ("Component Category Code", "Component Stock unit"),
    ("Quantity", "Recycled %"),
    ("Supplier PDS", "Supplier PCF external review"),
    ("RM EF Name", "RM DQR value"),
    ("Transformation Energy Name", "Transformation DQR value"),
    ("Freight Supplier Code", "Freight DQR value"),
    ("Use EF Name", "Use DQR value"),
    ("EoL EF Name", "EoL DQR value"),
]

COLUMN_GROUPS_MISSING_EF: list[tuple[str, str]] = [
    ("Component Category Code", "Component Pack unit box"),
    ("UVP description", "Recycled %"),
    ("RM AutoMatch EF Name", "RM AutoMatch DQR value"),
    ("Transfo AutoMatch Process Name", "Transfo AutoMatch PDS value"),
    ("UserValidation EF Yes/No", "UserValidation TEMP DQR"),
    ("Transfo UserValidation Process Name", "Transfo UserValidation PDS value"),
]

# Column renames: collecte column -> spec column name.
PCF_COLUMN_MAP: dict[str, str] = {
    "PCF Value": "PCF GHG Value",
    "PCF Unit": "PCF GHG Unit",
    "Pack Unit Box": "Product Pack Unit Box",
    "Category Code": "Component Category Code",
    "Category description": "Component Category description",
    "Pack unit box": "Component Pack unit box",
    "Supplier code": "Component Supplier Code",
    "Supplier Name": "Component Supplier Name",
    "Stock unit": "Component Stock unit",
    "Net Weight Unit": "Component Net Weight Unit",
    "Gross Weight Unit": "Component Gross Weight Unit",
    "Product Net Weight Unit": "Product Net Weight Unit",
    "Product Gross Weight Unit": "Product Gross Weight Unit",
    "Product Stock unit": "Product Stock unit",
    "Carbon category": "Component Carbon category",
    "RM GHG": "RM GHG Value",
    "Transformation GHG": "Transformation GHG Value",
    "Freight GHG": "Freight GHG Value",
}

DEFAULT_VALUES: dict[str, object] = {
    "PCF GHG Unit": "kgCO2e/component",
    "Product Net Weight Unit": "kg",
    "Product Gross Weight Unit": "kg",
    "Net Weight Unit": "kg",
    "Gross Weight Unit": "kg",
    "Component Net Weight Unit": "kg",
    "Component Gross Weight Unit": "kg",
    "RM GHG Unit": "kgCO2e/component",
    "RM EF PDS": 0,
    "RM Data PDS value": 1,
    "Transformation GHG Unit": "kgCO2e/component",
    "Freight GHG Unit": "kgCO2e/component",
    "Freight PDS value": 0,
    "Use GHG Unit": "kgCO2e/component",
    "EoL GHG Unit": "kgCO2e/component",
    "Use EF PDS": 0,
    "Use PDS value": 0,
    "EoL EF PDS": 0,
    "EoL PDS value": 0,
    "Supplier PCF external review": "No",
    "Supplier PCF source": "Supplier",
}

MISSING_EF_COLUMN_MAP: dict[str, str] = {
    "Category Code": "Component Category Code",
    "Category description": "Component Category description",
    "Carbon category": "Component Carbon category",
    "Pack unit box": "Component Pack unit box",
    "Supplier code": "Component Supplier Code",
    "Supplier Name": "Component Supplier Name",
    "pays": "Component Supplier Country",
    "RM AutoMatch EF Unit": "RM AutoMatch EF Unit",
}

MISSING_EF_FROM_MATCHING: dict[str, str] = {
    "Dataset ecoinvent": "RM AutoMatch EF Name",
    "Règle de matching": "RM AutoMatch EF Rationale",
    "FE proposé (kg CO2e/kg)": "RM AutoMatch EF Value",
    "Géographie": "RM AutoMatch EF Geography",
}


def _product_only_pcf_sheet(product_results: pd.DataFrame | None) -> pd.DataFrame:
    """PCF_Calculation sheet with product-level rows only (no component lines)."""
    df = product_results.copy() if product_results is not None else pd.DataFrame()
    rename = {"PCF Value": "PCF GHG Value", "PCF Unit": "PCF GHG Unit"}
    df = df.rename(columns=rename)
    out = pd.DataFrame(index=df.index)
    for name, _fmt in PCF_COLUMNS:
        if name in df.columns:
            out[name] = df[name]
        elif name in DEFAULT_VALUES:
            out[name] = DEFAULT_VALUES[name]
        else:
            out[name] = None
    return out[[name for name, _ in PCF_COLUMNS]]


def build_pcf_sheet(component_lines: pd.DataFrame,
                    product_results: pd.DataFrame | None = None) -> pd.DataFrame:
    """PCF_Calculation sheet: one row per (product, component) line."""
    df = component_lines.copy()
    df = df.rename(columns=PCF_COLUMN_MAP)
    out = pd.DataFrame(index=df.index)
    for name, _fmt in PCF_COLUMNS:
        if name in df.columns:
            out[name] = df[name]
        elif name in DEFAULT_VALUES:
            out[name] = DEFAULT_VALUES[name]
        else:
            out[name] = None
    if product_results is not None and "Flag level" in product_results.columns:
        flag_map = product_results.set_index("Product SKU")["Flag level"]
        out["Data validation flag"] = out["Product SKU"].map(flag_map)
        out["Data validation flag impact"] = out["Data validation flag"]
    return out[[name for name, _ in PCF_COLUMNS]]


def build_missing_ef_sheet(component_lines: pd.DataFrame,
                           matching: pd.DataFrame | None = None) -> pd.DataFrame:
    """MissingEF_Matching sheet: one row per component without an EF."""
    no_fe = component_lines[
        component_lines["RM EF Value"].isna()
        & component_lines["Component SKU"].notna()
    ].copy()
    if no_fe.empty:
        return pd.DataFrame(columns=[name for name, _ in MISSING_EF_COLUMNS])
    from .ecoinvent import missing_fe_components

    missing = missing_fe_components(component_lines)
    missing = missing.rename(columns=MISSING_EF_COLUMN_MAP)
    if matching is not None and not matching.empty:
        match_cols = matching.rename(columns=MISSING_EF_FROM_MATCHING)
        by_sku = match_cols.set_index("Component SKU")
        for dst in MISSING_EF_FROM_MATCHING.values():
            if dst in by_sku.columns:
                missing[dst] = missing["Component SKU"].map(by_sku[dst])
        if "Statut" in match_cols.columns:
            matched_skus = set(
                match_cols.loc[match_cols["Statut"] == "MATCHÉ", "Component SKU"]
            )
            missing["RM AutoMatch EF PDS"] = [
                0 if sku in matched_skus else None for sku in missing["Component SKU"]
            ]
    out = pd.DataFrame(index=missing.index)
    for name, _fmt in MISSING_EF_COLUMNS:
        if name in missing.columns:
            out[name] = missing[name]
        elif name in DEFAULT_VALUES:
            out[name] = DEFAULT_VALUES[name]
        else:
            out[name] = None
    return out[[name for name, _ in MISSING_EF_COLUMNS]]


def _column_index(columns: list[str], name: str) -> int:
    try:
        return columns.index(name)
    except ValueError as exc:
        raise ValueError(
            f"Grouping column '{name}' not found in the output columns"
        ) from exc


def _style_sheet(ws, columns: list[tuple[str, str | None]],
                 sheet_columns: list[str], groups: list[tuple[str, str]],
                 group_prefix: str = "SBM_PCF") -> None:
    """Apply the v0.97 sheet layout: mandatory row, header, groups, freeze."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    n_cols = len(sheet_columns)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    mandatory_fill = PatternFill("solid", fgColor="FFF2CC")

    # Row 1: mandatory marker for fields the user must fill.
    for idx, name in enumerate(sheet_columns, start=1):
        cell = ws.cell(row=1, column=idx,
                       value="Mandatory" if name in MANDATORY_FIELDS else None)
        cell.fill = mandatory_fill
        cell.font = Font(bold=True, size=9)
        cell.alignment = Alignment(horizontal="center")
    # Row 2: header.
    for idx, (name, fmt) in enumerate(columns, start=1):
        cell = ws.cell(row=2, column=idx, value=name)
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        if fmt:
            for row in range(3, ws.max_row + 1):
                ws.cell(row=row, column=idx).number_format = fmt
    ws.row_dimensions[2].height = 45.0
    # Column groups (outline levels).
    group_level = 1
    for start, end in groups:
        try:
            i0 = _column_index(sheet_columns, start)
            i1 = _column_index(sheet_columns, end)
        except ValueError:
            continue
        for i in range(min(i0, i1), max(i0, i1) + 1):
            ws.column_dimensions[get_column_letter(i + 1)].outline_level = group_level
        if i1 < i0:
            for i in range(max(i0, i1), min(i0, i1) - 1, -1):
                ws.column_dimensions[get_column_letter(i + 1)].outline_level = group_level
    ws.sheet_properties.outlinePr.summaryBelow = False
    # Freeze at C3: two header rows + two identifying columns.
    ws.freeze_panes = "C3"


def write_pcf_results(output_path: str | Path,
                      component_lines: pd.DataFrame,
                      product_results: pd.DataFrame | None = None,
                      matching: pd.DataFrame | None = None) -> Path:
    """Write the pcf_results.xlsx deliverable (PCF_Calculation + MissingEF_Matching)."""
    from openpyxl import load_workbook

    pcf_sheet = build_pcf_sheet(component_lines, product_results) if component_lines is not None else _product_only_pcf_sheet(product_results)
    missing_sheet = build_missing_ef_sheet(component_lines, matching) if component_lines is not None else pd.DataFrame(columns=[name for name, _ in MISSING_EF_COLUMNS])

    output_path = Path(output_path)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pcf_sheet.to_excel(writer, sheet_name="PCF_Calculation", index=False)
        missing_sheet.to_excel(writer, sheet_name="MissingEF_Matching", index=False)

    workbook = load_workbook(output_path)
    _style_sheet(workbook["PCF_Calculation"], PCF_COLUMNS,
                 [name for name, _ in PCF_COLUMNS], COLUMN_GROUPS_PCF)
    _style_sheet(workbook["MissingEF_Matching"], MISSING_EF_COLUMNS,
                 [name for name, _ in MISSING_EF_COLUMNS], COLUMN_GROUPS_MISSING_EF)
    workbook.save(output_path)
    return output_path
