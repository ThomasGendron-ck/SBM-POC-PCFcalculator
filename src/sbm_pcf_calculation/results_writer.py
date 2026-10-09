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

from .datashare import GROUPING
from .report import SYNTHESE_COLUMNS, build_synthese
from .results_spec import MANDATORY, MISSING_EF_COLUMNS, PCF_COLUMNS

GHOUPING_PATTERN = re.compile(r"Group between '?(.*?)'? and '?(.*?)'?\s*$")

# Bornes exactes des groupings de la spec v0.98 (Spec_CalculationFile,
# onglet PCF_Calculation, colonne J) : première et dernière colonne du groupe.
COLUMN_GROUPS_PCF: list[tuple[str, str]] = [
    ("Product Category Code", "Product Stock unit"),
    ("Component Category Code", "Component Pack unit box"),
    ("Quantity", "Recycled %"),
    ("Supplier PDS", "Supplier PCF external review"),
    ("RM EF Name", "RM DQR value"),
    ("Transformation Energy Name", "Transformation DQR value"),
    ("Freight Supplier Code", "Freight Transportation Mode"),
    ("Use EF Name", "Use DQR value"),
    ("EoL EF Name", "EoL DQR value"),
]

# Blocs de couleurs de l'onglet PCF_Calculation (conventions du Fichier de
# collecte : Produit vert foncé, Composants vert, Description du composant
# vert clair, Supplier PCF crème, Transformation pêche, Flag ambre, GROUPING
# gris ; phases de cycle de vie : palette de l'onglet Synthèse du rapport).
PCF_BLOCK_COLORS: dict[str, str] = {
    "Produit": "FF13501B",
    "Composants": "FF75A67C",
    "Description du composant": "FFA3C4A7",
    "Supplier PCF": "FFFBE3D6",
    "RM": "FFA6A6A6",
    "Transformation": "FFF6C6AD",
    "Freight": "FF6FC5E6",
    "Use": "FFF2AA84",
    "EoL": "FFC7C7C7",
    "Flag": "FFFFC000",
    GROUPING: "FFD9D9D9",
}
PCF_BLOCK_FONT_COLORS: dict[str, str] = {
    "Produit": "FFFFFFFF",
    GROUPING: "FF404040",
}


def _pcf_block(name: str) -> str:
    """Bloc de style d'une colonne PCF (couleur d'en-tête / GROUPING)."""
    if name in {
        "Product Details", "Component Details", "Component Material",
        "Supplier PCF", "RM GHG details", "Transformation Details",
        "Freight  Details", "Use GHG details", "EoL GHG details",
    }:
        return GROUPING
    if name in ("Data validation flag", "Data validation flag impact"):
        return "Flag"
    if name in ("Product SKU", "Product Designation", "PCF GHG Value",
               "PCF GHG Unit", "PDS Product", "DQR Product"):
        return "Produit"
    if name.startswith("Product "):
        return "Produit"
    if name.startswith("Component "):
        return "Composants"
    if name in ("Quantity", "Net Weight", "Net Weight Unit", "Gross Weight",
               "Gross Weight Unit", "Stock unit", "UVP description",
               "Raw Material (MB Product)", "Raw Material - Carbon Footprint",
               "Recycled %", "RM Data PDS value"):
        return "Description du composant"
    if name.startswith("Supplier "):
        return "Supplier PCF"
    if name.startswith("RM "):
        return "RM"
    if name.startswith("Transformation "):
        return "Transformation"
    if name.startswith("Freight "):
        return "Freight"
    if name.startswith("Use "):
        return "Use"
    if name.startswith("EoL "):
        return "EoL"
    return "Composants"


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
    "Prod_EF_Geography": "RM EF Geography",
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
    "design": "Component Designation",
    "sage": "Component Category Code",
    "cfdesc": "Component Category description",
    "carbon_cat": "Component Carbon category",
    "pack": "Component Pack unit box",
    "code_fournis": "Component Supplier Code",
    "fournis": "Component Supplier Name",
    "pays": "Component Supplier Country",
    "matiere": "Raw Material (MB Product)",
    "matiere_cf": "Raw Material - Carbon Footprint",
    "uvp": "UVP description",
    "recycle": "Recycled %",
}

MISSING_EF_FROM_MATCHING: dict[str, str] = {
    "Dataset ecoinvent": "RM AutoMatch EF Name",
    "Règle de matching": "RM AutoMatch EF Rationale",
    "FE proposé (kg CO2e/kg)": "RM AutoMatch EF Value",
    "Géographie": "RM AutoMatch EF Geography",
}


def _select_columns(df: pd.DataFrame, columns: list[tuple[str, str | None]]) -> pd.DataFrame:
    """Project df onto the spec columns in spec order, filling defaults.
    Built with a single concat to avoid DataFrame fragmentation (123+ columns)."""
    df = df.copy()
    parts = []
    for name, _fmt in columns:
        if name in df.columns:
            parts.append(df[[name]])
        elif name in DEFAULT_VALUES:
            parts.append(pd.DataFrame({name: [DEFAULT_VALUES[name]] * len(df)}, index=df.index))
        else:
            parts.append(pd.DataFrame({name: [None] * len(df)}, index=df.index))
    return pd.concat(parts, axis=1)


def _product_only_pcf_sheet(product_results: pd.DataFrame | None) -> pd.DataFrame:
    """PCF_Calculation sheet with product-level rows only (no component lines)."""
    df = product_results.copy() if product_results is not None else pd.DataFrame(index=[0])
    rename = {"PCF Value": "PCF GHG Value", "PCF Unit": "PCF GHG Unit"}
    df = df.rename(columns=rename)
    return _select_columns(df, PCF_COLUMNS)


def build_pcf_sheet(component_lines: pd.DataFrame,
                    product_results: pd.DataFrame | None = None) -> pd.DataFrame:
    """PCF_Calculation sheet: one row per (product, component) line."""
    df = component_lines.copy()
    df = df.rename(columns=PCF_COLUMN_MAP)
    if product_results is not None and "Flag level" in product_results.columns:
        flag_map = product_results.set_index("Product SKU")["Flag level"]
        df["Data validation flag"] = df["Product SKU"].map(flag_map)
        df["Data validation flag impact"] = df["Data validation flag"]
    return _select_columns(df, PCF_COLUMNS)


def build_missing_ef_sheet(component_lines: pd.DataFrame,
                           matching: pd.DataFrame | None = None) -> pd.DataFrame:
    """MissingEF_Matching sheet: one row per component without an EF."""
    no_fe = component_lines[
        component_lines["RM EF Value"].isna()
        & component_lines["Component SKU"].notna()
    ]
    if no_fe.empty:
        return pd.DataFrame(columns=[name for name, _ in MISSING_EF_COLUMNS])
    from .ecoinvent import missing_fe_components

    missing = missing_fe_components(component_lines)
    missing = missing.rename(columns=MISSING_EF_COLUMN_MAP)
    if "Component Supplier Country" not in missing.columns:
        missing["Component Supplier Country"] = None
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
    return _select_columns(missing, MISSING_EF_COLUMNS)


def _column_index(columns: list[str], name: str) -> int:
    try:
        return columns.index(name)
    except ValueError as exc:
        raise ValueError(
            f"Grouping column '{name}' not found in the output columns"
        ) from exc


HEADER_FONT_COLOR = "FFFFFF"
MANDATORY_FILL = "13501B"
PCF_TAB_FILL = "0B3041"
SYNTHESE_TAB_FILL = "FF1F497D"


def _style_sheet(ws, columns: list[tuple[str, str | None]],
                 sheet_columns: list[str], groups: list[tuple[str, str]],
                 block_of=_pcf_block,
                 block_colors: dict[str, str] | None = None,
                 block_font_colors: dict[str, str] | None = None) -> None:
    """Apply the v0.98 sheet layout: per-block header colors, vertical
    GROUPING columns, mandatory row, outline groups, freeze at C3."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    if block_colors is None:
        block_colors = PCF_BLOCK_COLORS
    if block_font_colors is None:
        block_font_colors = PCF_BLOCK_FONT_COLORS
    mandatory_fill = PatternFill("solid", fgColor=MANDATORY_FILL)
    grouping_fill = PatternFill("solid", fgColor=block_colors[GROUPING])

    data_end_row = max(ws.max_row, 3)
    grouping_blocks: list[int] = []

    # Row 1: exact "Mandatory field" value from the spec (column I).
    for idx, name in enumerate(sheet_columns, start=1):
        block = block_of(name)
        cell = ws.cell(row=1, column=idx, value=MANDATORY.get(name))
        if block == GROUPING:
            grouping_blocks.append(idx)
            cell.fill = grouping_fill
            cell.font = Font(bold=True, size=9, color=block_font_colors.get(GROUPING, "FF404040"))
        else:
            cell.fill = mandatory_fill
            cell.font = Font(bold=True, size=9, color=HEADER_FONT_COLOR)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    # Row 2: header, one fill color per block (Data Collection conventions).
    for idx, (name, _fmt) in enumerate(columns, start=1):
        block = block_of(name)
        cell = ws.cell(row=2, column=idx, value=name)
        cell.fill = PatternFill("solid", fgColor=block_colors[block])
        cell.font = Font(bold=True, color=block_font_colors.get(block, "FF000000"))
        if block == GROUPING:
            cell.alignment = Alignment(vertical="center", textRotation=90, wrap_text=True)
            ws.column_dimensions[get_column_letter(idx)].width = 3.43
        else:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30.0
    ws.row_dimensions[2].height = 60.0
    # GROUPING columns: fill grisé sur toute la hauteur des données.
    for idx in grouping_blocks:
        for row_idx in range(3, data_end_row + 1):
            ws.cell(row=row_idx, column=idx).fill = grouping_fill
    # Column groups (outline levels).
    group_level = 1
    for start, end in groups:
        try:
            i0 = _column_index(sheet_columns, start)
            i1 = _column_index(sheet_columns, end)
        except ValueError:
            continue
        for i in range(min(i0, i1), max(i0, i1) + 1):
            dim = ws.column_dimensions[get_column_letter(i + 1)]
            dim.outline_level = group_level
            dim.hidden = False
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False
    # Freeze at C3: two header rows + two identifying columns.
    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(columns))}{data_end_row}"
    # Largeurs de colonnes ajustees au contenu (en-tete + donnees, echantillon 200).
    for idx, name in enumerate(sheet_columns, start=1):
        if block_of(name) == GROUPING:
            continue
        longest = len(name)
        for row_idx in range(3, data_end_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is not None and not str(value).startswith("="):
                longest = max(longest, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 2, 10), 45)


def _missing_ef_block(name: str) -> str:
    """Bloc de style d'une colonne MissingEF_Matching (v0.98)."""
    if name in {
        "Component Details", "Component Material", "AutoMatch Details",
        "Transfo AutoMatch Details", "UserValidation", "Transfo UserValidation Details",
    }:
        return GROUPING
    if name in ("Data validation flag", "Data validation flag impact"):
        return "Flag"
    if name == "Component Supplier Country":
        return "Composants"
    if name.startswith("Component "):
        return "Composants"
    if name.startswith("Transfo AutoMatch"):
        if name.endswith("DQR value"):
            return "Calculé"
        return "Transfo AutoMatch"
    if name.startswith("Transfo UserValidation"):
        if name.endswith("DQR value"):
            return "Calculé"
        return "Transfo UserValidation"
    if name.startswith("UserValidation"):
        if name.endswith("DQR value"):
            return "Calculé"
        return "UserValidation"
    if name.startswith("RM AutoMatch"):
        if name.endswith("DQR value"):
            return "Calculé"
        return "AutoMatch"
    return "Description du composant"


def build_synthese_sheet(component_lines: pd.DataFrame) -> pd.DataFrame:
    """PCF_Synthese sheet: one row per product with the total PCF/DQR/PDS
    and the breakdown by lifecycle phase (no component detail)."""
    from .report import build_synthese as _build

    if component_lines is None or component_lines.empty:
        return pd.DataFrame(columns=SYNTHESE_COLUMNS)
    sheet = _build(component_lines)
    return sheet.reindex(columns=SYNTHESE_COLUMNS)


def _style_synthese_sheet(ws, synthese: pd.DataFrame) -> None:
    """Style de l'onglet PCF_Synthese : en-têtes de bloc produit en bleu,
    une couleur par phase du cycle de vie, formats nombre, groupings et
    volets figés (mêmes conventions que l'onglet Synthèse du rapport)."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="FF156082")
    phase_colors = {
        "Raw Material": "FFA3C4A7",
        "Packaging": "FFD1E1D3",
        "Transformation": "FFD1E1D3",
        "Freight": "FF6FC5E6",
        "Use": "FFF2AA84",
        "End of Life": "FFC7C7C7",
    }
    data_end_row = max(ws.max_row, 2)
    for col_idx, col in enumerate(SYNTHESE_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col)
        phase = col.split(" - ")[0] if " - " in col else None
        if col in ("Product SKU", "Product Designation", "PCF Value", "DQR", "PDS"):
            cell.fill = header_fill
            cell.font = Font(color="FFFFFFFF", bold=True)
        elif phase in phase_colors:
            cell.fill = PatternFill("solid", fgColor=phase_colors[phase])
            cell.font = Font(color="FF000000", bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        number_format = None
        if col.endswith("% PCF total"):
            number_format = "0.00%"
        elif col not in ("Product SKU", "Product Designation"):
            number_format = "#,##0.000"
        if number_format:
            for row_idx in range(2, data_end_row + 1):
                ws.cell(row=row_idx, column=col_idx).number_format = number_format
    for phase in ("Raw Material", "Packaging", "Transformation", "Freight", "Use", "End of Life"):
        first = f"{phase} - PDS Activity Data"
        last = f"{phase} - DQR"
        if first in SYNTHESE_COLUMNS and last in SYNTHESE_COLUMNS:
            for col_idx in range(
                SYNTHESE_COLUMNS.index(first) + 1,
                SYNTHESE_COLUMNS.index(last) + 2,
            ):
                ws.column_dimensions[get_column_letter(col_idx)].outlineLevel = 1
    ws.sheet_properties.outlinePr.summaryRight = True
    ws.row_dimensions[1].height = 60.0
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(SYNTHESE_COLUMNS))}{data_end_row}"
    ws.sheet_properties.tabColor = SYNTHESE_TAB_FILL
    # Largeurs ajustees au contenu (en-tete + donnees, echantillon 200).
    for col_idx, col in enumerate(SYNTHESE_COLUMNS, start=1):
        widths = [min(len(col), 18)]
        for value in synthese[col].head(200) if col in synthese.columns else []:
            if value is None or (isinstance(value, float) and pd.isna(value)):
                continue
            widths.append(len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(widths) + 2, 40)


def write_pcf_results(output_path: str | Path,
                      component_lines: pd.DataFrame,
                      product_results: pd.DataFrame | None = None,
                      matching: pd.DataFrame | None = None) -> Path:
    """Write the pcf_results.xlsx deliverable
    (PCF_Synthese + PCF_Calculation + MissingEF_Matching)."""
    from openpyxl import Workbook

    pcf_sheet = build_pcf_sheet(component_lines, product_results) if component_lines is not None else _product_only_pcf_sheet(product_results)
    missing_sheet = build_missing_ef_sheet(component_lines, matching) if component_lines is not None else pd.DataFrame(columns=[name for name, _ in MISSING_EF_COLUMNS])
    synthese_sheet = (
        build_synthese_sheet(component_lines)
        if component_lines is not None
        else pd.DataFrame(columns=SYNTHESE_COLUMNS)
    )

    from .ef_matching_writer import BLOCK_COLORS, BLOCK_FONT_COLORS

    output_path = Path(output_path)
    workbook = Workbook()
    workbook.remove(workbook.active)
    synthese_ws = workbook.create_sheet("PCF_Synthese")
    _write_sheet_rows(synthese_ws, synthese_sheet, [(c, None) for c in SYNTHESE_COLUMNS], start_row=2)
    _style_synthese_sheet(synthese_ws, synthese_sheet)
    for sheet_name, df, columns, groups in (
        ("PCF_Calculation", pcf_sheet, PCF_COLUMNS, COLUMN_GROUPS_PCF),
        ("MissingEF_Matching", missing_sheet, MISSING_EF_COLUMNS, COLUMN_GROUPS_MISSING_EF),
    ):
        ws = workbook.create_sheet(sheet_name)
        _write_sheet_rows(ws, df, columns)
        if sheet_name == "PCF_Calculation":
            _style_sheet(ws, columns, [name for name, _ in columns], groups)
        else:
            _style_sheet(
                ws, columns, [name for name, _ in columns], groups,
                block_of=_missing_ef_block,
                block_colors=BLOCK_COLORS,
                block_font_colors=BLOCK_FONT_COLORS,
            )
    workbook["PCF_Calculation"].sheet_properties.tabColor = PCF_TAB_FILL
    workbook.save(output_path)
    return output_path


def write_ef_matching(output_path: str | Path,
                      component_lines: pd.DataFrame,
                      matching: pd.DataFrame) -> Path:
    """Write the ef_matching.xlsx deliverable (MissingEF_Matching sheet,
    same layout as the spec sheet of pcf_results.xlsx)."""
    from openpyxl import Workbook

    sheet = build_missing_ef_sheet(component_lines, matching)
    output_path = Path(output_path)
    workbook = Workbook()
    ws = workbook.active
    ws.title = "MissingEF_Matching"
    _write_sheet_rows(ws, sheet, MISSING_EF_COLUMNS)
    from .ef_matching_writer import BLOCK_COLORS, BLOCK_FONT_COLORS
    _style_sheet(ws, MISSING_EF_COLUMNS,
                 [name for name, _ in MISSING_EF_COLUMNS], COLUMN_GROUPS_MISSING_EF,
                 block_of=_missing_ef_block,
                 block_colors=BLOCK_COLORS,
                 block_font_colors=BLOCK_FONT_COLORS)
    workbook.save(output_path)
    return output_path


def _write_sheet_rows(ws, df: pd.DataFrame,
                     columns: list[tuple[str, str | None]],
                     start_row: int = 3) -> None:
    """Write the data rows (from start_row) with per-column number formats."""
    values = df.values.tolist()
    for row_idx, row in enumerate(values, start=start_row):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    for col_idx, (_name, fmt) in enumerate(columns, start=1):
        if not fmt:
            continue
        for row_idx in range(start_row, start_row + len(values)):
            ws.cell(row=row_idx, column=col_idx).number_format = fmt
