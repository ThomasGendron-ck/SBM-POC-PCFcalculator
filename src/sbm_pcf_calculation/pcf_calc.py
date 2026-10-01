"""PCF calculation per product with quality flags (HIGH / MEDIUM / LOW).

Session-aware layer over collect.py and report.py: builds the component lines
from the loaded session, computes GHG per phase, the PCF per product, the DQR
and PDS scores, and attaches suspicion flags.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from .sources import PcfSession

HIGH_FLAGS = ("PAS DE FE", "PAS DE BOM", "PRODUIT INTROUVABLE", "FOURNISSEUR INTROUVABLE")
MEDIUM_FLAGS = ("PAS DE TRAJET FRET", "POIDS > 120%", "PAS DE MATIERE")
LOW_FLAGS = ("QTE NULLE",)


def build_component_lines(session: PcfSession, transformation_path: str | Path | None = None) -> pd.DataFrame:
    """One line per (product, component): quantity, weights, EF, freight.

    Uses the full collecte pipeline (collect.build_collecte) on the session's
    input directory; applies EF overrides when present."""
    from .collect import build_collecte

    input_dir = session.metadata.get("input_dir")
    if input_dir is None:
        raise ValueError("No input_dir in session metadata: reload sources first.")
    return build_collecte(
        input_dir,
        transformation=transformation_path,
        fe_overrides=session.ef_overrides,
    )


def compute_raw_material_ghg(component_lines: pd.DataFrame) -> pd.DataFrame:
    """GHG of raw materials and packaging per component line (already carried
    by the collecte pipeline in 'RM GHG' columns)."""
    return component_lines


def compute_freight_ghg(component_lines: pd.DataFrame) -> pd.DataFrame:
    """Freight GHG per component line (Freight GHG column of the collecte)."""
    return component_lines


def compute_transformation_ghg(component_lines: pd.DataFrame) -> pd.DataFrame:
    """Transformation GHG per component line (Transformation GHG column)."""
    return component_lines


def compute_pcf_per_product(component_lines: pd.DataFrame) -> pd.DataFrame:
    """PCF = sum of GHG per product; weighted DQR; PDS per phase.
    Returns the per-product summary (one row per product)."""
    from .report import build_synthese

    return build_synthese(component_lines)


def flag_suspect_products(product_results: pd.DataFrame,
                          component_lines: pd.DataFrame | None = None) -> pd.DataFrame:
    """Attach a quality flag level (HIGH / MEDIUM / LOW) per product based on
    the validation flags raised on its component lines."""
    if component_lines is not None and "Flag" in component_lines.columns:
        flags_by_product = {}
        for _, row in component_lines.iterrows():
            flag = row["Flag"]
            if pd.isna(flag) or not str(flag).strip():
                continue
            flags_by_product.setdefault(row["Product SKU"], []).extend(str(flag).split(" | "))
        levels = []
        for sku in product_results["Product SKU"]:
            flags = flags_by_product.get(sku, [])
            level = None
            for high in HIGH_FLAGS:
                if any(high in f for f in flags):
                    level = "HIGH"
                    break
            if level is None:
                level = next(
                    (m for m in MEDIUM_FLAGS if any(m in f for f in flags)),
                    None,
                )
            if level is None:
                level = next((l for l in LOW_FLAGS if any(l in f for f in flags)), None)
            levels.append(level)
        product_results = product_results.copy()
        product_results["Flag level"] = levels
    return product_results


def run_pcf_calculation(session: PcfSession,
                        transformation_path: str | Path | None = None,
                        reload_sources: bool = False) -> PcfSession:
    """Full calculation flow: component lines -> per-product PCF -> flags.
    Saves the results back into the session."""
    component_lines = build_component_lines(session, transformation_path)
    product_results = compute_pcf_per_product(component_lines)
    product_results = flag_suspect_products(product_results, component_lines)
    session.component_results = component_lines
    session.product_results = product_results
    from .cache import save_session

    save_session(session, session.metadata.get("work_dir", "."))
    return session
