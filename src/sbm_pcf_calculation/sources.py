"""Load every PCF calculation input from the SBM source files.

Source files (monthly-dated names tolerated, most recent file wins):
- Product list to compute (LM reference extract or spec "Produits LM" sheet)
- Product database (Masterbase products, full extract or ExtractPourPCF)
- Component database (CK_MaterialPurchase)
- Bill of materials (Masterbase BOM)
- Materials and emission factors (CK_EF_Packaging, carbon footprint study)
- Freight distances and emission factors (SBM Freight workbook)
- Optional ecoinvent LCIA base for EF matching
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from .collect import INPUT_PATTERNS, resolve_inputs
from .ecoinvent import load_lcia_gwp
from .io_sbm import load_sheet

HEADER_MB_PRODUCTS_FULL = 3


@dataclass
class PcfSession:
    """Holds all loaded inputs and step results, cached on disk between steps."""

    products: pd.DataFrame | None = None
    product_database: pd.DataFrame | None = None
    component_database: pd.DataFrame | None = None
    bom: pd.DataFrame | None = None
    materials_and_factors: pd.DataFrame | None = None
    freight_tables: dict[str, pd.DataFrame] | None = None
    freight_consolidated: pd.DataFrame | None = None
    lcia_base: pd.DataFrame | None = None
    filled_collection: pd.DataFrame | None = None
    ef_overrides: dict | None = None
    component_results: pd.DataFrame | None = None
    product_results: pd.DataFrame | None = None
    metadata: dict = field(default_factory=dict)


def load_product_list(sample_path: str | Path) -> pd.DataFrame:
    """Products to include in the calculation (SKU list, LM reference extract)."""
    from .io_sbm import load_sample_products

    return load_sample_products(sample_path)


def load_product_database(
    material_path: str | Path,
    full_product_extract: str | Path | None = None,
) -> pd.DataFrame:
    """Product master data. Uses the complete Masterbase products extract
    when provided, otherwise the MASTERBASE Products sheet of the material file."""
    if full_product_extract is not None:
        return load_sheet(full_product_extract, "MASTERBASE Products", HEADER_MB_PRODUCTS_FULL)
    return load_sheet(material_path, "MASTERBASE Products", 3)


def load_component_database(material_path: str | Path) -> pd.DataFrame:
    """Component master data: suppliers, EFs, recycled % (CK_MaterialPurchase)."""
    return load_sheet(material_path, "CK_MaterialPurchase", 12)


def load_bom(material_path: str | Path) -> pd.DataFrame:
    """Bill of materials linking products to components (alternatives, quantity)."""
    return load_sheet(material_path, "MASTERBASE BOM", 4)


def load_materials_and_factors(material_path: str | Path) -> pd.DataFrame:
    """Materials and their emission factors created during the carbon footprint
    study (CK_EF_Packaging + CK_MaterialPurchase EF columns)."""
    return load_sheet(material_path, "CK_EF_Packaging", 4)


def load_freight(freight_path: str | Path) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    """Upstream freight tables: distances, transport modes, emission factors."""
    from .collect import load_freight_tables

    return load_freight_tables(freight_path)


def load_lcia(lcia_path: str | Path) -> pd.DataFrame:
    """Emission factors per kg (GWP100 EF v3.1, ecoinvent cut-off)."""
    return load_lcia_gwp(str(lcia_path))


def load_all_sources(
    input_dir: str | Path,
    work_dir: str | Path,
    sample_path: str | Path | None = None,
    lcia_path: str | Path | None = None,
    full_product_extract: str | Path | None = None,
) -> PcfSession:
    """Load every input, resolve dated file patterns (latest wins) and save the
    session to the work directory. `sample_path` defaults to the resolved
    LM referencing file ("produits_lm" pattern)."""
    input_dir = Path(input_dir)
    resolved = resolve_inputs(input_dir)
    sample = Path(sample_path) if sample_path else resolved["produits_lm"]
    lcia = Path(lcia_path) if lcia_path else resolved.get("lcia")

    session = PcfSession(metadata={"input_dir": str(input_dir)})
    session.products = load_product_list(sample)
    session.product_database = load_product_database(resolved["material"], full_product_extract)
    session.component_database = load_component_database(resolved["material"])
    session.bom = load_bom(resolved["mb_bom"])
    session.materials_and_factors = load_materials_and_factors(resolved["material"])
    session.freight_tables, session.freight_consolidated = load_freight(resolved["freight"])
    if lcia is not None and Path(lcia).is_file():
        session.lcia_base = load_lcia(lcia)

    from .cache import save_session

    save_session(session, work_dir)
    return session


def resolve_patterns() -> dict[str, str]:
    """Expose the expected source file patterns (documentation helper)."""
    return dict(INPUT_PATTERNS)
