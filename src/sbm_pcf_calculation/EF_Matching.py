"""Missing emission factor matching against the ecoinvent database.

Thin session-aware layer over ecoinvent.py: identifies components without an
emission factor, proposes ecoinvent datasets and applies validated factors.
"""
from __future__ import annotations

import pandas as pd

from .ecoinvent import load_lcia_gwp, match_missing_fe
from .sources import PcfSession


def identify_missing_factors(session: PcfSession) -> pd.DataFrame:
    """Unique components with no emission factor, with supplier country.

    Requires the component results table to be present in the session
    (run the calculation step first, or provide component lines)."""
    if session.component_results is None:
        raise ValueError("No component results in session: run the calculation step first.")
    from .ecoinvent import missing_fe_components

    return missing_fe_components(session.component_results)


def match_factors_ecoinvent(missing: pd.DataFrame, lcia_base: pd.DataFrame) -> pd.DataFrame:
    """Propose one ecoinvent dataset per component (MATCH_RULES priority,
    geographic cascade country -> RER -> RoW -> GLO)."""
    return match_missing_fe(missing, lcia_base)


def build_factor_overrides(matching: pd.DataFrame) -> dict[str, dict]:
    """Convert validated (Statut == MATCHÉ / OUI) matching rows into
    calculation-ready EF overrides."""
    from .collect import build_fe_overrides

    return build_fe_overrides(matching)


def apply_validated_factors(session: PcfSession, validated: pd.DataFrame) -> PcfSession:
    """Inject validated emission factors into the session (ef_overrides)."""
    session.ef_overrides = build_factor_overrides(validated)
    return session


def run_ef_matching(session: PcfSession, lcia_path: str | None = None) -> pd.DataFrame:
    """Full matching flow: identify missing factors, load the LCIA base if the
    session does not hold it, and return the matching table."""
    if session.lcia_base is None:
        if lcia_path is None:
            raise ValueError("No LCIA base in session: provide lcia_path or reload sources.")
        session.lcia_base = load_lcia_gwp(lcia_path)
    missing = identify_missing_factors(session)
    return match_factors_ecoinvent(missing, session.lcia_base)
