"""Regression test: ef-match must accept both raw collecte and pre-aggregated input.

The ef-match command passes identify_missing_factors() output (unique
components, no 'RM EF Value' column) into match_missing_fe, which used to
re-aggregate it and raise KeyError: 'RM EF Value'.
"""
import pandas as pd

from sbm_pcf_calculation.ecoinvent import match_missing_fe, missing_fe_components


def _collecte() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Component SKU": ["C1", "C2"],
            "Component Designation": ["DOSEUR PP", "BIDON HDPE"],
            "Category Code": ["PDTMA", "PDTMA"],
            "Category description": ["RAW MATERIALS", "RAW MATERIALS"],
            "Raw Material": ["PP", "HDPE"],
            "Supplier code": ["EFR06853", "EDE00017"],
            "Supplier Name": ["S1", "S2"],
            "RM EF Value": [None, None],
        }
    )


def _lcia_base() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "name": ["polypropylene production", "polyethylene production, high density"],
            "geo": ["RER", "RER"],
            "prod": ["pp", "hdpe"],
            "unit": ["kg", "kg"],
            "amount": [1.0, 1.0],
            "gwp_per_unit": [1.8, 2.1],
        }
    )


def test_match_from_raw_collecte():
    matching = match_missing_fe(_collecte(), _lcia_base())
    assert matching["Statut"].tolist() == ["MATCHÉ", "MATCHÉ"]


def test_match_from_pre_aggregated_missing_list():
    missing = missing_fe_components(_collecte())
    matching = match_missing_fe(missing, _lcia_base())
    assert matching["Statut"].tolist() == ["MATCHÉ", "MATCHÉ"]
    assert matching["Component SKU"].tolist() == ["C1", "C2"]


def test_pre_aggregated_gives_same_results_as_raw():
    raw = match_missing_fe(_collecte(), _lcia_base())
    pre = match_missing_fe(missing_fe_components(_collecte()), _lcia_base())
    pd.testing.assert_frame_equal(raw, pre)
