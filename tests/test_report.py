"""Tests unitaires sur des données synthétiques (aucune dépendance aux fichiers SBM réels)."""

import pandas as pd
import pytest

from sbm_pcf_calculation.report import priority_missing, stats_missing_factors, unique_components
from sbm_pcf_calculation.extract import _agg_freight


@pytest.fixture
def relations():
    return pd.DataFrame(
        {
            "ID Unique Produit": ["P1", "P1", "P2", "P2"],
            "Nom": ["Produit 1", "Produit 1", "Produit 2", "Produit 2"],
            "Super Segment": ["A", "A", "B", "B"],
            "ID Unique Composant": ["C1", "C2", "C1", "C3"],
            "Quantité de composant dans produit": [2, 1, 3, 1],
            "Nom composant": ["Comp 1", "Comp 2", "Comp 1", "Comp 3"],
            "Type composant": ["MP", "PDTBO", "MP", "PDTEC"],
            "Emission - Production (kgCO2e)": [1.0, None, 1.0, 2.0],
            "Emission - Usage (kgCO2e)": [None, None, None, None],
            "Emission - Fin de Vie (kgCO2e)": [0.5, None, 0.5, None],
            "Fret Amont - IRIS - Emission (kgCO2e)": [0.1, 0.2, 0.1, None],
        }
    )


def test_agg_freight_aggregates_by_article():
    df = pd.DataFrame(
        {
            "Référence Article": ["A", "A", "B"],
            "Weigh Final (in to)": [10.0, 5.0, 3.0],
            "GHG_perunit (kgCO2e/kg)": [0.1, 0.2, 0.5],
        }
    )
    out = _agg_freight(df, "Référence Article", "Weigh Final (in to)", "GHG_perunit (kgCO2e/kg)")
    a = out[out["Référence Article"] == "A"].iloc[0]
    assert a["poids_total"] == 15.0
    assert a["emission_totale"] == pytest.approx(10 * 0.1 + 5 * 0.2)
    assert a["nb_lignes_achat"] == 2


def test_stats_missing_counts(relations):
    stats = stats_missing_factors(relations)
    prod_row = stats[stats["Phase"] == "Production"].iloc[0]
    assert prod_row["Nb manquant"] == 1
    assert prod_row["Nb relations"] == 4
    iris_row = stats[stats["Phase"].str.contains("IRIS")].iloc[0]
    assert iris_row["Nb manquant"] == 1


def test_unique_components_counts_products(relations):
    comps = unique_components(relations)
    c1 = comps[comps["ID Unique Composant"] == "C1"].iloc[0]
    assert c1["Nb produits concernés"] == 2
    assert c1["Emission totale connue (kgCO2e)"] == pytest.approx(2 * (1.0 + 0.5 + 0.1))


def test_priority_missing_filters_complete(relations):
    prio = priority_missing(relations)
    assert (prio["Nb FE manquants"] > 0).all()
    top1 = prio.iloc[0]
    assert top1["Nb produits concernés"] == prio["Nb produits concernés"].max()
