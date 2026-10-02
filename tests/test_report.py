"""Tests unitaires sur des données synthétiques (aucune dépendance aux fichiers SBM réels)."""

import pandas as pd
import pytest

from sbm_pcf_calculation.collect import agg_freight


def test_agg_freight_aggregates_by_article():
    df = pd.DataFrame(
        {
            "Référence Article": ["A", "A", "B"],
            "Weigh Final (in to)": [10.0, 5.0, 3.0],
            "GHG_perunit (kgCO2e/kg)": [0.1, 0.2, 0.5],
        }
    )
    out = agg_freight(df, "Référence Article", "Weigh Final (in to)", "GHG_perunit (kgCO2e/kg)")
    a = out[out["Référence Article"] == "A"].iloc[0]
    assert a["poids_total"] == 15.0
    assert a["emission_totale"] == pytest.approx(10 * 0.1 + 5 * 0.2)
    assert a["nb_lignes_achat"] == 2
