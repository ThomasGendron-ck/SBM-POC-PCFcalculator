"""Tests du matching ecoinvent : pays fournisseur, règles, produit principal, cascade géographique."""

import pandas as pd
import pytest

from sbm_pcf_calculation.ecoinvent import (
    _main_product,
    _pick_geo,
    match_missing_fe,
    pick_dataset,
    supplier_country,
)

pd = pytest.importorskip("pandas")


class TestSupplierCountry:
    def test_codes_fournisseurs(self):
        assert supplier_country("EFR06853") == "FR"
        assert supplier_country("XDE00017") == "DE"
        assert supplier_country("eit00001") == "IT"

    def test_codes_invalides(self):
        assert supplier_country(None) is None
        assert supplier_country("") is None
        assert supplier_country("FR") is None
        assert supplier_country("ABC") is None


def _base():
    return pd.DataFrame(
        [
            {"name": "market for zinc, primary", "geo": "GLO", "prod": "zinc", "unit": "kg", "amount": 1.0, "gwp_per_unit": 3.0},
            {"name": "zinc production", "geo": "RER", "prod": "zinc", "unit": "kg", "amount": 1.0, "gwp_per_unit": 2.5},
            {"name": "zinc production", "geo": "RoW", "prod": "zinc", "unit": "kg", "amount": 1.0, "gwp_per_unit": 2.9},
            {"name": "zinc production", "geo": "FR", "prod": "zinc", "unit": "kg", "amount": 1.0, "gwp_per_unit": 2.7},
            {"name": "ammonium sulfate production", "geo": "RoW", "prod": "zinc", "unit": "kg", "amount": 1.0, "gwp_per_unit": 0.28},
        ]
    )


class TestPickDataset:
    def test_cascade_pays_fournisseur(self):
        best = pick_dataset(_base(), [r"zinc production"], "FR")
        assert best["geo"] == "FR"
        assert best["gwp_per_unit"] == 2.7

    def test_cascade_rer_row_glo(self):
        assert pick_dataset(_base(), [r"zinc production"], "AD")["geo"] == "RER"
        assert pick_dataset(_base(), [r"zinc production"], None)["geo"] == "RER"
        base_no_rer = _base()[_base()["geo"] != "RER"]
        assert pick_dataset(base_no_rer, [r"zinc production"], None)["geo"] == "RoW"

    def test_produit_principal_non_coproduit(self):
        multi = pd.DataFrame(
            [
                {"name": "primary zinc production", "geo": "RoW", "prod": "zinc", "unit": "kg", "amount": 1.0, "gwp_per_unit": 2.9},
                {"name": "primary zinc production", "geo": "RoW", "prod": "ammonium sulfate", "unit": "kg", "amount": 1.0, "gwp_per_unit": 0.28},
            ]
        )
        best = _main_product(multi)
        assert best["prod"] == "zinc"
        assert best["gwp_per_unit"] == 2.9

    def test_dataset_market_prefere(self):
        pool = pd.DataFrame(
            [
                {"name": "1,3-dichloropropene to generic market for pesticide, unspecified", "geo": "GLO", "prod": "pesticide, unspecified", "unit": "kg", "amount": 1.0, "gwp_per_unit": 1.78},
                {"name": "market for pesticide, unspecified", "geo": "GLO", "prod": "pesticide, unspecified", "unit": "kg", "amount": 1.0, "gwp_per_unit": 10.35},
            ]
        )
        best = _main_product(pool)
        assert best["name"] == "market for pesticide, unspecified"


def _collecte():
    return pd.DataFrame(
        [
            {
                "Product SKU": "P1",
                "Component SKU": "Z1",
                "Component Designation": "SULFATE DE ZINC MONOHYDRATE",
                "Raw Material": "Matières Premières",
                "Supplier Name": "Fourni SA",
                "Supplier code": "EFR06853",
                "Category Code": "PDTMP",
                "Category description": "RAW MATERIAL",
                "RM EF Value": None,
            },
            {
                "Product SKU": "P1",
                "Component SKU": "W1",
                "Component Designation": "EAU DE VILLE",
                "Raw Material": "Matières Premières",
                "Supplier Name": "Autre",
                "Supplier code": "XIT00001",
                "Category Code": "PDTMP",
                "Category description": "RAW MATERIAL",
                "RM EF Value": None,
            },
            {
                "Product SKU": "P1",
                "Component SKU": "OK1",
                "Component Designation": "PE-HD GRANULE",
                "Raw Material": "EMBALLAGE",
                "Supplier Name": "Autre",
                "Supplier code": "EGB00001",
                "Category Code": "PDTEP",
                "Category description": "PACKAGING",
                "RM EF Value": 2.37,
            },
            {
                "Product SKU": "P1",
                "Component SKU": "N1",
                "Component Designation": "CHOSE INCONNUE XYZ",
                "Raw Material": "Matières Premières",
                "Supplier Name": "Inconnu",
                "Supplier code": None,
                "Category Code": "PDTMP",
                "Category description": "RAW MATERIAL",
                "RM EF Value": None,
            },
        ]
    )


class TestMatchMissingFe:
    def test_matching_complet(self):
        base = pd.DataFrame(
            [
                {"name": "market for tap water", "geo": "RoW", "prod": "tap water", "unit": "kg", "amount": 1.0, "gwp_per_unit": 0.0012},
                {"name": "market for tap water", "geo": "IT", "prod": "tap water", "unit": "kg", "amount": 1.0, "gwp_per_unit": 0.0013},
                {"name": "zinc production", "geo": "RER", "prod": "zinc", "unit": "kg", "amount": 1.0, "gwp_per_unit": 2.5},
                {"name": "zinc production", "geo": "RoW", "prod": "zinc", "unit": "kg", "amount": 1.0, "gwp_per_unit": 2.9},
                {"name": "primary zinc production", "geo": "RER", "prod": "zinc", "unit": "kg", "amount": 1.0, "gwp_per_unit": 2.5},
            ]
        )
        result = match_missing_fe(_collecte(), base)
        assert len(result) == 3

        zinc = result[result["Component SKU"] == "Z1"].iloc[0]
        assert zinc["Statut"] == "MATCHÉ"
        assert zinc["Dataset ecoinvent"] == "primary zinc production"
        assert zinc["Géographie"] == "RER"
        assert zinc["Pays fournisseur"] == "FR"
        assert zinc["FE proposé (kg CO2e/kg)"] == 2.5

        eau = result[result["Component SKU"] == "W1"].iloc[0]
        assert eau["Dataset ecoinvent"] == "market for tap water"
        assert eau["Géographie"] == "IT"

        inconnu = result[result["Component SKU"] == "N1"].iloc[0]
        assert inconnu["Statut"] == "NON MATCHÉ"

    def test_premiere_regle_gagne(self):
        base = pd.DataFrame(
            [{"name": "market for tap water", "geo": "RoW", "prod": "tap water", "unit": "kg", "amount": 1.0, "gwp_per_unit": 0.0012}]
        )
        result = match_missing_fe(_collecte(), base)
        zinc = result[result["Component SKU"] == "Z1"].iloc[0]
        assert zinc["Statut"] == "NON MATCHÉ"
