"""Unit tests for the SBM PCF calculation pipeline (synthetic data only)."""
import pandas as pd
import pytest

from sbm_pcf_calculation.activity import (
    activity_data_report,
    assess_missing_activity_data,
)
from sbm_pcf_calculation.cache import clear_session, load_session, save_session
from sbm_pcf_calculation.pdf_report import load_pcf_selection, render_report_html
from sbm_pcf_calculation.sources import PcfSession


@pytest.fixture
def filled_collection() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Product SKU": "P1",
                "Component SKU": "C1",
                "Supplier Code": "EFR01",
                "Supplier Name": "Fournisseur A",
                "Raw Material (MB Product)": "PE",
                "Net Weight": 0.5,
                "Scrap Rate": None,
                "Supplier PCF value": None,
            },
            {
                "Product SKU": "P2",
                "Component SKU": "C2",
                "Supplier Code": None,
                "Supplier Name": None,
                "Raw Material (MB Product)": None,
                "Net Weight": None,
                "Scrap Rate": 0.03,
                "Supplier PCF value": 1.2,
            },
        ]
    )


@pytest.fixture
def product_results() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Product SKU": "SORHOY15",
                "Product Designation": "Test product",
                "PCF Value": 0.693048,
                "DQR Product": 1.578581,
                "PDS Product": 0.0,
                "Raw Material GHG (kgCO2e)": 0.570,
                "Packaging GHG (kgCO2e)": 0.1165,
                "Transformation GHG (kgCO2e)": None,
                "Freight GHG (kgCO2e)": 0.0000149,
                "Quality flag": "PAS DE TRAJET FRET",
            }
        ]
    )


class TestActivity:
    def test_missing_fields_detected(self, filled_collection):
        missing = assess_missing_activity_data(filled_collection)
        assert set(missing["Field"]) >= {"Supplier Code", "Net Weight", "Supplier PCF value"}
        levels = dict(zip(missing["Field"], missing["Impact level"]))
        assert levels["Supplier Code"] == "HIGH"
        assert levels["Scrap Rate"] == "MEDIUM"

    def test_report_lists_missing_per_sku(self, filled_collection):
        report = activity_data_report(filled_collection)
        assert len(report) == 2
        assert "Supplier Code" in report.iloc[0]["Missing fields"]

    def test_sorted_by_impacted_products(self, filled_collection):
        missing = assess_missing_activity_data(filled_collection)
        assert missing["Products impacted"].is_monotonic_decreasing


class TestCache:
    def test_round_trip(self, tmp_path, filled_collection):
        session = PcfSession(products=filled_collection)
        save_session(session, tmp_path)
        loaded = load_session(tmp_path)
        assert loaded.products is not None
        assert len(loaded.products) == 2
        assert loaded.products.iloc[0]["Product SKU"] == "P1"

    def test_load_without_session_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            load_session(tmp_path)

    def test_clear(self, tmp_path, filled_collection):
        save_session(PcfSession(products=filled_collection), tmp_path)
        clear_session(tmp_path)
        with pytest.raises(FileNotFoundError):
            load_session(tmp_path)


class TestPdfReport:
    def test_render_html(self, product_results):
        html = render_report_html(product_results)
        assert "SORHOY15" in html
        assert "Raw Material" in html
        assert "82.2%" in html
        assert "PAS DE TRAJET FRET" in html

    def test_selection_filters_products(self, product_results):
        html = render_report_html(product_results, selected_skus=["OTHER"])
        assert "SORHOY15" not in html

    def test_load_pcf_selection_csv(self, tmp_path):
        pd.DataFrame({"SKU": ["A", "B", "A"]}).to_csv(tmp_path / "sel.csv", index=False)
        assert load_pcf_selection(tmp_path / "sel.csv") == ["A", "B"]


class TestSession:
    def test_default_session_is_empty(self):
        session = PcfSession()
        assert session.products is None
        assert session.product_results is None
        assert session.metadata == {}
