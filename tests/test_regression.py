"""Non-regression tests: PCF results must match the validated baseline.

The baseline (`baselines/pcf_results_baseline.csv`) stores a validated set of
results. When results change on purpose (new rule, new source data), regenerate
it with `save_baseline()` and review the diff before committing.
"""
from pathlib import Path

import pandas as pd
import pytest

from sbm_pcf_calculation.regression import (
    DEFAULT_BASELINE,
    check_regression,
    compare_with_baseline,
    load_baseline,
    save_baseline,
)


@pytest.fixture
def baseline_results() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Product SKU": "SORHOY15",
                "Product Designation": "Test product",
                "PCF Value": 0.693048,
                "DQR Product": 1.578581,
                "PDS Product": 0.0,
                "Flag level": None,
            },
            {
                "Product SKU": "PRODUCT2",
                "Product Designation": "Second product",
                "PCF Value": 1.234567,
                "DQR Product": 2.0,
                "PDS Product": 0.5,
                "Flag level": "HIGH",
            },
        ]
    )


class TestBaselineManagement:
    def test_save_and_load_round_trip(self, tmp_path, baseline_results):
        path = save_baseline(baseline_results, tmp_path / "baseline.csv")
        assert path.is_file()
        loaded = load_baseline(path)
        assert loaded["Product SKU"].tolist() == ["SORHOY15", "PRODUCT2"]

    def test_load_missing_baseline_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            load_baseline(tmp_path / "nowhere.csv")


class TestComparison:
    def test_identical_results_no_diff(self, baseline_results):
        report = compare_with_baseline(baseline_results.copy(), baseline_results.copy())
        assert report.empty

    def test_floating_point_noise_tolerated(self, baseline_results):
        current = baseline_results.copy()
        current.loc[current["Product SKU"] == "SORHOY15", "PCF Value"] = 0.693048 + 1e-12
        report = compare_with_baseline(current, baseline_results)
        assert report.empty

    def test_value_drift_detected(self, baseline_results):
        current = baseline_results.copy()
        current.loc[current["Product SKU"] == "SORHOY15", "PCF Value"] = 0.70
        report = compare_with_baseline(current, baseline_results)
        drift = report[(report["Type"] == "value drift") & (report["Product SKU"] == "SORHOY15")]
        assert not drift.empty
        assert drift.iloc[0]["Field"] == "PCF Value"

    def test_text_change_detected(self, baseline_results):
        current = baseline_results.copy()
        current.loc[current["Product SKU"] == "PRODUCT2", "Flag level"] = "LOW"
        report = compare_with_baseline(current, baseline_results)
        change = report[report["Type"] == "text change"]
        assert not change.empty

    def test_new_product_detected(self, baseline_results):
        current = pd.concat(
            [baseline_results, pd.DataFrame([{"Product SKU": "NEW1", "PCF Value": 0.1}])],
            ignore_index=True,
        )
        report = compare_with_baseline(current, baseline_results)
        assert not report[report["Type"] == "new product"].empty

    def test_missing_product_detected(self, baseline_results):
        current = baseline_results[baseline_results["Product SKU"] != "PRODUCT2"]
        report = compare_with_baseline(current, baseline_results)
        assert not report[report["Type"] == "missing product"].empty

    def test_nan_to_value_is_a_drift(self, baseline_results):
        current = baseline_results.copy()
        current.loc[current["Product SKU"] == "SORHOY15", "Flag level"] = "MEDIUM"
        report = compare_with_baseline(current, baseline_results)
        assert not report[(report["Product SKU"] == "SORHOY15") & (~report["Type"].isin(["new product", "missing product"]))].empty


class TestCommittedBaseline:
    def test_committed_baseline_is_valid_and_complete(self):
        """If a baseline is committed, it must be loadable and hold results.
        Regenerate with save_baseline() when the calculation rules change."""
        if not DEFAULT_BASELINE.is_file():
            pytest.skip("No committed baseline yet: create one with save_baseline()")
        baseline = load_baseline(DEFAULT_BASELINE)
        assert not baseline.empty
        assert "Product SKU" in baseline.columns
        assert "PCF Value" in baseline.columns

    def test_results_against_committed_baseline(self, baseline_results):
        """Template test: replace baseline_results by the current computation
        output to run the regression against the committed baseline."""
        if not DEFAULT_BASELINE.is_file():
            pytest.skip("No committed baseline yet")
        report = check_regression(baseline_results)
        assert report.empty, f"PCF results regressed:\n{report.to_string(index=False)}"
