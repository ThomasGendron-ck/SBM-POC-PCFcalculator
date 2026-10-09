"""CLI tests for the compute --check-baseline non-regression option."""
import pandas as pd
import pytest

from sbm_pcf_calculation import sbm_pcf_cli, regression
from sbm_pcf_calculation.regression import save_baseline


@pytest.fixture
def baseline():
    return pd.DataFrame(
        [
            {"Product SKU": "SORHOY15", "PCF Value": 0.693048, "DQR Product": 1.578581},
            {"Product SKU": "P2", "PCF Value": 1.5, "DQR Product": 2.0},
        ]
    )


@pytest.fixture
def fake_compute(monkeypatch):
    """Bypass session loading and PCF computation with canned results."""

    def _install(results):
        monkeypatch.setattr(
            sbm_pcf_cli,
            "_load_session",
            lambda args: _fake_session(results),
        )
        monkeypatch.setattr(
            "sbm_pcf_calculation.pcf_calc.run_pcf_calculation",
            lambda session, transformation_path=None, collection_path=None, reload_sources=False: session,
        )

    return _install


class TestCheckBaselineCli:
    def test_help_lists_check_baseline(self, capsys):
        with pytest.raises(SystemExit):
            sbm_pcf_cli.main(["compute", "--help"])
        out = capsys.readouterr().out
        assert "--check-baseline" in out

    def test_check_passes_on_identical_results(self, tmp_path, baseline, fake_compute):
        baseline_path = save_baseline(baseline, tmp_path / "baseline.csv")
        fake_compute(baseline)
        code = sbm_pcf_cli.main(
            ["--work-dir", str(tmp_path), "compute", "--check-baseline", str(baseline_path),
             "--output", str(tmp_path / "res.xlsx")]
        )
        assert code == 0

    def test_check_fails_on_drift(self, tmp_path, baseline, fake_compute, capsys):
        baseline_path = save_baseline(baseline, tmp_path / "baseline.csv")
        drifted = baseline.copy()
        drifted.loc[drifted["Product SKU"] == "P2", "PCF Value"] = 9.99
        fake_compute(drifted)
        code = sbm_pcf_cli.main(
            ["--work-dir", str(tmp_path), "compute", "--check-baseline", str(baseline_path),
             "--output", str(tmp_path / "res.xlsx")]
        )
        assert code == 2
        err = capsys.readouterr().err
        assert "REGRESSED" in err and "P2" in err

    def test_save_creates_default_baseline(self, tmp_path, baseline, fake_compute, monkeypatch):
        target = tmp_path / "pcf_results_baseline.csv"
        monkeypatch.setattr(regression, "DEFAULT_BASELINE", target)
        fake_compute(baseline)
        code = sbm_pcf_cli.main(
            ["--work-dir", str(tmp_path), "compute", "--check-baseline", "save",
             "--output", str(tmp_path / "res.xlsx")]
        )
        assert code == 0
        saved = pd.read_csv(target)
        assert list(saved["Product SKU"]) == ["SORHOY15", "P2"]


def _fake_session(results):
    from sbm_pcf_calculation.sources import PcfSession

    session = PcfSession(product_results=results)
    session.metadata["work_dir"] = "."
    return session


class TestCliErrorHandling:
    def test_missing_input_dir_reports_clean_error(self, tmp_path, capsys):
        code = sbm_pcf_cli.main(["load", "--input", str(tmp_path / "nope")])
        assert code == 1
        err = capsys.readouterr().err
        assert "Error:" in err and "Traceback" not in err

    def test_missing_source_file_reports_clean_error(self, tmp_path, capsys):
        (tmp_path / "input").mkdir()
        code = sbm_pcf_cli.main(["load", "--input", str(tmp_path / "input")])
        assert code == 1
        err = capsys.readouterr().err
        assert "Error:" in err and "Traceback" not in err

    def test_missing_session_reports_clean_error(self, tmp_path, capsys):
        code = sbm_pcf_cli.main(["--work-dir", str(tmp_path / "wd"), "compute"])
        assert code == 1
        err = capsys.readouterr().err
        assert "Error:" in err and "Traceback" not in err
