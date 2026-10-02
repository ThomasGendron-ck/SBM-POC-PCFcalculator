"""Tests for the cache source-file fingerprint (staleness warning)."""
import pandas as pd

from sbm_pcf_calculation.cache import (
    compute_sources_fingerprint,
    load_session,
    save_session,
)
from sbm_pcf_calculation.sources import PcfSession


def _session(input_dir):
    session = PcfSession(metadata={"input_dir": str(input_dir)})
    session.products = pd.DataFrame({"SKU": ["A"]})
    return session


def test_no_warning_when_sources_unchanged(tmp_path, capsys):
    inp, wd = tmp_path / "input", tmp_path / "work"
    inp.mkdir()
    (inp / "src.xlsx").write_bytes(b"data")
    save_session(_session(inp), wd)
    load_session(wd)
    assert "Warning" not in capsys.readouterr().err


def test_warning_when_source_changed(tmp_path, capsys):
    inp, wd = tmp_path / "input", tmp_path / "work"
    inp.mkdir()
    src = inp / "src.xlsx"
    src.write_bytes(b"data")
    save_session(_session(inp), wd)
    src.write_bytes(b"changed")
    load_session(wd)
    err = capsys.readouterr().err
    assert "Warning" in err and "src.xlsx" in err and "--reload" in err


def test_warning_when_source_added(tmp_path, capsys):
    inp, wd = tmp_path / "input", tmp_path / "work"
    inp.mkdir()
    save_session(_session(inp), wd)
    (inp / "new.xlsx").write_bytes(b"data")
    load_session(wd)
    err = capsys.readouterr().err
    assert "Warning" in err and "new.xlsx" in err


def test_fingerprint_ignores_non_excel_files(tmp_path):
    inp = tmp_path / "input"
    inp.mkdir()
    (inp / "notes.txt").write_text("x")
    assert compute_sources_fingerprint(inp) == {}
