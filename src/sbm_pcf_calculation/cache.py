"""Session cache: save/reload pipeline data between steps without re-reading Excel.

Each table of the session is persisted as parquet in the work directory, so any
step can be re-run quickly during development (no full source reload).
"""
from __future__ import annotations

import pickle
import sys
from dataclasses import fields
from pathlib import Path

import pandas as pd

from .sources import PcfSession

SESSION_FILE = "session.pkl"
TABLES_DIR = "tables"
SOURCES_FINGERPRINT = "sources_fingerprint"


def compute_sources_fingerprint(input_dir: str | Path) -> dict[str, list[int]]:
    """Stamp of every source file in the input directory: [size, mtime_ns]."""
    input_dir = Path(input_dir)
    fingerprint = {}
    if not input_dir.is_dir():
        return fingerprint
    for path in sorted(input_dir.rglob("*.xlsx")):
        stat = path.stat()
        fingerprint[str(path)] = [stat.st_size, stat.st_mtime_ns]
    return fingerprint


def _check_sources_fingerprint(session: PcfSession) -> None:
    """Warn if an Excel source changed since the session was saved."""
    saved = session.metadata.get(SOURCES_FINGERPRINT)
    input_dir = session.metadata.get("input_dir")
    if not saved or not input_dir:
        return
    current = compute_sources_fingerprint(input_dir)
    changed = [path for path, stamp in saved.items() if current.get(path) != stamp]
    added = sorted(set(current) - set(saved))
    if changed or added:
        details = ", ".join(changed + added)
        print(f"Warning: source files changed since the session was saved: {details}. "
              "Consider --reload to re-read them.", file=sys.stderr)


def _tables_path(work_dir: Path) -> Path:
    path = work_dir / TABLES_DIR
    path.mkdir(parents=True, exist_ok=True)
    return path


def save_session(session: PcfSession, work_dir: str | Path) -> None:
    """Persist the session to the work directory (parquet per table)."""
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    tables = _tables_path(work_dir)
    for field_obj in fields(session):
        value = getattr(session, field_obj.name)
        if value is None:
            continue
        if isinstance(value, pd.DataFrame):
            value.to_parquet(tables / f"{field_obj.name}.parquet", index=False)
        else:
            with open(tables / f"{field_obj.name}.pkl", "wb") as fh:
                pickle.dump(value, fh)
    if isinstance(session.metadata, dict) and session.metadata.get("input_dir"):
        session.metadata[SOURCES_FINGERPRINT] = compute_sources_fingerprint(session.metadata["input_dir"])
    with open(work_dir / SESSION_FILE, "wb") as fh:
        pickle.dump({}, fh)


def load_session(work_dir: str | Path) -> PcfSession:
    """Reload the cached session from the work directory. Raises FileNotFoundError
    if the session has never been saved."""
    work_dir = Path(work_dir)
    if not (work_dir / SESSION_FILE).is_file():
        raise FileNotFoundError(
            f"No cached session in {work_dir}. Run the load step first "
            "(load_all_sources) or pass reload_sources=True."
        )
    tables = work_dir / TABLES_DIR
    session = PcfSession()
    for field_obj in fields(session):
        parquet_path = tables / f"{field_obj.name}.parquet"
        pickle_path = tables / f"{field_obj.name}.pkl"
        if parquet_path.is_file():
            setattr(session, field_obj.name, pd.read_parquet(parquet_path))
        elif pickle_path.is_file():
            with open(pickle_path, "rb") as fh:
                setattr(session, field_obj.name, pickle.load(fh))
    _check_sources_fingerprint(session)
    return session


def clear_session(work_dir: str | Path) -> None:
    """Delete every cached table and the session marker."""
    work_dir = Path(work_dir)
    session_file = work_dir / SESSION_FILE
    if session_file.is_file():
        session_file.unlink()
    tables = work_dir / TABLES_DIR
    if tables.is_dir():
        for path in tables.iterdir():
            path.unlink()
