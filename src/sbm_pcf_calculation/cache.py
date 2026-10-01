"""Session cache: save/reload pipeline data between steps without re-reading Excel.

Each table of the session is persisted as parquet in the work directory, so any
step can be re-run quickly during development (no full source reload).
"""
from __future__ import annotations

import pickle
from dataclasses import fields
from pathlib import Path

import pandas as pd

from .sources import PcfSession

SESSION_FILE = "session.pkl"
TABLES_DIR = "tables"


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
