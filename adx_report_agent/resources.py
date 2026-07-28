from __future__ import annotations

import sys
from pathlib import Path


def resource_root() -> Path:
    """Return project root in source mode or PyInstaller resource root in frozen mode."""

    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parents[1]


def resource_path(path: str | Path) -> Path:
    candidate = Path(path)
    if candidate.is_absolute() or candidate.exists():
        return candidate
    return resource_root() / candidate


def runtime_root() -> Path:
    """Writable runtime directory for .env and generated reports."""

    if getattr(sys, "frozen", False):
        return Path.home() / "ADXReportAgent"
    return Path.cwd()
