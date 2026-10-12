"""Local semantic email search service for Outlook AI Assistant.

The service owns the embedding model and all SQLite storage. The C# Outlook
add-in never touches the database directly and never calls Outlook COM from
this process.

The version is read from the repository VERSION file when available so there
is one source of truth. When the service is installed standalone the file is
absent and the fallback below is used.
"""

from pathlib import Path

_FALLBACK_VERSION = "0.7.0"


def _read_version() -> str:
    """Read VERSION from the repo root, falling back to a constant.

    Layout when running from the repo:
        <repo>/VERSION
        <repo>/embedding_service/app/__init__.py   (this file)
    Layout when installed:
        <install_dir>/app/__init__.py
        <install_dir>/.installed-version           (written by the installer)
    """

    here = Path(__file__).resolve()
    candidates = [
        here.parent.parent.parent / "VERSION",       # repo root
        here.parent.parent / ".installed-version",   # install dir marker
    ]
    for candidate in candidates:
        try:
            if candidate.is_file():
                value = candidate.read_text(encoding="utf-8").strip()
                if value:
                    return value
        except OSError:
            continue
    return _FALLBACK_VERSION


__version__ = _read_version()
