"""Path-independence tests.

The service code may be installed to %LOCALAPPDATA%\\OutlookAiAssistant\\
embedding_service, while the release ZIP folder is deleted afterwards. To make
that safe, all per-user data must live under %LOCALAPPDATA%\\OutlookAiAssistant
and never next to the code. These tests pin that contract.
"""

from __future__ import annotations

import os
from pathlib import Path

from app.config import ServiceSettings, default_product_root


def test_default_product_root_is_under_localappdata(monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", r"C:\Users\test\AppData\Local")
    root = default_product_root()
    assert str(root).lower().endswith("outlookaiassistant")
    assert "appdata" in str(root).lower()


def test_data_paths_never_under_code_dir(tmp_path):
    settings = ServiceSettings(product_root=tmp_path)
    code_dir = Path(__file__).resolve().parent.parent
    for path in (
        settings.db_path,
        settings.model_cache_dir,
        settings.log_path,
        settings.token_path,
    ):
        resolved = Path(path).resolve()
        assert code_dir not in resolved.parents, (
            "%s must not live inside the code directory" % resolved
        )
        assert str(resolved).startswith(str(Path(tmp_path).resolve()))


def test_all_data_shares_one_product_root(tmp_path):
    settings = ServiceSettings(product_root=tmp_path)
    root = Path(tmp_path).resolve()
    assert Path(settings.db_path).resolve().parent.parent == root
    assert Path(settings.model_cache_dir).resolve() == root / "models"
    assert Path(settings.token_path).resolve().parent == root


def test_env_override_relocates_everything(tmp_path, monkeypatch):
    """OAA_PRODUCT_ROOT must move db, model cache, log and token together."""

    monkeypatch.setenv("OAA_PRODUCT_ROOT", str(tmp_path / "custom"))
    settings = ServiceSettings.from_env()
    base = (tmp_path / "custom").resolve()
    assert Path(settings.db_path).resolve().parent.parent == base
    assert Path(settings.model_cache_dir).resolve().parent == base
    assert Path(settings.token_path).resolve().parent == base
