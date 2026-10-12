"""API-level tests using FastAPI's TestClient.

Covers authentication, indexing, search, deletion and the health contract
that the C# add-in depends on.
"""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app.api import create_app
from app.config import ServiceSettings
from app.embedding_model import EmbeddingModel

from conftest import FakeEmbeddingModel


@pytest.fixture()
def client(tmp_path, monkeypatch):
    settings = ServiceSettings(
        product_root=tmp_path,
        db_path=tmp_path / "mail.db",
        model_cache_dir=tmp_path / "models",
        log_path=tmp_path / "logs" / "svc.log",
        token_path=tmp_path / "token",
    )
    app = create_app(settings, token="secret-token")

    # Replace the real model with the deterministic fake.
    fake = FakeEmbeddingModel()
    app.state.service.model = fake
    from app.indexing import Indexer
    from app.search import HybridSearch

    app.state.service.indexer = Indexer(
        app.state.service.database, fake, settings
    )
    app.state.service.search = HybridSearch(
        app.state.service.database, fake, settings
    )
    with TestClient(app) as test_client:
        yield test_client


def auth():
    return {"Authorization": "Bearer secret-token"}


def mail_payload(**kw):
    base = dict(
        store_id="S1",
        entry_id="E1",
        subject="Curtain wall",
        internet_message_id="m1",
        sender_email="client@example.com",
        body_text="can we edit the panels individually",
        received_at="2026-06-15T00:00:00Z",
        folder_path="Inbox",
    )
    base.update(kw)
    return base


def test_health_is_open(client):
    response = client.get("/health")
    assert response.status_code == 200
    body = response.json()
    assert body["status"] == "ok"
    assert "index" in body
    assert body["embedding_dim"] == 768


def test_search_requires_token(client):
    response = client.post("/search", json={"query": "x"})
    assert response.status_code == 401


def test_index_requires_token(client):
    response = client.post("/index/mail", json=mail_payload())
    assert response.status_code == 401


def test_wrong_token_rejected(client):
    response = client.post(
        "/search",
        json={"query": "x"},
        headers={"Authorization": "Bearer wrong"},
    )
    assert response.status_code == 401


def test_index_and_search_roundtrip(client):
    response = client.post("/index/mail", json=mail_payload(), headers=auth())
    assert response.status_code == 200
    assert response.json()["embedded"] is True

    response = client.post(
        "/search",
        json={"query": "幕墙 Panel 单独编辑"},
        headers=auth(),
    )
    assert response.status_code == 200
    body = response.json()
    assert body["results"]
    assert body["results"][0]["subject"] == "Curtain wall"
    assert body["results"][0]["entry_id"] == "E1"
    assert body["semantic_available"] is True


def test_index_requires_entry_or_store(client):
    response = client.post(
        "/index/mail",
        json=mail_payload(store_id="", entry_id=""),
        headers=auth(),
    )
    assert response.status_code == 400


def test_delete_endpoint(client):
    client.post("/index/mail", json=mail_payload(), headers=auth())
    response = client.post(
        "/index/delete",
        json={"store_id": "S1", "entry_id": "E1"},
        headers=auth(),
    )
    assert response.status_code == 200
    assert response.json()["deleted"] is True
    stats = client.get("/stats", headers=auth()).json()
    assert stats["mail_count"] == 0


def test_clear_index_endpoint(client):
    client.post("/index/mail", json=mail_payload(), headers=auth())
    response = client.delete("/index", headers=auth())
    assert response.status_code == 200
    stats = client.get("/stats", headers=auth()).json()
    assert stats["mail_count"] == 0
    assert stats["chunk_count"] == 0


def test_meta_endpoint(client):
    response = client.post(
        "/index/meta",
        json={"last_sync_at": "2026-06-15T10:00:00Z"},
        headers=auth(),
    )
    assert response.status_code == 200
    stats = client.get("/stats", headers=auth()).json()
    assert stats["last_sync_at"] == "2026-06-15T10:00:00Z"


def test_empty_query_rejected(client):
    response = client.post("/search", json={"query": "   "}, headers=auth())
    assert response.status_code == 400


def test_model_status_endpoint(client):
    response = client.get("/model/status", headers=auth())
    assert response.status_code == 200
    body = response.json()
    assert "available" in body
    assert "model_id" in body


def test_model_status_requires_token(client):
    assert client.get("/model/status").status_code == 401


def test_index_returns_503_when_model_unavailable(client):
    """A missing model must be reported as 503, not a silent success."""

    from app.embedding_model import EmbeddingModelUnavailable

    class BrokenModel:
        available = False
        load_error = "not downloaded"

        def load(self):
            return False

        def encode_documents(self, *a, **k):
            raise EmbeddingModelUnavailable("not downloaded")

        def encode_queries(self, *a, **k):
            raise EmbeddingModelUnavailable("not downloaded")

    app = client.app
    app.state.service.model = BrokenModel()
    from app.indexing import Indexer

    app.state.service.indexer = Indexer(
        app.state.service.database, BrokenModel(), app.state.service.settings
    )
    response = client.post("/index/mail", json=mail_payload(), headers=auth())
    assert response.status_code == 503
    # The message must be actionable, not an empty body.
    assert response.json()["detail"]
    # Nothing was written.
    stats = client.get("/stats", headers=auth()).json()
    assert stats["mail_count"] == 0


def test_bad_mail_does_not_crash_service(client):
    # Missing body is fine; the service must still respond.
    response = client.post(
        "/index/mail",
        json={"store_id": "S1", "entry_id": "E9", "subject": "Only subject"},
        headers=auth(),
    )
    assert response.status_code == 200
    # And the service keeps working afterward.
    assert client.get("/health").status_code == 200


def test_search_with_metadata_filters(client):
    client.post(
        "/index/mail",
        json=mail_payload(entry_id="E1", sender_email="richard@example.com"),
        headers=auth(),
    )
    client.post(
        "/index/mail",
        json=mail_payload(entry_id="E2", sender_email="other@example.com",
                          internet_message_id="m2"),
        headers=auth(),
    )
    response = client.post(
        "/search",
        json={"query": "panel edit", "senders": ["richard@example.com"]},
        headers=auth(),
    )
    body = response.json()
    assert body["filter_applied"] is True
    assert all(r["sender_email"] == "richard@example.com"
               for r in body["results"])


def test_health_reports_local_paths(client):
    """The add-in shows and opens these paths, so they must be present."""

    body = client.get("/health").json()
    assert "paths" in body
    paths = body["paths"]
    for key in ("model_dir", "data_dir", "index_dir", "database",
                "log_dir", "service_dir"):
        assert key in paths, "missing path hint: %s" % key
        assert paths[key], "empty path hint: %s" % key
    # Paths must never point at an Outlook or mail-content location.
    assert "mail.db" in paths["database"]


def test_health_reports_device_and_pid(client):
    body = client.get("/health").json()
    assert "device" in body
    assert "device_reason" in body
    assert "device_preference" in body
    assert "pid" in body
    assert isinstance(body["pid"], int)


def test_shutdown_endpoint_requires_token(client):
    assert client.post("/service/shutdown").status_code == 401


def test_shutdown_endpoint_requests_stop(client):
    calls = []
    client.app.state.service.shutdown_callback = lambda: calls.append(True)
    response = client.post("/service/shutdown", headers=auth())
    assert response.status_code == 200
    assert response.json()["stopping"] is True
    assert calls == [True]


def test_health_triggers_model_load_when_cached(client):
    """A running service must report 'model ready', not 'not ready'.

    The model loads lazily. Without a load on the health check, a freshly
    started service reported model_available=false and the add-in showed
    'model not ready' even though everything was fine.
    """

    calls = {"load": 0}

    class CachedModel:
        def __init__(self):
            self.available = False
            self.load_error = None
            self.resolved_device = "cpu"
            self.device_reason = "test"
            self.loading = False
            self.models_dir = None

        def is_cached(self):
            return True

        def load(self, offline_if_cached=True):
            calls["load"] += 1
            self.available = True
            return True

        def encode_queries(self, *a, **k):  # pragma: no cover
            raise AssertionError("not used")

    client.app.state.service.model = CachedModel()
    body = client.get("/health").json()
    assert calls["load"] == 1, "health must trigger a load when cached"
    assert body["model_available"] is True
    assert body["model_cached"] is True


def test_health_does_not_load_when_not_cached(client):
    """Do not attempt a load (and a download) when there is no cache."""

    calls = {"load": 0}

    class UncachedModel:
        available = False
        load_error = None
        resolved_device = "cpu"
        device_reason = "test"
        loading = False

        def is_cached(self):
            return False

        def load(self, offline_if_cached=True):  # pragma: no cover
            calls["load"] += 1
            return False

        def encode_queries(self, *a, **k):  # pragma: no cover
            raise AssertionError("not used")

    client.app.state.service.model = UncachedModel()
    body = client.get("/health").json()
    assert calls["load"] == 0
    assert body["model_available"] is False
    assert body["model_cached"] is False
