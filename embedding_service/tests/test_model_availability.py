"""Tests for the model-availability contract.

The most important guarantee here: when the embedding model is unavailable,
indexing must FAIL and must not write vector-less rows. A previous version
silently stored every mail without a vector, which made a 3000-mail index
look populated but return nothing.
"""

from __future__ import annotations

import pytest

from app.database import Database
from app.embedding_model import EmbeddingModelUnavailable
from app.indexing import Indexer
from app.search import HybridSearch

from conftest import FakeEmbeddingModel, make_mail


class BrokenModel(FakeEmbeddingModel):
    """A model that reports itself unavailable and cannot encode."""

    def __init__(self, error="model missing"):
        super().__init__()
        self.available = False
        self._error_text = error
        self.load_error = error

    def load(self):
        return False

    def encode_queries(self, queries, normalize=True):
        raise EmbeddingModelUnavailable(self._error_text)

    def encode_documents(self, documents, titles=None, normalize=True):
        raise EmbeddingModelUnavailable(self._error_text)


def test_index_fails_when_model_unavailable(settings, database):
    indexer = Indexer(database, BrokenModel("not downloaded"), settings)
    with pytest.raises(EmbeddingModelUnavailable):
        indexer.index_mail(make_mail(
            entry_id="E1", subject="Any", body_text="body"))
    # Crucially: nothing was written.
    assert database.stats()["mail_count"] == 0
    assert database.stats()["chunk_count"] == 0


def test_index_does_not_store_vectorless_rows(settings, database):
    """A failed build must not leave a partial, useless index behind."""

    indexer = Indexer(database, BrokenModel("no model"), settings)
    for i in range(5):
        with pytest.raises(EmbeddingModelUnavailable):
            indexer.index_mail(make_mail(
                entry_id="E%d" % i, subject="S%d" % i, body_text="b"))
    stats = database.stats()
    assert stats["mail_count"] == 0
    assert stats["embedded_chunk_count"] == 0


def test_existing_index_survives_model_loss(settings, database):
    """A working index must not be damaged when the model later goes away."""

    good = Indexer(database, FakeEmbeddingModel(), settings)
    good.index_mail(make_mail(
        entry_id="E1", subject="Curtain wall",
        body_text="can we edit the panels individually"))
    before = database.stats()
    assert before["embedded_chunk_count"] >= 1

    broken = Indexer(database, BrokenModel("gone"), settings)
    with pytest.raises(EmbeddingModelUnavailable):
        broken.index_mail(make_mail(
            entry_id="E2", subject="New", body_text="new body"))

    # The original mail is untouched and still searchable.
    after = database.stats()
    assert after["embedded_chunk_count"] == before["embedded_chunk_count"]
    searcher = HybridSearch(database, FakeEmbeddingModel(), settings)
    response = searcher.search("curtain wall panel edit")
    assert any(r.entry_id == "E1" for r in response.results)


def test_encoder_raises_when_not_loaded(tmp_path):
    """EmbeddingModel must raise, not return zeros, when unusable."""

    from app.embedding_model import EmbeddingModel

    model = EmbeddingModel("this/model/does/not/exist-xyz")
    # Force the unavailable state without touching the network.
    model._load_error = "not installed"  # noqa: SLF001
    with pytest.raises(EmbeddingModelUnavailable):
        model.encode_queries(["hello"])
    with pytest.raises(EmbeddingModelUnavailable):
        model.encode_documents(["hello"])


def test_is_cached_is_offline_and_does_not_raise(tmp_path):
    """is_cached must never touch the network and never raise.

    A fresh machine has no cache; a machine that already downloaded the
    model has one. Either way the call is a local check only.
    """

    from app.embedding_model import EmbeddingModel

    model = EmbeddingModel("this/model/does/not/exist-xyz")
    # Must return a bool, not raise, even for a bogus repo id.
    assert model.is_cached() in (True, False)


def test_load_prefers_offline_when_cached(monkeypatch):
    """When the snapshot is cached, load must pass local_files_only=True.

    This is what stops the service contacting Hugging Face on every start.
    """

    from app import embedding_model as mod

    captured = {}

    class FakeST:
        def __init__(self, model_id, **kwargs):
            captured["model_kwargs"] = kwargs.get("model_kwargs", {})

    model = mod.EmbeddingModel("google/embeddinggemma-2")
    # Pretend the snapshot is present and the runtime is importable.
    monkeypatch.setattr(model, "is_cached", lambda: True)
    monkeypatch.setattr(mod.EmbeddingModel, "_resolve_dtype",
                        lambda self, torch: "float32")

    import sys, types
    fake_torch = types.ModuleType("torch")
    fake_torch.float32 = "float32"
    fake_torch.bfloat16 = "bfloat16"
    monkeypatch.setitem(sys.modules, "torch", fake_torch)

    fake_st_module = types.ModuleType("sentence_transformers")
    fake_st_module.SentenceTransformer = FakeST
    monkeypatch.setitem(sys.modules, "sentence_transformers", fake_st_module)

    assert model.load() is True
    assert captured["model_kwargs"].get("local_files_only") is True
