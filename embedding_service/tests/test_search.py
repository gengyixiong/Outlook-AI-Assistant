"""Indexing + hybrid search tests, including Chinese-query/English-mail."""

from __future__ import annotations

import numpy as np

from app.database import ChunkRecord
from app.indexing import Indexer, IndexMailRequest, compute_content_hash
from app.search import HybridSearch

from conftest import FakeEmbeddingModel, make_mail


def _index(indexer, **kw):
    return indexer.index_mail(make_mail(**kw))


def test_index_mail_embeds_chunks(indexer, database):
    result = _index(
        indexer,
        entry_id="E1",
        subject="Curtain wall panels",
        internet_message_id="m1",
        sender_email="client@example.com",
        body_text="Can we edit the Panels individually? " * 30,
    )
    assert result.embedded is True
    assert result.chunk_count >= 1
    assert database.stats()["mail_count"] == 1


def test_unchanged_mail_is_skipped(indexer):
    payload = dict(
        entry_id="E1",
        subject="Same",
        internet_message_id="m1",
        body_text="unchanged body",
    )
    first = _index(indexer, **payload)
    assert first.skipped is False
    second = _index(indexer, **payload)
    assert second.skipped is True
    assert second.reason == "unchanged"


def test_changed_body_reembeds(indexer):
    _index(indexer, entry_id="E1", subject="S", body_text="version one")
    result = _index(indexer, entry_id="E1", subject="S", body_text="version two")
    assert result.skipped is False


def test_metadata_only_refresh_on_move(indexer, database):
    first = _index(indexer, entry_id="E1", subject="S", body_text="body",
                   folder_id="A", folder_path="Inbox")
    second = _index(indexer, entry_id="E1", subject="S", body_text="body",
                    folder_id="B", folder_path="Archive")
    assert second.skipped is True
    row = database.get_mail(first.mail_id)
    assert row["folder_path"] == "Archive"
    assert row["index_status"] == "moved"


def test_delete(indexer, database):
    _index(indexer, entry_id="E1", subject="S", body_text="body")
    assert indexer.delete_mail("STORE1", "E1") is True
    assert database.stats()["mail_count"] == 0


def test_semantic_search_cross_language(settings, database):
    """Chinese query must retrieve an English mail with no shared keywords."""

    indexer = Indexer(database, FakeEmbeddingModel(), settings)
    _index(
        indexer,
        entry_id="E1",
        subject="Curtain wall edit request",
        internet_message_id="m1",
        sender_email="india@example.com",
        body_text="Regarding the curtain wall edit, can we also edit the "
                  "Panels individually?",
    )
    _index(
        indexer,
        entry_id="E2",
        subject="Lunch menu",
        internet_message_id="m2",
        sender_email="other@example.com",
        body_text="Let us meet for lunch tomorrow.",
    )
    searcher = HybridSearch(database, FakeEmbeddingModel(), settings)
    response = searcher.search("之前印度客户问过幕墙玻璃 Panel 能不能单独编辑")
    assert response.results
    assert response.results[0].entry_id == "E1"
    assert response.semantic_available is True


def test_keyword_search_finds_exact_term(settings, database):
    indexer = Indexer(database, FakeEmbeddingModel(), settings)
    _index(indexer, entry_id="E1", subject="GstarBIM memory",
           body_text="GstarBIM used 24GB of memory")
    _index(indexer, entry_id="E2", subject="Unrelated",
           body_text="nothing to see here")
    searcher = HybridSearch(database, FakeEmbeddingModel(), settings)
    response = searcher.search("GstarBIM")
    ids = [r.entry_id for r in response.results]
    assert "E1" in ids
    assert response.keyword_available is True


def test_multiple_chunks_do_not_duplicate_results(settings, database):
    indexer = Indexer(database, FakeEmbeddingModel(), settings)
    long_body = "curtain wall panel edit request. " * 400
    _index(indexer, entry_id="E1", subject="Long", body_text=long_body)
    searcher = HybridSearch(database, FakeEmbeddingModel(), settings)
    response = searcher.search("curtain wall panel edit")
    mail_ids = [r.mail_id for r in response.results]
    assert len(mail_ids) == len(set(mail_ids))


def test_metadata_filter_in_search(settings, database):
    indexer = Indexer(database, FakeEmbeddingModel(), settings)
    _index(indexer, entry_id="E1", subject="Curtain", 
           sender_email="richard@example.com", body_text="curtain wall panel",
           received_at="2026-06-10T00:00:00Z")
    _index(indexer, entry_id="E2", subject="Curtain",
           sender_email="someone@example.com", body_text="curtain wall panel",
           received_at="2026-07-10T00:00:00Z")
    searcher = HybridSearch(database, FakeEmbeddingModel(), settings)
    response = searcher.search("curtain wall", senders=["richard@example.com"])
    assert response.filter_applied is True
    assert all(r.sender_email == "richard@example.com"
               for r in response.results)
    response = searcher.search(
        "curtain wall", received_from="2026-07-01", received_through="2026-07-31"
    )
    assert [r.entry_id for r in response.results] == ["E2"]


def test_empty_filter_result(settings, database):
    indexer = Indexer(database, FakeEmbeddingModel(), settings)
    _index(indexer, entry_id="E1", sender_email="a@b.com", body_text="x")
    searcher = HybridSearch(database, FakeEmbeddingModel(), settings)
    response = searcher.search("x", senders=["nobody@nowhere.com"])
    assert response.results == []
    assert response.filter_applied is True


def test_search_returns_snippet_and_metadata(settings, database):
    indexer = Indexer(database, FakeEmbeddingModel(), settings)
    _index(indexer, entry_id="E1", subject="Report",
           sender_name="Jane", sender_email="jane@example.com",
           folder_path="Inbox\\Reports", attachment_names="compare.xlsx",
           attachment_count=1,
           body_text="Excel product version comparison attached.")
    searcher = HybridSearch(database, FakeEmbeddingModel(), settings)
    response = searcher.search("Excel version comparison")
    assert response.results
    result = response.results[0]
    assert result.subject == "Report"
    assert result.sender_name == "Jane"
    assert result.folder_path == "Inbox\\Reports"
    assert result.attachment_names == "compare.xlsx"
    assert result.snippet


def test_keyword_only_when_model_unavailable(settings, database):
    """If the model cannot run, keyword search must still return results."""

    class UnavailableModel(FakeEmbeddingModel):
        def __init__(self):
            super().__init__()
            from app.embedding_model import EmbeddingModelUnavailable
            self._error = EmbeddingModelUnavailable("no model")

        def load(self):
            return False

        def encode_queries(self, queries, normalize=True):
            from app.embedding_model import EmbeddingModelUnavailable
            raise EmbeddingModelUnavailable("no model")

        def encode_documents(self, documents, titles=None, normalize=True):
            from app.embedding_model import EmbeddingModelUnavailable
            raise EmbeddingModelUnavailable("no model")

    # Index with the working fake first.
    indexer = Indexer(database, FakeEmbeddingModel(), settings)
    _index(indexer, entry_id="E1", subject="IFC4", body_text="IFC4 question")
    searcher = HybridSearch(database, UnavailableModel(), settings)
    response = searcher.search("IFC4")
    assert response.semantic_available is False
    # Index has vectors from the earlier run, so semantic scoring can still
    # run; here the model raises, so only keyword results remain.
    assert response.results
    assert response.results[0].entry_id == "E1"


def test_moved_and_deleted_mail_not_searchable(settings, database):
    indexer = Indexer(database, FakeEmbeddingModel(), settings)
    result = _index(indexer, entry_id="E1", subject="Gone",
                    body_text="unique gone marker")
    indexer.delete_mail("STORE1", "E1")
    searcher = HybridSearch(database, FakeEmbeddingModel(), settings)
    response = searcher.search("unique gone marker")
    assert result.mail_id not in [r.mail_id for r in response.results]
