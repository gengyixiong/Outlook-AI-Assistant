"""Database and consistency tests: unique keys, cascade delete, FTS, moves."""

from __future__ import annotations

from app.database import ChunkRecord, Database, MailRecord
from app.indexing import compute_content_hash


def _mail(store_id="S1", entry="E1", **kw):
    base = dict(store_id=store_id, entry_id=entry,)
    base.update(kw)
    return MailRecord(**base)


def test_upsert_and_fetch(database):
    mail = _mail(subject="Hello", body_text="body")
    mail.content_hash = compute_content_hash("Hello", "", "body", "")
    mail_id = database.upsert_mail(
        mail, [ChunkRecord(0, 0, "chunk text")], "2026-01-01T00:00:00Z"
    )
    assert mail_id > 0
    row = database.get_mail(mail_id)
    assert row["subject"] == "Hello"
    assert row["chunk_count"] == 1


def test_unique_entry_constraint(database):
    mail = _mail(subject="A")
    mail.content_hash = "h1"
    first = database.upsert_mail(mail, [ChunkRecord(0, 0, "a")])
    mail2 = _mail(subject="A updated")
    mail2.content_hash = "h2"
    second = database.upsert_mail(mail2, [ChunkRecord(0, 0, "b")])
    assert first == second
    assert database.get_mail(first)["subject"] == "A updated"
    assert database.stats()["mail_count"] == 1


def test_internet_message_id_dedupes_across_entry_change(database):
    mail = _mail(entry="E1", internet_message_id="msg-1", subject="Original")
    mail.content_hash = "h1"
    database.upsert_mail(mail, [ChunkRecord(0, 0, "x")])
    moved = _mail(entry="E2", internet_message_id="msg-1", subject="Original")
    moved.content_hash = "h1"
    database.upsert_mail(moved, [ChunkRecord(0, 0, "x")])
    # Same logical message must not create a second row.
    assert database.stats()["mail_count"] == 1
    assert database.get_mail_by_entry("S1", "E2") is not None


def test_cascade_delete_removes_chunks_and_fts(database):
    mail = _mail(subject="Delete me")
    mail.content_hash = "h"
    mail_id = database.upsert_mail(
        mail, [ChunkRecord(0, 0, "unique delete marker word")]
    )
    assert database.stats()["chunk_count"] == 1
    hits = database.keyword_search("marker", limit=10)
    assert any(h.mail_id == mail_id for h in hits)
    database.delete_mail(mail_id)
    assert database.stats()["chunk_count"] == 0
    assert database.keyword_search("marker", limit=10) == []


def test_no_orphan_vectors_after_delete(database):
    mail = _mail(subject="Orphan check")
    mail.content_hash = "h"
    import numpy as np

    mail_id = database.upsert_mail(
        mail,
        [ChunkRecord(0, 0, "text", embedding=np.ones(768, dtype=np.float32))],
    )
    database.delete_mail(mail_id)
    chunk_ids, _, matrix = database.all_embeddings()
    assert chunk_ids == []
    assert matrix.shape[0] == 0


def test_keyword_search_chinese_two_char(database):
    mail = _mail(subject="幕墙玻璃问题")
    mail.content_hash = "h"
    database.upsert_mail(mail, [ChunkRecord(0, 0, "幕墙玻璃面板需要单独编辑")])
    hits = database.keyword_search("幕墙", limit=10)
    assert len(hits) == 1
    hits2 = database.keyword_search("玻璃", limit=10)
    assert len(hits2) == 1


def test_keyword_search_short_technical_term(database):
    mail = _mail(subject="IFC4 认证")
    mail.content_hash = "h"
    database.upsert_mail(mail, [ChunkRecord(0, 0, "reseller asked about IFC4")])
    assert len(database.keyword_search("IFC4", limit=10)) == 1
    # "IFC" is not a prefix of the indexed token "ifc4"; exact-token FTS
    # deliberately does not expand to prefixes to avoid excessive recall.
    assert database.keyword_search("IFC", limit=10) == []
    # But a substring of length >= 3 inside a token is matched by the
    # auxiliary bigram/word tokens only when the tokenizer produced it.
    assert len(database.keyword_search("reseller", limit=10)) == 1


def test_keyword_search_no_syntax_error_on_punctuation(database):
    mail = _mail(subject="weird")
    mail.content_hash = "h"
    database.upsert_mail(mail, [ChunkRecord(0, 0, "value is 12-34-56")])
    # A raw hyphen would be an FTS5 error; our tokenizer must prevent that.
    assert database.keyword_search("12-34-56", limit=10)


def test_move_updates_folder(database):
    mail = _mail(subject="Move me", folder_id="A", folder_path="Inbox")
    mail.content_hash = "h"
    mail_id = database.upsert_mail(mail, [ChunkRecord(0, 0, "text")])
    database.update_tombstone(mail_id, "B", "Archive")
    row = database.get_mail(mail_id)
    assert row["folder_id"] == "B"
    assert row["folder_path"] == "Archive"
    assert row["index_status"] == "moved"


def test_metadata_filter(database):
    m1 = _mail(entry="E1", sender_email="richard@example.com",
               received_at="2026-06-01T00:00:00Z", attachment_count=1)
    m1.content_hash = "h1"
    m2 = _mail(entry="E2", sender_email="other@example.com",
               received_at="2026-07-01T00:00:00Z", attachment_count=0)
    m2.content_hash = "h2"
    database.upsert_mail(m1, [ChunkRecord(0, 0, "a")])
    database.upsert_mail(m2, [ChunkRecord(0, 0, "b")])
    ids = database.mail_ids_for(senders=["richard@example.com"])
    assert len(ids) == 1
    ids = database.mail_ids_for(has_attachments=True)
    assert len(ids) == 1
    ids = database.mail_ids_for(
        received_from="2026-06-15", received_through="2026-06-30"
    )
    assert ids == []
    ids = database.mail_ids_for(
        received_from="2026-06-01", received_through="2026-06-30"
    )
    assert len(ids) == 1


def test_clear_removes_everything(database):
    mail = _mail(subject="x")
    mail.content_hash = "h"
    database.upsert_mail(mail, [ChunkRecord(0, 0, "x")])
    database.clear()
    stats = database.stats()
    assert stats["mail_count"] == 0
    assert stats["chunk_count"] == 0


def test_reopen_preserves_index(settings):
    db1 = Database(settings.db_path, embedding_dim=768)
    mail = _mail(subject="Persist")
    mail.content_hash = "h"
    db1.upsert_mail(mail, [ChunkRecord(0, 0, "persist me")])
    db1.close()
    db2 = Database(settings.db_path, embedding_dim=768)
    assert db2.stats()["mail_count"] == 1
    db2.close()
