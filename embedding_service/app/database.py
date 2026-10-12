"""SQLite storage for email metadata, chunks, vectors and full-text search.

Design notes
------------
* One row per mail in ``mails`` keyed by ``(store_id, entry_id)``. Entry IDs
  change when a message moves folders, so ``internet_message_id`` plus a
  content hash is used for incremental identity.
* Chunks live in ``chunks`` with a foreign key to ``mails`` and cascade
  delete, so deleting a mail removes its vectors. No orphan vectors.
* Word search uses FTS5 over a *normalized token field* produced by
  ``text_utils.build_token_field``. The SQLite trigram tokenizer cannot match
  Chinese text or short terms, so we index pre-tokenized text with unicode61.
* Vectors are stored as float32 blobs in ``chunks.embedding`` and scored in
  Python with numpy. This keeps the schema simple and portable across SQLite
  builds and avoids a native vector extension.
* The whole database is written behind a process-wide lock and every mutation
  runs in a transaction, so an interrupted sync cannot leave half-written rows.
"""

from __future__ import annotations

import hashlib
import sqlite3
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np

from .text_utils import build_token_field

SCHEMA_VERSION = 1

_SCHEMA = """
PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;
PRAGMA synchronous = NORMAL;

CREATE TABLE IF NOT EXISTS meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS mails (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    store_id TEXT NOT NULL DEFAULT '',
    entry_id TEXT NOT NULL DEFAULT '',
    internet_message_id TEXT NOT NULL DEFAULT '',
    conversation_id TEXT NOT NULL DEFAULT '',
    subject TEXT NOT NULL DEFAULT '',
    sender_name TEXT NOT NULL DEFAULT '',
    sender_email TEXT NOT NULL DEFAULT '',
    recipients TEXT NOT NULL DEFAULT '',
    received_at TEXT NOT NULL DEFAULT '',
    sent_at TEXT NOT NULL DEFAULT '',
    last_modified_at TEXT NOT NULL DEFAULT '',
    folder_id TEXT NOT NULL DEFAULT '',
    folder_path TEXT NOT NULL DEFAULT '',
    attachment_names TEXT NOT NULL DEFAULT '',
    attachment_count INTEGER NOT NULL DEFAULT 0,
    body_text TEXT NOT NULL DEFAULT '',
    content_hash TEXT NOT NULL DEFAULT '',
    indexed_at TEXT NOT NULL DEFAULT '',
    index_status TEXT NOT NULL DEFAULT 'pending',
    chunk_count INTEGER NOT NULL DEFAULT 0,
    owner_mailbox TEXT NOT NULL DEFAULT ''
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_mails_entry
    ON mails(store_id, entry_id);

CREATE INDEX IF NOT EXISTS idx_mails_conv
    ON mails(conversation_id);

CREATE INDEX IF NOT EXISTS idx_mails_imid
    ON mails(internet_message_id);

CREATE INDEX IF NOT EXISTS idx_mails_received
    ON mails(received_at);

CREATE INDEX IF NOT EXISTS idx_mails_modified
    ON mails(last_modified_at);

CREATE INDEX IF NOT EXISTS idx_mails_status
    ON mails(index_status);

CREATE TABLE IF NOT EXISTS chunks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    mail_id INTEGER NOT NULL,
    chunk_index INTEGER NOT NULL,
    text TEXT NOT NULL DEFAULT '',
    token_text TEXT NOT NULL DEFAULT '',
    embedding BLOB,
    embedding_dim INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (mail_id) REFERENCES mails(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_chunks_mail
    ON chunks(mail_id);

CREATE UNIQUE INDEX IF NOT EXISTS idx_chunks_mail_index
    ON chunks(mail_id, chunk_index);

CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts USING fts5(
    token_text,
    content='chunks',
    content_rowid='id',
    tokenize='unicode61 remove_diacritics 2'
);

CREATE TRIGGER IF NOT EXISTS chunks_ai AFTER INSERT ON chunks BEGIN
    INSERT INTO chunks_fts(rowid, token_text)
    VALUES (new.id, new.token_text);
END;

CREATE TRIGGER IF NOT EXISTS chunks_ad AFTER DELETE ON chunks BEGIN
    INSERT INTO chunks_fts(chunks_fts, rowid, token_text)
    VALUES ('delete', old.id, old.token_text);
END;

CREATE TRIGGER IF NOT EXISTS chunks_au AFTER UPDATE ON chunks BEGIN
    INSERT INTO chunks_fts(chunks_fts, rowid, token_text)
    VALUES ('delete', old.id, old.token_text);
    INSERT INTO chunks_fts(rowid, token_text)
    VALUES (new.id, new.token_text);
END;
"""


@dataclass
class MailRecord:
    """Plain data for one indexed message."""

    store_id: str = ""
    entry_id: str = ""
    internet_message_id: str = ""
    conversation_id: str = ""
    subject: str = ""
    sender_name: str = ""
    sender_email: str = ""
    recipients: str = ""
    received_at: str = ""
    sent_at: str = ""
    last_modified_at: str = ""
    folder_id: str = ""
    folder_path: str = ""
    attachment_names: str = ""
    attachment_count: int = 0
    body_text: str = ""
    content_hash: str = ""
    owner_mailbox: str = ""
    id: Optional[int] = None
    chunk_count: int = 0
    index_status: str = "pending"

    def identity(self) -> str:
        """Stable key used for incremental comparison when EntryID changes."""

        if self.internet_message_id:
            return "imid:" + self.internet_message_id.lower()
        parts = "|".join([
            self.store_id,
            self.entry_id,
        ])
        return "entry:" + hashlib.sha1(parts.encode("utf-8")).hexdigest()


@dataclass
class ChunkRecord:
    mail_id: int
    chunk_index: int
    text: str
    embedding: Optional[np.ndarray] = None


@dataclass
class MailIdentity:
    """Minimal projection used by incremental sync decisions."""

    id: int
    store_id: str
    entry_id: str
    internet_message_id: str
    folder_id: str
    content_hash: str
    index_status: str


@dataclass
class KeywordHit:
    chunk_id: int
    mail_id: int
    rank: float
    score: float


class Database:
    """Thread-safe wrapper around a SQLite connection."""

    def __init__(self, path: Path, embedding_dim: int = 768) -> None:
        self.path = Path(path)
        self.embedding_dim = embedding_dim
        self._lock = threading.RLock()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._connection = sqlite3.connect(
            str(self.path),
            check_same_thread=False,
            isolation_level=None,
        )
        self._connection.row_factory = sqlite3.Row
        self._initialize()

    def _initialize(self) -> None:
        with self._lock:
            self._connection.executescript(_SCHEMA)
            self._set_meta("schema_version", str(SCHEMA_VERSION))
            self._set_meta("embedding_dim", str(self.embedding_dim))

    def close(self) -> None:
        with self._lock:
            self._connection.close()

    # ------------------------------------------------------------------
    # meta
    # ------------------------------------------------------------------
    def _set_meta(self, key: str, value: str) -> None:
        self._connection.execute(
            "INSERT INTO meta(key, value) VALUES(?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, value),
        )

    def get_meta(self, key: str) -> Optional[str]:
        with self._lock:
            row = self._connection.execute(
                "SELECT value FROM meta WHERE key = ?", (key,)
            ).fetchone()
            return row["value"] if row else None

    def set_meta(self, key: str, value: str) -> None:
        with self._lock:
            self._set_meta(key, value)

    # ------------------------------------------------------------------
    # mail lookup used by incremental sync
    # ------------------------------------------------------------------
    def load_identities(self) -> Dict[str, MailIdentity]:
        """Return every indexed mail keyed by store_id + entry_id.

        The C# scanner uses the returned data to decide which mails need
        embedding. Identity is also resolvable by internet message id via
        ``find_by_internet_message_id``.
        """

        with self._lock:
            rows = self._connection.execute(
                "SELECT id, store_id, entry_id, internet_message_id, "
                "folder_id, content_hash, index_status FROM mails"
            ).fetchall()
            result: Dict[str, MailIdentity] = {}
            for row in rows:
                identity = MailIdentity(
                    id=row["id"],
                    store_id=row["store_id"],
                    entry_id=row["entry_id"],
                    internet_message_id=row["internet_message_id"],
                    folder_id=row["folder_id"],
                    content_hash=row["content_hash"],
                    index_status=row["index_status"],
                )
                result[_entry_key(row["store_id"], row["entry_id"])] = identity
            return result

    def find_by_internet_message_id(
        self,
        internet_message_id: str,
    ) -> Optional[MailIdentity]:
        if not internet_message_id:
            return None
        with self._lock:
            row = self._connection.execute(
                "SELECT id, store_id, entry_id, internet_message_id, "
                "folder_id, content_hash, index_status FROM mails "
                "WHERE internet_message_id = ? COLLATE NOCASE LIMIT 1",
                (internet_message_id,),
            ).fetchone()
            return _identity_from_row(row) if row else None

    def existing_entry_keys(self) -> set:
        with self._lock:
            rows = self._connection.execute(
                "SELECT store_id, entry_id FROM mails"
            ).fetchall()
            return {_entry_key(r["store_id"], r["entry_id"]) for r in rows}

    # ------------------------------------------------------------------
    # writes
    # ------------------------------------------------------------------
    def upsert_mail(
        self,
        mail: MailRecord,
        chunks: Sequence[ChunkRecord],
        now_iso: Optional[str] = None,
    ) -> int:
        """Insert or replace one mail and its chunks atomically."""

        timestamp = now_iso or _utc_now()
        mail.indexed_at = timestamp
        mail.chunk_count = len(chunks)
        mail.index_status = mail.index_status or "indexed"

        with self._lock:
            self._connection.execute("BEGIN IMMEDIATE")
            try:
                existing = self._find_existing_row(mail)
                if existing is not None:
                    mail_id = existing["id"]
                    self._connection.execute(
                        """
                        UPDATE mails SET
                            store_id = ?, entry_id = ?,
                            internet_message_id = ?, conversation_id = ?,
                            subject = ?, sender_name = ?, sender_email = ?,
                            recipients = ?, received_at = ?, sent_at = ?,
                            last_modified_at = ?, folder_id = ?, folder_path = ?,
                            attachment_names = ?, attachment_count = ?,
                            body_text = ?, content_hash = ?, indexed_at = ?,
                            index_status = ?, chunk_count = ?, owner_mailbox = ?
                        WHERE id = ?
                        """,
                        (
                            mail.store_id, mail.entry_id,
                            mail.internet_message_id, mail.conversation_id,
                            mail.subject, mail.sender_name, mail.sender_email,
                            mail.recipients, mail.received_at, mail.sent_at,
                            mail.last_modified_at, mail.folder_id,
                            mail.folder_path, mail.attachment_names,
                            mail.attachment_count, mail.body_text,
                            mail.content_hash, mail.indexed_at,
                            mail.index_status, mail.chunk_count,
                            mail.owner_mailbox, mail_id,
                        ),
                    )
                    # Replace chunks (cascade keeps FTS in sync via triggers).
                    self._connection.execute(
                        "DELETE FROM chunks WHERE mail_id = ?", (mail_id,)
                    )
                else:
                    cursor = self._connection.execute(
                        """
                        INSERT INTO mails (
                            store_id, entry_id, internet_message_id,
                            conversation_id, subject, sender_name,
                            sender_email, recipients, received_at, sent_at,
                            last_modified_at, folder_id, folder_path,
                            attachment_names, attachment_count, body_text,
                            content_hash, indexed_at, index_status,
                            chunk_count, owner_mailbox
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            mail.store_id, mail.entry_id,
                            mail.internet_message_id, mail.conversation_id,
                            mail.subject, mail.sender_name, mail.sender_email,
                            mail.recipients, mail.received_at, mail.sent_at,
                            mail.last_modified_at, mail.folder_id,
                            mail.folder_path, mail.attachment_names,
                            mail.attachment_count, mail.body_text,
                            mail.content_hash, mail.indexed_at,
                            mail.index_status, mail.chunk_count,
                            mail.owner_mailbox,
                        ),
                    )
                    mail_id = int(cursor.lastrowid)

                for position, chunk in enumerate(chunks):
                    blob = None
                    dim = 0
                    if chunk.embedding is not None:
                        vector = np.asarray(chunk.embedding, dtype=np.float32)
                        blob = vector.tobytes()
                        dim = int(vector.shape[0])
                    self._connection.execute(
                        """
                        INSERT INTO chunks (
                            mail_id, chunk_index, text, token_text,
                            embedding, embedding_dim
                        ) VALUES (?,?,?,?,?,?)
                        """,
                        (
                            mail_id,
                            chunk.chunk_index if chunk.chunk_index >= 0
                            else position,
                            chunk.text,
                            build_token_field(chunk.text),
                            blob,
                            dim,
                        ),
                    )

                mail.id = mail_id
                self._connection.execute("COMMIT")
                return mail_id
            except Exception:
                self._connection.execute("ROLLBACK")
                raise

    def delete_mail(self, mail_id: int) -> None:
        with self._lock:
            self._connection.execute("BEGIN IMMEDIATE")
            try:
                self._connection.execute(
                    "DELETE FROM mails WHERE id = ?", (mail_id,)
                )
                self._connection.execute("COMMIT")
            except Exception:
                self._connection.execute("ROLLBACK")
                raise

    def delete_mails(self, mail_ids: Sequence[int]) -> int:
        if not mail_ids:
            return 0
        with self._lock:
            self._connection.execute("BEGIN IMMEDIATE")
            try:
                placeholders = ",".join("?" for _ in mail_ids)
                cursor = self._connection.execute(
                    "DELETE FROM mails WHERE id IN (%s)" % placeholders,
                    tuple(mail_ids),
                )
                self._connection.execute("COMMIT")
                return cursor.rowcount
            except Exception:
                self._connection.execute("ROLLBACK")
                raise

    def update_tombstone(
        self,
        mail_id: int,
        folder_id: str,
        folder_path: str,
        last_modified_at: str = "",
    ) -> None:
        """Record a move so the next search still finds the mail."""

        with self._lock:
            self._connection.execute(
                "UPDATE mails SET folder_id = ?, folder_path = ?, "
                "last_modified_at = CASE WHEN ? != '' THEN ? "
                "ELSE last_modified_at END, index_status = 'moved' "
                "WHERE id = ?",
                (folder_id, folder_path, last_modified_at,
                 last_modified_at, mail_id),
            )

    # ------------------------------------------------------------------
    # reads
    # ------------------------------------------------------------------
    def _find_existing_row(self, mail: MailRecord) -> Optional[sqlite3.Row]:
        if mail.internet_message_id:
            row = self._connection.execute(
                "SELECT * FROM mails WHERE internet_message_id = ? "
                "COLLATE NOCASE LIMIT 1",
                (mail.internet_message_id,),
            ).fetchone()
            if row is not None:
                return row
        return self._connection.execute(
            "SELECT * FROM mails WHERE store_id = ? AND entry_id = ? LIMIT 1",
            (mail.store_id, mail.entry_id),
        ).fetchone()

    def get_mail(self, mail_id: int) -> Optional[sqlite3.Row]:
        with self._lock:
            return self._connection.execute(
                "SELECT * FROM mails WHERE id = ?", (mail_id,)
            ).fetchone()

    def get_mail_by_entry(
        self,
        store_id: str,
        entry_id: str,
    ) -> Optional[sqlite3.Row]:
        with self._lock:
            return self._connection.execute(
                "SELECT * FROM mails WHERE store_id = ? AND entry_id = ? "
                "LIMIT 1",
                (store_id, entry_id),
            ).fetchone()

    def all_embeddings(self) -> Tuple[List[int], List[int], np.ndarray]:
        """Load every chunk embedding for brute-force scoring.

        Returns (chunk_ids, mail_ids, matrix). Chunks without an embedding are
        skipped. The matrix is float32 and L2-normalized.
        """

        with self._lock:
            rows = self._connection.execute(
                "SELECT id, mail_id, embedding, embedding_dim FROM chunks "
                "WHERE embedding IS NOT NULL AND embedding_dim > 0"
            ).fetchall()
        chunk_ids: List[int] = []
        mail_ids: List[int] = []
        vectors: List[np.ndarray] = []
        for row in rows:
            vector = np.frombuffer(row["embedding"], dtype=np.float32)
            if vector.shape[0] != row["embedding_dim"]:
                continue
            chunk_ids.append(row["id"])
            mail_ids.append(row["mail_id"])
            vectors.append(vector)
        if not vectors:
            return [], [], np.zeros((0, self.embedding_dim), dtype=np.float32)
        matrix = np.vstack(vectors).astype(np.float32)
        norms = np.linalg.norm(matrix, axis=1, keepdims=True)
        norms[norms == 0] = 1.0
        matrix = matrix / norms
        return chunk_ids, mail_ids, matrix

    def keyword_search(
        self,
        query: str,
        limit: int,
        restrict_mail_ids: Optional[Sequence[int]] = None,
    ) -> List[KeywordHit]:
        """FTS5 search over the normalized token field."""

        tokens = _fts_query_from(query)
        if not tokens:
            return []
        with self._lock:
            try:
                rows = self._connection.execute(
                    """
                    SELECT c.id AS chunk_id, c.mail_id AS mail_id,
                           bm25(chunks_fts) AS rank
                    FROM chunks_fts
                    JOIN chunks c ON c.id = chunks_fts.rowid
                    WHERE chunks_fts MATCH ?
                    ORDER BY rank
                    LIMIT ?
                    """,
                    (tokens, limit * 4),
                ).fetchall()
            except sqlite3.OperationalError:
                return []
        allowed = set(restrict_mail_ids) if restrict_mail_ids else None
        hits: List[KeywordHit] = []
        for row in rows:
            mail_id = row["mail_id"]
            if allowed is not None and mail_id not in allowed:
                continue
            # bm25 returns smaller-is-better (often negative); convert to a
            # positive score where larger is better.
            rank = float(row["rank"])
            hits.append(KeywordHit(
                chunk_id=row["chunk_id"],
                mail_id=mail_id,
                rank=rank,
                score=-rank,
            ))
            if len(hits) >= limit:
                break
        return hits

    def mail_ids_for(
        self,
        senders: Optional[Sequence[str]] = None,
        received_from: Optional[str] = None,
        received_through: Optional[str] = None,
        has_attachments: Optional[bool] = None,
    ) -> Optional[List[int]]:
        """Return mail ids matching structured metadata constraints.

        Returns ``None`` when no constraint was supplied, meaning "no filter".
        An empty list means the constraints matched nothing.
        """

        clauses: List[str] = []
        params: List[object] = []
        if senders:
            lowered = [s.lower() for s in senders if s]
            if lowered:
                sub = []
                for sender in lowered:
                    sub.append(
                        "(lower(sender_email) = ? OR "
                        "lower(sender_email) LIKE ? OR "
                        "lower(sender_name) = ? OR "
                        "lower(sender_name) LIKE ?)"
                    )
                    params.extend([
                        sender, "%" + sender + "%", sender,
                        "%" + sender + "%",
                    ])
                clauses.append("(" + " OR ".join(sub) + ")")
        if received_from:
            clauses.append("received_at >= ?")
            params.append(received_from)
        if received_through:
            clauses.append("received_at <= ?")
            params.append(received_through + "\uffff")
        if has_attachments is True:
            clauses.append("attachment_count > 0")
        elif has_attachments is False:
            clauses.append("attachment_count = 0")

        if not clauses:
            return None

        sql = "SELECT id FROM mails WHERE " + " AND ".join(clauses)
        with self._lock:
            rows = self._connection.execute(sql, tuple(params)).fetchall()
        return [row["id"] for row in rows]

    def stats(self) -> Dict[str, object]:
        with self._lock:
            mail_count = self._connection.execute(
                "SELECT COUNT(*) AS c FROM mails"
            ).fetchone()["c"]
            chunk_count = self._connection.execute(
                "SELECT COUNT(*) AS c FROM chunks"
            ).fetchone()["c"]
            embedded = self._connection.execute(
                "SELECT COUNT(*) AS c FROM chunks WHERE embedding IS NOT NULL"
            ).fetchone()["c"]
            with_attachments = self._connection.execute(
                "SELECT COUNT(*) AS c FROM mails WHERE attachment_count > 0"
            ).fetchone()["c"]
        size_bytes = 0
        for suffix in ("", "-wal", "-shm"):
            candidate = Path(str(self.path) + suffix)
            if candidate.exists():
                size_bytes += candidate.stat().st_size
        return {
            "mail_count": int(mail_count),
            "chunk_count": int(chunk_count),
            "embedded_chunk_count": int(embedded),
            "mails_with_attachments": int(with_attachments),
            "size_bytes": int(size_bytes),
            "database_path": str(self.path),
            "last_sync_at": self.get_meta("last_sync_at") or "",
            "model_id": self.get_meta("model_id") or "",
            "embedding_dim": int(self.get_meta("embedding_dim") or 0),
        }

    def clear(self) -> None:
        with self._lock:
            self._connection.execute("BEGIN IMMEDIATE")
            try:
                self._connection.execute("DELETE FROM chunks")
                self._connection.execute("DELETE FROM mails")
                self._connection.execute(
                    "INSERT INTO chunks_fts(chunks_fts) VALUES('rebuild')"
                )
                self._connection.execute("COMMIT")
            except Exception:
                self._connection.execute("ROLLBACK")
                raise


def _entry_key(store_id: str, entry_id: str) -> str:
    return (store_id or "") + "\x1f" + (entry_id or "")


def _identity_from_row(row: sqlite3.Row) -> MailIdentity:
    return MailIdentity(
        id=row["id"],
        store_id=row["store_id"],
        entry_id=row["entry_id"],
        internet_message_id=row["internet_message_id"],
        folder_id=row["folder_id"],
        content_hash=row["content_hash"],
        index_status=row["index_status"],
    )


def _fts_query_from(query: str) -> str:
    """Build a safe FTS5 MATCH expression from user keyword text."""

    from .text_utils import keyword_tokens

    tokens = keyword_tokens(query)
    if not tokens:
        return ""
    # Quote every token to avoid FTS5 syntax errors on terms like "IFC4-1".
    quoted = []
    seen = set()
    for token in tokens:
        if token in seen:
            continue
        seen.add(token)
        quoted.append('"' + token.replace('"', '""') + '"')
    return " OR ".join(quoted)


def _utc_now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
