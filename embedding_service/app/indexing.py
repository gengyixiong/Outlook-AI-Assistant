"""Indexing orchestration: chunk, embed, and upsert one mail at a time.

The C# add-in sends one mail per request so that:
* it can show progress and cancel between messages;
* a single bad message cannot corrupt the whole batch;
* no long-running transaction spans the whole mailbox.
"""

from __future__ import annotations

import hashlib
import logging
import time
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence

import numpy as np

from .body_cleaner import clean_email_body
from .config import ServiceSettings
from .database import ChunkRecord, Database, MailRecord
from .embedding_model import EmbeddingModel, EmbeddingModelUnavailable
from .text_utils import chunk_text, deduplicate_chunks

logger = logging.getLogger("oaa.indexing")


@dataclass
class IndexMailRequest:
    store_id: str
    entry_id: str
    subject: str = ""
    internet_message_id: str = ""
    conversation_id: str = ""
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
    html_body: str = ""
    owner_mailbox: str = ""


@dataclass
class IndexMailResult:
    mail_id: int
    chunk_count: int
    embedded: bool
    content_hash: str
    skipped: bool = False
    reason: str = ""


def compute_content_hash(
    subject: str,
    sender_email: str,
    body_text: str,
    attachment_names: str,
) -> str:
    normalized = "|".join([
        (subject or "").strip().lower(),
        (sender_email or "").strip().lower(),
        (body_text or "").strip(),
        (attachment_names or "").strip().lower(),
    ])
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


class Indexer:
    def __init__(
        self,
        database: Database,
        model: EmbeddingModel,
        settings: ServiceSettings,
    ) -> None:
        self._db = database
        self._model = model
        self._settings = settings

    def index_mail(
        self,
        request: IndexMailRequest,
        force_reembed: bool = False,
    ) -> IndexMailResult:
        cleaned = clean_email_body(request.body_text, request.html_body)
        content_hash = compute_content_hash(
            request.subject,
            request.sender_email,
            cleaned,
            request.attachment_names,
        )

        existing = self._db.get_mail_by_entry(
            request.store_id, request.entry_id
        )
        if existing is None and request.internet_message_id:
            identity = self._db.find_by_internet_message_id(
                request.internet_message_id
            )
            if identity is not None:
                existing = self._db.get_mail(identity.id)

        if (
            existing is not None
            and not force_reembed
            and existing["content_hash"] == content_hash
            and existing["chunk_count"] > 0
            and self._mail_has_embeddings(existing["id"])
        ):
            # Metadata-only refresh: the message may have moved folders.
            if (
                existing["folder_id"] != request.folder_id
                or existing["folder_path"] != request.folder_path
            ):
                self._db.update_tombstone(
                    existing["id"],
                    request.folder_id,
                    request.folder_path,
                    request.last_modified_at,
                )
            return IndexMailResult(
                mail_id=existing["id"],
                chunk_count=existing["chunk_count"],
                embedded=True,
                content_hash=content_hash,
                skipped=True,
                reason="unchanged",
            )

        mail = MailRecord(
            store_id=request.store_id,
            entry_id=request.entry_id,
            internet_message_id=request.internet_message_id,
            conversation_id=request.conversation_id,
            subject=request.subject,
            sender_name=request.sender_name,
            sender_email=request.sender_email,
            recipients=request.recipients,
            received_at=request.received_at,
            sent_at=request.sent_at,
            last_modified_at=request.last_modified_at,
            folder_id=request.folder_id,
            folder_path=request.folder_path,
            attachment_names=request.attachment_names,
            attachment_count=request.attachment_count,
            body_text=cleaned,
            content_hash=content_hash,
            owner_mailbox=request.owner_mailbox,
        )

        chunk_texts = self._build_chunks(request, cleaned)

        # The semantic index is only useful with vectors. If the model is
        # unavailable we must fail loudly instead of writing thousands of
        # vector-less rows that make the index look populated but return
        # nothing. The add-in surfaces this error and stops the build.
        if not self._model.available and not self._model.load():
            raise EmbeddingModelUnavailable(
                self._model.load_error
                or "Embedding model is unavailable. Download it first."
            )

        vectors = self._model.encode_documents(
            chunk_texts,
            titles=[request.subject] * len(chunk_texts),
        )
        embedded = True

        chunks: List[ChunkRecord] = []
        for position, text in enumerate(chunk_texts):
            vector: Optional[np.ndarray] = None
            if position < vectors.shape[0]:
                vector = vectors[position]
            chunks.append(ChunkRecord(
                mail_id=existing["id"] if existing is not None else 0,
                chunk_index=position,
                text=text,
                embedding=vector,
            ))

        mail.index_status = "indexed"
        mail_id = self._db.upsert_mail(mail, chunks, _utc_now())
        return IndexMailResult(
            mail_id=mail_id,
            chunk_count=len(chunks),
            embedded=embedded,
            content_hash=content_hash,
        )

    def _build_chunks(
        self,
        request: IndexMailRequest,
        cleaned_body: str,
    ) -> List[str]:
        """Build chunk texts: subject header + body, deduplicated.

        The subject is prepended to each chunk so that a mail whose relevance
        lives only in the subject is still retrievable. It is included once per
        chunk, not repeated inside the body, so it cannot be over-weighted.
        """

        settings = self._settings
        parts: List[str] = []
        subject = (request.subject or "").strip()
        if cleaned_body:
            parts.extend(chunk_text(
                cleaned_body,
                target_tokens=settings.chunk_target_tokens,
                min_tokens=settings.chunk_min_tokens,
                max_tokens=settings.chunk_max_tokens,
                overlap_tokens=settings.chunk_overlap_tokens,
            ))

        if not parts and subject:
            parts = [subject]

        parts = deduplicate_chunks(parts)

        attachment_names = (request.attachment_names or "").strip()
        result: List[str] = []
        for part in parts:
            header = subject if subject else ""
            if attachment_names:
                header = (header + " | attachments: " + attachment_names).strip(
                    " |"
                )
            text = (header + "\n" + part).strip() if header else part
            result.append(text)

        if not result:
            # Very short or empty mail: index the subject so it stays findable.
            fallback = subject or request.sender_email or "(empty message)"
            result = [fallback]
        return result

    def _mail_has_embeddings(self, mail_id: int) -> bool:
        with self._db._lock:  # noqa: SLF001
            row = self._db._connection.execute(  # noqa: SLF001
                "SELECT COUNT(*) AS c FROM chunks "
                "WHERE mail_id = ? AND embedding IS NOT NULL",
                (mail_id,),
            ).fetchone()
        return bool(row and row["c"] > 0)

    def delete_mail(self, store_id: str, entry_id: str) -> bool:
        row = self._db.get_mail_by_entry(store_id, entry_id)
        if row is None and entry_id:
            # EntryID changed after a move; try to find by nothing else here.
            return False
        if row is None:
            return False
        self._db.delete_mail(row["id"])
        return True

    def stats(self) -> Dict[str, object]:
        return self._db.stats()


def _utc_now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
