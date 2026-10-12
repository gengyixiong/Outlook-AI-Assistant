"""Test suite for the local semantic search service.

A deterministic fake embedding model replaces EmbeddingGemma 2 so the whole
pipeline (chunking, storage, FTS, vectors, RRF, API) can be verified without
downloading multi-GB weights. The fake produces vectors from a small keyword
space so that semantically related test mails overlap as intended.
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

import numpy as np
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


# ----------------------------------------------------------------------
# fake embedding model
# ----------------------------------------------------------------------
# A small hand-built "concept" vocabulary. Each concept maps to a dimension;
# documents/queries containing a concept's keywords get weight there. This
# makes the fake behave like a real semantic model for these fixed tests.
_CONCEPTS = {
    0: ["curtain", "wall", "panel", "panels", "幕墙", "玻璃", "edit", "编辑"],
    1: ["ifc", "ifc4", "certification", "认证", "代理商", "reseller"],
    2: ["revit", "link", "材质", "material", "import", "导入"],
    3: ["view", "视图", "rvt", "选择"],
    4: ["memory", "gstarbim", "内存", "24gb"],
    5: ["training", "channel", "渠道", "培训", "richard"],
    6: ["excel", "attachment", "附件", "version", "compare", "对比"],
}


class FakeEmbeddingModel:
    def __init__(self, dim: int = 768):
        self.embedding_dim = dim
        self.available = True
        self.load_error = None
        self._dim = dim

    def load(self) -> bool:
        return True

    def _vector(self, text: str) -> np.ndarray:
        lowered = (text or "").lower()
        vector = np.zeros(self._dim, dtype=np.float32)
        for axis, keywords in _CONCEPTS.items():
            if any(keyword in lowered for keyword in keywords):
                vector[axis] = 1.0
        # Add a weak token-hash component so distinct texts are not identical.
        for token in set(lowered.split()):
            index = (hash(token) % 200) + 10
            vector[index] += 0.05
        # Normalize like the real model.
        norm = np.linalg.norm(vector)
        if norm > 0:
            vector = vector / norm
        # Never emit an all-zero vector, which would break cosine scoring.
        if not vector.any():
            vector[200] = 1.0
        return vector

    def encode_queries(self, queries, normalize=True):
        return np.vstack([self._vector(q) for q in queries]).astype(np.float32)

    def encode_documents(self, documents, titles=None, normalize=True):
        vectors = []
        for index, document in enumerate(documents):
            title = titles[index] if titles and index < len(titles) else ""
            vectors.append(self._vector(f"{title} {document}"))
        return np.vstack(vectors).astype(np.float32)


@pytest.fixture()
def settings(tmp_path):
    from app.config import ServiceSettings

    return ServiceSettings(
        product_root=tmp_path,
        db_path=tmp_path / "mail.db",
        model_cache_dir=tmp_path / "models",
        log_path=tmp_path / "logs" / "svc.log",
        token_path=tmp_path / "token",
        embedding_dim=768,
    )


@pytest.fixture()
def database(settings):
    from app.database import Database

    db = Database(settings.db_path, embedding_dim=768)
    yield db
    db.close()


@pytest.fixture()
def indexer(settings, database):
    from app.indexing import Indexer

    return Indexer(database, FakeEmbeddingModel(), settings)


@pytest.fixture()
def searcher(settings, database):
    from app.search import HybridSearch

    return HybridSearch(database, FakeEmbeddingModel(), settings)


def make_mail(**overrides):
    from app.indexing import IndexMailRequest

    base = dict(
        store_id="STORE1",
        entry_id="",
        subject="",
        internet_message_id="",
        conversation_id="",
        sender_name="",
        sender_email="",
        recipients="",
        received_at="2026-06-15T09:00:00Z",
        sent_at="2026-06-15T08:59:00Z",
        last_modified_at="2026-06-15T09:00:00Z",
        folder_id="FOLDER_INBOX",
        folder_path="Inbox",
        attachment_names="",
        attachment_count=0,
        body_text="",
        html_body="",
        owner_mailbox="me@example.com",
    )
    base.update(overrides)
    return IndexMailRequest(**base)
