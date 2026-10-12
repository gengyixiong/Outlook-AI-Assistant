"""End-to-end test against the REAL EmbeddingGemma 2 model.

Skipped automatically unless the runtime and model are present. Run with:
    python -m pytest tests/test_real_model.py -q -s --run-real-model
"""

from __future__ import annotations

import pytest

from app.config import ServiceSettings
from app.embedding_model import EmbeddingModel
from app.indexing import Indexer
from app.search import HybridSearch
from app.database import Database

from conftest import make_mail


def _model_available() -> bool:
    try:
        import torch  # noqa: F401
        import sentence_transformers  # noqa: F401
    except Exception:
        return False
    return True


pytestmark = pytest.mark.skipif(
    not _model_available(),
    reason="embedding runtime not installed",
)


@pytest.fixture(scope="module")
def real_model():
    model = EmbeddingModel(
        "google/embeddinggemma-2", device="cpu", dtype="float32",
        embedding_dim=768,
    )
    if not model.load():
        pytest.skip("model could not be loaded: %s" % model.load_error)
    return model


def test_real_model_dimensions_and_norm(real_model):
    import numpy as np

    query = real_model.encode_queries(["curtain wall panel edit"])
    document = real_model.encode_documents(
        ["can we edit the panels individually"], ["Curtain wall edit"]
    )
    assert query.shape == (1, 768)
    assert document.shape == (1, 768)
    assert abs(float(np.linalg.norm(query[0])) - 1.0) < 1e-3
    assert abs(float(np.linalg.norm(document[0])) - 1.0) < 1e-3


def test_real_cross_language_retrieval(real_model, tmp_path):
    """Chinese queries must rank the semantically matching English mail first."""

    settings = ServiceSettings(
        product_root=tmp_path,
        db_path=tmp_path / "mail.db",
        model_cache_dir=tmp_path / "models",
        log_path=tmp_path / "log.txt",
        token_path=tmp_path / "token",
    )
    db = Database(settings.db_path, embedding_dim=768)
    indexer = Indexer(db, real_model, settings)
    searcher = HybridSearch(db, real_model, settings)

    mails = [
        dict(entry_id="E1", internet_message_id="m1",
             sender_email="india@example.com", subject="Re: Curtain wall edit",
             body_text="Regarding the curtain wall edit, can we also edit the "
                       "Panels individually?"),
        dict(entry_id="E2", internet_message_id="m2",
             sender_email="reseller@example.com",
             subject="IFC4 certification",
             body_text="We need the IFC4 certification for the reseller."),
        dict(entry_id="E3", internet_message_id="m3",
             sender_email="user@example.com", subject="Material loss",
             body_text="After importing the Revit Link the material is lost."),
        dict(entry_id="D1", internet_message_id="d1",
             sender_email="noise@example.com", subject="Lunch",
             body_text="Lunch is at noon in the canteen."),
    ]
    for mail in mails:
        indexer.index_mail(make_mail(**mail))

    cases = [
        ("之前印度客户问过幕墙玻璃 Panel 能不能单独编辑的邮件", "E1"),
        ("代理商询问 IFC4 认证", "E2"),
        ("Revit Link 导入后材质丢失", "E3"),
    ]
    for query, expected in cases:
        response = searcher.search(query, top_k=5)
        ranked = [r.entry_id for r in response.results]
        print("%s -> %s (expected %s)" % (query, ranked, expected))
        assert response.results, "no results for %r" % query
        assert expected in ranked[:3], (
            "%s not in top 3 for %r: %s" % (expected, query, ranked)
        )
    db.close()


def test_real_long_mail_chunking(real_model, tmp_path):
    settings = ServiceSettings(
        product_root=tmp_path,
        db_path=tmp_path / "mail.db",
        model_cache_dir=tmp_path / "models",
        log_path=tmp_path / "log.txt",
        token_path=tmp_path / "token",
    )
    db = Database(settings.db_path, embedding_dim=768)
    indexer = Indexer(db, real_model, settings)
    # A long thread whose relevant fact appears only near the end.
    filler = "This paragraph is unrelated project coordination text. " * 200
    body = filler + " By the way, GstarBIM used 24GB of memory on the client."
    result = indexer.index_mail(make_mail(
        entry_id="L1", internet_message_id="l1", subject="Long thread",
        body_text=body,
    ))
    assert result.chunk_count >= 2
    searcher = HybridSearch(db, real_model, settings)
    response = searcher.search("GstarBIM 内存占用 24GB")
    assert any(r.entry_id == "L1" for r in response.results)
    db.close()
