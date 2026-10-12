"""Hybrid retrieval: semantic + keyword, fused with Reciprocal Rank Fusion.

Why RRF: cosine similarity and BM25 scores live on different scales and are
not comparable. RRF combines only ranks, which needs no score calibration and
is robust to one retriever returning poor absolute scores.

Mail-level aggregation: a long thread can produce many chunks that all score
well. We aggregate chunk scores into a mail score so one mail cannot occupy
several result slots, and we keep the best-matching chunk as the evidence
snippet.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np

from .config import ServiceSettings
from .database import Database
from .embedding_model import EmbeddingModel, EmbeddingModelUnavailable

logger = logging.getLogger("oaa.search")


@dataclass
class SearchResult:
    mail_id: int
    store_id: str
    entry_id: str
    internet_message_id: str
    subject: str
    sender_name: str
    sender_email: str
    received_at: str
    folder_path: str
    attachment_names: str
    attachment_count: int
    snippet: str
    score: float
    semantic_rank: Optional[int] = None
    keyword_rank: Optional[int] = None


@dataclass
class SearchResponse:
    query: str
    results: List[SearchResult] = field(default_factory=list)
    semantic_available: bool = False
    keyword_available: bool = False
    candidate_count: int = 0
    filter_applied: bool = False
    diagnostics: Dict[str, object] = field(default_factory=dict)


class HybridSearch:
    def __init__(
        self,
        database: Database,
        model: EmbeddingModel,
        settings: ServiceSettings,
    ) -> None:
        self._db = database
        self._model = model
        self._settings = settings

    def search(
        self,
        query: str,
        top_k: Optional[int] = None,
        senders: Optional[Sequence[str]] = None,
        received_from: Optional[str] = None,
        received_through: Optional[str] = None,
        has_attachments: Optional[bool] = None,
    ) -> SearchResponse:
        final_k = top_k or self._settings.final_top_k
        response = SearchResponse(query=query)

        restrict = self._db.mail_ids_for(
            senders=senders,
            received_from=received_from,
            received_through=received_through,
            has_attachments=has_attachments,
        )
        response.filter_applied = restrict is not None
        if restrict is not None and not restrict:
            return response

        semantic = self._semantic_search(query, restrict)
        keyword = self._keyword_search(query, restrict)
        response.semantic_available = semantic is not None
        response.keyword_available = len(keyword) > 0

        fused = self._fuse(semantic or [], keyword)
        response.candidate_count = len(fused)
        response.results = self._materialize(fused, final_k)
        response.diagnostics = {
            "semantic_candidates": len(semantic or []),
            "keyword_candidates": len(keyword),
            "fused_candidates": len(fused),
            "final_count": len(response.results),
        }
        return response

    # ------------------------------------------------------------------
    def _semantic_search(
        self,
        query: str,
        restrict: Optional[Sequence[int]],
    ) -> Optional[List[Tuple[int, int, float]]]:
        """Return (mail_id, chunk_id, cosine) sorted best-first."""

        chunk_ids, mail_ids, matrix = self._db.all_embeddings()
        if matrix.shape[0] == 0:
            return [] if self._model.available else None
        try:
            query_vector = self._model.encode_queries([query])[0]
        except EmbeddingModelUnavailable as ex:
            logger.warning("Semantic search unavailable: %s", ex)
            return None
        if query_vector.shape[0] != matrix.shape[1]:
            logger.warning(
                "Embedding dimension mismatch: query=%d index=%d",
                query_vector.shape[0], matrix.shape[1],
            )
            return None

        scores = matrix @ query_vector
        allowed = set(restrict) if restrict is not None else None

        # Aggregate chunk scores per mail: max score plus a small bonus for
        # additional supporting chunks, so a mail with several relevant chunks
        # ranks slightly above a single-chunk match.
        best: Dict[int, Tuple[float, int, List[float]]] = {}
        for index in range(scores.shape[0]):
            mail_id = mail_ids[index]
            if allowed is not None and mail_id not in allowed:
                continue
            score = float(scores[index])
            entry = best.get(mail_id)
            if entry is None or score > entry[0]:
                previous = entry[2] if entry else []
                best[mail_id] = (score, chunk_ids[index], previous + [score])
            else:
                entry[2].append(score)

        ranked: List[Tuple[int, int, float]] = []
        for mail_id, (top_score, chunk_id, all_scores) in best.items():
            # Small aggregation bonus, capped so it cannot overturn a clearly
            # better match.
            support = min(0.02, 0.005 * (len(all_scores) - 1))
            ranked.append((mail_id, chunk_id, top_score + support))
        ranked.sort(key=lambda item: item[2], reverse=True)
        return ranked[: self._settings.semantic_top_k]

    def _keyword_search(
        self,
        query: str,
        restrict: Optional[Sequence[int]],
    ) -> List[Tuple[int, int, float]]:
        hits = self._db.keyword_search(
            query,
            limit=self._settings.keyword_top_k,
            restrict_mail_ids=restrict,
        )
        ranked: List[Tuple[int, int, float]] = []
        seen_mails = set()
        for hit in hits:
            # Keep one chunk per mail for the keyword ranking; the best chunk
            # is the first because BM25 already sorted ascending.
            if hit.mail_id in seen_mails:
                continue
            seen_mails.add(hit.mail_id)
            ranked.append((hit.mail_id, hit.chunk_id, hit.score))
            if len(ranked) >= self._settings.keyword_top_k:
                break
        return ranked

    def _fuse(
        self,
        semantic: List[Tuple[int, int, float]],
        keyword: List[Tuple[int, int, float]],
    ) -> List[Tuple[int, float, Optional[int], Optional[int], int]]:
        """RRF: sum 1/(k + rank) over both retrievers, per mail."""

        k = self._settings.rrf_k
        scores: Dict[int, float] = {}
        semantic_rank: Dict[int, int] = {}
        keyword_rank: Dict[int, int] = {}
        best_chunk: Dict[int, int] = {}

        for rank, (mail_id, chunk_id, _) in enumerate(semantic, start=1):
            scores[mail_id] = scores.get(mail_id, 0.0) + 1.0 / (k + rank)
            semantic_rank[mail_id] = rank
            best_chunk.setdefault(mail_id, chunk_id)

        for rank, (mail_id, chunk_id, _) in enumerate(keyword, start=1):
            scores[mail_id] = scores.get(mail_id, 0.0) + 1.0 / (k + rank)
            keyword_rank[mail_id] = rank
            best_chunk.setdefault(mail_id, chunk_id)

        fused = [
            (
                mail_id,
                score,
                semantic_rank.get(mail_id),
                keyword_rank.get(mail_id),
                best_chunk.get(mail_id, 0),
            )
            for mail_id, score in scores.items()
        ]
        fused.sort(key=lambda item: item[1], reverse=True)
        return fused

    def _materialize(
        self,
        fused: List[Tuple[int, float, Optional[int], Optional[int], int]],
        limit: int,
    ) -> List[SearchResult]:
        results: List[SearchResult] = []
        for mail_id, score, sem_rank, kw_rank, chunk_id in fused[:limit]:
            row = self._db.get_mail(mail_id)
            if row is None:
                continue
            snippet = self._snippet(row["id"], chunk_id)
            results.append(SearchResult(
                mail_id=mail_id,
                store_id=row["store_id"],
                entry_id=row["entry_id"],
                internet_message_id=row["internet_message_id"],
                subject=row["subject"],
                sender_name=row["sender_name"],
                sender_email=row["sender_email"],
                received_at=row["received_at"],
                folder_path=row["folder_path"],
                attachment_names=row["attachment_names"],
                attachment_count=row["attachment_count"],
                snippet=snippet,
                score=round(float(score), 6),
                semantic_rank=sem_rank,
                keyword_rank=kw_rank,
            ))
        return results

    def _snippet(self, mail_id: int, chunk_id: int) -> str:
        if chunk_id:
            with self._db._lock:  # noqa: SLF001 - internal, same module family
                row = self._db._connection.execute(  # noqa: SLF001
                    "SELECT text FROM chunks WHERE id = ?", (chunk_id,)
                ).fetchone()
            if row is not None and row["text"]:
                return _shorten(row["text"])
        with self._db._lock:  # noqa: SLF001
            row = self._db._connection.execute(  # noqa: SLF001
                "SELECT text FROM chunks WHERE mail_id = ? "
                "ORDER BY chunk_index LIMIT 1",
                (mail_id,),
            ).fetchone()
        return _shorten(row["text"]) if row is not None else ""


def _shorten(text: str, limit: int = 320) -> str:
    value = " ".join((text or "").split())
    if len(value) <= limit:
        return value
    return value[:limit] + "…"
