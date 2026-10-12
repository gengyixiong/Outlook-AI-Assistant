"""FastAPI application exposing the local semantic search API.

Endpoints
---------
GET  /health                 service + model + index status
POST /index/mail             index or refresh one mail
POST /index/delete           remove one mail from the index
DELETE /index                delete the whole local index
POST /search                 hybrid search, returns real indexed mails
GET  /stats                  index statistics
POST /index/meta             record sync metadata (last sync time, model)

Every request except /health requires the bearer token from the token file.
The service never returns mail content for arbitrary ids other than through
/search, and never returns embeddings.
"""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import List, Optional

from fastapi import Depends, FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field

from .config import ServiceSettings
from .database import Database
from .embedding_model import EmbeddingModel, EmbeddingModelUnavailable
from .indexing import Indexer, IndexMailRequest
from .search import HybridSearch

logger = logging.getLogger("oaa.api")


# ----------------------------------------------------------------------
# request / response models
# ----------------------------------------------------------------------
class MailPayload(BaseModel):
    store_id: str = ""
    entry_id: str = ""
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
    force_reembed: bool = False


class DeletePayload(BaseModel):
    store_id: str = ""
    entry_id: str = ""


class SearchPayload(BaseModel):
    query: str
    top_k: Optional[int] = None
    senders: List[str] = Field(default_factory=list)
    received_from: Optional[str] = None
    received_through: Optional[str] = None
    has_attachments: Optional[bool] = None


class MetaPayload(BaseModel):
    last_sync_at: Optional[str] = None
    model_id: Optional[str] = None


class ResultModel(BaseModel):
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


class SearchResponseModel(BaseModel):
    query: str
    results: List[ResultModel]
    semantic_available: bool
    keyword_available: bool
    candidate_count: int
    filter_applied: bool
    diagnostics: dict


# ----------------------------------------------------------------------
# application factory
# ----------------------------------------------------------------------
class ServiceState:
    def __init__(self, settings: ServiceSettings) -> None:
        self.settings = settings
        self.database = Database(settings.db_path, settings.embedding_dim)
        self.model = EmbeddingModel(
            settings.model_id,
            device=settings.device,
            dtype=settings.dtype,
            embedding_dim=settings.embedding_dim,
        )
        self.indexer = Indexer(self.database, self.model, settings)
        self.search = HybridSearch(self.database, self.model, settings)
        self.index_busy = False


def create_app(
    settings: Optional[ServiceSettings] = None,
    token: Optional[str] = None,
) -> FastAPI:
    settings = settings or ServiceSettings.from_env()
    settings.ensure_directories()

    state = ServiceState(settings)
    state.database.set_meta("model_id", settings.model_id)

    auth_token = token or settings.token_path.read_text(
        encoding="utf-8"
    ).strip() if settings.token_path.exists() else token

    app = FastAPI(title="Outlook AI Assistant Semantic Search", version="0.5.0")
    app.state.service = state

    def require_token(request: Request) -> None:
        if not auth_token:
            # No token configured (developer/test mode): allow loopback only.
            client = request.client.host if request.client else ""
            if client not in ("127.0.0.1", "::1", "localhost", "testclient"):
                raise HTTPException(status_code=403, detail="forbidden")
            return
        header = request.headers.get("authorization", "")
        provided = header[7:].strip() if header.lower().startswith("bearer ") \
            else ""
        from .auth import tokens_match

        if not tokens_match(auth_token, provided):
            raise HTTPException(status_code=401, detail="unauthorized")

    @app.get("/health")
    def health() -> dict:
        stats = state.database.stats()
        return {
            "status": "ok",
            "model_id": settings.model_id,
            "model_available": state.model.available,
            "model_error": state.model.load_error,
            "device": settings.device,
            "dtype": settings.dtype,
            "embedding_dim": settings.embedding_dim,
            "index_busy": state.index_busy,
            "index": stats,
            "version": "0.5.0",
            # Local path hints so the add-in can show and open them. These are
            # filesystem paths only; they contain no mail content.
            "paths": {
                "model_dir": str(settings.model_cache_dir),
                "data_dir": str(settings.product_root),
                "index_dir": str(settings.db_path.parent),
                "database": str(settings.db_path),
                "log_dir": str(settings.log_path.parent),
                "service_dir": str(Path(__file__).resolve().parent.parent),
            },
        }

    @app.get("/model/status", dependencies=[Depends(require_token)])
    def model_status() -> dict:
        """Report whether the model can be used, and why not if it cannot."""

        # Attempt a load so the caller gets the real reason (network, gated
        # repo, missing runtime) instead of an empty error string.
        if not state.model.available:
            state.model.load()
        return {
            "model_id": settings.model_id,
            "available": state.model.available,
            "error": state.model.load_error or "",
        }

    @app.post("/model/download", dependencies=[Depends(require_token)])
    def download_model() -> dict:
        """Explicitly download and cache the model.

        This is the step a fresh machine needs before indexing can work. It
        only downloads; the add-in can then call /health to see availability.
        """

        ok = state.model.download()
        if not ok:
            raise HTTPException(
                status_code=502,
                detail=state.model.load_error
                or "Model download failed.",
            )
        loaded = state.model.load()
        return {
            "downloaded": True,
            "available": loaded,
            "error": state.model.load_error or "",
        }

    @app.post("/index/mail", dependencies=[Depends(require_token)])
    def index_mail(payload: MailPayload) -> dict:
        if not payload.store_id and not payload.entry_id:
            raise HTTPException(
                status_code=400, detail="store_id or entry_id is required"
            )
        # Fail before touching the database when the model is unusable, so a
        # broken setup cannot fill the index with vector-less rows.
        if not state.model.available and not state.model.load():
            raise HTTPException(
                status_code=503,
                detail=state.model.load_error
                or "Embedding model is unavailable. Download it first.",
            )
        payload_data = payload.model_dump()
        force_reembed = bool(payload_data.pop("force_reembed", False))
        request = IndexMailRequest(**payload_data)
        try:
            result = state.indexer.index_mail(
                request, force_reembed=force_reembed
            )
        except EmbeddingModelUnavailable as ex:
            # The model disappeared mid-build; stop with a clear status.
            raise HTTPException(status_code=503, detail=str(ex))
        except Exception as ex:  # keep the add-in alive on a single bad mail
            logger.exception("Indexing failed for entry %s", payload.entry_id)
            raise HTTPException(status_code=500, detail=str(ex))
        return {
            "mail_id": result.mail_id,
            "chunk_count": result.chunk_count,
            "embedded": result.embedded,
            "content_hash": result.content_hash,
            "skipped": result.skipped,
            "reason": result.reason,
        }

    @app.post("/index/delete", dependencies=[Depends(require_token)])
    def delete_mail(payload: DeletePayload) -> dict:
        deleted = state.indexer.delete_mail(payload.store_id, payload.entry_id)
        return {"deleted": deleted}

    @app.delete("/index", dependencies=[Depends(require_token)])
    def clear_index() -> dict:
        state.database.clear()
        return {"cleared": True}

    @app.get("/stats", dependencies=[Depends(require_token)])
    def stats() -> dict:
        return state.database.stats()

    @app.post("/index/meta", dependencies=[Depends(require_token)])
    def update_meta(payload: MetaPayload) -> dict:
        if payload.last_sync_at:
            state.database.set_meta("last_sync_at", payload.last_sync_at)
        if payload.model_id:
            state.database.set_meta("model_id", payload.model_id)
        return {"ok": True}

    @app.post(
        "/search",
        response_model=SearchResponseModel,
        dependencies=[Depends(require_token)],
    )
    def search(payload: SearchPayload) -> SearchResponseModel:
        if not payload.query.strip():
            raise HTTPException(status_code=400, detail="query is required")
        try:
            response = state.search.search(
                payload.query,
                top_k=payload.top_k,
                senders=payload.senders,
                received_from=payload.received_from,
                received_through=payload.received_through,
                has_attachments=payload.has_attachments,
            )
        except Exception as ex:
            logger.exception("Search failed")
            raise HTTPException(status_code=500, detail=str(ex))
        return SearchResponseModel(
            query=response.query,
            results=[
                ResultModel(
                    mail_id=r.mail_id,
                    store_id=r.store_id,
                    entry_id=r.entry_id,
                    internet_message_id=r.internet_message_id,
                    subject=r.subject,
                    sender_name=r.sender_name,
                    sender_email=r.sender_email,
                    received_at=r.received_at,
                    folder_path=r.folder_path,
                    attachment_names=r.attachment_names,
                    attachment_count=r.attachment_count,
                    snippet=r.snippet,
                    score=r.score,
                    semantic_rank=r.semantic_rank,
                    keyword_rank=r.keyword_rank,
                )
                for r in response.results
            ],
            semantic_available=response.semantic_available,
            keyword_available=response.keyword_available,
            candidate_count=response.candidate_count,
            filter_applied=response.filter_applied,
            diagnostics=response.diagnostics,
        )

    @app.exception_handler(Exception)
    async def unhandled(request: Request, exc: Exception) -> JSONResponse:
        logger.exception("Unhandled service error")
        return JSONResponse(
            status_code=500, content={"detail": "internal error"}
        )

    return app
