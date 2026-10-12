"""EmbeddingGemma 2 text encoder wrapper.

The model is loaded text-only (vision and audio encoders disabled) and runs on
CPU with float32 by default. ``float16`` is never used because the model's
activation range exceeds float16's dynamic range and silently degrades.

Queries and documents use the model's official asymmetric prompts, and
documents embed the real email subject as their title, which the model card
says significantly improves retrieval accuracy.
"""

from __future__ import annotations

import logging
import threading
from typing import List, Optional, Sequence

import numpy as np

from .config import DOCUMENT_TITLE_TEMPLATE, QUERY_PROMPT

logger = logging.getLogger("oaa.embedding")


class EmbeddingModelUnavailable(RuntimeError):
    """Raised when the model cannot be loaded or used."""


class EmbeddingModel:
    """Lazy, thread-safe wrapper around SentenceTransformer."""

    def __init__(
        self,
        model_id: str,
        device: str = "cpu",
        dtype: str = "float32",
        embedding_dim: int = 768,
    ) -> None:
        self.model_id = model_id
        self.device = device
        self.dtype = dtype
        self.embedding_dim = embedding_dim
        self._model = None
        self._load_error: Optional[str] = None
        self._lock = threading.RLock()

    # ------------------------------------------------------------------
    @property
    def available(self) -> bool:
        return self._model is not None

    @property
    def load_error(self) -> Optional[str]:
        return self._load_error

    def load(self) -> bool:
        """Load the model. Returns False when the runtime is missing."""

        with self._lock:
            if self._model is not None:
                return True
            try:
                import torch
                from sentence_transformers import SentenceTransformer
            except Exception as ex:  # pragma: no cover - environment dependent
                self._load_error = (
                    "Python embedding runtime is not installed: %s" % ex
                )
                logger.error(self._load_error)
                return False

            try:
                resolved_dtype = self._resolve_dtype(torch)
                # Disable the vision and audio encoders: this service only
                # embeds text. config_kwargs is the officially documented way
                # to skip those modules (270M vs 740M parameters).
                self._model = SentenceTransformer(
                    self.model_id,
                    device=self.device,
                    config_kwargs={
                        "vision_config": None,
                        "audio_config": None,
                    },
                    model_kwargs={"dtype": resolved_dtype},
                )
                self._load_error = None
                logger.info(
                    "Loaded embedding model %s on %s (%s)",
                    self.model_id, self.device, resolved_dtype,
                )
                return True
            except Exception as ex:  # pragma: no cover - environment dependent
                self._model = None
                self._load_error = "Failed to load embedding model: %s" % ex
                logger.error(self._load_error)
                return False

    def _resolve_dtype(self, torch_module):
        """Map the configured dtype to a torch dtype.

        Only bfloat16 and float32 are valid; float16 is rejected explicitly.
        CPU always uses float32 because bfloat16 is slow without native
        support.
        """

        requested = (self.dtype or "float32").lower()
        if requested == "float16":
            logger.warning("float16 is unsafe for EmbeddingGemma 2; using float32")
            requested = "float32"
        if self.device == "cpu":
            requested = "float32"
        if requested in ("bfloat16", "bf16"):
            return torch_module.bfloat16
        return torch_module.float32

    # ------------------------------------------------------------------
    def encode_queries(
        self,
        queries: Sequence[str],
        normalize: bool = True,
    ) -> np.ndarray:
        if not self.load():
            raise EmbeddingModelUnavailable(
                self._load_error or "Embedding model is unavailable."
            )
        with self._lock:
            vectors = self._model.encode(
                list(queries),
                prompt=QUERY_PROMPT,
                normalize_embeddings=normalize,
                convert_to_numpy=True,
                show_progress_bar=False,
            )
        return _truncate_and_normalize(vectors, self.embedding_dim)

    def encode_documents(
        self,
        documents: Sequence[str],
        titles: Optional[Sequence[str]] = None,
        normalize: bool = True,
    ) -> np.ndarray:
        if not self.load():
            raise EmbeddingModelUnavailable(
                self._load_error or "Embedding model is unavailable."
            )
        prompts = self._document_prompts(documents, titles)
        with self._lock:
            vectors = self._model.encode(
                list(documents),
                prompt=prompts,
                normalize_embeddings=normalize,
                convert_to_numpy=True,
                show_progress_bar=False,
            )
        return _truncate_and_normalize(vectors, self.embedding_dim)

    @staticmethod
    def _document_prompts(
        documents: Sequence[str],
        titles: Optional[Sequence[str]],
    ) -> List[str]:
        """Build per-document prompts embedding the real email subject.

        The model card notes that ``prompt_name="Document"`` only supplies
        ``title: none``; titled documents must be formatted manually.
        """

        result: List[str] = []
        for index in range(len(documents)):
            title = ""
            if titles is not None and index < len(titles):
                title = (titles[index] or "").strip().replace("|", " ")
            if not title:
                title = "none"
            result.append(DOCUMENT_TITLE_TEMPLATE.format(title=title))
        return result


def _truncate_and_normalize(
    vectors: np.ndarray,
    embedding_dim: int,
) -> np.ndarray:
    """Apply MRL truncation then re-normalize (slicing breaks unit length)."""

    matrix = np.asarray(vectors, dtype=np.float32)
    if matrix.ndim == 1:
        matrix = matrix.reshape(1, -1)
    if embedding_dim and matrix.shape[1] > embedding_dim:
        matrix = matrix[:, :embedding_dim]
    norms = np.linalg.norm(matrix, axis=1, keepdims=True)
    norms[norms == 0] = 1.0
    return (matrix / norms).astype(np.float32)
