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
from typing import Callable, List, Optional, Sequence

import numpy as np

from .config import DOCUMENT_TITLE_TEMPLATE, QUERY_PROMPT

logger = logging.getLogger("oaa.embedding")


class EmbeddingModelUnavailable(RuntimeError):
    """Raised when the model cannot be loaded or used."""


class _ProgressReporter:
    """Factory for a tqdm subclass that forwards progress to a callback.

    huggingface_hub uses ``tqdm_class`` as a *class* (it calls
    ``cls(...)`` and sometimes wraps it), so this must return a real
    subclass, not an instance. The callback is carried on the class because
    the hub constructs the bars itself.
    """

    @staticmethod
    def make(report: Optional[Callable[[str], None]]):
        try:
            from tqdm.auto import tqdm as _tqdm
        except Exception:  # pragma: no cover - tqdm ships with HF hub
            return _FallbackProgress.with_report(report)

        import time as _time

        class _ForwardingTqdm(_tqdm):  # type: ignore[misc]
            _oaa_report = report
            _oaa_last = 0.0
            _OAA_INTERVAL = 2.0

            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self._oaa_name = (
                    args[1] if len(args) > 1
                    else kwargs.get("desc", "") or "files"
                )
                self._oaa_last = 0.0

            def update(self, n=1):
                super().update(n)
                self._oaa_maybe_report()

            def _oaa_maybe_report(self):
                if self._oaa_report is None:
                    return
                now = _time.time()
                if now - self._oaa_last < _ForwardingTqdm._OAA_INTERVAL:
                    return
                type(self)._oaa_last = now
                total = getattr(self, "total", None)
                done = getattr(self, "n", 0)
                if total:
                    percent = 100.0 * done / total
                    try:
                        self._oaa_report(
                            "  %s: %d/%d (%.1f%%)"
                            % (self._oaa_name, done, total, percent)
                        )
                    except Exception:
                        pass

        return _ForwardingTqdm


class _FallbackProgress:
    """Minimal tqdm replacement used when tqdm is not installed."""

    _oaa_report = None

    def __init__(self, *args, **kwargs):
        self.total = kwargs.get("total")
        self.n = 0
        self.desc = kwargs.get("desc", "")

    @classmethod
    def with_report(cls, report):
        return cls

    def update(self, n=1):
        self.n += n
        report = type(self)._oaa_report
        if report and self.total:
            try:
                report("  %s: %d/%d" % (self.desc or "files", self.n, self.total))
            except Exception:
                pass

    def close(self):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    def set_description(self, value):
        self.desc = value


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
                self._load_error = self._humanize_load_error(ex)
                logger.error(self._load_error)
                return False

    def download(self, progress: Optional[Callable[[str], None]] = None) -> bool:
        """Download and cache the model without keeping it loaded.

        This is the explicit download step the add-in exposes so a fresh
        machine does not silently fail every indexing attempt. ``progress`` is
        called with human-readable status lines so the CLI and the UI can show
        what is happening on a multi-minute download. Returns True when the
        snapshot is present locally. Errors are reported through
        ``load_error``.
        """

        with self._lock:
            try:
                from huggingface_hub import snapshot_download
            except Exception as ex:  # pragma: no cover
                self._load_error = (
                    "huggingface_hub is not installed: %s" % ex
                )
                return False

            report = progress or (lambda _message: None)
            try:
                report("Checking local cache for %s ..." % self.model_id)
                # tqdm_class writes a live progress bar to stderr. We also
                # forward throttled plain status lines, which are what a
                # cmd.exe window and the UI log display reliably (\r-based
                # bars do not render well in every host).
                _FallbackProgress._oaa_report = report
                reporter = _ProgressReporter.make(report)
                # One worker keeps a single, readable progress bar. Parallel
                # workers each print their own bar and overwrite each other in
                # a plain console window.
                snapshot_download(
                    repo_id=self.model_id,
                    tqdm_class=reporter,
                    max_workers=1,
                )
                self._load_error = None
                report("Download complete.")
                return True
            except Exception as ex:  # pragma: no cover - network dependent
                self._load_error = self._humanize_load_error(ex)
                logger.error(self._load_error)
                report("FAILED: %s" % self._load_error)
                return False

    def _humanize_load_error(self, ex: Exception) -> str:
        """Turn a raw model/hub error into an actionable message."""

        text = str(ex)
        lowered = text.lower()
        if "401" in text or "403" in text or "gated" in lowered \
                or "invalid username or password" in lowered \
                or "unauthorized" in lowered:
            return (
                "Cannot download the embedding model (%s). The repository "
                "may be gated or require a Hugging Face token. Run the "
                "download-model step with network access, or set HF_TOKEN. "
                "Original error: %s" % (self.model_id, text)
            )
        if "connection" in lowered or "timed out" in lowered \
                or "couldn't connect" in lowered or "network" in lowered:
            return (
                "Cannot reach Hugging Face to download the embedding model "
                "(%s). Check the network or proxy, then retry the model "
                "download. Original error: %s" % (self.model_id, text)
            )
        if "no module named" in lowered or "modulenotfound" in lowered:
            return (
                "The embedding runtime is incomplete: %s. Re-run "
                "setup-venv.cmd." % text
            )
        return "Failed to load embedding model: %s" % text

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
