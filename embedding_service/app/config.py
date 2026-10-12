"""Service configuration.

All paths default under %LOCALAPPDATA%\\OutlookAiAssistant so that model
files, the database, logs and settings stay separated. macOS/Linux paths are
only used for the test suite.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


def default_product_root() -> Path:
    """Return the per-user product directory shared with the C# add-in."""

    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA")
        if base:
            return Path(base) / "OutlookAiAssistant"
    # Non-Windows fallback used by tests and developer machines.
    xdg = os.environ.get("XDG_DATA_HOME")
    if xdg:
        return Path(xdg) / "OutlookAiAssistant"
    return Path.home() / ".local" / "share" / "OutlookAiAssistant"


def apply_huggingface_cache_environment(model_dir: Path) -> None:
    """Point Hugging Face at our own model directory.

    Without this, snapshot_download and SentenceTransformer silently use the
    default cache (%USERPROFILE%\\.cache\\huggingface), so the configured
    model directory stays empty even though the model is downloaded. Setting
    HF_HOME here makes everything land under our own path.

    Only set it when the caller has not already chosen a location, so an
    explicit HF_HOME still wins.
    """

    try:
        model_dir = Path(model_dir)
        hf_root = model_dir / "huggingface"
        hf_root.mkdir(parents=True, exist_ok=True)
        # HF_HOME is the root that contains hub/ and other caches.
        os.environ.setdefault("HF_HOME", str(hf_root))
    except OSError:
        # A read-only or unusual path must not break startup; the default
        # cache will be used instead.
        pass


# The model officially requires these prompts. Queries and documents use
# different (asymmetric) prefixes; document titles are injected manually
# because prompt_name="Document" only supplies "title: none".
QUERY_PROMPT = "task: search result | query: "
DOCUMENT_TITLE_TEMPLATE = "title: {title} | text: "


@dataclass
class ServiceSettings:
    """Runtime settings for the embedding service."""

    model_id: str = "google/embeddinggemma-2"
    # 768 is the native output dimension. MRL truncation (512/256/128) is
    # supported by the model but disabled by default to keep quality high.
    embedding_dim: int = 768
    # CPU float32 is the safe default on machines without a CUDA GPU.
    # The model is incompatible with float16 on every backend.
    device: str = "auto"
    dtype: str = "auto"
    host: str = "127.0.0.1"
    port: int = 8765

    product_root: Path = field(default_factory=default_product_root)
    db_path: Optional[Path] = None
    model_cache_dir: Optional[Path] = None
    log_path: Optional[Path] = None
    token_path: Optional[Path] = None

    # Chunking. EmbeddingGemma 2 has an 8K token window; we keep chunks well
    # inside it so that retrieval stays precise.
    chunk_target_tokens: int = 512
    chunk_min_tokens: int = 300
    chunk_max_tokens: int = 800
    chunk_overlap_tokens: int = 64

    # Retrieval candidate pool sizes before fusion.
    semantic_top_k: int = 40
    keyword_top_k: int = 40
    final_top_k: int = 10

    # RRF constant. 60 is the value from the original RRF paper.
    rrf_k: int = 60

    def __post_init__(self) -> None:
        self.product_root = Path(self.product_root)
        if self.db_path is None:
            self.db_path = self.product_root / "semantic-index" / "mail.db"
        if self.model_cache_dir is None:
            self.model_cache_dir = self.product_root / "models"
        if self.log_path is None:
            self.log_path = self.product_root / "logs" / "semantic-service.log"
        if self.token_path is None:
            self.token_path = self.product_root / "semantic-service-token"
        self.db_path = Path(self.db_path)
        self.model_cache_dir = Path(self.model_cache_dir)
        self.log_path = Path(self.log_path)
        self.token_path = Path(self.token_path)

    def ensure_directories(self) -> None:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.model_cache_dir.mkdir(parents=True, exist_ok=True)
        self.log_path.parent.mkdir(parents=True, exist_ok=True)
        # Point Hugging Face at our own model directory so the configured path
        # is where the model actually lives.
        apply_huggingface_cache_environment(self.model_cache_dir)

    @classmethod
    def from_env(cls) -> "ServiceSettings":
        settings = cls()
        if os.environ.get("OAA_MODEL_ID"):
            settings.model_id = os.environ["OAA_MODEL_ID"]
        if os.environ.get("OAA_DEVICE"):
            settings.device = os.environ["OAA_DEVICE"]
        if os.environ.get("OAA_DTYPE"):
            settings.dtype = os.environ["OAA_DTYPE"]
        if os.environ.get("OAA_PORT"):
            settings.port = int(os.environ["OAA_PORT"])
        if os.environ.get("OAA_PRODUCT_ROOT"):
            settings.product_root = Path(os.environ["OAA_PRODUCT_ROOT"])
            settings.db_path = None
            settings.model_cache_dir = None
            settings.log_path = None
            settings.token_path = None
            settings.__post_init__()
        return settings