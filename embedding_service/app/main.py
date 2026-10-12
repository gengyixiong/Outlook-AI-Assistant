"""Command line entry point for the local semantic search service.

Usage:
    python -m app.main                 # start the server on 127.0.0.1:8765
    python -m app.main --download-model # download/cache the model, then exit
    python -m app.main --check         # load the model and report status
    python -m app.main --index-check   # report index statistics, then exit
"""

from __future__ import annotations

import argparse
import logging
import sys

from .api import create_app
from .auth import load_or_create_token
from .config import ServiceSettings


def configure_logging(settings: ServiceSettings, verbose: bool) -> None:
    settings.ensure_directories()
    handlers = []
    if sys.stderr is not None and hasattr(sys.stderr, "write"):
        try:
            handlers.append(logging.StreamHandler(sys.stderr))
        except Exception:
            pass
    try:
        handlers.append(logging.FileHandler(settings.log_path, encoding="utf-8"))
    except OSError:
        pass
    if not handlers:
        handlers.append(logging.NullHandler())
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        handlers=handlers,
    )
    if not verbose:
        # The download/check CLI prints its own clean progress on stdout;
        # silence the very chatty HTTP and model-loading libraries so the
        # console output stays readable.
        for name in (
            "httpx", "httpcore", "urllib3", "huggingface_hub",
            "filelock", "sentence_transformers", "transformers",
        ):
            logging.getLogger(name).setLevel(logging.WARNING)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Local semantic search service")
    parser.add_argument("--download-model", action="store_true",
                        help="download and cache the embedding model, then exit")
    parser.add_argument("--check", action="store_true",
                        help="load model and exit")
    parser.add_argument("--index-check", action="store_true",
                        help="print index stats and exit")
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--host", default=None)
    parser.add_argument("--port", type=int, default=None)
    args = parser.parse_args(argv)

    # Silence the Hugging Face symlink-mode warning on Windows so the CLI
    # output stays readable. It is informational only.
    import os as _os
    _os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")
    _os.environ.setdefault("HF_HUB_DISABLE_PROGRESS_BARS", "0")
    _os.environ.setdefault("TRANSFORMERS_NO_ADVISORY_WARNINGS", "1")

    settings = ServiceSettings.from_env()
    if args.host:
        settings.host = args.host
    if args.port:
        settings.port = args.port
    configure_logging(settings, args.verbose)

    app = create_app(settings)
    state = app.state.service

    if args.download_model:
        print("=" * 60)
        print("Downloading the embedding model:", settings.model_id)
        print("This is a one-time download of about 1.5 GB.")
        print("Cache location:", settings.model_cache_dir)
        print("=" * 60)
        print()

        def report(message):
            # flush so a cmd.exe window shows progress immediately.
            print(message, flush=True)

        ok = state.model.download(progress=report)
        print()
        if not ok:
            print("FAILED:", state.model.load_error)
            return 3

        print("Verifying that the model loads ...")
        loaded = state.model.load()
        print("model available:", loaded)
        if not loaded:
            print("warning:", state.model.load_error or "")
            return 3

        print("SUCCESS: the model is ready. You can now start the service.")
        return 0

    if args.index_check:
        import json

        print(json.dumps(state.database.stats(), indent=2))
        return 0

    if args.check:
        ok = state.model.load()
        print("model_id:", settings.model_id)
        print("available:", ok)
        print("error:", state.model.load_error or "")
        print("device:", settings.device, "dtype:", settings.dtype)
        return 0 if ok else 2

    token = load_or_create_token(settings.token_path)
    # Re-create the app with the token so requests are authenticated.
    app = create_app(settings, token=token)
    settings.ensure_directories()
    import uvicorn

    logging.getLogger("oaa.api").info(
        "Starting semantic search service on %s:%d",
        settings.host, settings.port,
    )
    uvicorn.run(
        app,
        host=settings.host,
        port=settings.port,
        log_level="info",
        log_config=None,
        access_log=False,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
