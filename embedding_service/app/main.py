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
    handlers = [logging.StreamHandler(sys.stderr)]
    try:
        handlers.append(logging.FileHandler(settings.log_path, encoding="utf-8"))
    except OSError:
        pass
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        handlers=handlers,
    )


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

    settings = ServiceSettings.from_env()
    if args.host:
        settings.host = args.host
    if args.port:
        settings.port = args.port
    configure_logging(settings, args.verbose)

    app = create_app(settings)
    state = app.state.service

    if args.download_model:
        print("Downloading embedding model:", settings.model_id)
        print("This is a one-time download (about 1.5 GB) and is cached at:")
        print("  ", settings.model_cache_dir)
        ok = state.model.download()
        if not ok:
            print("FAILED:", state.model.load_error)
            return 3
        loaded = state.model.load()
        print("downloaded: True")
        print("available:", loaded)
        if not loaded:
            print("warning:", state.model.load_error or "")
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
        access_log=False,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
