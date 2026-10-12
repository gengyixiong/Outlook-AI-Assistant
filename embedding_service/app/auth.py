"""Loopback-only service authentication.

The service binds to 127.0.0.1 so it is unreachable from the LAN. A shared
token adds a second layer so that other local processes cannot freely read
mail content from the database. The token is stored in a file readable only by
the current user (mode 0600 on POSIX; on Windows the per-user profile ACL
already restricts it).
"""

from __future__ import annotations

import hmac
import os
import secrets
from pathlib import Path
from typing import Optional


def load_or_create_token(path: Path) -> str:
    path = Path(path)
    if path.exists():
        token = path.read_text(encoding="utf-8").strip()
        if token:
            return token
    token = secrets.token_urlsafe(32)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(token, encoding="utf-8")
    try:
        os.chmod(path, 0o600)
    except OSError:
        # Windows does not implement POSIX modes; the user profile ACL applies.
        pass
    return token


def tokens_match(expected: str, provided: Optional[str]) -> bool:
    if not expected:
        return False
    if not provided:
        return False
    return hmac.compare_digest(expected, provided)
