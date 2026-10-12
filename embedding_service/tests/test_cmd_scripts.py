"""Static guards for the Windows batch launcher scripts.

cmd.exe has brittle parsing rules. Two mistakes have already broken the
installer once each:

1. an unescaped parenthesis inside an `if (...)` block, which aborts the
   script with "was unexpected at this time"
2. a nested double quote inside a `for /f` command

These tests parse the .cmd files and fail on the patterns that break.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

SERVICE_DIR = Path(__file__).resolve().parent.parent
REPO_ROOT = SERVICE_DIR.parent
CMD_FILES = sorted(SERVICE_DIR.glob("*.cmd"))
INSTALLER_CMD_FILES = sorted(
    (REPO_ROOT / "desktop" / "installer").glob("*.cmd")
)
ALL_CMD_FILES = CMD_FILES + INSTALLER_CMD_FILES


def test_installer_cmd_files_exist():
    names = {p.name for p in INSTALLER_CMD_FILES}
    assert "Install.cmd" in names
    assert "Uninstall.cmd" in names


def test_installer_checks_for_its_ps1():
    """Running from inside a ZIP extracts incompletely; fail with guidance."""

    for name, ps1 in (("Install.cmd", "install.ps1"),
                      ("Uninstall.cmd", "uninstall.ps1")):
        path = REPO_ROOT / "desktop" / "installer" / name
        text = path.read_text(encoding="utf-8", errors="replace")
        assert ps1 in text
        assert "if not exist" in text.lower(), (
            "%s must check that %s exists next to it" % (name, ps1)
        )


def test_cmd_files_exist():
    names = {p.name for p in CMD_FILES}
    for expected in (
        "setup-venv.cmd",
        "download-model.cmd",
        "start-service.cmd",
        "install-service.cmd",
    ):
        assert expected in names, "%s is missing" % expected


@pytest.mark.parametrize("cmd_file", ALL_CMD_FILES, ids=lambda p: p.name)
def test_no_unescaped_parens_in_echo_inside_if(cmd_file):
    """`echo (x)` inside an `if (...)` block breaks cmd.exe parsing."""

    lines = cmd_file.read_text(encoding="utf-8", errors="replace").splitlines()
    depth = 0
    problems = []
    for number, line in enumerate(lines, start=1):
        stripped = line.strip()
        if stripped.startswith("REM") or stripped.startswith("::"):
            continue

        # Track entry/exit of parenthesized blocks.
        opens = stripped.count("(")
        closes = stripped.count(")")
        entering = depth == 0 and re.match(r"^if\b.*\($", stripped)
        was_inside = depth > 0
        if entering or (was_inside and opens > closes):
            depth += opens - closes

        if was_inside and stripped.lower().startswith("echo"):
            # Any unescaped paren in the echoed text is a parse hazard.
            text = stripped[4:]
            if "(" in text or ")" in text:
                if "^(" not in text and "^)" not in text:
                    problems.append((number, stripped))

        if depth > 0 and closes > opens and stripped == ")":
            depth = 0

    assert not problems, (
        "%s echoes unescaped parentheses inside an if block: %s"
        % (cmd_file.name, problems)
    )


@pytest.mark.parametrize("cmd_file", ALL_CMD_FILES, ids=lambda p: p.name)
def test_no_script_closes_without_pause(cmd_file):
    """Every user-facing script must pause so its window does not vanish."""

    text = cmd_file.read_text(encoding="utf-8", errors="replace").lower()
    if cmd_file.name == "start-service.cmd":
        # The service script pauses on exit, not on success.
        assert "pause" in text
    else:
        assert "pause" in text, "%s must pause so the user can read output" % cmd_file.name


@pytest.mark.parametrize("cmd_file", ALL_CMD_FILES, ids=lambda p: p.name)
def test_crlf_line_endings(cmd_file):
    """cmd.exe can mis-handle LF-only files on some systems."""

    raw = cmd_file.read_bytes()
    # Allow files that are entirely LF (git may normalize), but flag a mix.
    has_crlf = b"\r\n" in raw
    has_lone_lf = re.search(rb"(?<!\r)\n", raw) is not None
    assert not (has_crlf and has_lone_lf), (
        "%s mixes CRLF and LF line endings" % cmd_file.name
    )
