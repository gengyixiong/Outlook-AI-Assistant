"""Version single-source-of-truth tests.

The product version must come from one file. Earlier it was hardcoded in
three places (AssemblyInfo.cs, build.ps1, app/__init__.py) and had already
drifted (the package said 0.6.8 while the service reported 0.5.0).
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

SERVICE_DIR = Path(__file__).resolve().parent.parent
REPO_ROOT = SERVICE_DIR.parent
VERSION_FILE = REPO_ROOT / "VERSION"


def test_version_file_exists_and_is_semver():
    assert VERSION_FILE.is_file(), "VERSION file is missing"
    value = VERSION_FILE.read_text(encoding="utf-8").strip()
    assert re.match(r"^\d+\.\d+\.\d+$", value), "VERSION must be X.Y.Z"


def test_service_reports_the_version_file():
    import app

    expected = VERSION_FILE.read_text(encoding="utf-8").strip()
    assert app.__version__ == expected, (
        "the service version must come from VERSION"
    )


def test_assembly_info_does_not_hardcode_version():
    """The DLL version is generated from VERSION at build time."""

    path = (REPO_ROOT / "desktop" / "src" / "OutlookAiAssistant"
            / "Properties" / "AssemblyInfo.cs")
    text = path.read_text(encoding="utf-8")
    assert "AssemblyVersion(" not in text, (
        "AssemblyVersion must not be hardcoded; it is generated from VERSION"
    )
    assert "AssemblyFileVersion(" not in text


def test_build_script_reads_version_file():
    path = REPO_ROOT / "desktop" / "build" / "build.ps1"
    text = path.read_text(encoding="utf-8")
    assert "VERSION" in text
    assert "$productVersion" in text
    # The zip name must be derived, not hardcoded to a specific version.
    assert re.search(r'Outlook-AI-Assistant-v\$productVersion\.zip', text), (
        "the package name must come from VERSION"
    )


def test_setup_script_exists_and_reads_version():
    setup = REPO_ROOT / "desktop" / "installer" / "setup.ps1"
    assert setup.is_file(), "unified setup.ps1 is missing"
    text = setup.read_text(encoding="utf-8-sig")
    assert "VERSION" in text
    # It must record what it installed for the next run to compare.
    assert ".installed-version" in text
    # Key safety property: never destroy user data.
    assert "保留" in text or "preserve" in text.lower()
