"""Tests for the model download progress reporting.

A previous version called huggingface_hub's snapshot_download without a
working tqdm_class, which raised "'function' object is not iterable" and made
the download step fail. These tests pin the contract: make() must return a
real tqdm-compatible class, and progress must be forwarded.
"""

from __future__ import annotations

import time

import pytest

from app.embedding_model import _FallbackProgress, _ProgressReporter


def test_make_returns_a_class_not_an_instance():
    """huggingface_hub calls tqdm_class(...) itself, so it must be a class."""

    cls = _ProgressReporter.make(lambda _m: None)
    assert isinstance(cls, type), "make() must return a class"
    # It must be constructible the way the hub constructs it.
    bar = cls(total=3, desc="files")
    bar.update(1)
    bar.close()


def test_progress_is_forwarded():
    lines = []
    cls = _ProgressReporter.make(lines.append)
    bar = cls(total=2, desc="model.safetensors")
    # Force two updates with a sleep so the throttle (2s) does not skip both.
    bar.update(1)
    time.sleep(2.05)
    bar.update(1)
    bar.close()
    assert lines, "progress lines must be emitted"
    assert any("model.safetensors" in line for line in lines)


def test_progress_reporter_survives_bad_callback():
    """A throwing callback must never break the download."""

    def boom(_message):
        raise RuntimeError("callback gone")

    cls = _ProgressReporter.make(boom)
    bar = cls(total=1, desc="x")
    time.sleep(2.05)
    bar.update(1)  # must not raise
    bar.close()


def test_fallback_progress_records_report():
    lines = []
    _FallbackProgress._oaa_report = lines.append
    cls = _FallbackProgress.with_report(lines.append)
    # Simulate the hub constructing and updating it.
    bar = cls(total=2, desc="f")
    bar.update(1)
    bar.close()
    assert lines
    _FallbackProgress._oaa_report = None
