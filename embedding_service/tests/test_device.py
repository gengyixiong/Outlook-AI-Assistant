
"""Device selection tests: NVIDIA GPU when available, CPU otherwise."""

from __future__ import annotations

import pytest

from app.embedding_model import detect_device


def test_explicit_cpu_is_respected():
    device, reason = detect_device("cpu")
    assert device == "cpu"
    assert "configuration" in reason


def test_auto_falls_back_to_cpu_without_cuda(monkeypatch):
    """No CUDA (the common case) must not raise and must pick CPU."""

    import sys, types

    fake = types.ModuleType("torch")
    fake.cuda = types.SimpleNamespace(is_available=lambda: False)
    monkeypatch.setitem(sys.modules, "torch", fake)

    device, reason = detect_device("auto")
    assert device == "cpu"
    # A CPU-only torch build is the common cause; the message must say so
    # instead of a vague "no GPU", so the fix is obvious.
    assert "cpu-only" in reason.lower() or "cuda" in reason.lower()


def test_auto_uses_cuda_when_available(monkeypatch):
    import sys, types

    fake = types.ModuleType("torch")
    fake.cuda = types.SimpleNamespace(
        is_available=lambda: True,
        get_device_name=lambda index: "NVIDIA GeForce RTX 4090",
    )
    monkeypatch.setitem(sys.modules, "torch", fake)

    device, reason = detect_device("auto")
    assert device == "cuda"
    assert "RTX 4090" in reason


def test_cuda_requested_but_missing_falls_back(monkeypatch):
    import sys, types

    fake = types.ModuleType("torch")
    fake.cuda = types.SimpleNamespace(is_available=lambda: False)
    monkeypatch.setitem(sys.modules, "torch", fake)

    device, reason = detect_device("cuda")
    assert device == "cpu"
    assert "not available" in reason.lower()


def test_missing_torch_falls_back_to_cpu(monkeypatch):
    import sys

    monkeypatch.setitem(sys.modules, "torch", None)
    device, reason = detect_device("auto")
    assert device == "cpu"
