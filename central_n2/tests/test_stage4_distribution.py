from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

from core.updater import SemVersion, UpdateManager


class FakeResponse:
    def __init__(self, payload: bytes):
        self.payload = payload
        self.offset = 0

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def read(self, size: int = -1) -> bytes:
        if self.offset >= len(self.payload):
            return b""
        if size < 0:
            size = len(self.payload) - self.offset
        chunk = self.payload[self.offset : self.offset + size]
        self.offset += len(chunk)
        return chunk


def test_semver_prerelease_ordering():
    assert SemVersion.parse("5.1.0-beta.1") < SemVersion.parse("5.1.0")
    assert SemVersion.parse("5.1.0-beta.2") > SemVersion.parse("5.1.0-beta.1")
    assert SemVersion.parse("v6.0.0") > SemVersion.parse("5.99.99")


def test_updater_rejects_asset_path_traversal(tmp_path: Path):
    manager = UpdateManager("owner/repo", "5.0.0")
    with pytest.raises(ValueError):
        manager.download_asset(
            {
                "name": "../CentralN2.exe",
                "browser_download_url": "https://example.test/file",
            },
            tmp_path,
        )


def test_updater_download_is_verified_and_atomic(
    tmp_path: Path,
    monkeypatch,
):
    payload = b"verified-central-n2"
    digest = hashlib.sha256(payload).hexdigest()

    monkeypatch.setattr(
        "urllib.request.urlopen",
        lambda request, timeout=None: FakeResponse(payload),
    )

    manager = UpdateManager("owner/repo", "5.0.0")
    target = manager.download_asset(
        {
            "name": "CentralN2.zip",
            "browser_download_url": "https://example.test/CentralN2.zip",
            "size": len(payload),
            "digest": f"sha256:{digest}",
        },
        tmp_path,
    )

    assert target.read_bytes() == payload
    assert not list(tmp_path.glob("*.part"))


def test_updater_removes_partial_file_on_digest_failure(
    tmp_path: Path,
    monkeypatch,
):
    payload = b"tampered"
    monkeypatch.setattr(
        "urllib.request.urlopen",
        lambda request, timeout=None: FakeResponse(payload),
    )

    manager = UpdateManager("owner/repo", "5.0.0")
    with pytest.raises(ValueError, match="SHA-256"):
        manager.download_asset(
            {
                "name": "CentralN2.zip",
                "browser_download_url": "https://example.test/CentralN2.zip",
                "size": len(payload),
                "digest": "sha256:" + ("0" * 64),
            },
            tmp_path,
        )

    assert not (tmp_path / "CentralN2.zip").exists()
    assert not list(tmp_path.glob("*.part"))
