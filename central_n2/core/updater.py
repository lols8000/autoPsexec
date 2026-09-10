from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

_SEMVER_RE = re.compile(
    r"^v?"
    r"(0|[1-9]\d*)\."
    r"(0|[1-9]\d*)\."
    r"(0|[1-9]\d*)"
    r"(?:-([0-9A-Za-z.-]+))?"
    r"(?:\+[0-9A-Za-z.-]+)?$"
)


@dataclass(frozen=True, slots=True)
class SemVersion:
    major: int
    minor: int
    patch: int
    prerelease: tuple[int | str, ...] = ()

    @classmethod
    def parse(cls, value: str) -> "SemVersion":
        match = _SEMVER_RE.fullmatch(value.strip())
        if not match:
            raise ValueError(f"Versão SemVer inválida: {value!r}")

        prerelease: list[int | str] = []
        raw_pre = match.group(4)
        if raw_pre:
            for token in raw_pre.split("."):
                if token.isdigit():
                    prerelease.append(int(token))
                else:
                    prerelease.append(token)

        return cls(
            int(match.group(1)),
            int(match.group(2)),
            int(match.group(3)),
            tuple(prerelease),
        )

    def _compare_prerelease(self, other: "SemVersion") -> int:
        if not self.prerelease and not other.prerelease:
            return 0
        if not self.prerelease:
            return 1
        if not other.prerelease:
            return -1

        for left, right in zip(self.prerelease, other.prerelease):
            if left == right:
                continue
            if isinstance(left, int):
                if isinstance(right, str):
                    return -1
                return -1 if left < right else 1

            if isinstance(right, int):
                return 1

            return -1 if left < right else 1

        if len(self.prerelease) == len(other.prerelease):
            return 0
        return -1 if len(self.prerelease) < len(other.prerelease) else 1

    def __lt__(self, other: object) -> bool:
        if not isinstance(other, SemVersion):
            return NotImplemented
        core_self = (self.major, self.minor, self.patch)
        core_other = (other.major, other.minor, other.patch)
        if core_self != core_other:
            return core_self < core_other
        return self._compare_prerelease(other) < 0

    def __le__(self, other: object) -> bool:
        if not isinstance(other, SemVersion):
            return NotImplemented
        return self == other or self < other

    def __gt__(self, other: object) -> bool:
        if not isinstance(other, SemVersion):
            return NotImplemented
        return not self <= other

    def __ge__(self, other: object) -> bool:
        if not isinstance(other, SemVersion):
            return NotImplemented
        return not self < other


@dataclass(slots=True)
class ReleaseInfo:
    current: str
    latest: str
    update_available: bool
    html_url: str | None
    assets: list[dict[str, Any]]


class UpdateManager:
    def __init__(
        self,
        repository: str,
        current_version: str,
        *,
        timeout: int = 10,
    ) -> None:
        if "/" not in repository:
            raise ValueError("repository deve estar no formato owner/name")
        self.repository = repository
        self.current_version = current_version.lstrip("v")
        self.timeout = max(1, int(timeout))
        self._current_semver = SemVersion.parse(self.current_version)

    @staticmethod
    def _version_tuple(value: str) -> SemVersion:
        """Compatibilidade com testes/código antigo; agora usa SemVer real."""
        return SemVersion.parse(value)

    def check_latest(self) -> ReleaseInfo:
        request = urllib.request.Request(
            (
                "https://api.github.com/repos/"
                f"{self.repository}/releases/latest"
            ),
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": "CentralN2",
            },
        )
        with urllib.request.urlopen(
            request,
            timeout=self.timeout,
        ) as response:
            data = json.loads(response.read().decode("utf-8"))

        latest_raw = str(data.get("tag_name") or "").strip()
        if not latest_raw:
            raise ValueError("Release do GitHub sem tag_name.")

        latest = latest_raw.lstrip("v")
        latest_semver = SemVersion.parse(latest)

        return ReleaseInfo(
            current=self.current_version,
            latest=latest,
            update_available=latest_semver > self._current_semver,
            html_url=data.get("html_url"),
            assets=list(data.get("assets") or []),
        )

    @staticmethod
    def _safe_asset_name(asset: dict[str, Any]) -> str:
        raw = str(asset.get("name") or "").strip()
        if not raw:
            raise ValueError("Asset sem nome.")
        if Path(raw).name != raw or raw in {".", ".."}:
            raise ValueError("Nome de asset inválido.")
        if any(char in raw for char in ("/", "\\", "\x00")):
            raise ValueError("Nome de asset contém separadores inválidos.")
        return raw

    @staticmethod
    def _expected_sha256(asset: dict[str, Any]) -> str | None:
        digest = str(asset.get("digest") or "").strip().lower()
        if not digest:
            return None
        if not digest.startswith("sha256:"):
            return None

        value = digest.split(":", 1)[1]
        if not re.fullmatch(r"[0-9a-f]{64}", value):
            raise ValueError("Digest SHA-256 do asset é inválido.")
        return value

    @staticmethod
    def _validate_https(url: str) -> None:
        parsed = urllib.parse.urlparse(url)
        if parsed.scheme.casefold() != "https" or not parsed.netloc:
            raise ValueError(
                "O asset de atualização deve usar uma URL HTTPS válida."
            )

    def download_asset(
        self,
        asset: dict[str, Any],
        directory: str | Path,
    ) -> Path:
        directory_path = Path(directory)
        directory_path.mkdir(parents=True, exist_ok=True)

        name = self._safe_asset_name(asset)
        target = directory_path / name
        url = str(asset.get("browser_download_url") or "").strip()
        if not url:
            raise ValueError("Asset sem browser_download_url.")
        self._validate_https(url)

        expected_size_raw = asset.get("size")
        expected_size: int | None = None
        if expected_size_raw is not None and str(expected_size_raw).strip():
            try:
                expected_size = int(str(expected_size_raw))
            except ValueError as exc:
                raise ValueError(
                    "Tamanho publicado do asset é inválido."
                ) from exc
        if expected_size is not None and expected_size < 0:
            raise ValueError("Tamanho publicado do asset é inválido.")

        expected_sha256 = self._expected_sha256(asset)
        sha256 = hashlib.sha256()
        downloaded = 0
        temp_path: Path | None = None

        request = urllib.request.Request(
            url,
            headers={"User-Agent": "CentralN2"},
        )

        try:
            with tempfile.NamedTemporaryFile(
                mode="wb",
                dir=directory_path,
                prefix=f".{name}.",
                suffix=".part",
                delete=False,
            ) as temporary:
                temp_path = Path(temporary.name)
                with urllib.request.urlopen(
                    request,
                    timeout=max(self.timeout, 60),
                ) as response:
                    while True:
                        chunk = response.read(1024 * 1024)
                        if not chunk:
                            break
                        temporary.write(chunk)
                        sha256.update(chunk)
                        downloaded += len(chunk)

                temporary.flush()
                os.fsync(temporary.fileno())

            if expected_size is not None and downloaded != expected_size:
                raise ValueError(
                    "Tamanho baixado diverge do tamanho publicado: "
                    f"{downloaded} != {expected_size} bytes."
                )

            actual_sha256 = sha256.hexdigest()
            if (
                expected_sha256 is not None
                and actual_sha256 != expected_sha256
            ):
                raise ValueError(
                    "SHA-256 do download diverge do digest publicado."
                )

            os.replace(temp_path, target)
            temp_path = None
            return target
        finally:
            if temp_path is not None:
                try:
                    temp_path.unlink(missing_ok=True)
                except OSError:
                    pass
