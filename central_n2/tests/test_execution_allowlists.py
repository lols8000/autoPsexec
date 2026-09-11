from __future__ import annotations

from pathlib import Path

from core.result import CommandResult
from modules.certificates import CertificatesModule
from modules.packages import PackagesModule
from modules.registry_actions import RegistryActionsModule


class RecordingExecutor:
    logger = None

    def __init__(self) -> None:
        self.commands: list[tuple[str, str, str]] = []

    def execute_mutating_cmd(
        self,
        host: str,
        command: str,
        timeout: int | None = None,
    ) -> CommandResult:
        self.commands.append(("cmd", host, command))
        return CommandResult(
            True,
            command,
            host,
            transport="local",
        )

    def execute_mutating_powershell_json(
        self,
        host: str,
        script: str,
        timeout: int | None = None,
    ) -> CommandResult:
        self.commands.append(("powershell", host, script))
        return CommandResult(
            True,
            script,
            host,
            data={"Applied": True, "Imported": True},
            transport="local",
        )


def test_package_module_rejects_unknown_catalog_key():
    module = PackagesModule(
        RecordingExecutor(),
        {"packages": {}},
    )

    result = module.install("PC01", "not-allowed")

    assert result.success is False
    assert "não encontrado" in result.stderr


def test_package_module_rejects_unsupported_type(tmp_path: Path):
    source = tmp_path / "payload.ps1"
    source.write_text("Write-Host test", encoding="utf-8")
    module = PackagesModule(
        RecordingExecutor(),
        {
            "packages": {
                "script": {
                    "source": str(source),
                    "type": "ps1",
                }
            }
        },
    )

    result = module.install("PC01", "script")

    assert result.success is False
    assert "Tipo permitido" in result.stderr


def test_certificate_module_rejects_private_key_container(tmp_path: Path):
    source = tmp_path / "certificate.pfx"
    source.write_bytes(b"not-a-real-pfx")
    module = CertificatesModule(
        RecordingExecutor(),
        {
            "certificates": {
                "private": {
                    "source": str(source),
                    "store": "My",
                }
            }
        },
    )

    result = module.import_certificate("PC01", "private")

    assert result.success is False
    assert ".cer/.crt" in result.stderr


def test_certificate_module_rejects_unapproved_store(tmp_path: Path):
    source = tmp_path / "certificate.cer"
    source.write_bytes(b"public-certificate")
    module = CertificatesModule(
        RecordingExecutor(),
        {
            "certificates": {
                "ca": {
                    "source": str(source),
                    "store": "SecretStore",
                }
            }
        },
    )

    result = module.import_certificate("PC01", "ca")

    assert result.success is False
    assert "Store não permitido" in result.stderr


def test_registry_actions_reject_non_hklm_paths():
    executor = RecordingExecutor()
    module = RegistryActionsModule(
        executor,
        {
            "registry_actions": {
                "bad": {
                    "path": r"HKCU:\Software\Empresa",
                    "name": "Enabled",
                    "type": "DWord",
                    "value": 1,
                    "mode": "set",
                }
            }
        },
    )

    result = module.apply("PC01", "bad")

    assert result.success is False
    assert "Somente HKLM" in result.stderr
    assert executor.commands == []


def test_registry_action_uses_only_allowlisted_payload():
    executor = RecordingExecutor()
    module = RegistryActionsModule(
        executor,
        {
            "registry_actions": {
                "policy": {
                    "path": r"HKLM:\SOFTWARE\Empresa\Produto",
                    "name": "Enabled",
                    "type": "DWord",
                    "value": 1,
                    "mode": "set",
                }
            }
        },
    )

    result = module.apply("PC01", "policy")

    assert result.success is True
    assert len(executor.commands) == 1
    mode, host, script = executor.commands[0]
    assert mode == "powershell"
    assert host == "PC01"
    assert "HKLM:\\SOFTWARE\\Empresa\\Produto" in script
    assert "Enabled" in script
    assert "New-ItemProperty" in script
