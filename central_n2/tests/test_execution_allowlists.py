from __future__ import annotations

from pathlib import Path

from core.result import CommandResult
from modules.certificates import CertificatesModule
from modules.file_ops import FileOperationsModule
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



class LocalRecordingExecutor(RecordingExecutor):
    @staticmethod
    def is_local(host: str) -> bool:
        return True

    def execute_powershell_json(
        self,
        host: str,
        script: str,
        timeout: int | None = None,
    ) -> CommandResult:
        self.commands.append(("powershell-read", host, script))
        return CommandResult(
            True,
            script,
            host,
            data={
                "Exists": True,
                "Value": 1,
                "ValueKind": "DWord",
            },
            transport="local",
        )


def test_package_copy_uses_direct_path_for_local_target(tmp_path: Path):
    source = tmp_path / "app.msi"
    source.write_bytes(b"package")
    destination = tmp_path / "remote" / "app.msi"

    executor = LocalRecordingExecutor()
    module = PackagesModule(
        executor,
        {
            "packages": {
                "app": {
                    "source": str(source),
                    "remote_path": str(destination),
                    "type": "msi",
                    "cleanup": False,
                }
            }
        },
    )

    result = module.install("localhost", "app")

    assert result.success is True
    assert destination.read_bytes() == b"package"
    assert all("\\\\localhost\\" not in command for _, _, command in executor.commands)


def test_certificate_copy_uses_direct_path_for_local_target(tmp_path: Path):
    source = tmp_path / "ca.cer"
    source.write_bytes(b"certificate")
    destination = tmp_path / "remote" / "ca.cer"

    executor = LocalRecordingExecutor()
    module = CertificatesModule(
        executor,
        {
            "certificates": {
                "ca": {
                    "source": str(source),
                    "remote_path": str(destination),
                    "store": "Root",
                }
            }
        },
    )

    result = module.import_certificate("localhost", "ca")

    assert result.success is True
    assert destination.read_bytes() == b"certificate"
    assert "\\\\localhost\\" not in executor.commands[-1][2]


def test_registry_rollback_restores_original_value_kind():
    executor = RecordingExecutor()
    module = RegistryActionsModule(
        executor,
        {
            "registry_actions": {
                "policy": {
                    "path": r"HKLM:\SOFTWARE\Empresa",
                    "name": "Enabled",
                    "type": "String",
                    "value": "new",
                    "mode": "set",
                }
            }
        },
    )

    result = module.rollback(
        "PC01",
        "policy",
        {
            "Exists": True,
            "Value": 1,
            "ValueKind": "DWord",
        },
    )

    assert result.success is True
    script = executor.commands[-1][2]
    assert "-PropertyType DWord" in script
    assert "-Value 1" in script



def test_file_operations_reject_path_outside_allowed_roots():
    executor = RecordingExecutor()
    module = FileOperationsModule(
        executor,
        allowed_roots=[r"C:\CentralN2", r"C:\Temp"],
    )

    try:
        module.remove_file(
            "PC01",
            r"C:\Windows\System32\kernel32.dll",
        )
    except ValueError as exc:
        assert "fora das raízes permitidas" in str(exc)
    else:
        raise AssertionError("Caminho fora da allowlist deveria falhar")

    assert executor.commands == []


def test_file_operations_reject_parent_traversal():
    executor = RecordingExecutor()
    module = FileOperationsModule(
        executor,
        allowed_roots=[r"C:\CentralN2"],
    )

    try:
        module.ensure_directory(
            "PC01",
            r"C:\CentralN2\..\Windows\Temp",
        )
    except ValueError as exc:
        assert "'..'" in str(exc)
    else:
        raise AssertionError("Traversal deveria falhar")

    assert executor.commands == []


def test_file_operations_accept_scoped_path():
    executor = RecordingExecutor()
    module = FileOperationsModule(
        executor,
        allowed_roots=[r"C:\CentralN2"],
    )

    result = module.ensure_directory(
        "PC01",
        r"C:\CentralN2\Packages",
    )

    assert result.success is True
    assert len(executor.commands) == 1
    assert r"C:\CentralN2\Packages" in executor.commands[0][2]
