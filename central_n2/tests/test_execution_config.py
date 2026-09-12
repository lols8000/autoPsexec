from __future__ import annotations

from execution.config_validation import validate_execution_configuration


def test_invalid_catalog_entries_are_disabled_without_blocking_valid_ones():
    settings = {
        "packages": {
            "good": {
                "source": r"\\server\share\good.msi",
                "type": "msi",
                "timeout_seconds": 900,
            },
            "bad": {
                "source": "",
                "type": "ps1",
            },
        },
        "certificates": {
            "good": {
                "source": r"\\server\share\ca.cer",
                "store": "Root",
            },
            "private": {
                "source": r"\\server\share\private.pfx",
                "store": "My",
            },
        },
        "registry_actions": {
            "good": {
                "path": r"HKLM:\SOFTWARE\Empresa",
                "name": "Enabled",
                "type": "DWord",
                "value": 1,
                "mode": "set",
            },
            "bad": {
                "path": r"HKCU:\Software\Empresa",
                "name": "Enabled",
                "type": "DWord",
                "value": 1,
                "mode": "set",
            },
        },
    }

    report = validate_execution_configuration(settings)

    assert "good" in report.settings["packages"]
    assert "bad" not in report.settings["packages"]

    assert "good" in report.settings["certificates"]
    assert "private" not in report.settings["certificates"]

    assert "good" in report.settings["registry_actions"]
    assert "bad" not in report.settings["registry_actions"]

    assert report.valid is False
    assert len(report.issues) == 3


def test_invalid_file_roots_fall_back_to_safe_defaults():
    report = validate_execution_configuration(
        {
            "execution": {
                "file_roots": [
                    r"..\Windows",
                    r"relative\path",
                ]
            }
        }
    )

    roots = report.settings["execution"]["file_roots"]

    assert r"C:\CentralN2" in roots
    assert r"C:\Temp" in roots
    assert all(".." not in root for root in roots)
    assert any(
        issue.section == "execution.file_roots"
        for issue in report.issues
    )


def test_valid_catalog_configuration_remains_unchanged():
    settings = {
        "packages": {
            "agent": {
                "source": r"\\server\share\agent.msi",
                "type": "msi",
                "timeout_seconds": 1200,
            }
        },
        "certificates": {
            "root": {
                "source": r"\\server\share\root.cer",
                "store": "Root",
            }
        },
        "registry_actions": {
            "policy": {
                "path": r"HKLM:\SOFTWARE\Empresa",
                "name": "Enabled",
                "type": "DWord",
                "value": 1,
                "mode": "set",
            }
        },
        "execution": {
            "file_roots": [
                r"C:\CentralN2",
                r"D:\Suporte",
            ]
        },
    }

    report = validate_execution_configuration(settings)

    assert report.valid is True
    assert report.enabled_counts == {
        "packages": 1,
        "certificates": 1,
        "registry_actions": 1,
        "file_roots": 2,
    }
    assert report.settings["packages"]["agent"]["type"] == "msi"
    assert report.settings["execution"]["file_roots"] == [
        r"C:\CentralN2",
        r"D:\Suporte",
    ]
