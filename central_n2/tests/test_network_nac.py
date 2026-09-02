from pathlib import Path

import pytest

from modules.network_nac import (
    AclPlanner,
    MacEntry,
    MacTableParser,
    NetworkNACModule,
    OuiAllowlist,
    PROFILES,
    is_locally_administered,
    mac_to_intelbras,
    normalize_mac,
    normalize_prefix,
    prefix_to_serie3000_acl,
)


def test_normalize_mac_accepts_common_formats():
    assert normalize_mac("00-11-22-33-44-55") == "00:11:22:33:44:55"
    assert normalize_mac("0011.2233.4455") == "00:11:22:33:44:55"
    assert normalize_mac("001122334455") == "00:11:22:33:44:55"


def test_normalize_mac_rejects_invalid_value():
    with pytest.raises(ValueError):
        normalize_mac("00:11:22")


def test_normalize_prefix():
    assert normalize_prefix("00-11-22") == "00:11:22"
    with pytest.raises(ValueError):
        normalize_prefix("00:11")


def test_intelbras_format_and_serie3000_wildcard():
    assert mac_to_intelbras("00:11:22:33:44:55") == "0011.2233.4455"
    base, wildcard = prefix_to_serie3000_acl("00:11:22")
    assert base == "0011.2200.0000"
    assert wildcard == "0000.00FF.FFFF"


def test_locally_administered_detection():
    assert is_locally_administered("02:11:22:33:44:55") is True
    assert is_locally_administered("00:11:22:33:44:55") is False


def test_parser_handles_colon_and_dotted_mac_formats():
    text = """
    10  0011.2233.4455  dynamic  ge1/0/3
    20  AA:BB:CC:DD:EE:FF dynamic gi1/0/7
    """
    entries = MacTableParser.parse(text)
    assert len(entries) == 2
    assert entries[0].mac == "00:11:22:33:44:55"
    assert entries[0].port.lower() == "ge1/0/3"
    assert entries[0].vlan == 10
    assert entries[1].mac == "AA:BB:CC:DD:EE:FF"
    assert entries[1].vlan == 20


def test_parser_deduplicates_same_entry():
    text = """
    10 0011.2233.4455 dynamic ge1/0/3
    10 0011.2233.4455 dynamic ge1/0/3
    """
    assert len(MacTableParser.parse(text)) == 1


def make_allowlist(tmp_path: Path) -> Path:
    path = tmp_path / "oui.json"
    path.write_text(
        """
        {
          "manufacturers": [
            {"name": "Fabricante A", "prefixes": ["00:11:22"]},
            {"name": "Fabricante B", "prefixes": ["AA:BB:CC"]}
          ],
          "exact_macs": [
            {"name": "Exceção", "mac": "DE:AD:BE:EF:00:01"}
          ]
        }
        """,
        encoding="utf-8",
    )
    return path


def test_allowlist_classifies_oui_and_exact_mac(tmp_path):
    allowlist = OuiAllowlist(make_allowlist(tmp_path))
    oui = allowlist.classify(MacEntry("00:11:22:33:44:55", "ge1/0/1", 10))
    exact = allowlist.classify(MacEntry("DE:AD:BE:EF:00:01", "ge1/0/2", 10))
    denied = allowlist.classify(MacEntry("10:20:30:40:50:60", "ge1/0/3", 10))

    assert oui.authorized is True
    assert oui.manufacturer == "Fabricante A"
    assert exact.authorized is True
    assert exact.reason == "exceção por MAC exato"
    assert denied.authorized is False


def test_audit_summary_groups_by_port(tmp_path):
    module = NetworkNACModule(make_allowlist(tmp_path))
    text = """
    10 0011.2233.4455 dynamic ge1/0/1
    10 1020.3040.5060 dynamic ge1/0/2
    """
    summary = module.audit_summary(text)
    assert summary["total"] == 2
    assert summary["authorized"] == 1
    assert summary["unauthorized"] == 1
    assert summary["by_port"]["ge1/0/1"]["authorized"] == 1
    assert summary["by_port"]["ge1/0/2"]["unauthorized"] == 1


def test_serie3000_prefix_plan_has_permits_before_port_binding():
    planner = AclPlanner(PROFILES["intelbras_serie3000"])
    commands = planner.build_prefix_allowlist(
        ["00:11:22", "AA:BB:CC"],
        ["ge1/0/1", "ge1/0/2"],
        acl_id=2001,
    )
    assert commands[0] == "configure terminal"
    assert commands[1] == "mac access-list standard 2001"
    assert "10 permit 0011.2200.0000 0000.00FF.FFFF" in commands
    assert "20 permit AABB.CC00.0000 0000.00FF.FFFF" in commands
    assert "interface ge1/0/1" in commands
    assert "mac access-group 2001 in" in commands


def test_empty_prefix_allowlist_is_refused():
    planner = AclPlanner(PROFILES["intelbras_serie3000"])
    with pytest.raises(ValueError, match="deny-by-default vazia"):
        planner.build_prefix_allowlist([], ["ge1/0/1"], acl_id=2001)


def test_prefix_policy_refused_for_unvalidated_profile():
    planner = AclPlanner(PROFILES["intelbras_s2050g_a"])
    with pytest.raises(ValueError, match="não possui geração automática"):
        planner.build_prefix_allowlist(["00:11:22"], ["ge1"], acl_id=200)


def test_serie3000_acl_id_range_is_enforced():
    planner = AclPlanner(PROFILES["intelbras_serie3000"])
    with pytest.raises(ValueError, match="2001-3000"):
        planner.build_prefix_allowlist(["00:11:22"], ["ge1/0/1"], acl_id=200)
