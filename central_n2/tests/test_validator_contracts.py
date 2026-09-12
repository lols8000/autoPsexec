from __future__ import annotations

import pytest

from core.result import CommandResult
from execution.validators import (
    after_field_matches_parameter,
    command_completed,
    command_field_true,
    field_equals,
    postcheck_succeeded,
    service_running,
    service_stopped,
)
from remediation import ValidationStatus
from remediation.validators import (
    _result_data,
    validate_cleanup,
    validate_gpupdate,
    validate_spooler,
    validate_windows_update_reset,
)


def result(
    *,
    success: bool = True,
    data=None,
    stderr: str = "",
    stdout: str = "",
    indeterminate: bool = False,
) -> CommandResult:
    item = CommandResult(
        success=success,
        command="test",
        host="PC01",
        data=data,
        stderr=stderr,
        stdout=stdout,
    )
    if indeterminate:
        item.mark_indeterminate("estado incerto")
    return item


@pytest.mark.parametrize(
    ("command", "expected"),
    [
        (result(indeterminate=True), ValidationStatus.UNKNOWN),
        (result(success=True), ValidationStatus.PASS),
        (result(success=False, stderr="boom"), ValidationStatus.FAIL),
    ],
)
def test_command_completed_contract(command, expected):
    validation = command_completed(None, command, {"after": True}, {})

    assert validation.status is expected
    assert validation.evidence == {"after": True}


def test_command_completed_uses_default_failure_message():
    validation = command_completed(
        None,
        result(success=False),
        None,
        {},
    )

    assert validation.status is ValidationStatus.FAIL
    assert validation.message == "A execução falhou."


@pytest.mark.parametrize(
    ("after", "command", "expected"),
    [
        ({"State": "Ready"}, result(), ValidationStatus.PASS),
        (
            result(data={"State": "Ready"}),
            result(),
            ValidationStatus.PASS,
        ),
        ({"State": "Other"}, result(), ValidationStatus.FAIL),
        ({}, result(), ValidationStatus.UNKNOWN),
        ({}, result(success=False, stderr="probe failed"), ValidationStatus.FAIL),
    ],
)
def test_field_equals_handles_dict_command_result_and_missing_fields(
    after,
    command,
    expected,
):
    validator = field_equals(
        "State",
        "Ready",
        pass_message="ok",
        fail_message="not ready",
    )

    validation = validator(None, command, after, {})

    assert validation.status is expected


def test_field_equals_indeterminate_has_priority():
    validator = field_equals(
        "State",
        "Ready",
        pass_message="ok",
        fail_message="not ready",
    )

    validation = validator(
        None,
        result(indeterminate=True),
        {"State": "Ready"},
        {},
    )

    assert validation.status is ValidationStatus.UNKNOWN


@pytest.mark.parametrize(
    ("validator", "state", "expected"),
    [
        (service_running, "Running", ValidationStatus.PASS),
        (service_running, "Stopped", ValidationStatus.FAIL),
        (service_stopped, "Stopped", ValidationStatus.PASS),
        (service_stopped, "Running", ValidationStatus.FAIL),
    ],
)
def test_service_state_validators(validator, state, expected):
    validation = validator(
        None,
        result(),
        {"Status": state},
        {},
    )

    assert validation.status is expected


@pytest.mark.parametrize("validator", [service_running, service_stopped])
def test_service_state_indeterminate_is_unknown(validator):
    validation = validator(
        None,
        result(indeterminate=True),
        {"Status": "Running"},
        {},
    )

    assert validation.status is ValidationStatus.UNKNOWN


@pytest.mark.parametrize("validator", [service_running, service_stopped])
def test_service_state_falls_back_to_command_contract_without_dict(validator):
    passed = validator(None, result(success=True), None, {})
    failed = validator(
        None,
        result(success=False, stderr="service error"),
        None,
        {},
    )

    assert passed.status is ValidationStatus.PASS
    assert failed.status is ValidationStatus.FAIL


@pytest.mark.parametrize(
    ("command", "expected"),
    [
        (result(data={"Done": True}), ValidationStatus.PASS),
        (result(data={"Done": False}), ValidationStatus.FAIL),
        (result(data={}), ValidationStatus.UNKNOWN),
        (
            result(success=False, data={}, stderr="command failed"),
            ValidationStatus.FAIL,
        ),
    ],
)
def test_command_field_true_contract(command, expected):
    validator = command_field_true(
        "Done",
        pass_message="done",
        fail_message="not done",
    )

    validation = validator(None, command, None, {})

    assert validation.status is expected


def test_command_field_true_indeterminate_is_unknown():
    validator = command_field_true(
        "Done",
        pass_message="done",
        fail_message="not done",
    )

    validation = validator(
        None,
        result(data={"Done": True}, indeterminate=True),
        None,
        {},
    )

    assert validation.status is ValidationStatus.UNKNOWN


@pytest.mark.parametrize(
    ("after", "command", "expected"),
    [
        ({"Mode": "AUTO"}, result(), ValidationStatus.PASS),
        (
            result(data={"Mode": "auto"}),
            result(),
            ValidationStatus.PASS,
        ),
        ({"Mode": "manual"}, result(), ValidationStatus.FAIL),
        ({}, result(), ValidationStatus.UNKNOWN),
        ({}, result(success=False, stderr="failed"), ValidationStatus.FAIL),
    ],
)
def test_after_field_matches_parameter_contract(after, command, expected):
    validator = after_field_matches_parameter(
        "Mode",
        "mode",
        pass_message="matched",
        fail_message="mismatch",
    )

    validation = validator(
        None,
        command,
        after,
        {"mode": "auto"},
    )

    assert validation.status is expected


def test_after_field_matches_parameter_indeterminate_is_unknown():
    validator = after_field_matches_parameter(
        "Mode",
        "mode",
        pass_message="matched",
        fail_message="mismatch",
    )

    validation = validator(
        None,
        result(indeterminate=True),
        {"Mode": "auto"},
        {"mode": "auto"},
    )

    assert validation.status is ValidationStatus.UNKNOWN


@pytest.mark.parametrize(
    ("command", "after", "expected"),
    [
        (
            result(),
            result(success=True, data={"ok": True}),
            ValidationStatus.PASS,
        ),
        (
            result(),
            result(success=True, data=None, stdout="healthy"),
            ValidationStatus.PASS,
        ),
        (
            result(),
            result(success=False, stderr="probe failed"),
            ValidationStatus.UNKNOWN,
        ),
        (
            result(success=False),
            result(success=False, stderr="probe failed"),
            ValidationStatus.FAIL,
        ),
        (
            result(),
            {"_probe_success": True},
            ValidationStatus.PASS,
        ),
        (
            result(),
            {"_probe_success": False, "_error": "offline"},
            ValidationStatus.UNKNOWN,
        ),
        (
            result(success=False),
            {"_probe_success": False},
            ValidationStatus.FAIL,
        ),
        (result(), {"other": 1}, ValidationStatus.UNKNOWN),
        (
            result(success=False, stderr="action failed"),
            None,
            ValidationStatus.FAIL,
        ),
    ],
)
def test_postcheck_succeeded_contract(command, after, expected):
    validation = postcheck_succeeded(None, command, after, {})

    assert validation.status is expected


def test_postcheck_succeeded_indeterminate_is_unknown():
    validation = postcheck_succeeded(
        None,
        result(indeterminate=True),
        {"_probe_success": True},
        {},
    )

    assert validation.status is ValidationStatus.UNKNOWN


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ({"a": 1}, {"a": 1}),
        (result(data={"a": 1}), {"a": 1}),
        (result(data="not-a-dict"), {}),
        (None, {}),
    ],
)
def test_result_data_normalizes_supported_payloads(value, expected):
    assert _result_data(value) == expected


@pytest.mark.parametrize(
    ("command", "after", "expected"),
    [
        (result(), {"Status": "Running"}, ValidationStatus.PASS),
        (
            result(indeterminate=True),
            {"Status": "Stopped"},
            ValidationStatus.UNKNOWN,
        ),
        (result(), {"Status": "Stopped"}, ValidationStatus.FAIL),
    ],
)
def test_validate_spooler_contract(command, after, expected):
    validation = validate_spooler(None, command, after)

    assert validation.status is expected


@pytest.mark.parametrize(
    ("command", "expected"),
    [
        (
            result(success=True, data={"RecoveredGB": 1.25}),
            ValidationStatus.PASS,
        ),
        (
            result(indeterminate=True, data={}),
            ValidationStatus.UNKNOWN,
        ),
        (result(success=False, data={}), ValidationStatus.FAIL),
        (result(success=True, data={}), ValidationStatus.FAIL),
    ],
)
def test_validate_cleanup_contract(command, expected):
    validation = validate_cleanup(None, command, None)

    assert validation.status is expected


@pytest.mark.parametrize(
    ("command", "expected"),
    [
        (
            result(
                success=True,
                data={"RestoredOriginalRunningServices": True},
            ),
            ValidationStatus.PASS,
        ),
        (
            result(
                success=False,
                data={"RestoredOriginalRunningServices": False},
                indeterminate=True,
            ),
            ValidationStatus.UNKNOWN,
        ),
        (
            result(success=True, data={}),
            ValidationStatus.UNKNOWN,
        ),
        (
            result(
                success=True,
                data={"RestoredOriginalRunningServices": False},
            ),
            ValidationStatus.FAIL,
        ),
    ],
)
def test_validate_windows_update_reset_contract(command, expected):
    validation = validate_windows_update_reset(None, command, None)

    assert validation.status is expected


@pytest.mark.parametrize(
    ("command", "after", "expected"),
    [
        (
            result(success=False, stderr="gpupdate failed"),
            None,
            ValidationStatus.FAIL,
        ),
        (
            result(success=False, indeterminate=True),
            None,
            ValidationStatus.UNKNOWN,
        ),
        (
            result(),
            result(success=True, stdout="Applied GPOs"),
            ValidationStatus.PASS,
        ),
        (
            result(),
            {"_probe_success": True, "_stdout": "Applied GPOs"},
            ValidationStatus.PASS,
        ),
        (
            result(),
            result(success=False, stderr="gpresult failed"),
            ValidationStatus.UNKNOWN,
        ),
        (result(), {}, ValidationStatus.UNKNOWN),
    ],
)
def test_validate_gpupdate_contract(command, after, expected):
    validation = validate_gpupdate(None, command, after)

    assert validation.status is expected
