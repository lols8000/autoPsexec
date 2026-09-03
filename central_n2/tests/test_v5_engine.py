from core.redaction import redact
from diagnostics.correlation import CorrelationEngine
from diagnostics.engine import DiagnosticEngine
from reports.builder import SupportReportBuilder
from storage.diff import diff_values

def test_redaction_masks_secrets_recursively():
    safe=redact({"token":"abc","nested":{"message":"password=123","safe":"ok"}})
    assert safe["token"]=="***" and "123" not in safe["nested"]["message"] and safe["nested"]["safe"]=="ok"

def test_diagnostic_engine_and_correlation_storage_pressure():
    findings=DiagnosticEngine().evaluate({"DiskFreePercent":3,"DiskActivePercent":99})
    assert {"DISK_CRITICAL","DISK_SATURATION"}.issubset({f.id for f in findings})
    assert any(d.code=="STORAGE_PRESSURE" and d.confidence=="alta" for d in CorrelationEngine().correlate(findings))

def test_recursive_diff_reports_nested_changes():
    changes=diff_values({"a":{"b":1},"x":2},{"a":{"b":3},"y":4})
    assert {"a.b","x","y"}.issubset({x["path"] for x in changes})

def test_report_builder_is_structured():
    r=SupportReportBuilder().build(host="PC01",problem="lento",actions=["coleta"])
    assert r["host"]=="PC01" and r["actions"]==["coleta"]
