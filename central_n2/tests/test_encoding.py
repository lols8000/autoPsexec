from __future__ import annotations

from core.executor import RemoteExecutor


def test_decode_utf8_preserves_portuguese_accents():
    text = "não conseguiu se conectar à solicitação do serviço"
    raw = text.encode("utf-8")
    assert RemoteExecutor._decode_output(raw, preferred="utf-8") == text


def test_decode_cp850_preserves_portuguese_accents():
    text = "não conseguiu se conectar à solicitação do serviço"
    raw = text.encode("cp850")
    assert RemoteExecutor._decode_output(raw, preferred="cp850") == text


def test_decode_utf16_bom():
    text = "reinicialização necessária"
    raw = text.encode("utf-16")
    assert RemoteExecutor._decode_output(raw) == text


def test_powershell_prefix_forces_utf8():
    prefix = RemoteExecutor._powershell_utf8_prefix()
    assert "Console]::OutputEncoding" in prefix
    assert "$OutputEncoding" in prefix
    assert "UTF8Encoding" in prefix


def test_decoder_does_not_insert_replacement_character():
    text = "serviço não disponível"
    raw = text.encode("cp850")
    decoded = RemoteExecutor._decode_output(raw, preferred="cp850")
    assert "�" not in decoded
    assert decoded == text
