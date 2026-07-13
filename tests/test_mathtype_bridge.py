import json
import subprocess
from pathlib import Path

import pytest

from word_document_server.core import mathtype_bridge


def test_invoke_bridge_sends_one_utf8_json_request(tmp_path, monkeypatch):
    executable = tmp_path / "MathTypeBridge.exe"
    executable.touch()
    captured = {}

    def fake_run(command, **kwargs):
        captured["command"] = command
        captured.update(kwargs)
        return subprocess.CompletedProcess(
            command,
            0,
            stdout=json.dumps({"ok": True, "result": {"equations": []}}),
            stderr="",
        )

    monkeypatch.setattr(mathtype_bridge.subprocess, "run", fake_run)

    result = mathtype_bridge.invoke_bridge(
        "list_equations",
        executable=executable,
        filename="paper.docx",
    )

    assert result == {"equations": []}
    assert captured["command"] == [str(executable)]
    assert json.loads(captured["input"]) == {
        "command": "list_equations",
        "filename": "paper.docx",
    }
    assert captured["encoding"] == "utf-8"
    assert captured["timeout"] == 30
    assert captured["check"] is False


def test_invoke_bridge_surfaces_structured_bridge_error(tmp_path, monkeypatch):
    executable = tmp_path / "MathTypeBridge.exe"
    executable.touch()

    monkeypatch.setattr(
        mathtype_bridge.subprocess,
        "run",
        lambda *args, **kwargs: subprocess.CompletedProcess(
            args[0],
            1,
            stdout=json.dumps(
                {
                    "ok": False,
                    "error": {
                        "code": "equation_changed",
                        "message": "The equation changed after it was read.",
                    },
                }
            ),
            stderr="",
        ),
    )

    with pytest.raises(mathtype_bridge.MathTypeBridgeError) as error:
        mathtype_bridge.invoke_bridge("replace_equation", executable=executable)

    assert error.value.code == "equation_changed"
    assert str(error.value) == "The equation changed after it was read."


def test_invoke_bridge_rejects_invalid_json_response(tmp_path, monkeypatch):
    executable = tmp_path / "MathTypeBridge.exe"
    executable.touch()
    monkeypatch.setattr(
        mathtype_bridge.subprocess,
        "run",
        lambda *args, **kwargs: subprocess.CompletedProcess(
            args[0], 1, stdout="not-json", stderr="native failure"
        ),
    )

    with pytest.raises(mathtype_bridge.MathTypeBridgeError) as error:
        mathtype_bridge.invoke_bridge("list_equations", executable=executable)

    assert error.value.code == "invalid_bridge_response"
    assert "native failure" in str(error.value)


def test_invoke_bridge_reports_timeout(tmp_path, monkeypatch):
    executable = tmp_path / "MathTypeBridge.exe"
    executable.touch()

    def time_out(*args, **kwargs):
        raise subprocess.TimeoutExpired(args[0], kwargs["timeout"])

    monkeypatch.setattr(mathtype_bridge.subprocess, "run", time_out)

    with pytest.raises(mathtype_bridge.MathTypeBridgeError) as error:
        mathtype_bridge.invoke_bridge("list_equations", executable=executable)

    assert error.value.code == "bridge_timeout"


def test_default_executable_is_packaged_next_to_the_python_module():
    path = mathtype_bridge.bridge_executable_path()

    assert path.name == "MathTypeBridge.exe"
    assert path.parent.name == "bin"
    assert path.is_absolute()


def test_missing_bridge_executable_has_build_instruction(tmp_path):
    missing = Path(tmp_path) / "MathTypeBridge.exe"

    with pytest.raises(mathtype_bridge.MathTypeBridgeError) as error:
        mathtype_bridge.invoke_bridge("list_equations", executable=missing)

    assert error.value.code == "bridge_not_built"
    build_script = (
        Path(mathtype_bridge.__file__).resolve().parents[1]
        / "mathtype_bridge"
        / "build_mathtype_bridge.ps1"
    )
    assert str(build_script) in str(error.value)
    assert "scripts/build_mathtype_bridge.ps1" not in str(error.value)


def test_packaged_bridge_validates_and_canonicalizes_mathml():
    first = '<math xmlns="http://www.w3.org/1998/Math/MathML">\n  <mi>x</mi>\n</math>'
    second = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>'

    first_result = mathtype_bridge.invoke_bridge("validate_mathml", mathml=first)
    second_result = mathtype_bridge.invoke_bridge("validate_mathml", mathml=second)

    assert first_result["canonical_mathml"] == second_result["canonical_mathml"]
    assert first_result["mathml_sha256"] == second_result["mathml_sha256"]
    assert len(first_result["mathml_sha256"]) == 64


def test_packaged_bridge_rejects_dtd_mathml():
    mathml = (
        '<!DOCTYPE math [<!ENTITY value "x">]>'
        '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>&value;</mi></math>'
    )

    with pytest.raises(mathtype_bridge.MathTypeBridgeError) as error:
        mathtype_bridge.invoke_bridge("validate_mathml", mathml=mathml)

    assert error.value.code == "invalid_mathml"


def test_packaged_bridge_preserves_unicode_math_symbols():
    mathml = (
        '<math xmlns="http://www.w3.org/1998/Math/MathML">'
        "<mo>&#x222B;</mo><mo>&#x2211;</mo>"
        "</math>"
    )

    result = mathtype_bridge.invoke_bridge("validate_mathml", mathml=mathml)

    assert "∫" in result["canonical_mathml"]
    assert "∑" in result["canonical_mathml"]


def test_packaged_bridge_hash_is_namespace_prefix_independent():
    default_namespace = (
        '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>'
    )
    prefixed = (
        '<mml:math xmlns:mml="http://www.w3.org/1998/Math/MathML">'
        "<mml:mi>x</mml:mi></mml:math>"
    )

    default_result = mathtype_bridge.invoke_bridge(
        "validate_mathml", mathml=default_namespace
    )
    prefixed_result = mathtype_bridge.invoke_bridge("validate_mathml", mathml=prefixed)

    assert default_result["canonical_mathml"] == prefixed_result["canonical_mathml"]
    assert default_result["mathml_sha256"] == prefixed_result["mathml_sha256"]


def test_csharp_bridge_has_deterministic_com_ownership():
    source = (
        Path("word_document_server/mathtype_bridge/MathTypeBridge.cs")
        .read_text(encoding="utf-8")
    )

    assert "internal sealed class EquationInventory : IDisposable" in source
    assert "internal sealed class EquationReference : IDisposable" in source
    assert source.count("using (EquationInventory") >= 8
    assert "foreach (object firstStoryObject in _document.StoryRanges)" not in source
    assert "candidate.Kind == original.Kind" in source
    assert "FinalReleaseComObject" not in source
    assert source.count("Marshal.ReleaseComObject") == 1
    assert "catch (InvalidComObjectException)" in source
    assert 'command == "set_equation_text"' not in source
    assert 'command == "clone_equation"' not in source
