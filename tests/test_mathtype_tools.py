import pytest

from word_document_server.tools import mathtype_tools


def test_list_mathtype_equations_delegates_to_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload))
        or {"equations": []},
    )

    result = mathtype_tools.word_live_list_mathtype_equations("paper.docx")

    assert result == {"equations": []}
    assert calls == [("list_equations", {"filename": "paper.docx"})]


def test_get_mathtype_equation_delegates_to_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload))
        or {"mathml": "<math />"},
    )

    result = mathtype_tools.word_live_get_mathtype_equation(
        "ole:1:42:Equation.DSMT4", "paper.docx"
    )

    assert result == {"mathml": "<math />"}
    assert calls == [
        (
            "get_equation",
            {
                "equation_id": "ole:1:42:Equation.DSMT4",
                "filename": "paper.docx",
            },
        )
    ]


def test_replace_mathtype_equation_requires_expected_hash(monkeypatch):
    monkeypatch.setattr(
        mathtype_tools, "invoke_bridge", lambda command, **payload: payload
    )

    with pytest.raises(ValueError, match="expected_mathml_sha256"):
        mathtype_tools.word_live_replace_mathtype_equation(
            "ole:1:42:Equation.DSMT4",
            '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>',
            "",
        )


def test_replace_mathtype_equation_sends_mathml_and_hash(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload))
        or {"verified": True},
    )
    mathml = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>y</mi></math>'
    digest = "a" * 64

    result = mathtype_tools.word_live_replace_mathtype_equation(
        "ole:1:42:Equation.DSMT4", mathml, digest, "paper.docx"
    )

    assert result == {"verified": True}
    assert calls == [
        (
            "replace_equation",
            {
                "equation_id": "ole:1:42:Equation.DSMT4",
                "mathml": mathml,
                "expected_mathml_sha256": digest,
                "filename": "paper.docx",
            },
        )
    ]


def test_replace_tex_requires_expected_hash():
    with pytest.raises(ValueError, match="expected_mathml_sha256"):
        mathtype_tools.word_live_replace_mathtype_equation_tex(
            "ole:inline:1.0:10:1:Equation.DSMT4", "x^2", ""
        )


def test_replace_tex_passes_hash_to_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload))
        or {"equation_id": "new"},
    )

    digest = "a" * 64
    result = mathtype_tools.word_live_replace_mathtype_equation_tex(
        "eq-1", r"\alpha", digest, "paper.docx"
    )

    assert result == {"equation_id": "new"}
    assert calls == [
        (
            "replace_equation_tex",
            {
                "equation_id": "eq-1",
                "tex": r"\alpha",
                "expected_mathml_sha256": digest,
                "filename": "paper.docx",
                "timeout": 120,
            },
        )
    ]


def test_probe_delegates_to_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(
        mathtype_tools,
        "invoke_bridge",
        lambda command, **payload: calls.append((command, payload)) or {"formats": []},
    )

    result = mathtype_tools.word_live_probe_mathtype_equation("eq-1", "paper.docx")

    assert result == {"formats": []}
    assert calls == [("probe_equation", {"equation_id": "eq-1", "filename": "paper.docx"})]


def test_all_mathtype_tools_registered_with_docstring_description():
    import pathlib

    source = pathlib.Path("word_document_server/main.py").read_text(encoding="utf-8")
    for name in (
        "word_live_list_mathtype_equations",
        "word_live_get_mathtype_equation",
        "word_live_dump_mathtype_equations",
        "word_live_dump_mathtype_document",
        "word_live_delete_mathtype_equation",
        "word_live_replace_mathtype_equation_tex",
        "word_live_probe_mathtype_equation",
        "word_live_replace_mathtype_equation",
    ):
        assert f"def {name}(" in source, name
        assert f"description=mathtype_tools.{name}.__doc__" in source, name


def test_packaged_addin_scripts_exist():
    import pathlib

    bridge_dir = pathlib.Path("word_document_server/mathtype_bridge")
    for script in (
        "build_mathtype_bridge.ps1",
        "install_mathtype_addin.ps1",
        "uninstall_mathtype_addin.ps1",
    ):
        assert (bridge_dir / script).is_file(), script
