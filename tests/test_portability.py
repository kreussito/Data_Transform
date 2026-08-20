"""Running anywhere — specification_v1.md §12.

Two of these tests exist because of one specific, nasty failure mode: on Windows,
``stdout`` falls back to **cp1252** whenever the output is redirected to a file or a
pipe, and neither ``⟦`` nor ``→`` exists in cp1252. The run would complete, write the
workbook, write both logs — and then raise ``UnicodeEncodeError`` on the *summary*. The
user sees a traceback and concludes the run failed, with the finished output sitting on
disk beside them.
"""

from __future__ import annotations

import ast
import io
import pathlib
import shutil
import subprocess
import sys

import pytest

from datatransform.console import safe, use_utf8

ROOT = pathlib.Path(__file__).resolve().parents[1]
PACKAGE = ROOT / "datatransform"
MEXICO = ROOT / "Intake_Mexico_v1.xlsx"

# Characters the tool prints that a Windows ANSI code page cannot represent.
NON_CP1252 = ("⟦", "⟧", "→")


# ───────────────────────────────────────────── the cp1252 trap

def test_the_tool_prints_characters_cp1252_cannot_hold():
    """If this ever stops being true the two tests below are no longer needed."""
    for ch in NON_CP1252:
        with pytest.raises(UnicodeEncodeError):
            ch.encode("cp1252")


def test_a_redirected_windows_stdout_does_not_lose_the_run(tmp_path):
    """The real shape of it: cp1252, and a summary line full of anchors."""
    stream = io.TextIOWrapper(io.BytesIO(), encoding="cp1252", newline="")
    line = "06. EQ Aggs  processed  ⟦ZONES⟧ → 52 zones"

    with pytest.raises(UnicodeEncodeError):
        stream.write(line)          # what used to happen
        stream.flush()

    stream = io.TextIOWrapper(io.BytesIO(), encoding="cp1252", newline="")
    stream.write(safe(line, stream))     # what happens now
    stream.flush()
    written = stream.buffer.getvalue().decode("cp1252")
    assert "06. EQ Aggs" in written and "52 zones" in written
    assert "\\u27e6" in written           # the anchor is shown, not silently dropped


def test_safe_leaves_utf8_streams_completely_alone():
    stream = io.TextIOWrapper(io.BytesIO(), encoding="utf-8")
    assert safe("⟦ZONES⟧ → 52", stream) == "⟦ZONES⟧ → 52"


def test_use_utf8_is_harmless_where_it_cannot_apply():
    """Under pytest the streams are captured objects; this must not raise."""
    use_utf8()


def test_the_cli_survives_a_cp1252_pipe(tmp_path):
    """End to end, in a real subprocess, with the encoding Windows would pick."""
    source = tmp_path / MEXICO.name
    shutil.copy2(MEXICO, source)

    env = {
        "PYTHONPATH": str(ROOT),
        "PYTHONIOENCODING": "cp1252",     # exactly what a redirected Windows pipe does
        "PATH": "/usr/bin:/bin",
        "SYSTEMROOT": "C:\\Windows",
    }
    result = subprocess.run(
        [sys.executable, "-m", "datatransform", str(source),
         "-o", str(tmp_path / "out.xlsx"), "--log-dir", str(tmp_path / "logs")],
        capture_output=True, text=True, encoding="cp1252", errors="replace", env=env,
    )
    assert result.returncode == 0, result.stderr
    assert "Traceback" not in result.stderr
    assert "06. EQ Aggs" in result.stdout
    assert (tmp_path / "out.xlsx").exists()


# ──────────────────────────────────────── everything written stays UTF-8

def test_every_text_file_the_tool_writes_declares_utf8():
    """Windows would otherwise write the logs in the ANSI code page and lose the
    anchors on the way out as well as on the way in."""
    offenders = []
    for path in sorted(PACKAGE.glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            name = getattr(node.func, "id", None) or getattr(node.func, "attr", None)
            if name not in {"open", "write_text", "read_text", "FileHandler"}:
                continue
            mode = ""
            if name == "open" and len(node.args) > 1:
                mode = getattr(node.args[1], "value", "") or ""
            if "b" in mode:
                continue                       # binary: no encoding to declare
            if not any(k.arg == "encoding" for k in node.keywords):
                offenders.append(f"{path.name}:{node.lineno} {name}()")
    assert not offenders, "these need encoding='utf-8': " + ", ".join(offenders)


# ──────────────────────────────────────────────── no platform assumptions

def test_a_path_with_spaces_and_a_different_cwd(tmp_path, monkeypatch):
    """'C:\\Users\\Marie Lakreuss\\My Documents\\...' — the normal Windows case.

    Nothing here uses literal separators (the only ``/`` strings in the package are zip
    entry names, which OOXML fixes as ``/`` on every platform), but the way to know that
    is to run from somewhere awkward rather than to grep for it.
    """
    from datatransform.runner import run

    workspace = tmp_path / "My Treaty Packs" / "renewal 2026"
    workspace.mkdir(parents=True)
    source = workspace / "Intake Mexico v1.xlsx"
    shutil.copy2(MEXICO, source)

    monkeypatch.chdir(tmp_path)
    report = run(source, workspace / "out put.xlsx", workspace / "run logs")

    assert report.ok, [o.detail for o in report.outcomes if o.status == "error"]
    assert (workspace / "out put.xlsx").exists()
    assert list((workspace / "run logs").glob("*_process.log"))


def test_the_output_path_defaults_beside_the_source(tmp_path):
    """Given no -o, the tool must not write into the current directory — which on a
    shared drive may be somewhere the user cannot write at all."""
    from datatransform.runner import run

    source = tmp_path / "pack" / "Intake.xlsx"
    source.parent.mkdir()
    shutil.copy2(MEXICO, source)
    report = run(source, None, tmp_path / "logs")
    assert pathlib.Path(report.output).parent == source.parent


def test_the_only_runtime_dependency_is_openpyxl():
    """Everything else is the standard library, which is what makes 'copy the folder'
    a real installation route on a locked-down machine."""
    import tomllib

    meta = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert meta["project"]["dependencies"] == ["openpyxl>=3.1"]
    assert meta["project"]["requires-python"] == ">=3.10"
