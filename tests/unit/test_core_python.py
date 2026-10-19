"""The core's Python files (src/core/python/) compile under the Python running the tests,
and the build compiles each in whole; bootstrap.py's helpers, run from the text the build
compiled in.

The build compiles their text into the core, which compiles it again when Python starts
(docs/architecture.md, "The core's Python files"). CI runs this under every supported
Python, so a syntax only a newer Python accepts fails here, not in Excel.
"""

import io
import os
import sys
from pathlib import Path

import pytest

FOLDER = Path(__file__).resolve().parents[2] / "src" / "core" / "python"
FILES = sorted(FOLDER.glob("*.py"))


def test_the_core_files_are_there():
    assert {"bootstrap.py", "call.py", "excel_error.py"} <= {p.name for p in FILES}


@pytest.mark.parametrize("path", FILES, ids=lambda p: p.name)
def test_compiles(path):
    compile(path.read_bytes(), f"<pyinex/{path.name}>", "exec", dont_inherit=True)


def test_the_build_compiles_in_each_file_whole(testing):
    assert testing.sources == {p.name: p.read_bytes().decode("utf-8") for p in FILES}


@pytest.fixture
def bootstrap(testing, monkeypatch):
    """bootstrap.py run from the text the build compiled in, as the core runs it, with
    standard output and error streams of its own to set up: (its names, the two streams)."""
    streams = [io.TextIOWrapper(io.BytesIO(), encoding="cp1252") for _ in range(2)]
    monkeypatch.setattr(sys, "stdout", streams[0])
    monkeypatch.setattr(sys, "stderr", streams[1])
    names = {}
    exec(compile(testing.sources["bootstrap.py"], "<pyinex/bootstrap.py>", "exec", dont_inherit=True), names)
    return names, streams


def test_bootstrap_writes_output_in_utf8_a_line_at_a_time(bootstrap):
    _, streams = bootstrap
    for stream in streams:
        assert (stream.encoding, stream.errors, stream.line_buffering) == ("utf-8", "backslashreplace", True)


def test_bootstrap_info(bootstrap):
    info = bootstrap[0]["info"]
    assert info("version") == sys.version
    assert info("prefix") == sys.prefix
    assert info("venv") is (sys.prefix != sys.base_prefix)


@pytest.mark.parametrize("kept_by", [None, "PYTHONPATH"])
def test_bootstrap_fix_up_takes_the_host_s_folder_off_the_path(bootstrap, monkeypatch, tmp_path, kept_by):
    host = tmp_path / "Office"
    monkeypatch.setattr(sys, "path", ["", str(tmp_path / "lib"), str(host), str(host) + os.sep])
    monkeypatch.setattr(sys, "_base_executable", sys.executable, raising=False)
    if kept_by:
        monkeypatch.setenv("PYTHONPATH", str(host))
    else:
        monkeypatch.delenv("PYTHONPATH", raising=False)
    bootstrap[0]["fix_up"](str(tmp_path / "python.exe"), str(host / "EXCEL.EXE"))
    assert sys._base_executable == str(tmp_path / "python.exe")
    if kept_by:
        assert sys.path == ["", str(tmp_path / "lib"), str(host), str(host) + os.sep]
    else:
        assert sys.path == ["", str(tmp_path / "lib")]
