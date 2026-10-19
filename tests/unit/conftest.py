"""Fixtures for the unit tests: load the test-only extension module built with the add-in.

_pyinex_testing.pyd is compiled from the same conversion sources as pyinex_core.dll and
exposes them to pytest, so these tests need no Excel and no fake Excel. It comes from a
build of the x64-release preset; set PYINEX_BUILD_DIR to use another build's bin folder.
These tests run only on Windows, except test_examples.py, which needs only openpyxl, and
the test_funcs*.py files, which need only Python (and NumPy, for test_funcs_annotations.py).

Values cross the module as descriptions of Excel's XLOPER12 values, as tuples:

    ("num", 1.5)  ("str", "text")  ("bool", True)  ("err", 42)  ("int", 7)
    ("nil",)  ("missing",)  ("multi", rows, columns, [item, ...])  ("raw", xltype)

("raw", xltype) is a value of that type with nothing in it, for types Pyinex does not
convert.
"""

import importlib.util
import os
import sys
import types
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]

if sys.platform != "win32":
    collect_ignore = [p.name for p in Path(__file__).parent.glob("test_*.py") if p.name not in ("test_examples.py", "test_funcs.py", "test_funcs_names.py", "test_funcs_reload.py", "test_funcs_signatures.py", "test_funcs_annotations.py")]


@pytest.fixture(scope="session")
def testing():
    """The _pyinex_testing module, loaded from the build's bin folder by its full path."""
    bin_dir = Path(os.environ.get("PYINEX_BUILD_DIR", ROOT / "build" / "x64-release" / "bin")).resolve()
    path = bin_dir / "_pyinex_testing.pyd"
    if not path.is_file():
        pytest.fail(f"{path} not found: build the x64-release preset first, or set PYINEX_BUILD_DIR")
    spec = importlib.util.spec_from_file_location("_pyinex_testing", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def describe(value):
    """The description of a plain Python value as Excel would pass it."""
    if isinstance(value, tuple) and value and isinstance(value[0], str) and value[0] in {
        "num", "str", "bool", "err", "int", "nil", "missing", "multi", "raw"
    }:
        return value  # Already a description.
    if value is None:
        return ("nil",)
    if isinstance(value, bool):
        return ("bool", value)
    if isinstance(value, (int, float)):
        return ("num", float(value))
    if isinstance(value, str):
        return ("str", value)
    if isinstance(value, (list, tuple)):  # Rows of values.
        rows = [list(r) for r in value]
        return ("multi", len(rows), len(rows[0]), [describe(v) for r in rows for v in r])
    raise TypeError(f"cannot describe {value!r}")


@pytest.fixture
def pycall(testing):
    """PyCall through the add-in's own code: pycall(path, function, *args) -> description.

    Arguments are plain values or descriptions; keywords are reload_on_change and folder.
    """

    def run(path, function, *args, **keywords):
        return testing.call(describe(str(path)), describe(function), *[describe(a) for a in args], **keywords)

    return run


@pytest.fixture
def script(tmp_path):
    """Writes a script: script(text, name="model.py", folder=tmp_path) -> its path."""
    import textwrap

    def write(text, name="model.py", folder=None):
        path = (folder or tmp_path) / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(textwrap.dedent(text).encode("utf-8"))  # \n, not \r\n: sizes are exact.
        return path

    return write


PYTHON_FILES = ROOT / "src" / "core" / "python"


@pytest.fixture
def core(monkeypatch):
    """call.py and funcs.py run in one namespace, as the core runs them, under plain Python,
    with a stand-in pyinex module offering the decorator, Range, and a _stopped that says
    Excel never stopped a call (test_funcs_range.py stands in for the rest)."""
    names = {}
    for name in ("call.py", "funcs.py"):
        exec(compile((PYTHON_FILES / name).read_bytes(), f"<pyinex/{name}>", "exec", dont_inherit=True), names)
    monkeypatch.setitem(sys.modules, "pyinex", types.SimpleNamespace(func=names["func"], Range=names.get("Range"), _stopped=lambda: False))
    return names


@pytest.fixture
def load(core, script):
    """Writes a script and returns load_functions' rows for it."""

    def run(text, name="model.py"):
        return core["load_functions"](str(script(text, name=name)))

    return run
