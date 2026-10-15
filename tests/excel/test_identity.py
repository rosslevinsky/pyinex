"""Which Python Excel embedded: the one the driver chose for this run.

The session checks it before any test runs and stops on a mismatch (conftest.py); this
test records what was found. Run by hand, without the driver's --excel-run, nothing says
which Python was meant, and it skips; the driver always passes one, and counts a skip as a
failure.
"""

import os
import subprocess
import sys
from pathlib import Path

import pytest


def test_excel_embedded_the_python_this_run_chose(embedded, expected):
    print(f"Excel embedded {embedded['executable']}, Python {embedded['version']}, venv {embedded['venv']}", file=sys.stderr)
    if expected is None:
        pytest.skip("run without the driver's --excel-run: no Python was chosen to compare with")
    assert str(embedded["executable"]).lower() == str(expected["executable"]).lower()
    assert embedded["version"] == expected["version"]
    assert embedded["venv"] == expected["venv"]


def test_a_virtual_environment_sees_only_its_own_packages(app, expected, tmp_path):
    # The driver's virtual environments hold NumPy alone; every base Python has pytest.
    if expected is None:
        pytest.skip("run without the driver's --excel-run: no Python was chosen to compare with")
    # pytest must be in the base Python's own packages for its absence to show isolation.
    # The base is the run's Python, or for a virtual environment the one its pyvenv.cfg
    # names, which need not be the Python this process runs under.
    base = Path(expected["executable"])
    if expected["venv"]:
        cfg = (base.parent.parent / "pyvenv.cfg").read_text(encoding="utf-8")
        home = next(line.split("=", 1)[1].strip() for line in cfg.splitlines() if line.split("=", 1)[0].strip() == "home")
        base = Path(home) / "python.exe"
    check = "import importlib.util, sysconfig, pathlib, sys; sys.exit(not pathlib.Path(importlib.util.find_spec('pytest').origin).is_relative_to(sysconfig.get_paths()['purelib']))"
    assert subprocess.run([str(base), "-I", "-c", check]).returncode == 0, f"pytest is not in {base}'s own packages"
    found = app.Evaluate('PyInfo("venv")')
    assert found is expected["venv"]
    script = tmp_path / "isolation.py"
    script.write_text("import importlib.util\n\ndef f(name):\n    return importlib.util.find_spec(name) is not None\n", encoding="utf-8")
    assert app.Evaluate(f'PyCall("{script}", "f", "numpy")') is True
    assert app.Evaluate(f'PyCall("{script}", "f", "pytest")') is (not expected["venv"])


def test_files_open_with_the_encoding_python_exe_uses(app, expected, tmp_path):
    # docs/architecture.md, "Scripts": UTF-8 mode, on by default from 3.15, and the encoding
    # open() uses by default are the run's python.exe's, so a script reads a file alike in
    # Excel and at the command line. PYTHONUTF8 overrides the default; python.exe is run
    # with Excel's value of it, which Windows, not this process, gave Excel.
    if expected is None:
        pytest.skip("run without the driver's --excel-run: no Python was chosen to compare with")
    script = tmp_path / "encoding.py"
    script.write_text(
        "import locale, os, sys\n\ndef f():\n"
        "    return f'{sys.flags.utf8_mode}|{locale.getpreferredencoding(False)}|{os.environ.get(\"PYTHONUTF8\")}'\n",
        encoding="utf-8",
    )
    utf8_mode, encoding, override = app.Evaluate(f'PyCall("{script}", "f")').split("|", 2)
    env = {k: v for k, v in os.environ.items() if k.upper() != "PYTHONUTF8"}
    if override != "None":
        env["PYTHONUTF8"] = override
    report = "import sys, locale; print(sys.flags.utf8_mode, locale.getpreferredencoding(False))"
    want = subprocess.run([expected["executable"], "-c", report], env=env, capture_output=True, text=True, check=True).stdout.split()
    assert [utf8_mode, encoding] == want
    if override == "None" and tuple(int(n) for n in expected["version"].split()[0].split(".")[:2]) >= (3, 15):
        assert want == ["1", "utf-8"]
