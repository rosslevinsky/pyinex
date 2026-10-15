"""The add-in embeds the Python it chose, started as that python.exe would start.

docs/architecture.md, "Choosing a Python" and "Virtual environments".
"""

import ast
import os
import subprocess
import sys

from conftest import HELPERS, evaluate, host_env

# What a Python reports about how it started. sys.path is compared after its first entry:
# the current folder, which `python -c` puts first, and in the embedded Python the folder
# of scripts/helpers.py, which PyCall put first when it loaded that script. UTF-8 mode and
# the encoding open() uses by default: on by default from 3.15, so a script reads files
# alike in Excel and at the command line only if the embedded Python follows suit.
STATE = (
    "[sys.version, sys.executable, sys._base_executable, sys.prefix, sys.exec_prefix, sys.base_prefix, "
    "sys.path[{start}:], sys.flags.utf8_mode, __import__('locale').getpreferredencoding(False)]"
)


def started_as(python: str, env: dict[str, str]) -> list:
    """How the python.exe itself starts, run with the environment the host has."""
    out = subprocess.run(
        [python, "-c", "import sys; print(repr(" + STATE.format(start=1) + "))"],
        env=env,
        capture_output=True,
        text=True,
        check=True,
    ).stdout
    return ast.literal_eval(out)


def embedded_state(host) -> list:
    assert evaluate(host, "sys.path[0]", "import sys") == str(HELPERS.parent)
    return ast.literal_eval(evaluate(host, STATE.format(start=1)))


def test_reports_the_chosen_python(host):
    assert evaluate(host, "sys.version", "import sys") == sys.version
    assert os.path.normcase(evaluate(host, "sys.executable")) == os.path.normcase(sys.executable)


def test_starts_as_its_python_exe_would(host):
    assert embedded_state(host) == started_as(sys.executable, host_env())


def test_python_is_running(host):
    assert host.call("PyInfo", "status") == "running"
    assert host.stderr == ""


def test_eval_returns_text_and_errors(host):
    from hostdriver import VALUE, Err

    assert evaluate(host, "'é ✓ 😀' * 2") == "é ✓ 😀é ✓ 😀"
    assert evaluate(host, "x + 1", "x = 41") == "42"
    assert evaluate(host, "x") == "41", "the namespace is kept between calls"
    assert evaluate(host, 1.0) == Err(VALUE), "an expression that is not text"
    assert evaluate(host, "1 / 0") == Err(VALUE)
    assert "ZeroDivisionError" in host.stderr
    assert evaluate(host, "0", "raise SystemExit(3)") == Err(VALUE), "SystemExit does not end Excel"
    assert evaluate(host, "'still running'") == "still running"
