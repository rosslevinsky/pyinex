"""Fixtures for the fake-Excel tests: locate the built binaries and start a host.

The binaries come from a build of the x64-release preset; set PYINEX_BUILD_DIR to use
another build's bin folder. These tests run only on Windows.

Every host starts with an environment chosen here, not the one pytest inherited: no
PYTHON... or PYINEX... variables, a PATH of only Windows' own folders, and PYINEX_PYTHON
naming the Python running pytest. So each test run embeds the Python it runs under, and
five runs under five Pythons test five Pythons. Discovery tests change this on purpose.
"""

import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]

if sys.platform != "win32":
    collect_ignore_glob = ["test_*.py"]


def _bin() -> Path:
    # Absolute, so tests that change the current folder still find the binaries.
    return Path(os.environ.get("PYINEX_BUILD_DIR", ROOT / "build" / "x64-release" / "bin")).resolve()


def _built(name: str) -> Path:
    path = _bin() / name
    if not path.is_file():
        pytest.fail(f"{path} not found: build the x64-release preset first, or set PYINEX_BUILD_DIR")
    return path


def host_env(**changes: str | None) -> dict[str, str]:
    """The environment a host starts with, with changes applied; None removes a variable."""
    env = {
        k: v
        for k, v in os.environ.items()
        if not k.upper().startswith(("PYTHON", "PYINEX", "__PYVENV")) and k.upper() not in ("PATH", "VIRTUAL_ENV")
    }
    windows = os.environ.get("SystemRoot", r"C:\Windows")
    env["PATH"] = f"{windows}\\System32;{windows}"
    env["PYINEX_PYTHON"] = sys.executable
    for name, value in changes.items():
        if value is None:
            env.pop(name, None)
        else:
            env[name] = value
    return env


HELPERS = Path(__file__).resolve().parent / "scripts" / "helpers.py"


def evaluate(host, expression, setup: str | None = None):
    """Runs setup, then returns str() of the expression, inside the add-in's Python, through
    PyCall and scripts/helpers.py. Names setup defines are kept between calls."""
    return host.call("PyCall", str(HELPERS), "evaluate", expression, *([] if setup is None else [setup]))


def attempt(host, expression, setup: str | None = None):
    """Like evaluate, but an exception comes back as "Type: message" instead of #VALUE!."""
    return host.call("PyCall", str(HELPERS), "attempt", expression, *([] if setup is None else [setup]))


@pytest.fixture(scope="session")
def xllhost() -> Path:
    return _built("xllhost.exe")


@pytest.fixture(scope="session")
def pyinex_xll() -> Path:
    xll = _built("pyinex.xll")
    _built("pyinex_core.dll")
    # A pyinex.ini beside the add-in would override PYINEX_PYTHON's absence in discovery
    # tests; tests that need one copy the add-in elsewhere.
    if (xll.parent / "pyinex.ini").exists():
        pytest.fail(f"{xll.parent / 'pyinex.ini'} must not exist: it would change which Python the tests embed")
    return xll


@pytest.fixture(scope="session")
def bad_xll() -> Path:
    return _built("badxll.xll")


@pytest.fixture(scope="session")
def fakepython() -> Path:
    return _built("fakepython.exe")


@pytest.fixture(scope="session", autouse=True)
def embedded_python_is_the_one_running_pytest(xllhost, pyinex_xll):
    """Fails the whole run unless the add-in embeds the Python running pytest."""
    from hostdriver import Host

    with Host(xllhost, pyinex_xll, env=host_env()) as h:
        h.open()
        status = h.call("PyInfo", "status")
        assert status == "running", status
        version = evaluate(h, "tuple(__import__('sys').version_info)")
        executable = evaluate(h, "__import__('sys').executable")
    assert version == str(tuple(sys.version_info)), f"embedded Python {version}, pytest runs under {sys.version_info}"
    assert os.path.normcase(executable) == os.path.normcase(sys.executable)


@pytest.fixture
def host(xllhost, pyinex_xll):
    """A fake Excel with pyinex.xll open."""
    from hostdriver import Host

    with Host(xllhost, pyinex_xll, env=host_env()) as h:
        h.open()
        yield h


@pytest.fixture
def bad_host(xllhost, bad_xll):
    """A fake Excel with the deliberately wrong add-in open."""
    from hostdriver import Host

    with Host(xllhost, bad_xll, env=host_env()) as h:
        h.open()
        yield h
