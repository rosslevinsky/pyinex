"""Choosing a Python: where Pyinex looks, virtual environments, inherited variables, and
each refusal (docs/architecture.md, "Choosing a Python").

Installations the runner cannot have — too old, 32-bit, broken — are crafted in temporary
folders, some with fakepython.exe standing in for python.exe.
"""

import os
import shutil
import struct
import subprocess
import sys
from pathlib import Path

import _winapi
import pytest

from conftest import evaluate, host_env
from hostdriver import VALUE, Err, Host
from test_embedded_python import embedded_state, started_as

# The Python running the tests, as a plain installation even if pytest runs in a venv.
BASE = Path(getattr(sys, "_base_executable", sys.executable))
PYTHON_DLL = Path(_winapi.GetModuleFileName(sys.dllhandle))
WINDOWS = os.environ.get("SystemRoot", r"C:\Windows")
SYSTEM_PATH = f"{WINDOWS}\\System32;{WINDOWS}"


@pytest.fixture
def start(xllhost, pyinex_xll):
    """start(xll=None, before_open=None, **env changes): a host with the add-in open.

    PYINEX_PYTHON is unset unless given. before_open(host) runs before the add-in opens.
    """
    hosts = []

    def _start(xll=None, before_open=None, **changes):
        changes.setdefault("PYINEX_PYTHON", None)
        h = Host(xllhost, xll or pyinex_xll, env=host_env(**changes))
        hosts.append(h)
        if before_open:
            before_open(h)
        h.open()
        return h

    yield _start
    for h in hosts:
        h.quit()


@pytest.fixture
def addin_copy(tmp_path, pyinex_xll):
    """The add-in and its core copied to a folder of their own, for a pyinex.ini beside them."""
    folder = tmp_path / "addin"
    folder.mkdir()
    for name in ("pyinex.xll", "pyinex_core.dll"):
        shutil.copy2(pyinex_xll.parent / name, folder / name)
    return folder / "pyinex.xll"


def write_ini(xll: Path, encoding="utf-8", **python) -> None:
    lines = ["[python]"] + [f"{k} = {v}" for k, v in python.items()]
    (xll.parent / "pyinex.ini").write_text("\n".join(lines) + "\n", encoding=encoding)


def executable(host) -> str:
    return os.path.normcase(evaluate(host, "__import__('sys').executable"))


def same(path) -> str:
    return os.path.normcase(str(path))


def importable(host, module: str) -> bool:
    return evaluate(host, f"__import__('importlib.util').util.find_spec({module!r}) is not None") == "True"


@pytest.fixture(scope="session")
def venvs(tmp_path_factory):
    """Two virtual environments made from BASE: an ordinary one and one with
    --system-site-packages, each holding a module only it has."""
    root = tmp_path_factory.mktemp("venvs")
    made = {}
    for name, extra in (("plain", []), ("system", ["--system-site-packages"])):
        folder = root / name
        subprocess.run([str(BASE), "-m", "venv", "--without-pip", *extra, str(folder)], check=True)
        (folder / "Lib" / "site-packages" / "venv_only_module.py").write_text("", encoding="utf-8")
        made[name] = folder / "Scripts" / "python.exe"
    return made


@pytest.fixture
def user_site(tmp_path):
    """A user site-packages (PYTHONUSERBASE) holding a module only it has."""
    base = tmp_path / "userbase"
    site_packages = subprocess.run(
        [str(BASE), "-c", "import site; print(site.getusersitepackages())"],
        env=host_env(PYTHONUSERBASE=str(base)),
        capture_output=True,
        text=True,
        check=True,
    ).stdout.strip()
    Path(site_packages).mkdir(parents=True)
    (Path(site_packages) / "user_only_module.py").write_text("", encoding="utf-8")
    return str(base)


# Where Pyinex looks ------------------------------------------------------------------


def test_first_python_on_path(start):
    h = start(PATH=f"{BASE.parent};{SYSTEM_PATH}")
    assert h.call("PyInfo", "status") == "running"
    assert executable(h) == same(BASE)


def test_path_skips_the_store_placeholder(start, tmp_path, fakepython):
    store = tmp_path / "Microsoft" / "WindowsApps"
    store.mkdir(parents=True)
    shutil.copy2(fakepython, store / "python.exe")
    h = start(PATH=f"{store};{BASE.parent};{SYSTEM_PATH}")
    assert executable(h) == same(BASE)


# The registry, as python.org's installer writes it (PEP 514), in the host's own registry.

# The running Python's entry: a plain "3.N" names an x64 build.
TAG = f"3.{sys.version_info.minor}"
TOO_OLD = "pyinex-probe:version=3.9"  # A fake that says this is refused if it is chosen.


def registered(*entries):
    """A before_open for start() that records each (hive, version, executable, folder)."""
    return lambda h: [h.register_python(*entry) for entry in entries]


def test_registry_python_when_path_has_none(start):
    h = start(PATH=SYSTEM_PATH, before_open=registered(("user", TAG, BASE, None)))
    assert h.call("PyInfo", "status") == "running"
    assert executable(h) == same(BASE)


def test_registry_install_folder_without_an_executable_path(start):
    h = start(PATH=SYSTEM_PATH, before_open=registered(("machine", TAG, None, BASE.parent)))
    assert executable(h) == same(BASE)


def test_registry_comes_after_path(start, tmp_path, fakepython):
    other = fake(tmp_path, fakepython)
    h = start(PATH=f"{BASE.parent};{SYSTEM_PATH}", FAKEPYTHON_OUTPUT=TOO_OLD,
              before_open=registered(("user", "3.99", other, None)))
    assert executable(h) == same(BASE)


def test_registry_newest_version_by_number(start, tmp_path, fakepython):
    old = fake(tmp_path, fakepython)
    # As text, "3.9" sorts after "3.1N"; as a version it is older.
    h = start(PATH=SYSTEM_PATH, FAKEPYTHON_OUTPUT=TOO_OLD,
              before_open=registered(("user", "3.9", old, None), ("user", TAG, BASE, None)))
    assert executable(h) == same(BASE)


def test_registry_skips_other_builds_and_missing_files(start, tmp_path, fakepython):
    other = fake(tmp_path, fakepython)
    entries = [("user", name, other, None) for name in ("3.99-32", "3.99-arm64", "3.99t", "4.0", "3.x")]
    entries += [("user", "3.98", tmp_path / "gone" / "python.exe", None), ("user", TAG, BASE, None)]
    h = start(PATH=SYSTEM_PATH, FAKEPYTHON_OUTPUT=TOO_OLD, before_open=registered(*entries))
    assert executable(h) == same(BASE)


def test_registry_skips_the_store_placeholder(start, tmp_path, fakepython):
    store = fake(tmp_path / "Microsoft" / "WindowsApps", fakepython)
    h = start(PATH=SYSTEM_PATH, FAKEPYTHON_OUTPUT=TOO_OLD,
              before_open=registered(("user", "3.99", store, None), ("machine", TAG, BASE, None)))
    assert executable(h) == same(BASE)


def test_registry_user_before_machine_at_the_same_version(start, tmp_path, fakepython):
    other = fake(tmp_path, fakepython)
    h = start(PATH=SYSTEM_PATH, FAKEPYTHON_OUTPUT=TOO_OLD,
              before_open=registered(("machine", TAG, other, None), ("user", TAG, BASE, None)))
    assert executable(h) == same(BASE)


def test_registry_newer_machine_entry_before_older_user_entry(start, tmp_path, fakepython):
    old = fake(tmp_path, fakepython)
    h = start(PATH=SYSTEM_PATH, FAKEPYTHON_OUTPUT=TOO_OLD,
              before_open=registered(("user", "3.9", old, None), ("machine", TAG, BASE, None)))
    assert executable(h) == same(BASE)


def test_pyinex_python_names_a_folder(start):
    h = start(PYINEX_PYTHON=str(BASE.parent), PATH=SYSTEM_PATH)
    assert executable(h) == same(BASE)


def test_pyinex_python_comes_before_the_ini_and_path(start, addin_copy, venvs):
    write_ini(addin_copy, executable=BASE)
    h = start(addin_copy, PYINEX_PYTHON=str(venvs["plain"]), PATH=f"{BASE.parent};{SYSTEM_PATH}")
    assert executable(h) == same(venvs["plain"])


def test_the_ini_comes_before_path(start, addin_copy, venvs):
    write_ini(addin_copy, executable=venvs["plain"])
    h = start(addin_copy, PATH=f"{BASE.parent};{SYSTEM_PATH}")
    assert executable(h) == same(venvs["plain"])


def test_ini_path_relative_to_the_addin(start, addin_copy, venvs):
    write_ini(addin_copy, executable=os.path.relpath(venvs["plain"], addin_copy.parent))
    h = start(addin_copy)
    assert executable(h) == same(venvs["plain"])


@pytest.mark.parametrize("encoding", ["utf-8", "utf-8-sig", "utf-16"])
def test_ini_names_a_path_beyond_the_ansi_code_page(start, addin_copy, tmp_path, encoding):
    # A path no 8-bit code page holds, in a pyinex.ini saved as editors save it: UTF-8 with
    # or without its byte-order mark, or UTF-16. The folder is a junction to BASE's.
    link = tmp_path / "Bücher ✓ 😀"
    subprocess.run(["cmd", "/c", "mklink", "/J", str(link), str(BASE.parent)], check=True, capture_output=True)
    write_ini(addin_copy, encoding=encoding, executable=link / BASE.name)
    h = start(addin_copy)
    assert h.call("PyInfo", "status") == "running"
    assert executable(h) == same(link / BASE.name)


@pytest.mark.parametrize(
    "text",
    [
        "[python]\rexecutable = {base}\r",  # Carriage returns alone end lines too.
        '[python]\nexecutable = " {base} "\n',  # Spaces inside quotes go as Windows' do.
        "; a comment\n[other]\nexecutable = elsewhere\n[Python]\nExecutable={base}\n",  # Names ignore case.
    ],
)
def test_ini_is_read_by_windows_rules(start, addin_copy, text):
    (addin_copy.parent / "pyinex.ini").write_text(text.format(base=BASE), encoding="utf-8", newline="")
    h = start(addin_copy)
    assert executable(h) == same(BASE)


def test_ini_reads_only_the_first_python_section(start, addin_copy, tmp_path):
    # As Windows does: a second [python] section is not read, so PATH chooses.
    text = f"[python]\nverbose = 0\n[python]\nexecutable = {tmp_path / 'missing' / 'python.exe'}\n"
    (addin_copy.parent / "pyinex.ini").write_text(text, encoding="utf-8")
    h = start(addin_copy, PATH=f"{BASE.parent};{SYSTEM_PATH}")
    assert executable(h) == same(BASE)


def test_ini_verbose_sets_pythonverbose(start, addin_copy):
    write_ini(addin_copy, executable=BASE, verbose=1)
    h = start(addin_copy)
    assert evaluate(h, "__import__('os').environ.get('PYTHONVERBOSE')") == "1"
    assert evaluate(h, "__import__('sys').flags.verbose") == "1"


# Virtual environments -----------------------------------------------------------------


def test_venv_starts_as_its_python_exe_would(start, venvs):
    for python in venvs.values():
        h = start(PYINEX_PYTHON=str(python))
        assert embedded_state(h) == started_as(str(python), host_env(PYINEX_PYTHON=None))


@pytest.fixture(scope="session")
def base_only_module() -> str:
    """A module in the base installation's site-packages: pip, which python.org installs
    there and which the venvs, made --without-pip, lack."""
    import importlib.util
    import sysconfig

    spec = importlib.util.find_spec("pip")
    purelib = sysconfig.get_path("purelib", vars={"base": sys.base_prefix, "platbase": sys.base_prefix})
    if spec is None or not same(spec.origin).startswith(same(purelib)):
        pytest.skip(f"pip is not in {purelib}")
    return "pip"


def test_venv_packages_and_isolation(start, venvs, base_only_module, tmp_path):
    # An empty user site, so only base site-packages could supply the base module.
    h = start(PYINEX_PYTHON=str(venvs["plain"]), PYTHONUSERBASE=str(tmp_path))
    assert same(evaluate(h, "__import__('sys').prefix")) == same(venvs["plain"].parent.parent)
    assert importable(h, "venv_only_module")
    assert not importable(h, base_only_module)


def test_system_site_packages_venv_sees_base_packages(start, venvs, base_only_module, tmp_path):
    h = start(PYINEX_PYTHON=str(venvs["system"]), PYTHONUSERBASE=str(tmp_path))
    assert importable(h, "venv_only_module")
    assert importable(h, base_only_module)


@pytest.mark.parametrize(
    "python, sees_user_site",
    [("base", True), ("system", True), ("plain", False)],
)
def test_user_site_packages(start, venvs, user_site, python, sees_user_site):
    exe = BASE if python == "base" else venvs[python]
    h = start(PYINEX_PYTHON=str(exe), PYTHONUSERBASE=user_site)
    assert importable(h, "user_only_module") == sees_user_site


# Inherited environment variables -------------------------------------------------------


def test_pythonhome_is_removed(start, tmp_path):
    h = start(PYINEX_PYTHON=str(BASE), PYTHONHOME=str(tmp_path / "elsewhere"))
    assert h.call("PyInfo", "status") == "running"
    assert evaluate(h, "__import__('os').environ.get('PYTHONHOME')") == "None"
    assert "Pyinex: ignoring PYTHONHOME" in h.stderr


def test_pythonpath_is_honored(start, tmp_path):
    (tmp_path / "pythonpath_module.py").write_text("", encoding="utf-8")
    h = start(PYINEX_PYTHON=str(BASE), PYTHONPATH=str(tmp_path))
    assert importable(h, "pythonpath_module")
    assert f"Pyinex: PYTHONPATH is set, adding to sys.path: {tmp_path}" in h.stderr


def test_newer_than_tested_python_is_used_with_a_note(start, tmp_path, fakepython):
    fake = tmp_path / "python.exe"
    shutil.copy2(fakepython, fake)
    output = f"pyinex-probe:version=3.99|pyinex-probe:gil_disabled=0|pyinex-probe:dll={PYTHON_DLL}|pyinex-probe:base_executable={BASE}"
    # Python 3.10 looks for its library beside its program, which here is the fake's, or else
    # in the registry, which in the host is empty; 3.11 on look beside its DLL.
    library = {"PYTHONPATH": f"{BASE.parent / 'Lib'};{BASE.parent / 'DLLs'}"} if sys.version_info < (3, 11) else {}
    h = start(PYINEX_PYTHON=str(fake), FAKEPYTHON_OUTPUT=output, **library)
    assert h.call("PyInfo", "status") == "running"
    assert "Pyinex: Python 3.99 is newer than any Pyinex was tested with (3.15); using it anyway." in h.stderr


# Refusals ----------------------------------------------------------------------------


def crafted_program(path: Path, machine: int) -> Path:
    """A file with just enough of a Windows program's header to name its CPU."""
    path.parent.mkdir(parents=True, exist_ok=True)
    dos = bytearray(64)
    dos[0:2] = b"MZ"
    dos[0x3C:0x40] = struct.pack("<I", 64)
    path.write_bytes(bytes(dos) + b"PE\0\0" + struct.pack("<H", machine) + bytes(18))
    return path


def broken_install(folder: Path) -> Path:
    """A copy of the running Python's program and DLLs with no standard library. Its ._pth
    file stops it finding the real installation's library through the registry."""
    folder.mkdir()
    for name in ["python.exe", PYTHON_DLL.name, "python3.dll", "vcruntime140.dll", "vcruntime140_1.dll"]:
        if (BASE.parent / name).exists():
            shutil.copy2(BASE.parent / name, folder / name)
    (folder / f"{PYTHON_DLL.stem}._pth").write_text("missing-library\n", encoding="utf-8")
    return folder / "python.exe"


def fake(tmp_path, fakepython, name="python.exe") -> Path:
    path = tmp_path / "fake" / name
    path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(fakepython, path)
    return path


def venv_without_base(tmp_path, fakepython) -> Path:
    exe = fake(tmp_path, fakepython)
    (exe.parent / "pyvenv.cfg").write_text(f"home = {tmp_path / 'gone'}\n", encoding="utf-8")
    return exe


# Each case: a function making (environment changes, the reason Pyinex should give).
REFUSALS = {
    "missing file": lambda t, f: ({"PYINEX_PYTHON": str(t / "nowhere" / "python.exe")}, "file not found"),
    "store placeholder": lambda t, f: (
        {"PYINEX_PYTHON": str(fake(t / "Microsoft" / "WindowsApps", f))},
        "it is the Microsoft Store placeholder, not an installed Python; install Python from python.org",
    ),
    "free-threaded by name": lambda t, f: (
        {"PYINEX_PYTHON": str(fake(t, f, "python3.13t.exe"))},
        "free-threaded builds are not supported",
    ),
    "x86": lambda t, f: (
        {"PYINEX_PYTHON": str(crafted_program(t / "x86" / "python.exe", 0x14C))},
        "it is a 32-bit (x86) build; Excel needs the x64 build",
    ),
    "arm64": lambda t, f: (
        {"PYINEX_PYTHON": str(crafted_program(t / "arm64" / "python.exe", 0xAA64))},
        "it is an ARM64 build; Excel needs the x64 build",
    ),
    "venv base gone": lambda t, f: (
        {"PYINEX_PYTHON": str(venv_without_base(t, f))},
        f"its pyvenv.cfg names base Python {t / 'gone'}, which does not exist",
    ),
    "too old": lambda t, f: (
        {"PYINEX_PYTHON": str(fake(t, f)), "FAKEPYTHON_OUTPUT": "pyinex-probe:version=3.9"},
        "version 3.9 is too old; Pyinex needs 3.10 or later",
    ),
    "free-threaded when run": lambda t, f: (
        {
            "PYINEX_PYTHON": str(fake(t, f)),
            "FAKEPYTHON_OUTPUT": f"pyinex-probe:version=3.14|pyinex-probe:gil_disabled=1|pyinex-probe:dll={PYTHON_DLL}"
            f"|pyinex-probe:base_executable={BASE}",
        },
        "free-threaded builds are not supported",
    ),
    "exits with an error after reporting": lambda t, f: (
        {
            "PYINEX_PYTHON": str(fake(t, f)),
            "FAKEPYTHON_OUTPUT": f"pyinex-probe:version=3.14|pyinex-probe:gil_disabled=0|pyinex-probe:dll={PYTHON_DLL}"
            f"|pyinex-probe:base_executable={BASE}|Traceback: something failed afterwards",
            "FAKEPYTHON_EXIT": "1",
        },
        "it failed to start (exit code 1): Traceback: something failed afterwards",
    ),
    "fails to start": lambda t, f: (
        {"PYINEX_PYTHON": str(fake(t, f)), "FAKEPYTHON_OUTPUT": "Fatal Python error: boom", "FAKEPYTHON_EXIT": "1"},
        "it failed to start (exit code 1): Fatal Python error: boom",
    ),
    "broken installation": lambda t, f: (
        {"PYINEX_PYTHON": str(broken_install(t / "broken"))},
        "it failed to start (exit code ",
    ),
}


@pytest.mark.parametrize("case", REFUSALS)
def test_refusal(start, tmp_path, fakepython, case):
    changes, reason = REFUSALS[case](tmp_path, fakepython)
    h = start(**changes)
    assert_refused(h, f"cannot use Python {changes['PYINEX_PYTHON']}: {reason}")


def test_refusal_nothing_on_path(start):
    h = start(PATH=SYSTEM_PATH)
    assert_refused(
        h, "cannot find Python: no python.exe on PATH, and none registered. Choose one with PYINEX_PYTHON or pyinex.ini."
    )


def test_refusal_only_the_store_placeholder_on_path(start, tmp_path, fakepython):
    store = fake(tmp_path / "Microsoft" / "WindowsApps", fakepython).parent
    h = start(PATH=f"{store};{SYSTEM_PATH}")
    assert_refused(
        h,
        "cannot find Python: no python.exe on PATH, and none registered; skipped the Microsoft Store placeholder at "
        f"{store / 'python.exe'}.",
    )


def test_refusal_python_already_loaded(start):
    h = start(PYINEX_PYTHON=str(BASE), before_open=lambda h: h.load(PYTHON_DLL))
    assert_refused(
        h,
        f"cannot use Python {BASE}: another add-in has already loaded Python from {PYTHON_DLL}; "
        "Pyinex cannot share Excel with it",
    )


def test_refusal_is_retried_when_the_addin_reopens(start, addin_copy):
    write_ini(addin_copy, executable=addin_copy.parent / "missing" / "python.exe")
    h = start(addin_copy)
    assert h.call("PyInfo", "status").endswith("file not found. Choose another with PYINEX_PYTHON or pyinex.ini.")
    write_ini(addin_copy, executable=BASE)
    assert "file not found" in h.call("PyInfo", "status"), "retried before the add-in reopened"
    h.close()
    h.open()
    assert h.call("PyInfo", "status") == "running"


def assert_refused(h, message: str) -> None:
    """The refusal's message, and each function's result after it (docs/architecture.md,
    "Refusals")."""
    status = h.call("PyInfo", "status")
    assert status.startswith(message), status
    assert f"Pyinex: {status}\n" in h.stderr.replace("\r\n", "\n")
    assert evaluate(h, "1") == Err(VALUE)
    assert h.call("PyInfo", "pyinex") == str(h.xll)
    for item in ("python", "version", "executable", "prefix", "venv", "addin_version"):
        assert h.call("PyInfo", item) == Err(VALUE), item
    table = h.call("PyInfo")
    assert table[0] == ("status", status) and table[2] == ("pyinex", str(h.xll))
    assert [value for name, value in table if name not in ("status", "pyinex")] == [Err(VALUE)] * 6
    assert h.call("PyConsole", True) is True
    assert h.call("PyReloadOnChange", False) is False
    assert h.stderr.count("Pyinex: cannot") == 1, "the refusal was repeated"


@pytest.fixture
def in_folder(tmp_path):
    """Runs the test, and the hosts it starts, with tmp_path as the current folder."""
    cwd = os.getcwd()
    os.chdir(tmp_path)
    yield tmp_path
    os.chdir(cwd)


def test_the_check_run_ignores_modules_in_the_current_folder(start, in_folder):
    # `python -c` would import this before the real sysconfig (docs/architecture.md,
    # "Refusals"); the embedded Python never looks in the current folder.
    (in_folder / "sysconfig.py").write_text("raise SystemExit('imported the stray sysconfig.py')\n", encoding="utf-8")
    h = start(PYINEX_PYTHON=str(BASE))
    assert h.call("PyInfo", "status") == "running"


def test_the_check_run_resolves_relative_paths_as_python_will(start, in_folder):
    # With PYTHONPATH=".", Python imports this broken encodings module at start-up, from
    # Excel's current folder. The check run must meet it too, and refuse, rather than pass
    # elsewhere and leave Python to end the process.
    (in_folder / "encodings.py").write_text("raise SystemExit('broken encodings')\n", encoding="utf-8")
    h = start(PYINEX_PYTHON=str(BASE), PYTHONPATH=".")
    assert h.call("PyInfo", "status").startswith(f"cannot use Python {BASE}: it failed to start (exit code ")


def test_a_check_run_with_much_output_is_read_while_it_runs(start, tmp_path):
    # More output at start-up than a pipe holds: read only after the run ends, it would
    # wait forever on the full pipe, and Pyinex would time out after 60 s.
    (tmp_path / "sitecustomize.py").write_text(
        "import sys\nsys.stdout.write('x' * 2_000_000)\nsys.stdout.flush()\n", encoding="utf-8"
    )
    h = start(PYINEX_PYTHON=str(BASE), PYTHONPATH=str(tmp_path))
    assert h.call("PyInfo", "status") == "running"


def test_a_missing_core_is_retried_when_the_addin_reopens(start, addin_copy, pyinex_xll):
    core = addin_copy.parent / "pyinex_core.dll"
    core.unlink()
    h = start(addin_copy, PYINEX_PYTHON=str(BASE))
    assert h.call("PyInfo", "status").startswith(f"cannot load {core}")
    shutil.copy2(pyinex_xll.parent / "pyinex_core.dll", core)
    h.close()
    h.open()
    assert h.call("PyInfo", "status") == "running"


def test_a_relative_folder_on_path_gives_an_absolute_executable(start, in_folder):
    relative = os.path.relpath(BASE.parent, in_folder)
    h = start(PATH=f"{relative};{SYSTEM_PATH}")
    assert executable(h) == same(BASE)


def test_pythonpath_naming_the_host_folder_keeps_it(start, tmp_path, xllhost, pyinex_xll):
    # Pyinex drops the host program's folder from sys.path, as python.exe would not have
    # it, unless PYTHONPATH names it.
    folder = tmp_path / "excel"
    folder.mkdir()
    shutil.copy2(xllhost, folder / "xllhost.exe")
    (folder / "beside_the_host.py").write_text("", encoding="utf-8")
    hosts = []
    try:
        for pythonpath, found in ((None, False), (str(folder), True)):
            h = Host(folder / "xllhost.exe", pyinex_xll, env=host_env(PYTHONPATH=pythonpath))
            hosts.append(h)
            h.open()
            assert importable(h, "beside_the_host") == found
    finally:
        for h in hosts:
            h.quit()
