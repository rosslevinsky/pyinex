"""Runs the real-Excel tests (tests/excel) once per Python, each in a fresh Excel.

Run at the signed-in, unlocked Windows desktop, with Excel closed, from a checkout of the
commit the add-in was built from, under a Python with pytest and pywin32:

    py -3.14 scripts\\excel_tests.py C:\\path\\to\\pyinex.xll [--out FOLDER]
        [--python 3.12 --python venv ...] [--tests tests\\excel\\test_examples.py ...]
        [--timeout SECONDS] [--keep-going]

For each of Python 3.10-3.15, a virtual environment made from 3.15, and another in a
folder whose name holds spaces and Unicode (or only the runs given with --python), it
copies the add-in and pyinex_core.dll into a folder of their own with a pyinex.ini naming
that Python, and runs pytest on the tests, which start a new Excel, load that copy, and
check Excel embedded that Python before anything else. Then, under 3.15, the runs that
need an Excel of their own: a pyinex.ini naming no Python, PyConsole as the first call,
each PyInfo item as the first call, and the add-in unloaded and reloaded; each runs only
the tests marked for it. It asks for no input and opens no window that waits for a person.

A run passes only if every test passed: a failure, an error, a skipped test, no tests, a
run that reached its timeout, an Excel that had to be ended, or a setting that could not
be put back each fail it. The driver exits 0 only if every run passed. A run whose Excel
had to be ended, that reached its timeout, or in which no tests ran stops the driver there,
since the runs left would almost always meet the same: the add-in or the machine, not one
Python, is at fault. Failed tests do not stop it, since which Pythons they fail under says
something of why. --keep-going runs the rest whatever happens.

It records the settings the tests may change (Windows' app mode, Excel's options and the
add-ins it loads at start-up) before each run, and puts back and checks them after, even
when a run fails, reaches its timeout, or is stopped with Ctrl+C. Everything it writes is
in a temporary folder, removed at the end, and in --out: a JUnit XML file per run,
summary.json, and log.txt, the whole log.

It refuses to start, saying why, at a locked screen or outside the desktop session, while
an Excel is running, or when the environment Excel will get sets a variable that would
choose or steer its Python. docs/testing.md has the details.
"""

import argparse
import ctypes
import datetime
import hashlib
import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
import xml.etree.ElementTree as ElementTree
from ctypes import wintypes
from pathlib import Path

if sys.platform == "win32":
    import winreg

ROOT = Path(__file__).resolve().parent.parent
PYTHONS = ["3.10", "3.11", "3.12", "3.13", "3.14", "3.15"]
VENV = "venv"  # The run in a virtual environment made from the newest Python.
UNICODE_VENV = "unicode-venv"  # The same, in a folder whose name holds spaces and Unicode.
VENV_FOLDERS = {VENV: "venv", UNICODE_VENV: "venv Bücher ✓ 😀"}
# Runs that need an Excel of their own, each once under the newest Python, by the kind of
# run their tests are marked for (tests/excel/conftest.py): a refused Python, PyConsole as
# the first call, each PyInfo item as the first call, and the add-in unloaded and reloaded.
FIRST_ITEMS = ["status", "python", "pyinex", "version", "executable", "prefix", "venv", "addin_version"]
OWN_EXCEL = {"refused": "refused", "console-first": "console_first", "lifetime": "lifetime"}
OWN_EXCEL.update({f"first-{item}": "first_call" for item in FIRST_ITEMS})
RUNS = [*PYTHONS, VENV, UNICODE_VENV, *OWN_EXCEL]

# Variables that would choose or steer Excel's Python if the user's environment set them:
# PYINEX_PYTHON wins over pyinex.ini, and the rest change what Python loads.
REFUSED_VARIABLES = ["PYINEX_PYTHON", "PYTHONHOME", "PYTHONPATH", "PYTHONSTARTUP", "PYTHONUSERBASE"]

# The settings a test may change, each a whole registry key under HKEY_CURRENT_USER:
# Windows' app mode, Excel's options (multi-threaded recalculation among them), and the
# add-ins Excel loads at start-up (Options' OPEN values and the Add-in Manager's list).
PERSONALIZE = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
SETTINGS = [
    PERSONALIZE,
    r"Software\Microsoft\Office\16.0\Excel\Options",
    r"Software\Microsoft\Office\16.0\Excel\Add-in Manager",
]

PROBE = "import json, sys, importlib.util; print(json.dumps({'executable': sys.executable, 'version': sys.version, 'numpy': importlib.util.find_spec('numpy') is not None}))"


class Log:
    """Prints a line and appends it to log.txt."""

    def __init__(self, path):
        self.file = open(path, "a", encoding="utf-8")
        self.lock = threading.Lock()

    def __call__(self, text=""):
        with self.lock:
            print(text, flush=True)
            self.file.write(text + "\n")
            self.file.flush()


# Settings: recorded before a run, put back and checked after.


def read_key(path, root=None):
    """A key's values as {name: (value, type)}, or None if the key does not exist."""
    root = winreg.HKEY_CURRENT_USER if root is None else root
    try:
        key = winreg.OpenKey(root, path, 0, winreg.KEY_READ)
    except FileNotFoundError:
        return None
    with key:
        values, i = {}, 0
        while True:
            try:
                name, value, kind = winreg.EnumValue(key, i)
            except OSError:
                return values
            values[name] = (value, kind)
            i += 1


def record(keys, root=None):
    return {path: read_key(path, root) for path in keys}


def restore(recorded, root=None):
    """Puts each key's values back as recorded, then reads them again; returns a problem
    for each key that is not as it was."""
    root = winreg.HKEY_CURRENT_USER if root is None else root
    problems = []
    for path, before in recorded.items():
        try:
            now = read_key(path, root)
            if now == before:
                continue
            if before is None:
                with winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS) as key:
                    for name in now:
                        winreg.DeleteValue(key, name)
                try:
                    winreg.DeleteKey(root, path)
                except OSError:
                    pass  # It has subkeys of its own; its values are gone, which is what was recorded.
            else:
                with winreg.CreateKeyEx(root, path, 0, winreg.KEY_ALL_ACCESS) as key:
                    for name in set(now or {}) - set(before):
                        winreg.DeleteValue(key, name)
                    for name, (value, kind) in before.items():
                        if (now or {}).get(name) != (value, kind):
                            winreg.SetValueEx(key, name, 0, kind, value)
            after = read_key(path, root)
        except OSError as e:
            problems.append(f"could not put back HKCU\\{path}: {e}")
            continue
        if after != before and not (before is None and not after):
            problems.append(f"HKCU\\{path} is not as it was after putting it back")
        if path == PERSONALIZE:
            announce_mode()
    return problems


def announce_mode():
    """Tells open windows the app mode may have changed, as Windows' Settings does."""
    result = wintypes.DWORD()
    ctypes.windll.user32.SendMessageTimeoutW(
        0xFFFF, 0x001A, 0, "ImmersiveColorSet", 0x0002, 5000, ctypes.byref(result)
    )  # HWND_BROADCAST, WM_SETTINGCHANGE, SMTO_ABORTIFHUNG.


# Processes.

SYNCHRONIZE = 0x00100000
PROCESS_TERMINATE = 0x0001
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
WAIT_TIMEOUT = 0x102


def open_process(pid):
    """A handle to a running process, or None. While it is held, Windows gives the process's
    id to no other process, so ending it through the handle cannot end another."""
    kernel32 = ctypes.windll.kernel32
    kernel32.OpenProcess.restype = wintypes.HANDLE
    return kernel32.OpenProcess(SYNCHRONIZE | PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION, False, pid) or None


def created(handle):
    """When the process started, in FILETIME units, or None. With its pid, it names one
    process: Windows may give a pid to a new process once the old one has ended."""
    kernel32 = ctypes.windll.kernel32
    kernel32.GetProcessTimes.argtypes = [wintypes.HANDLE, *[ctypes.POINTER(wintypes.FILETIME)] * 4]
    times = [wintypes.FILETIME() for _ in range(4)]
    if not kernel32.GetProcessTimes(handle, *[ctypes.byref(t) for t in times]):
        return None
    return (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime


def end_process(handle, wait_seconds):
    """Waits for a process to end, then ends it, and closes the handle; True if it had to be
    ended."""
    kernel32 = ctypes.windll.kernel32
    kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    try:
        if kernel32.WaitForSingleObject(handle, int(wait_seconds * 1000)) != WAIT_TIMEOUT:
            return False
        kernel32.TerminateProcess(handle, 1)
        kernel32.WaitForSingleObject(handle, 10_000)
        return True
    finally:
        kernel32.CloseHandle(handle)


def excel_running():
    out = subprocess.run(
        ["tasklist", "/fi", "imagename eq excel.exe", "/fo", "csv", "/nh"], capture_output=True, text=True
    ).stdout
    return "excel.exe" in out.lower()


# Before starting.


def screen_unlocked():
    """True at an unlocked desktop: a locked screen, or a session with no desktop (SSH, a
    service), cannot open the desktop that receives input."""
    user32 = ctypes.windll.user32
    user32.OpenInputDesktop.restype = wintypes.HANDLE
    desktop = user32.OpenInputDesktop(0, False, 0x0100)  # DESKTOP_SWITCHDESKTOP
    if not desktop:
        return False
    try:
        return bool(user32.SwitchDesktop(desktop))
    finally:
        user32.CloseDesktop(desktop)


def user_environment():
    """The variables Excel will get: the user's own over the machine's, and this process's,
    which a person may have set in this window."""
    found = {}
    for root, path in (
        (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
        (winreg.HKEY_CURRENT_USER, "Environment"),
    ):
        for name, (value, _) in (read_key(path, root) or {}).items():
            found[name.upper()] = value
    return found


def refusals(xll):
    problems = []
    if sys.platform != "win32":
        return ["it runs only on Windows"]
    if not xll.is_file() or not (xll.parent / "pyinex_core.dll").is_file():
        problems.append(f"{xll} or pyinex_core.dll beside it not found")
    if not screen_unlocked():
        problems.append("the screen is locked, or this is not the signed-in desktop session (SSH, a service): Excel needs the desktop")
    environment = user_environment()
    for name in REFUSED_VARIABLES:
        if os.environ.get(name) or environment.get(name):
            problems.append(f"{name} is set in the environment Excel would get; it would choose or steer Excel's Python")
    if excel_running():
        problems.append("Excel is running: close it first (the driver neither ends an Excel it did not start nor can put settings back around one)")
    for module in ("pytest", "win32com"):
        if subprocess.run([sys.executable, "-c", f"import {module}"], capture_output=True).returncode:
            problems.append(f"{module} is not installed for {sys.executable}")
    return problems


# One run.


def probe(command):
    done = subprocess.run([*command, "-c", PROBE], capture_output=True, text=True)
    if done.returncode:
        return None
    return json.loads(done.stdout)


def make_venv(temp, log, name):
    """A virtual environment made from the newest Python, with NumPy, in the folder name;
    its probe or None."""
    folder = temp / "venvs" / name  # Not temp / <run name>, the run's own folder.
    for step in (
        ["py", f"-{PYTHONS[-1]}", "-m", "venv", str(folder)],
        [str(folder / "Scripts" / "python.exe"), "-m", "pip", "install", "--quiet", "--no-cache-dir", "--disable-pip-version-check", "numpy"],
    ):
        done = subprocess.run(step, capture_output=True, text=True)
        if done.returncode:
            log(f"could not make the virtual environment: {' '.join(step)}\n{done.stdout}{done.stderr}")
            return None
    return probe([str(folder / "Scripts" / "python.exe")])


def junit_counts(path):
    counts = {"tests": 0, "failures": 0, "errors": 0, "skipped": 0}
    try:
        root = ElementTree.parse(path).getroot()
    except (OSError, ElementTree.ParseError):
        return None
    for suite in [root] if root.tag == "testsuite" else root.iter("testsuite"):
        for k in counts:
            counts[k] += int(suite.get(k, 0))
    return counts


def excel_report(folder):
    """What the session wrote of its Excel (pid, version, build), or {} before it has."""
    try:
        return json.loads((folder / "excel.json").read_text("utf-8"))
    except (OSError, ValueError):
        return {}


def excel_process(path):
    """A handle to the Excel a report names, or None: only while its pid is still that
    Excel, the process started when the session says, so once pytest has exited and Excel
    with it, another process given the same pid is never taken for it."""
    try:
        report = json.loads(path.read_text("utf-8"))
    except (OSError, ValueError):
        return None
    handle = open_process(report["pid"]) if "pid" in report else None
    if handle and (report.get("created") is None or created(handle) != report["created"]):
        ctypes.windll.kernel32.CloseHandle(wintypes.HANDLE(handle))
        return None
    return handle


def take_excels(folder, excels):
    """Adds to excels, by report, a handle to each Excel the session has reported and not yet
    taken: its own (excel.json) and any other a test started ("excel <name>.json"). A report
    whose Excel was ended maps to None, so it is not taken again."""
    for path in folder.glob("excel*.json"):
        if path.name not in excels and (handle := excel_process(path)):
            excels[path.name] = handle


def run_one(name, python, args, xll, temp, out, log):
    """Runs the tests once in a fresh Excel embedding one Python, or, for a refused run,
    refusing one; returns its result."""
    result = {"run": name, "python": python and python["version"], "executable": python and python["executable"]}
    if python is None:
        result["problems"] = ["this Python could not be found or run"]
        return result
    problems = []
    if not python["numpy"]:
        problems.append(f"NumPy is not installed for {python['executable']}")
    folder = temp / name
    addin = folder / "addin"
    addin.mkdir(parents=True)
    for source in (xll, xll.parent / "pyinex_core.dll"):
        copy = addin / source.name
        shutil.copyfile(source, copy)
        try:
            os.remove(f"{copy}:Zone.Identifier")  # What Unblock-File does.
        except FileNotFoundError:
            pass
    kind = OWN_EXCEL.get(name, "normal")
    executable = folder / "no Python here" / "python.exe" if kind == "refused" else python["executable"]
    # The Unicode run names its Python relative to the add-in, as pyinex.ini allows.
    named = os.path.relpath(executable, addin) if name == UNICODE_VENV else executable
    (addin / "pyinex.ini").write_text(f"[python]\nexecutable = {named}\n", encoding="utf-8")
    expected = {"executable": str(executable), "version": python["version"], "venv": name in VENV_FOLDERS, "kind": kind}
    if kind == "first_call":
        expected["first"] = name.removeprefix("first-")
    if kind == "refused":
        expected["retry"] = python["executable"]  # The Python its test then names, and reopens.
    (folder / "expected.json").write_text(json.dumps(expected), "utf-8")
    junit = out / f"{name}.xml"
    command = [
        sys.executable, "-m", "pytest", *args.tests,
        "--excel", str(addin / xll.name), "--excel-run", str(folder),
        "--rootdir", str(ROOT), "--junitxml", str(junit), "--basetemp", str(folder / "pytest"),
        "-p", "no:cacheprovider", "-p", "no:faulthandler", "-v", "-rA",
    ]  # fmt: skip
    # No faulthandler: it prints a whole stack for each failed call to an Excel the
    # driver has ended, which buries the result.
    # No PYTEST_ variable reaches pytest: PYTEST_ADDOPTS=-k ... would run a subset that passes.
    environment = {k: v for k, v in os.environ.items() if not k.upper().startswith("PYTEST_")}
    environment.update(PYTHONDONTWRITEBYTECODE="1", PYTHONIOENCODING="utf-8")
    recorded = record(SETTINGS)
    log(f"\n=== {name}: Python {python['version'].split()[0]} ({python['executable']})")
    started = time.monotonic()
    process = subprocess.Popen(
        command, cwd=ROOT, env=environment, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace"
    )
    reader = threading.Thread(target=lambda: [log(line.rstrip("\n")) for line in process.stdout], daemon=True)
    reader.start()
    interrupted = timed_out = forced = False
    excels = {}  # Handles to the session's Excels, each taken as soon as it is reported.
    try:
        while process.poll() is None:
            take_excels(folder, excels)
            if time.monotonic() - started > args.timeout:
                timed_out = True
                log(f"{name}: no result after {args.timeout} seconds; ending its Excel")
                break
            time.sleep(0.5)
    except KeyboardInterrupt:
        interrupted = True
    previous = signal.signal(signal.SIGINT, signal.SIG_IGN)  # A second Ctrl+C must not stop the cleanup.
    try:
        if interrupted:
            log(f"{name}: stopped (Ctrl+C); ending its Excel and putting settings back")
        if timed_out or interrupted:
            deadline = time.monotonic() + 60
            while process.poll() is None and time.monotonic() < deadline:
                take_excels(folder, excels)
                for report, excel in excels.items():
                    if excel:
                        forced |= end_process(excel, 0)
                        excels[report] = None
                time.sleep(0.5)
            if process.poll() is None:
                process.kill()
                log(f"{name}: pytest did not finish; ended it")
        process.wait()
        reader.join(10)
        # A short run can start and quit its Excel between two looks, and Ctrl+C can come
        # before the first: take each now if it is still running.
        take_excels(folder, excels)
        # Excel quits once pytest, which held it, has exited. It sometimes takes a minute,
        # with no add-in loaded as with one (seen in testing), so each gets 90 seconds.
        for excel in excels.values():
            if excel and end_process(excel, 90):
                forced = True
        if excel_running():
            problems.append("an Excel is still running that the driver could not end: close it, then check the settings, which it may change as it quits; no further runs")
            result["excel_left"] = True
        # A test may leave a file to appear only if something ran as Excel exited, which
        # nothing in the add-in may (tests/excel/test_lifetime.py).
        for ran in sorted(folder.glob("after-exit *")):
            problems.append(f"at Excel's exit: {ran.name.removeprefix('after-exit ')}")
        problems += restore(recorded)
    finally:
        signal.signal(signal.SIGINT, previous)
    report = excel_report(folder)
    counts = junit_counts(junit)
    if counts is None:
        problems.append("pytest wrote no results")
        counts = {"tests": 0, "failures": 0, "errors": 0, "skipped": 0}
    result.update(counts)
    result.update(excel=report.get("version"), excel_build=report.get("build"), seconds=round(time.monotonic() - started))
    if process.returncode != 0:
        problems.append(f"pytest exited {process.returncode}")
    if counts["tests"] == 0:
        problems.append("no tests ran")
    if counts["skipped"]:
        problems.append(f"{counts['skipped']} skipped: a run that did not test everything is not a pass")
    if timed_out:
        problems.append(f"reached its timeout ({args.timeout} seconds)")
    if forced:
        problems.append("Excel had to be ended")
    # What the runs left would meet too, whatever their Python (the module's docstring).
    result["stops"] = forced or timed_out or counts["tests"] == 0
    if interrupted:
        problems.append("stopped with Ctrl+C")
        result["interrupted"] = True
    result["problems"] = problems
    return result


# The whole run.


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def windows_version():
    info = read_key(r"SOFTWARE\Microsoft\Windows NT\CurrentVersion", winreg.HKEY_LOCAL_MACHINE) or {}
    machine = (read_key(r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment", winreg.HKEY_LOCAL_MACHINE) or {}).get(
        "PROCESSOR_ARCHITECTURE", ("?", 0)
    )[0]
    build = info.get("CurrentBuild", ("?", 0))[0]
    name = "Windows 11" if build.isdigit() and int(build) >= 22000 else info.get("ProductName", ("Windows", 0))[0]
    return f"{name} {info.get('DisplayVersion', ('', 0))[0]} (build {build}.{info.get('UBR', ('?', 0))[0]}), {machine}"


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("xll", type=Path, help="path to pyinex.xll, with pyinex_core.dll beside it")
    parser.add_argument("--out", type=Path, help="folder for the results (default: a new one under %%TEMP%%)")
    parser.add_argument("--python", action="append", choices=RUNS, metavar="RUN", help=f"run only these (repeatable): {', '.join(RUNS)}")
    parser.add_argument("--tests", nargs="+", default=["tests/excel"], help="what pytest runs (default: tests/excel)")
    parser.add_argument("--timeout", type=int, default=3600, help="seconds each run may take (default: 3600)")
    parser.add_argument("--keep-going", action="store_true",
                        help="run the rest after a run whose Excel had to be ended, that reached its timeout, or in which no tests ran")
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    xll = args.xll.resolve()
    problems = refusals(xll)
    if problems:
        for p in problems:
            print(f"refusing to start: {p}", file=sys.stderr)
        return 2
    out = (args.out or Path(tempfile.mkdtemp(prefix="pyinex-excel-results-"))).resolve()
    out.mkdir(parents=True, exist_ok=True)
    log = Log(out / "log.txt")
    temp = Path(tempfile.mkdtemp(prefix="pyinex-excel-"))
    summary = {
        "started": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "pyinex.xll": sha256(xll),
        "pyinex_core.dll": sha256(xll.parent / "pyinex_core.dll"),
        "windows": windows_version(),
        "runs": [],
    }
    log(f"Pyinex real-Excel tests: {xll}\n  pyinex.xll      {summary['pyinex.xll']}\n  pyinex_core.dll {summary['pyinex_core.dll']}")
    log(f"  {summary['windows']}; results in {out}")
    code = 0
    try:
        for name in args.python or RUNS:
            if name in VENV_FOLDERS:
                python = make_venv(temp, log, VENV_FOLDERS[name])
            else:
                python = probe(["py", f"-{name if name in PYTHONS else PYTHONS[-1]}"])
            result = run_one(name, python, args, xll, temp, out, log)
            summary["runs"].append(result)
            if result["problems"]:
                code = 1
            if result.get("interrupted"):
                raise KeyboardInterrupt
            if result.get("excel_left"):
                break
            if result.get("stops") and not args.keep_going:
                summary["stopped"] = name
                log(f"\n{name}: stopping here, since the runs left would almost always fail the same way; --keep-going runs them")
                break
    except KeyboardInterrupt:
        summary["stopped"] = "Ctrl+C"
        code = 130
    finally:
        summary["excel"] = next((f"{r['excel']} (build {r['excel_build']})" for r in summary["runs"] if r.get("excel")), None)
        (out / "summary.json").write_text(json.dumps(summary, indent=2), "utf-8")
        log("\n=== Results")
        for r in summary["runs"]:
            tests = r.get("tests", 0)
            counts = f"{tests} test{'' if tests == 1 else 's'}, {r.get('failures', 0)} failed, {r.get('errors', 0)} errors, {r.get('skipped', 0)} skipped"
            log(f"{'ok  ' if not r['problems'] else 'FAIL'}  {r['run']:<5} {counts}" + "".join(f"\n        {p}" for p in r["problems"]))
        if code == 130:
            log("stopped with Ctrl+C before every run finished: not a pass")
        elif summary.get("stopped"):
            log(f"stopped after {summary['stopped']}, before every run finished: not a pass")
        elif code == 0 and len(summary["runs"]) == len(args.python or RUNS):
            log("every run passed")
        shutil.rmtree(temp, ignore_errors=True)
        log(f"results: {out}")
        log(f"exit code {code}")
    return code


if __name__ == "__main__":
    sys.exit(main())
