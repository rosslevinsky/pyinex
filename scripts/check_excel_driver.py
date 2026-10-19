"""Checks the real-Excel driver (scripts/excel_tests.py) fails, refuses and cleans up as it should.

Run at the signed-in, unlocked Windows desktop, with Excel closed, like the driver:

    py -3.14 scripts\\check_excel_driver.py C:\\path\\to\\pyinex.xll

It runs the driver, under the newest Python only, on each deliberately bad test in
tests/excel/driver_checks/: one that fails, one that skips, one whose PyCall never
returns, one that leaves a second Excel of its own running, and one that changes Windows'
app mode and then makes such a call, the hanging ones with a short timeout; and that last
one again, stopped with Ctrl+C. The one that fails and the one that leaves an Excel also
run with a second run asked for: the driver carries on after failed tests, stops after an
Excel it had to end, and with --keep-going carries on after that too. Then it starts the
driver with PYTHONPATH set, and with an Excel open, both of which it must refuse. After
each it checks the driver's exit code and what it reported, that no Excel is left running,
and that the settings the driver guards are as they were. It asks for no input, and exits
0 only if every check passed.
"""

import argparse
import ctypes
import gc
import json
import os
import signal
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path

sys.dont_write_bytecode = True  # Nothing compiled is written into scripts/.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import excel_tests as driver  # noqa: E402

CHECKS = Path("tests") / "excel" / "driver_checks"


def mode():
    return (driver.read_key(driver.PERSONALIZE) or {}).get("AppsUseLightTheme")


def run_driver(xll, out, *extra, environment=None, stop_when=None):
    """Runs the driver, printing what it prints; returns its exit code, its output and its
    summary.json. stop_when, if given, is polled; once it is true, Ctrl+C is pressed."""
    command = [sys.executable, str(driver.ROOT / "scripts" / "excel_tests.py"), str(xll), "--python", driver.PYTHONS[-1], "--out", str(out), *extra]
    process = subprocess.Popen(
        command, cwd=driver.ROOT, env=environment, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace"
    )
    lines = []
    reader = threading.Thread(target=lambda: [lines.append(line) or print(line, end="", flush=True) for line in process.stdout], daemon=True)
    reader.start()
    if stop_when is not None:
        deadline = time.monotonic() + 300
        while process.poll() is None and not stop_when() and time.monotonic() < deadline:
            time.sleep(0.5)
        if process.poll() is None:
            time.sleep(5)  # Into the call that never returns.
            press_ctrl_c()
    process.wait()
    reader.join(10)
    try:
        summary = json.loads((out / "summary.json").read_text("utf-8"))
    except (OSError, ValueError):
        summary = None
    return process.returncode, "".join(lines), summary


def press_ctrl_c():
    """What pressing Ctrl+C in this window does: every process on the console gets it.
    This one ignores it for the moment it takes to arrive."""
    previous = signal.signal(signal.SIGINT, signal.SIG_IGN)
    try:
        ctypes.windll.kernel32.GenerateConsoleCtrlEvent(0, 0)  # CTRL_C_EVENT, this console.
        time.sleep(1)
    finally:
        signal.signal(signal.SIGINT, previous)


def problems_of(summary):
    return [p for r in (summary or {}).get("runs", []) for p in r["problems"]]


def open_excel():
    """An Excel this script started, and a handle to its process."""
    import win32process
    from win32com.client import DispatchEx

    app = DispatchEx("Excel.Application")
    return app, driver.open_process(win32process.GetWindowThreadProcessId(app.Hwnd)[1])


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("xll", type=Path, help="path to pyinex.xll, with pyinex_core.dll beside it")
    args = parser.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    xll = args.xll.resolve()
    refusals = driver.refusals(xll)
    if refusals:
        for p in refusals:
            print(f"refusing to start: {p}", file=sys.stderr)
        return 2
    results = Path(tempfile.mkdtemp(prefix="pyinex-driver-checks-"))
    outcomes = []

    def check(name, run, expected_code, expected_text, runs=None):
        settings, light = driver.record(driver.SETTINGS), mode()
        print(f"\n##### {name}", flush=True)
        code, output, summary = run(results / name)
        wrong = []
        if code != expected_code:
            wrong.append(f"exit code {code}, expected {expected_code}")
        if expected_text not in output + "\n".join(problems_of(summary)):
            wrong.append(f"did not report {expected_text!r}")
        if expected_code != 2 and summary is None:
            wrong.append("wrote no summary.json")
        if runs is not None and summary is not None and len(summary["runs"]) != runs:
            wrong.append(f"{len(summary['runs'])} runs, expected {runs}")
        if driver.excel_running():
            wrong.append("an Excel is still running")
        if mode() != light:
            wrong.append("the app mode is not as it was")
        if driver.record(driver.SETTINGS) != settings:
            wrong.append("a setting the driver guards is not as it was")
        outcomes.append((name, code, wrong))

    def tests(file, *extra, **options):
        return lambda out: run_driver(xll, out, "--tests", str(CHECKS / file), *extra, **options)

    check("fails", tests("test_fails.py"), 1, "pytest exited 1")
    check("skips", tests("test_skips.py"), 1, "skipped: a run that did not test everything")
    check("hangs", tests("test_hangs.py", "--timeout", "60"), 1, "reached its timeout")
    check("second Excel left", tests("test_leaves_a_second_excel.py"), 1, "Excel had to be ended")
    # Of two runs asked for: failed tests run the second; an Excel ended stops before it,
    # unless --keep-going.
    check("fails, two runs", tests("test_fails.py", "--python", driver.VENV), 1, "pytest exited 1", runs=2)
    check("second Excel left, two runs", tests("test_leaves_a_second_excel.py", "--python", driver.VENV), 1,
          "stopped after", runs=1)
    check("second Excel left, keep going", tests("test_leaves_a_second_excel.py", "--python", driver.VENV, "--keep-going"),
          1, "Excel had to be ended", runs=2)
    check("changes mode, timeout", tests("test_changes_mode_then_hangs.py", "--timeout", "60"), 1, "reached its timeout")
    light = mode()
    check("changes mode, Ctrl+C", tests("test_changes_mode_then_hangs.py", "--timeout", "600", stop_when=lambda: mode() != light), 130, "stopped with Ctrl+C")
    environment = dict(os.environ, PYTHONPATH=r"C:\nowhere")
    check("PYTHONPATH set", lambda out: run_driver(xll, out, environment=environment), 2, "PYTHONPATH is set")

    def with_excel_open(out):
        recorded = driver.record(driver.SETTINGS)
        app, process = open_excel()
        try:
            return run_driver(xll, out)
        finally:
            app.Quit()
            app = None
            gc.collect()
            if process:
                driver.end_process(process, 30)
            driver.restore(recorded)  # Whatever this Excel wrote of its own.

    check("Excel open", with_excel_open, 2, "Excel is running")

    print("\n=== Driver checks")
    for name, code, wrong in outcomes:
        print(f"{'ok  ' if not wrong else 'FAIL'}  {name:<22} exit {code}" + "".join(f"\n        {w}" for w in wrong))
    failed = any(wrong for _, _, wrong in outcomes)
    print("the driver failed, refused and cleaned up as it should" if not failed else "the driver did not behave as it should")
    print(f"results: {results}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
