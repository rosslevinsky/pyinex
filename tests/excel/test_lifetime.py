"""The add-in's life in real Excel (docs/architecture.md, "Lifetime"): unloaded with Python
running, and loaded again. It runs in an Excel of its own, the driver's lifetime run, since
unloading the add-in changes the whole Excel.

The add-in is unloaded as Excel's UNREGISTER does it when given the add-in's path, which
calls xlAutoClose as unticking it in the Add-ins dialog does, and loaded again with
RegisterXLL, which calls xlAutoOpen. (The dialog's own route, Application.AddIns2, will
not take an add-in loaded with RegisterXLL.)
"""

from pathlib import Path

import pytest
from conftest import EXCEL_PID, Console, E, modules, python_dlls, shown


FLUSHES = """import atexit
import sys

import pyinex

LOG = {log!r}


class Recorded:
    # sys.stdout, recording each flush and whether pyinex.caller() worked then.
    def __init__(self, stream):
        self.stream = stream

    def write(self, text):
        return self.stream.write(text)

    def flush(self):
        # Outside a call, all three must refuse; in one, caller answers.
        refused = []
        for ask in (pyinex.caller, pyinex.sheet, pyinex.esc_pressed):
            try:
                ask()
            except RuntimeError:
                refused.append(ask.__name__)
        if len(refused) == 3:
            seen = 'outside a call: RuntimeError'
        elif not refused:
            seen = 'in a call'
        else:
            seen = 'only some refused: ' + ', '.join(refused)
        with open(LOG, 'a', encoding='utf-8') as f:
            f.write(seen + '\\n')
        self.stream.flush()


class Marks:
    # Writes the file named if Python finalizes it, which Pyinex never does.
    def __init__(self, path):
        self.path = path

    def __del__(self):
        open(self.path, 'w').close()


n = 0


def count():
    global n
    if not isinstance(sys.stdout, Recorded):
        sys.stdout = Recorded(sys.stdout)
        # Neither may run at Excel's exit: Pyinex never finalizes Python. The driver looks
        # for the files once Excel has ended.
        atexit.register(lambda: open({atexit_ran!r}, 'w').close())
        sys.marks = Marks({finalizer_ran!r})
    n += 1
    return n
"""


@pytest.mark.kind("lifetime")
def test_before_python_starts_unloading_takes_the_addin_out(excel, pytestconfig):
    # First in the lifetime run: nothing has started Python or made the console, so
    # nothing has pinned the add-in, and Windows unloads it.
    pid = pytestconfig.stash[EXCEL_PID]
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    assert python_dlls(pid) == []
    assert modules(pid, r"pyinex\.xll")
    assert excel.ExecuteExcel4Macro(f'UNREGISTER("{xll}")') is True
    assert modules(pid, r"pyinex\.xll") == [], "the add-in is still loaded"
    assert excel.RegisterXLL(str(xll))
    assert modules(pid, r"pyinex\.xll")


@pytest.mark.kind("lifetime")
def test_the_console_alone_keeps_the_addin_loaded(excel, pytestconfig):
    # Second: the console window, made before Python starts, holds the add-in's code, so
    # unloading leaves the add-in in Excel, and the window still Pyinex's.
    pid = pytestconfig.stash[EXCEL_PID]
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    assert python_dlls(pid) == []
    assert excel.Evaluate("PyConsole(FALSE)") is False
    assert excel.ExecuteExcel4Macro(f'UNREGISTER("{xll}")') is True
    assert modules(pid, r"pyinex\.xll"), "the add-in was unloaded under its console window"
    assert Console.find(pid) is not None
    assert excel.RegisterXLL(str(xll))
    assert excel.Evaluate("PyConsole(TRUE)") is True and Console.find(pid).shown
    assert excel.Evaluate("PyConsole(FALSE)") is False
    assert python_dlls(pid) == []


@pytest.mark.kind("lifetime")
def test_unloading_and_loading_again_keeps_python_and_its_scripts(excel, pytestconfig, tmp_path, desktop, run_folder):
    pid = pytestconfig.stash[EXCEL_PID]
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    log = tmp_path / "flushes.txt"
    script = tmp_path / "kept.py"
    after_exit = run_folder if run_folder else tmp_path
    script.write_text(
        FLUSHES.format(
            log=str(log),
            atexit_ran=str(after_exit / "after-exit atexit ran"),
            finalizer_ran=str(after_exit / "after-exit finalizer ran"),
        ),
        encoding="utf-8",
    )
    count = f'PyCall("{script}", "count")'
    assert excel.Evaluate(count) == 1.0
    loaded = python_dlls(pid)
    assert loaded
    log.unlink(missing_ok=True)

    assert excel.ExecuteExcel4Macro(f'UNREGISTER("{xll}")') is True
    # xlAutoClose flushed Python's output, outside any call, where pyinex refuses.
    assert log.read_text("utf-8").splitlines()[-1] == "outside a call: RuntimeError"
    assert isinstance(shown(excel.Evaluate(count)), E), "PyCall still works"
    assert isinstance(shown(excel.Evaluate('PyInfo("status")')), E)
    assert python_dlls(pid) == loaded, "Python is not still running"
    assert Console.find(pid) is not None, "the console window went with the add-in"
    assert modules(pid, r"pyinex\.xll"), "the add-in was unloaded, though Python had started"

    assert excel.RegisterXLL(str(xll))  # xlAutoOpen again: the same interpreter and scripts.
    assert excel.Evaluate('PyInfo("status")') == "running"
    # 2: the module was kept, and no call ran while the add-in was unloaded.
    assert excel.Evaluate(count) == 2.0, "the script's module state did not survive"
    assert python_dlls(pid) == loaded
    # The console window is still Pyinex's own: placing it applies as it is shown.
    left, top, _, _ = desktop
    assert excel.Evaluate(f"PyConsole(TRUE, FALSE, {left + 40}, {top + 50}, 400, 300)") is True
    assert Console.find(pid).rect == (left + 40, top + 50, 400, 300)
    assert excel.Evaluate("PyConsole(FALSE)") is False
