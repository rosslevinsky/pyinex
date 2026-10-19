"""The add-in's life in real Excel (docs/architecture.md, "Lifetime"): unloaded with Python
running, and loaded again; unloaded with direct functions queued to come back; and a quit
cancelled at the save prompt, with direct functions registered and with none, after which a
workbook opened still loads its scripts. It runs in an Excel of its own, the driver's lifetime run, since
unloading the add-in changes the whole Excel.

The add-in is unloaded as Excel's UNREGISTER does it when given the add-in's path, which
calls xlAutoClose as unticking it in the Add-ins dialog does, and loaded again with
RegisterXLL, which calls xlAutoOpen. (The dialog's own route, Application.AddIns2, will
not take an add-in loaded with RegisterXLL.)
"""

import time
from pathlib import Path

import pytest
from conftest import EXCEL_PID, Console, E, bring_to_front, modules, python_dlls, shown


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


# Direct functions across closing (docs/architecture.md, "Lifetime" and "The queue").

LOADED = "import pyinex\n\n@pyinex.func\ndef {0}(x):\n    return '{0}: ' + str(x)\n"


def console_text(pid):
    window = Console.find(pid)
    return window.text if window else ""


def wait_until(check, timeout=20.0):
    """check()'s value once it is true, or its last value after timeout."""
    deadline = time.monotonic() + timeout
    while not (value := check()) and time.monotonic() < deadline:
        time.sleep(0.1)
    return value


def evaluated(excel, formula):
    """A check giving formula's value, or None while it is an error."""

    def check():
        value = shown(excel.Evaluate(formula))
        return None if isinstance(value, E) else value

    return check


def slots(excel, xll):
    return sorted(p for dll, p, _ in excel.RegisteredFunctions if Path(dll).resolve() == xll and p.startswith("PyFunc"))


@pytest.mark.kind("lifetime")
def test_unloaded_with_work_queued_and_loaded_again(excel, pytestconfig, tmp_path):
    # Closing queues each direct function to come back, in case Excel carries on; here
    # Excel unloads the add-in, which the queue sees, and loading it again starts clean.
    pid = pytestconfig.stash[EXCEL_PID]
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    workbook = excel.Workbooks.Add()
    try:
        workbook.SaveAs(str(tmp_path / "unloaded.xlsx"), FileFormat=51)
        (tmp_path / "life_unloaded.py").write_text(LOADED.format("LifeUnloaded"), "utf-8")
        cell = workbook.Worksheets(1).Range("A1")
        cell.Formula2 = '=PyLoad("life_unloaded.py")'
        assert cell.Value2 == "1 function"
        assert wait_until(evaluated(excel, "LifeUnloaded(1)")) == "LifeUnloaded: 1.0"
        assert excel.ExecuteExcel4Macro(f'UNREGISTER("{xll}")') is True
        said = "Pyinex: the add-in was unloaded, so 1 function will not be registered again"
        assert said in wait_until(lambda: said in console_text(pid) and console_text(pid))
        assert isinstance(shown(excel.Evaluate("LifeUnloaded(1)")), E)
        # Loading the add-in calculates the PyLoad cells of every workbook open then
        # (docs/architecture.md, "When a workbook opens"), so the cell is emptied first: what
        # is checked here is that nothing queued from before comes back.
        formula = cell.Formula2
        cell.ClearContents()
        assert excel.RegisterXLL(str(xll))
        time.sleep(1)  # Longer than the queue would take: nothing from before comes back.
        assert slots(excel, xll) == []
        assert isinstance(shown(excel.Evaluate("LifeUnloaded(1)")), E)
        assert console_text(pid).count("registered LifeUnloaded") == 1
        # The PyLoad cell's next calculation registers it again, in the first slot.
        cell.Formula2 = formula
        assert wait_until(evaluated(excel, "LifeUnloaded(2)")) == "LifeUnloaded: 2.0"
        assert slots(excel, xll) == ["PyFunc0"]
    finally:
        workbook.Close(SaveChanges=False)


def retried(call, timeout=30.0):
    """call(), tried again while Excel rejects COM calls, as it does while it is busy: as
    RPC_E_CALL_REJECTED, or, while its prompt to save is still closing, as 0x800AC472."""
    import pywintypes

    deadline = time.monotonic() + timeout
    while True:
        try:
            return call()
        except pywintypes.com_error as e:
            if e.hresult not in (-2147418111, -2146777998) or time.monotonic() > deadline:
                raise
            time.sleep(0.2)


def find_save_prompt(pid, timeout=20.0):
    """The window of Excel's prompt to save changes as it quits, or None."""
    import win32gui
    import win32process

    def look():
        found = []

        def visit(hwnd, _):
            if (win32process.GetWindowThreadProcessId(hwnd)[1] == pid and win32gui.IsWindowVisible(hwnd)
                    and win32gui.GetClassName(hwnd) in ("NUIDialog", "#32770")):
                found.append(hwnd)

        win32gui.EnumWindows(visit, None)
        return found[0] if found else None

    return wait_until(look, timeout)


def cancel_quit(excel, pid):
    """Quits Excel, which has a changed workbook, and cancels at its prompt to save: Excel
    calls xlAutoClose before the prompt, and then carries on."""
    import threading

    import pythoncom
    import win32com.client
    from test_direct_functions import VK_ESCAPE, send

    excel.DisplayAlerts = True
    stream = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, excel._oleobj_)
    outcome = {}

    def quit_excel():
        pythoncom.CoInitialize()
        other = None
        try:
            other = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(stream, pythoncom.IID_IDispatch))
            other.Quit()
            outcome["returned"] = True
        except pythoncom.com_error as e:
            outcome["error"] = e
        finally:
            other = None
            pythoncom.CoUninitialize()

    quitting = threading.Thread(target=quit_excel, daemon=True)
    quitting.start()
    prompt = find_save_prompt(pid)
    assert prompt, "Excel did not ask whether to save"
    assert bring_to_front(prompt), "the prompt could not be made the foreground window"
    send((VK_ESCAPE, False), (VK_ESCAPE, True))  # Cancel.
    quitting.join(30)
    assert not quitting.is_alive(), "Quit did not return after the prompt was cancelled"
    assert outcome == {"returned": True}, outcome
    retried(lambda: setattr(excel, "DisplayAlerts", False))


@pytest.mark.kind("lifetime")
def test_a_quit_cancelled_at_the_save_prompt_brings_the_functions_back(excel, pytestconfig, tmp_path):
    # Excel calls xlAutoClose before it asks whether to save; cancelled there, it carries on
    # with the add-in loaded, and the queue registers the direct functions again.
    pid = pytestconfig.stash[EXCEL_PID]
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    workbook = excel.Workbooks.Add()
    try:
        workbook.SaveAs(str(tmp_path / "cancelled.xlsx"), FileFormat=51)
        (tmp_path / "life_cancelled.py").write_text(LOADED.format("LifeAlpha") + LOADED.format("LifeBeta"), "utf-8")
        sheet = workbook.Worksheets(1)
        sheet.Range("A1").Formula2 = '=PyLoad("life_cancelled.py")'
        assert wait_until(evaluated(excel, "LifeBeta(1)")) == "LifeBeta: 1.0"
        before = slots(excel, xll)
        sheet.Range("B1").Value2 = "changed"  # So Excel asks whether to save.
        cancel_quit(excel, pid)
        # The functions come back without the add-in being opened again, each in its own slot.
        assert wait_until(lambda: console_text(pid).count("registered LifeBeta") == 2), console_text(pid)
        assert wait_until(evaluated(excel, "LifeAlpha(2)")) == "LifeAlpha: 2.0"
        assert wait_until(evaluated(excel, "LifeBeta(3)")) == "LifeBeta: 3.0"
        assert slots(excel, xll) == before
        assert sheet.Range("A1").Value2 == "2 functions"
    finally:
        retried(lambda: setattr(excel, "DisplayAlerts", False))
        retried(lambda: workbook.Close(SaveChanges=False))


@pytest.mark.kind("lifetime")
def test_after_a_cancelled_quit_with_nothing_registered_a_workbook_opened_loads(excel, pytestconfig, tmp_path):
    # With no direct function to register again, a cancelled quit still leaves Pyinex
    # following workbooks as they open (docs/architecture.md, "Lifetime").
    pid = pytestconfig.stash[EXCEL_PID]
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    # Unloaded and loaded again, so that no direct function from before is registered.
    assert excel.ExecuteExcel4Macro(f'UNREGISTER("{xll}")') is True
    assert excel.RegisterXLL(str(xll))
    assert wait_until(lambda: slots(excel, xll) == []), slots(excel, xll)
    saved, changed, reopened = tmp_path / "reopened.xlsx", excel.Workbooks.Add(), None
    try:
        workbook = excel.Workbooks.Add()
        workbook.SaveAs(str(saved), FileFormat=51)
        workbook.Worksheets(1).Range("A1").Formula2 = '=PyLoad("life_reopened.py")'  # Not written yet: #VALUE!.
        workbook.Close(SaveChanges=True)
        marker = tmp_path / "life_reopened.loaded"
        (tmp_path / "life_reopened.py").write_text(
            f"import pathlib\n\npathlib.Path({str(marker)!r}).write_text('loaded')\n\n" + LOADED.format("LifeReopened"),
            "utf-8")
        changed.Worksheets(1).Range("A1").Value2 = "changed"  # So Excel asks whether to save.
        assert slots(excel, xll) == []
        cancel_quit(excel, pid)
        reopened = retried(lambda: excel.Workbooks.Open(str(saved)))
        assert wait_until(marker.exists), "the script did not load as the workbook opened"
        assert wait_until(evaluated(excel, "LifeReopened(1)")) == "LifeReopened: 1.0"
    finally:
        retried(lambda: setattr(excel, "DisplayAlerts", False))
        for book in (reopened, changed):
            if book is not None:
                retried(lambda: book.Close(SaveChanges=False))
