"""PyLoad and the queue in real Excel (docs/architecture.md, "PyLoad" and "The queue"): the
direct-functions test group 5 (paths, each way a script fails, the fixed file, a script
loaded twice and from two workbooks, a changed argument, one script through PyLoad and
PyCall), and group 7 but a workbook reopened in a new Excel: registration queued during a
calculation, a long cell edit and a modal dialog, two Excels, no workbook open, and
Protected View.

A name registered stays registered for the rest of the session's Excel, so every test's
functions have names of their own, starting "Pl".
"""

import gc
import shutil
import time
from pathlib import Path

import pytest
from conftest import EXCEL_PID, E, Console, report_excel, shown
from test_direct_functions import VK_CONTROL, VK_ESCAPE, chord, send

GOOD = "import pyinex\n\n@pyinex.func\ndef {0}(x):\n    return '{0}: ' + str(x)\n"
WAITING = "Pyinex: Excel cannot register functions yet"
VK_F2 = 0x71


def when_registered(app, formula, timeout=15.0, macro=False):
    """formula's value once it is no longer #NAME?, or #NAME? after timeout; with macro,
    calculated as Excel 4 macros are, which works where Evaluate gives #N/A (a Protected
    View window active)."""
    deadline = time.monotonic() + timeout
    while True:
        value = shown(app.ExecuteExcel4Macro(formula) if macro else app.Evaluate(formula))
        if value != E("#NAME?") or time.monotonic() > deadline:
            return value
        time.sleep(0.1)


def slots(app, xll):
    """The numbered exports Excel has registered for the add-in."""
    return sorted(p for dll, p, _ in app.RegisteredFunctions if Path(dll).resolve() == xll and p.startswith("PyFunc"))


@pytest.fixture
def xll(pytestconfig):
    return Path(pytestconfig.getoption("--excel")).resolve()


# Group 5: loading.


def test_a_relative_path_is_relative_to_the_workbooks_folder(app, book, cleared):
    book.script("pl_relative.py", GOOD.format("PlRelative"))
    assert book.formula("A1", '=PyLoad("pl_relative.py")') == "1 function"
    assert when_registered(app, "PlRelative(2)") == "PlRelative: 2.0"
    assert book.formula("B1", "=PlRelative(3)") == "PlRelative: 3.0"
    assert "Pyinex: registered PlRelative from the script pl_relative.py" in cleared(holding="registered PlRelative").text


def test_an_absolute_path(app, book, tmp_path):
    script = tmp_path / "elsewhere" / "pl_absolute.py"
    script.parent.mkdir()
    script.write_text(GOOD.format("PlAbsolute") + "\n" + GOOD.format("PlAbsoluteToo"), encoding="utf-8")
    assert book.formula("A1", f'=PyLoad("{script}")') == "2 functions"
    assert when_registered(app, "PlAbsoluteToo(1)") == "PlAbsoluteToo: 1.0"
    assert book.formula("B1", "=PlAbsolute(4)") == "PlAbsolute: 4.0"


@pytest.mark.parametrize(
    "text, said",
    [
        (None, "Pyinex: cannot find the script"),
        ("def f(:\n", "SyntaxError"),
        ("import pyinex\n\n@pyinex.func\ndef PlImportFails(x):\n    return x\n\nraise ValueError('at import')\n",
         "ValueError: at import"),
    ],
    ids=["missing", "syntax error", "exception at import"],
)
def test_a_script_that_fails_to_load_is_reported(app, book, cleared, text, said):
    if text is not None:
        book.script("pl_broken.py", text)
    assert book.formula("A1", '=PyLoad("pl_broken.py")') == E("#VALUE!")
    shown_text = cleared(holding=said).text
    assert said in shown_text and "pl_broken.py" in shown_text
    assert shown(app.Evaluate("PlImportFails(1)")) == E("#NAME?")


@pytest.mark.parametrize("statement, said", [("raise ValueError('top level')", "ValueError: top level"),
                                             ("raise SystemExit(5)", "SystemExit: 5"),
                                             ("raise KeyboardInterrupt", "KeyboardInterrupt")],
                         ids=["ValueError", "SystemExit", "KeyboardInterrupt"])
def test_a_script_that_raises_registers_nothing_until_it_is_fixed(app, book, cleared, statement, said):
    name = "PlLost" + said.split(":")[0]
    script = book.script("pl_lost.py", GOOD.format(name) + "\n" + statement + "\n")
    assert book.formula("A1", f'=PyLoad("{script}")') == E("#VALUE!")
    assert said in cleared(holding=said).text
    time.sleep(1)  # Longer than the queue would take.
    assert shown(app.Evaluate(f"{name}(1)")) == E("#NAME?")
    book.script("pl_lost.py", GOOD.format(name))
    assert book.formula("A1", f'=PyLoad("{script}")') == "1 function"  # Calculated again.
    assert when_registered(app, f"{name}(1)") == f"{name}: 1.0"


def test_a_script_loaded_twice_and_from_two_workbooks_registers_once(app, book, workbook, xll):
    book.script("pl_twice.py", GOOD.format("PlTwice"))
    before = slots(app, xll)
    assert book.formula("A1", '=PyLoad("pl_twice.py")') == "1 function"
    assert when_registered(app, "PlTwice(1)") == "PlTwice: 1.0"
    assert book.formula("A2", '=PyLoad("pl_twice.py")') == "1 function"
    other = workbook(book.folder / "other.xlsx")
    assert other.formula("A1", '=PyLoad("pl_twice.py")') == "1 function"
    time.sleep(1)
    assert len(slots(app, xll)) == len(before) + 1
    # Its functions work in every open workbook.
    assert other.formula("B1", "=PlTwice(2)") == "PlTwice: 2.0"
    assert book.formula("B1", "=PlTwice(3)") == "PlTwice: 3.0"


def test_pyload_calculates_again_when_its_argument_changes(app, book):
    book.script("pl_arg_a.py", GOOD.format("PlArgA"))
    book.script("pl_arg_b.py", GOOD.format("PlArgB") + "\n" + GOOD.format("PlArgC"))
    book.sheet.Range("C1").Value2 = "pl_arg_a.py"
    assert book.formula("A1", "=PyLoad(C1)") == "1 function"
    assert when_registered(app, "PlArgA(1)") == "PlArgA: 1.0"
    book.sheet.Range("C1").Value2 = "pl_arg_b.py"  # Only the argument changes.
    assert book.value("A1") == "2 functions"
    assert when_registered(app, "PlArgC(2)") == "PlArgC: 2.0"


def test_one_script_through_pyload_and_pycall_is_one_module(app, book):
    book.script("pl_counted.py", "import pyinex\n\nn = 0\n\n@pyinex.func\ndef PlCount():\n    global n\n"
                                 "    n += 1\n    return n\n")
    assert book.formula("A1", '=PyLoad("pl_counted.py")') == "1 function"
    first = when_registered(app, "PlCount()")
    assert book.formula("B1", '=PyCall("pl_counted.py", "PlCount")') == first + 1
    assert book.formula("C1", "=PlCount()") == first + 2


# Group 7: timing and Excel's states.


def test_registration_queued_during_a_full_recalculation(app, book):
    book.script("pl_calc.py", GOOD.format("PlCalcA"))
    assert book.formula("A1", '=PyLoad("pl_calc.py")') == "1 function"
    assert when_registered(app, "PlCalcA(1)") == "PlCalcA: 1.0"
    book.script("pl_calc.py", GOOD.format("PlCalcA") + "\n" + GOOD.format("PlCalcB"))
    app.CalculateFull()  # The PyLoad cell calculates again, and queues PlCalcB.
    assert book.value("A1") == "2 functions"
    assert when_registered(app, "PlCalcB(2)") == "PlCalcB: 2.0"


# The script's top level, run as PyLoad calculates, sends keys to Excel and waits for them
# to reach Excel's input, so that Excel handles them before the queue's timer: the
# registration is queued while Excel is busy with what they started.
PRESSES = """import ctypes, time

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort), ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong), ("dwExtraInfo", ctypes.c_size_t)]

class INPUT(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT), ("padding", ctypes.c_byte * 32)]
    _anonymous_ = ("u",)
    _fields_ = [("type", ctypes.c_ulong), ("u", _U)]

keys = {keys!r}
events = (INPUT * len(keys))()
for i, (key, up) in enumerate(keys):
    events[i].type, events[i].ki.wVk, events[i].ki.dwFlags = 1, key, 2 if up else 0
ctypes.windll.user32.SendInput(len(keys), events, ctypes.sizeof(INPUT))
time.sleep(0.5)

import pyinex

@pyinex.func
def {name}(x):
    return '{name}: ' + str(x)
"""


def console_text(pid):
    window = Console.find(pid)
    return window.text if window else ""


def wait_for_text(pid, text, timeout=20.0):
    deadline = time.monotonic() + timeout
    while text not in console_text(pid) and time.monotonic() < deadline:
        time.sleep(0.1)
    return console_text(pid)


@pytest.mark.parametrize(
    "name, keys",
    [("PlEdited", [(VK_F2, False), (VK_F2, True)]),  # A cell edited.
     ("PlInDialog", chord(VK_CONTROL, ord("1")))],    # Format Cells, a modal dialog.
    ids=["cell edit", "modal dialog"],
)
def test_registration_waits_while_excel_is_busy(app, book, cleared, excel_in_front, pytestconfig, name, keys):
    pid = pytestconfig.stash[EXCEL_PID]
    book.sheet.Range("C3").Select()
    book.script(f"{name}.py", PRESSES.format(keys=keys, name=name))
    # No COM call reaches Excel until the edit or dialog ends: only the console is read.
    book.put("A1", f'=PyLoad("{name}.py")')
    text = wait_for_text(pid, WAITING, timeout=10)
    time.sleep(3)  # Twelve times the queue's interval.
    text = console_text(pid)
    assert WAITING in text, text
    assert f"registered {name}" not in text, "registered while Excel was busy"
    send((VK_ESCAPE, False), (VK_ESCAPE, True))
    assert f"Pyinex: registered {name} from the script" in wait_for_text(pid, f"registered {name}")
    assert when_registered(app, f"{name}(1)") == f"{name}: 1.0"
    assert book.value("A1") == "1 function"


def test_with_no_workbook_open(app, tmp_path):
    for workbook in list(app.Workbooks):
        workbook.Close(SaveChanges=False)
    assert app.Workbooks.Count == 0
    script = tmp_path / "pl_no_book.py"
    script.write_text(GOOD.format("PlNoBook"), encoding="utf-8")
    assert app.Evaluate(f'PyLoad("{script}")') == "1 function"
    assert when_registered(app, "PlNoBook(1)") == "PlNoBook: 1.0"


def protected_view(app, folder):
    """Opens a saved workbook in Protected View; returns its window."""
    saved = app.Workbooks.Add()
    saved.SaveAs(str(folder / "protected.xlsx"), FileFormat=51)
    saved.Close()
    return app.ProtectedViewWindows.Open(str(folder / "protected.xlsx"))


def test_with_only_a_protected_view_window_open(app, tmp_path):
    # Its window's object has no Application, so the queue passes it over. (Evaluate gives
    # #N/A while such a window is active, so formulas are calculated as Excel 4 macros are.)
    for workbook in list(app.Workbooks):
        workbook.Close(SaveChanges=False)
    window = protected_view(app, tmp_path)
    try:
        assert app.Workbooks.Count == 0 and app.ProtectedViewWindows.Count == 1
        script = tmp_path / "pl_protected.py"
        script.write_text(GOOD.format("PlProtected"), encoding="utf-8")
        assert app.ExecuteExcel4Macro(f'PyLoad("{script}")') == "1 function"
        assert when_registered(app, "PlProtected(1)", macro=True) == "PlProtected: 1.0"
    finally:
        window.Close()


def test_with_a_protected_view_window_in_front_of_a_workbook(app, book, tmp_path):
    folder = tmp_path / "protected"
    folder.mkdir()
    window = protected_view(app, folder)
    try:
        window.Activate()
        book.script("pl_behind.py", GOOD.format("PlBehind"))
        assert book.formula("A1", '=PyLoad("pl_behind.py")') == "1 function"
        assert when_registered(app, "PlBehind(1)", macro=True) == "PlBehind: 1.0"
    finally:
        window.Close()


@pytest.fixture
def second_excel(excel, embedded, xll, run_folder, tmp_path):
    """Another Excel, with a copy of the add-in of its own, ended after the test."""
    import pywintypes
    import win32process
    from win32com.client import DispatchEx

    folder = tmp_path / "second addin"
    folder.mkdir()
    for source in (xll, xll.parent / "pyinex_core.dll"):
        shutil.copyfile(source, folder / source.name)
    (folder / "pyinex.ini").write_text(f"[python]\nexecutable = {embedded['executable']}\n", "utf-8")
    other = DispatchEx("Excel.Application")
    report_excel(run_folder, "second", win32process.GetWindowThreadProcessId(other.Hwnd)[1])
    try:
        other.Visible = True
        other.DisplayAlerts = False
        assert other.RegisterXLL(str(folder / xll.name))
        yield other
    finally:
        try:
            for workbook in list(other.Workbooks):
                workbook.Close(SaveChanges=False)
            other.Quit()
        except pywintypes.com_error:
            pass
        other = None
        gc.collect()


def test_two_excels_each_register_their_own(app, book, second_excel, tmp_path):
    book.script("pl_first_excel.py", GOOD.format("PlFirstExcel"))
    theirs = second_excel.Workbooks.Add()
    theirs.SaveAs(str(tmp_path / "second.xlsx"), FileFormat=51)
    (tmp_path / "pl_second_excel.py").write_text(GOOD.format("PlSecondExcel"), encoding="utf-8")
    assert book.formula("A1", '=PyLoad("pl_first_excel.py")') == "1 function"
    theirs.Worksheets(1).Range("A1").Formula2 = '=PyLoad("pl_second_excel.py")'
    assert theirs.Worksheets(1).Range("A1").Value2 == "1 function"
    assert when_registered(app, "PlFirstExcel(1)") == "PlFirstExcel: 1.0"
    assert when_registered(second_excel, "PlSecondExcel(1)") == "PlSecondExcel: 1.0"
    # Each in its own Excel only.
    time.sleep(1)
    assert shown(app.Evaluate("PlSecondExcel(1)")) == E("#NAME?")
    assert shown(second_excel.Evaluate("PlFirstExcel(1)")) == E("#NAME?")
