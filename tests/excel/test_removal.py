"""Removing direct functions, and the pool's limits, in real Excel (docs/architecture.md,
"Removing a function" and "A full pool"): what removal leaves, asserted before and after
with the checked reader of Excel's registration table (conftest.registered_types) and in
Insert Function; and the direct-functions test group 9: the pool filled
exactly, one function more, and many cycles of registering and removing.

A removal with a step of the name-clearing recipe failing cannot be brought about in real
Excel; tests/host/test_removal.py makes each step fail in the fake Excel. Real Excel's
answers to a failed unregistering and a failed deleting of the name were checked with an
add-in of its own before this was built.
"""

import gc
import shutil
import time
from pathlib import Path

import pytest
from conftest import EXCEL_PID, E, Console, registered_types, report_excel, shown
from test_direct_functions import VK_ESCAPE, VK_MENU, VK_SHIFT, VK_F3, accessible_texts, chord, dialog, send
from test_pyload import when_registered
from test_reload_functions import ONE, when

POOL = 1000


@pytest.fixture
def xll(pytestconfig):
    return Path(pytestconfig.getoption("--excel")).resolve()


def insert_function_lists(pid, category_key):
    """The texts Insert Function shows for the first category starting with category_key
    (Shift+F3, then Alt+C and the key)."""
    send(*chord(VK_SHIFT, VK_F3))
    window = dialog(pid, "Insert Function")
    assert window, "the Insert Function dialog did not open"
    try:
        send(*chord(VK_MENU, ord("C")))
        send(*chord(ord(category_key)))
        time.sleep(1)
        return accessible_texts(window)
    finally:
        for _ in range(3):  # The first may only close the category's list.
            send(*chord(VK_ESCAPE))
            if not dialog(pid, "Insert Function", timeout=2, present=False):
                break


LEAVING = '''import pyinex

@pyinex.func(category="Zz Removal")
def RmLeaving(a):
    """Removed soon."""
    return "leaving " + str(a)

@pyinex.func
def RmStaying(a):
    return a
'''


def test_what_removal_leaves(app, book, xll, excel_in_front, pytestconfig):
    pid = pytestconfig.stash[EXCEL_PID]
    book.script("rm_leaves.py", LEAVING)
    assert book.formula("A9", '=PyLoad("rm_leaves.py")') == "2 functions"
    assert when_registered(app, "RmLeaving(0)") == "leaving 0.0"
    count = len(registered_types(app, xll))
    book.put("A1", 1)
    assert book.formula("B1", "=RmLeaving(A1)") == "leaving 1.0"
    book.sheet.Range("C3").Select()
    assert "RmLeaving, Removed soon." in insert_function_lists(pid, "Z")

    book.script("rm_leaves.py", "import pyinex\n\n@pyinex.func\ndef RmStaying(a):\n    return a\n")
    assert book.formula("A9", '=PyLoad("rm_leaves.py")') == "1 function"
    assert when(lambda: len(registered_types(app, xll)) == count - 1)  # Gone from the table.
    book.sheet.Range("C3").Select()
    assert not [t for t in insert_function_lists(pid, "Z") if "RmLeaving" in t]  # And from Insert Function.
    assert book.value("B1") == "leaving 1.0"  # A cell keeps its value until it calculates ...
    book.put("A1", 2)
    assert book.value("B1") == E("#VALUE!")  # ... then gives #VALUE!,
    assert book.formula("B2", "=RmLeaving(3)") == E("#VALUE!")  # as does one typed afresh.
    assert book.formula("B3", "=RmStaying(3)") == 3.0


def test_many_cycles_of_registering_and_removing(app, book, xll):
    book.script("rm_cycle.py", ONE.format("RmCycle0", "a", "'RmCycle0'"))
    assert book.formula("A9", '=PyLoad("rm_cycle.py")') == "1 function"
    assert when_registered(app, "RmCycle0(1)") == "RmCycle0"
    count = len(registered_types(app, xll))
    book.put("B1", "=RmCycle0(1)")
    for n in range(1, 51):
        book.script("rm_cycle.py", ONE.format(f"RmCycle{n}", "a", f"'RmCycle{n}'"))
        assert book.formula("A9", '=PyLoad("rm_cycle.py")') == "1 function"
        assert when_registered(app, f"RmCycle{n}(1)") == f"RmCycle{n}"
        assert len(registered_types(app, xll)) == count  # The slot freed is the one taken.
        app.CalculateFull()
        assert book.value("B1") == E("#VALUE!")  # The first name never reaches a later function.
        assert book.formula("B2", f"=RmCycle{n - 1}(1)") == E("#VALUE!")


@pytest.fixture
def full_excel(excel, embedded, xll, run_folder, tmp_path):
    """An Excel of its own whose add-in, a copy, lists in its pyinex.ini a script declaring
    exactly as many functions as the pool has slots; ended after the test."""
    import pywintypes
    import win32process
    from win32com.client import DispatchEx

    folder = tmp_path / "full addin"
    folder.mkdir()
    for source in (xll, xll.parent / "pyinex_core.dll"):
        shutil.copyfile(source, folder / source.name)
    (folder / "full.py").write_text(
        "import pyinex\n\nfor i in range(%d):\n    globals()['f%%d' %% i] = pyinex.func(name='Full_%%d' %% i)(lambda x, i=i: i)\n" % POOL,
        encoding="utf-8")
    (folder / "pyinex.ini").write_text(f"[python]\nexecutable = {embedded['executable']}\nscripts = full.py\n", "utf-8")
    other = DispatchEx("Excel.Application")
    pid = win32process.GetWindowThreadProcessId(other.Hwnd)[1]
    report_excel(run_folder, "full", pid)
    try:
        other.Visible = True
        other.DisplayAlerts = False
        assert other.RegisterXLL(str(folder / xll.name))
        other.Workbooks.Add()  # A workbook window, through which the queue reaches this Excel.
        yield other, pid, folder / xll.name
    finally:
        try:
            for workbook in list(other.Workbooks):
                workbook.Close(SaveChanges=False)
            other.Quit()
        except pywintypes.com_error:
            pass
        other = None
        gc.collect()


def test_the_pool_filled_exactly_then_one_function_more(full_excel, tmp_path):
    app, pid, xll = full_excel
    assert len(registered_types(app, xll)) == POOL
    assert shown(app.Evaluate("Full_0(1)")) == 0.0 and shown(app.Evaluate(f"Full_{POOL - 1}(1)")) == POOL - 1
    window = Console.find(pid)
    assert window is None or "slots are taken" not in window.text
    script = tmp_path / "one_more.py"
    script.write_text(ONE.format("OneMore", "a", "'one more'"), encoding="utf-8")
    assert app.Evaluate(f'PyLoad("{script}")') == "1 function"
    said = f"Pyinex: cannot register OneMore from the script {script}: all {POOL} of Pyinex's function slots are taken"
    deadline = time.monotonic() + 15
    while not ((window := Console.find(pid)) and said in window.text) and time.monotonic() < deadline:
        time.sleep(0.1)
    assert window is not None and said in window.text
    assert len(registered_types(app, xll)) == POOL
    assert shown(app.Evaluate("OneMore(1)")) == E("#NAME?")
    assert shown(app.Evaluate("Full_0(1)")) == 0.0 and shown(app.Evaluate(f"Full_{POOL - 1}(1)")) == POOL - 1
