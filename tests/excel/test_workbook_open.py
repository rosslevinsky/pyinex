"""A workbook's PyLoad scripts load when it opens, in real Excel (docs/architecture.md, "When
a workbook opens"): a workbook saved in the session's Excel, closed, then opened in a new
Excel. The script runs, shown by a file its top level writes, before any input is changed
or any recalculation asked for; then a changed input recalculates the cell that uses its
function. With a literal path, a path in a cell, the PyLoad cell on a hidden or very hidden
sheet, a volatile function in the workbook, a workbook made from a template, the
workbook already open when the add-in loads, and an Excel with no workbook open, whose
queue waits without a word until one opens; and a workbook opened with Excel's events off,
which is missed until Ctrl+Alt+F9 or PyRepair loads it.

Every new Excel loads a copy of the add-in, whose pyinex.ini names the session's Python.
Functions have names of their own, starting "Wo", since each also registers in the
session's Excel.
"""

import gc
import shutil
import time
from pathlib import Path

import pytest
from conftest import Console, report_excel, shown

SCRIPT = """import pathlib

import pyinex

pathlib.Path(__file__).with_name({marker!r}).write_text("loaded", encoding="utf-8")

@pyinex.func
def {name}(x):
    return x * 2
"""


def when(condition, timeout=20.0):
    """Whether condition() became true within timeout."""
    deadline = time.monotonic() + timeout
    while not condition():
        if time.monotonic() > deadline:
            return False
        time.sleep(0.1)
    return True


def saved_workbook(app, folder, name, cell=False, hidden=False, volatile=False, template=False):
    """Writes the script name.py, and saves the workbook name.xlsx using it from the session's
    Excel, then closes it: A1 holds 3 and B1 =name(A1), 6; D1 of the first sheet, or of a
    second one hidden (hidden True) or very hidden ("very"), loads the script, with its path in
    the formula or in C1. With template, it is saved as the template name.xltx, its PyLoad
    path whole, since a workbook made from it has no folder. Returns the workbook's path and the
    file the script writes as it loads, which is then deleted."""
    marker = folder / f"{name}.loaded"
    (folder / f"{name}.py").write_text(SCRIPT.format(marker=marker.name, name=name), encoding="utf-8")
    path = folder / f"{name}.{'xltx' if template else 'xlsx'}"
    workbook = app.Workbooks.Add()
    try:
        while workbook.Worksheets.Count > 1:
            workbook.Worksheets(workbook.Worksheets.Count).Delete()
        sheet = workbook.Worksheets(1)
        loader = workbook.Worksheets.Add(After=sheet) if hidden else sheet
        workbook.SaveAs(str(path), FileFormat=54 if template else 51)  # .xltx, .xlsx. Saved first, for the relative path.
        sheet.Range("A1").Value2 = 3
        if cell:
            loader.Range("C1").Value2 = f"{name}.py"
        script = folder / f"{name}.py" if template else f"{name}.py"
        loader.Range("D1").Formula2 = "=PyLoad(C1)" if cell else f'=PyLoad("{script}")'
        assert shown(loader.Range("D1").Value2) == "1 function"
        assert when(lambda: shown(app.Evaluate(f"{name}(1)")) == 2.0)
        sheet.Range("B1").Formula2 = f"={name}(A1)" + ("+NOW()*0" if volatile else "")
        assert shown(sheet.Range("B1").Value2) == 6.0
        if hidden:
            loader.Visible = 2 if hidden == "very" else 0  # xlSheetVeryHidden, xlSheetHidden
        workbook.Save()
    finally:
        workbook.Close(SaveChanges=False)
    marker.unlink()
    return path, marker


@pytest.fixture
def new_excel(excel, embedded, pytestconfig, run_folder, tmp_path):
    """new_excel(before=None, empty=False): a new Excel, with a copy of the add-in of its own
    loaded after the workbook before is opened, or, with none, after a new workbook is made,
    so that the add-in reaches Excel and follows its events before the test opens anything;
    with empty, with no workbook open. Each is quit after the test."""
    import pywintypes
    import win32process
    from win32com.client import DispatchEx

    xll = Path(pytestconfig.getoption("--excel")).resolve()
    folder = tmp_path / "new addin"
    folder.mkdir()
    for source in (xll, xll.parent / "pyinex_core.dll"):
        shutil.copyfile(source, folder / source.name)
    (folder / "pyinex.ini").write_text(f"[python]\nexecutable = {embedded['executable']}\n", "utf-8")
    started = []

    def start(before=None, empty=False):
        other = DispatchEx("Excel.Application")
        report_excel(run_folder, f"workbook open {len(started)}", win32process.GetWindowThreadProcessId(other.Hwnd)[1])
        started.append(other)
        other.Visible = True
        other.DisplayAlerts = False
        if before:
            other.Workbooks.Open(str(before))
        elif not empty:
            other.Workbooks.Add()
        assert other.RegisterXLL(str(folder / xll.name))
        if not before:
            time.sleep(2)  # The queue tries every 250 ms; once it reaches Excel it follows its events.
        return other

    try:
        yield start
    finally:
        for other in started:
            try:
                for workbook in list(other.Workbooks):
                    workbook.Close(SaveChanges=False)
                other.Quit()
            except pywintypes.com_error:
                pass
        started.clear()
        other = None
        gc.collect()


def check_cells(book):
    """B1, =name(A1) with A1 3, shows 6, then follows A1."""
    sheet = book.Worksheets("Sheet1")
    assert when(lambda: shown(sheet.Range("B1").Value2) == 6.0), shown(sheet.Range("B1").Value2)
    sheet.Range("A1").Value2 = 5
    assert shown(sheet.Range("B1").Value2) == 10.0


@pytest.mark.parametrize(
    "name, options",
    [
        ("WoLiteral", {}),
        ("WoInACell", {"cell": True}),
        ("WoHidden", {"hidden": True}),
        ("WoVeryHidden", {"hidden": "very"}),
        ("WoVolatile", {"volatile": True}),
    ],
    ids=["literal path", "path in a cell", "hidden sheet", "very hidden sheet", "volatile function"],
)
def test_a_workbook_opened_in_a_new_excel_loads_its_scripts(app, new_excel, tmp_path, name, options):
    path, marker = saved_workbook(app, tmp_path, name, **options)
    other = new_excel()
    book = other.Workbooks.Open(str(path))
    # Nothing is asked of Excel until the script has run.
    assert when(marker.exists), "the script did not run as the workbook opened"
    # Its PyLoad cell, entered again, counts as a change.
    assert when(lambda: not book.Saved), "the workbook does not count as changed"
    check_cells(book)


def test_a_workbook_made_from_a_template_loads_its_scripts(app, new_excel, tmp_path):
    # Excel reports it as NewWorkbook, not WorkbookOpen.
    path, marker = saved_workbook(app, tmp_path, "WoTemplate", template=True)
    other = new_excel()
    book = other.Workbooks.Add(str(path))
    assert when(marker.exists), "the script did not run as the workbook was made"
    check_cells(book)


def test_a_workbook_open_before_the_addin_loads_loads_its_scripts(app, new_excel, tmp_path):
    path, marker = saved_workbook(app, tmp_path, "WoAlreadyOpen")
    other = new_excel(before=path)
    assert when(marker.exists), "the script did not run when the add-in loaded"
    check_cells(other.Workbooks(path.name))


def test_with_no_workbook_open_the_queue_waits_without_a_word_then_follows(app, new_excel, tmp_path):
    import win32process

    path, marker = saved_workbook(app, tmp_path, "WoNoneOpen")
    other = new_excel(empty=True)
    assert other.Workbooks.Count == 0
    time.sleep(3)  # The queue tries every 250 ms, finding no workbook window.
    # Nothing was written: the add-in makes its console window only for output.
    assert Console.find(win32process.GetWindowThreadProcessId(other.Hwnd)[1]) is None
    book = other.Workbooks.Open(str(path))
    assert when(marker.exists), "the script did not run as the workbook opened"
    check_cells(book)


@pytest.mark.parametrize("workaround", ["Ctrl+Alt+F9", "PyRepair"])
def test_a_workbook_opened_with_events_off_loads_by_the_workaround(app, new_excel, tmp_path, workaround):
    path, marker = saved_workbook(app, tmp_path, "WoEventsOff" + ("Full" if workaround == "Ctrl+Alt+F9" else "Repair"))
    other = new_excel()
    other.EnableEvents = False  # As a macro may, around its work.
    try:
        book = other.Workbooks.Open(str(path))
    finally:
        other.EnableEvents = True
    time.sleep(3)
    assert not marker.exists(), "Excel reported a workbook opened with its events off"  # The documented gap.
    if workaround == "Ctrl+Alt+F9":
        other.CalculateFull()
    else:
        other.Run("PyRepair")
    assert when(marker.exists), f"{workaround} did not load the script"
    check_cells(book)
