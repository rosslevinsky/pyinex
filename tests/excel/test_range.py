"""pyinex.Range in real Excel (docs/architecture.md, "pyinex.Range"; docs/conversion.md,
"References"): every kind of reference in Excel's own syntax, several areas, a workbook not
open, a deleted sheet, values refused, a range kept after its call, a kept range whose sheet
is gone or renamed, or replaced by another of its name, a range returned, the refusals outside a call and on a script's thread,
cells not yet calculated giving the stop and then the right result with no call into Excel
after it, the recalculation the documents state, and reloads that change an argument from a
value to a reference and back.

A name registered stays registered for the rest of the session's Excel, so every function
here has a name starting "Xr". One script holds most of them, loaded once by PyLoad in a
workbook of the module's own, ranges.xlsx, whose Sheet1 holds B2:C3 = 1, "", (empty), "x";
D1 = 10 and D3 = 30; E5 = TRUE; and whose sheet Data holds A1:B1 = "data", 2. A second
workbook, other.xlsx, is open beside it; a third, closed.xlsx, is saved and closed.
"""

import time
from collections import Counter

import pytest
from conftest import E, Book, registered_types, shown
from test_pyload import when_registered
from test_reload_functions import ONE, loaded, when, xll  # noqa: F401 (xll: a fixture)

SCRIPT = """import threading, time
import pyinex

def attempt(f):
    try:
        return repr(f())
    except Exception as e:
        return f'{type(e).__name__}: {e}'

@pyinex.func
def XrShown(r: pyinex.Range):
    return repr([r.workbook, r.sheet, r.address(), r.address(r1c1=True), r.address(full=True), r.rows, r.columns])

@pyinex.func
def XrValues(r: pyinex.Range):
    return repr([r.values(), r.empty()])

@pyinex.func
def XrAreas(r: pyinex.Range):
    return repr([r.address(), [(a.address(), a.values()) for a in r.areas], attempt(r.values)])

kept = []

@pyinex.func
def XrKeep(r: pyinex.Range):
    kept.append(r)
    return r.address(full=True)

@pyinex.func
def XrKept(x):
    r = kept[-1]
    return repr([r.address(full=True), r.rows, r.values()])

@pyinex.func
def XrReturned(r: pyinex.Range):
    return r

@pyinex.func
def XrKeptRead(x):
    r = kept[-1]
    return repr([attempt(r.values), attempt(r.areas[0].values)])

@pyinex.func
def XrSameSheet(r: pyinex.Range):
    # Whether Excel gave r's sheet the ID of the kept range's.
    return r._id == kept[-1]._id

later = []

@pyinex.func
def XrReadLater(r: pyinex.Range):
    def read():
        time.sleep(0.5)
        later.append(attempt(r.values))
    threading.Thread(target=read).start()
    return 'reading later'

@pyinex.func
def XrReadLaterResult(x):
    return later[-1] if later else 'not yet'

@pyinex.func
def XrOnThread(r: pyinex.Range):
    out = []
    t = threading.Thread(target=lambda: out.append(attempt(r.values)))
    t.start()
    t.join()
    return out[0]

calls = []

@pyinex.func
def XrCounted(r: pyinex.Range):
    calls.append(r.address())
    return repr(r.values())

@pyinex.func
def XrCalls(x):
    return len(calls)

events = []

@pyinex.func
def XrSum(r: pyinex.Range):
    return total(r)

kept_sums = []

@pyinex.func
def XrKeepSum(r: pyinex.Range):
    kept_sums.append(r)
    return r.address()

@pyinex.func
def XrSumKept(x):
    # Reads a range kept from another cell's call, so its cells are none of this formula's.
    return total(kept_sums[-1]) if kept_sums else 'none kept'

def total(r):
    try:
        values = r.values()
    except RuntimeError as e:
        events.append('stopped: ' + str(e))
        events.append(attempt(r.values))
        events.append(attempt(pyinex.caller))
        raise
    result = sum(v or 0 for row in values for v in row)
    events.append(f'read {result}')
    return result

@pyinex.func
def XrEvents(clear):
    shown = repr(events)
    if clear:
        events.clear()
    return shown
"""

SHOWN = "['{0}', '{1}', '{2}', '{3}', '[{0}]{1}!{2}', {4}, {5}]"


@pytest.fixture(scope="module")
def sheet(excel, embedded, tmp_path_factory):
    folder = tmp_path_factory.mktemp("ranges")
    (folder / "xr_ranges.py").write_text(SCRIPT, encoding="utf-8")
    closed = excel.Workbooks.Add()
    closed.Worksheets(1).Range("A1:B1").Value = ("closed", 5)
    closed.SaveAs(str(folder / "closed.xlsx"), FileFormat=51)  # 51: .xlsx
    closed.Close(SaveChanges=False)
    other = excel.Workbooks.Add()
    while other.Worksheets.Count > 1:
        other.Worksheets(other.Worksheets.Count).Delete()
    other.Worksheets(1).Name = "Data"
    other.Worksheets(1).Range("A1").Value = "other book"
    other.SaveAs(str(folder / "other.xlsx"), FileFormat=51)
    workbook = excel.Workbooks.Add()
    while workbook.Worksheets.Count > 1:
        workbook.Worksheets(workbook.Worksheets.Count).Delete()
    workbook.Worksheets.Add().Name = "Data"  # Before Sheet1, the one the tests use.
    workbook.Worksheets("Sheet1").Activate()
    workbook.SaveAs(str(folder / "ranges.xlsx"), FileFormat=51)
    book = Book(excel, workbook)
    book.sheet = workbook.Worksheets("Sheet1")
    try:
        for cell, value in (("B2", 1), ("C2", '=""'), ("C3", "x"), ("D1", 10), ("D3", 30), ("E5", True)):
            book.put(cell, value)
        workbook.Worksheets("Data").Range("A1:B1").Value = ("data", 2)
        workbook.Names.Add("XrRates", "=Sheet1!$B$2:$C$3")
        book.put("A1", '=PyLoad("xr_ranges.py")')
        assert when_registered(excel, 'XrAreas(B2)') != E("#NAME?")
        book.folder = folder
        yield book
    finally:
        workbook.Close(SaveChanges=False)
        other.Close(SaveChanges=False)


@pytest.mark.parametrize(
    "reference, shown_as",
    [
        ("B2", SHOWN.format("ranges.xlsx", "Sheet1", "B2", "R2C2", 1, 1)),
        ("B2:C3", SHOWN.format("ranges.xlsx", "Sheet1", "B2:C3", "R2C2:R3C3", 2, 2)),
        ("D:D", SHOWN.format("ranges.xlsx", "Sheet1", "D1:D1048576", "R1C4:R1048576C4", 1048576, 1)),
        ("5:5", SHOWN.format("ranges.xlsx", "Sheet1", "A5:XFD5", "R5C1:R5C16384", 1, 16384)),
        ("Data!A1:B1", SHOWN.format("ranges.xlsx", "Data", "A1:B1", "R1C1:R1C2", 1, 2)),
        ("[other.xlsx]Data!A1", SHOWN.format("other.xlsx", "Data", "A1", "R1C1", 1, 1)),
        ("XrRates", SHOWN.format("ranges.xlsx", "Sheet1", "B2:C3", "R2C2:R3C3", 2, 2)),
    ],
    ids=["cell", "range", "whole column", "whole row", "another sheet", "another workbook", "named range"],
)
def test_each_reference_gives_its_workbook_sheet_address_and_size(sheet, reference, shown_as):
    assert sheet.formula("G1", f"=XrShown({reference})") == shown_as


@pytest.mark.parametrize(
    "reference, values",
    [
        ("B2", "[1.0, False]"),
        ("B2:C3", "[((1.0, ''), (None, 'x')), ((False, False), (True, False))]"),
        ("Z99", "[None, True]"),
        ("D:D", "[((10.0,), (None,), (30.0,)), ((False,), (True,), (False,))]"),
        ("5:5", "[((None, None, None, None, True),), ((True, True, True, True, False),)]"),
        ("Y:Y", "[None, True]"),
        ("Data!A1:B1", "[(('data', 2.0),), ((False, False),)]"),
        ("[other.xlsx]Data!A1", "['other book', False]"),
        ("XrRates", "[((1.0, ''), (None, 'x')), ((False, False), (True, False))]"),
    ],
    ids=["cell", "range", "empty cell", "whole column", "whole row", "empty column", "another sheet",
         "another workbook", "named range"],
)
def test_values_and_empty_cells(sheet, reference, values):
    assert sheet.formula("G1", f"=XrValues({reference})") == values


def test_several_areas(sheet):
    assert sheet.formula("G1", "=XrAreas((B2:C3,E5))") == repr([
        "B2:C3,E5",
        [("B2:C3", ((1.0, ""), (None, "x"))), ("E5", True)],
        "ValueError: B2:C3,E5 has 2 areas: use its areas",
    ])


@pytest.mark.parametrize(
    "argument, said",
    [
        ("5", "5.0"),
        ('"B2"', "the text 'B2'"),
        ("{1,2}", "a range of 1 × 2 cells"),
        ("B2+1", "2.0"),
        ("'{folder}\\[closed.xlsx]Sheet1'!A1:B1", "a range of 1 × 2 cells"),
    ],
    ids=["number", "text", "array constant", "calculation", "workbook not open"],
)
def test_a_value_that_is_not_a_reference_is_refused(sheet, cleared, argument, said):
    assert sheet.formula("G1", f"=XrShown({argument.replace('{folder}', str(sheet.folder))})") == E("#VALUE!")
    message = f"Pyinex: XrShown: argument 1 (r): pyinex.Range takes a reference, not {said}"
    window = cleared(holding=message)
    assert window is not None and message in window.text, window and window.text


def test_a_deleted_sheet_is_refused(sheet, cleared):
    gone = sheet.workbook.Worksheets.Add()
    gone.Name = "Gone"
    sheet.sheet.Activate()
    try:
        assert sheet.formula("G1", "=XrShown(Gone!A1)") == SHOWN.format("ranges.xlsx", "Gone", "A1", "R1C1", 1, 1)
    finally:
        sheet.excel.DisplayAlerts = False
        gone.Delete()
        sheet.excel.DisplayAlerts = True
    assert sheet.value("G1") == E("#VALUE!")
    message = "Pyinex: XrShown: argument 1 (r): pyinex.Range takes a reference, not the error #REF!"
    window = cleared(holding=message)
    assert window is not None and message in window.text, window and window.text


def test_a_range_kept_after_its_call(sheet):
    assert sheet.formula("G1", "=XrKeep(Data!A1:B1)") == "[ranges.xlsx]Data!A1:B1"
    assert sheet.formula("G2", "=XrKept(G1)") == repr(["[ranges.xlsx]Data!A1:B1", 1, (("data", 2.0),)])


@pytest.mark.parametrize("closing", [False, "renamed", True], ids=["sheet deleted", "sheet renamed", "workbook closed"])
def test_a_kept_range_whose_sheet_is_gone_is_not_read(sheet, closing):
    excel = sheet.excel
    if closing is True:
        book = excel.Workbooks.Add()
        book.SaveAs(str(sheet.folder / "kept.xlsx"), FileFormat=51)
        gone = book.Worksheets(1)
        name = f"[kept.xlsx]{gone.Name}"
        reference = f"{name}!A1"
    else:
        gone = sheet.workbook.Worksheets.Add()
        gone.Name = "Kept"
        name, reference = "[ranges.xlsx]Kept", "Kept!A1"
    gone.Range("A1").Value = "kept"
    sheet.workbook.Activate()
    sheet.sheet.Activate()
    try:
        assert sheet.formula("G1", f"=XrKeep({reference})") == f"{name}!A1"
        sheet.sheet.Range("G1").ClearContents()
    finally:
        excel.DisplayAlerts = False
        if closing is True:
            book.Close(SaveChanges=False)
        elif closing:
            gone.Name = "KeptRenamed"
        else:
            gone.Delete()
        excel.DisplayAlerts = True
    added = excel.Workbooks.Add()  # Its sheet may be given the ID Excel freed.
    try:
        added.Worksheets(1).Range("A1").Value = "another sheet"
        sheet.workbook.Activate()
        refused = (f"RuntimeError: pyinex.Range: Excel no longer has the sheet {name} open; it was deleted or "
                   "renamed, or its workbook closed")
        assert sheet.formula("G2", "=XrKeptRead(0)") == repr([refused] * 2)
    finally:
        sheet.sheet.Range("G2").ClearContents()
        added.Close(SaveChanges=False)
        if closing == "renamed":
            excel.DisplayAlerts = False
            gone.Delete()
            excel.DisplayAlerts = True


def test_a_kept_range_whose_sheet_is_replaced_under_its_name_is_read_only_with_its_id(sheet):
    excel, sheets = sheet.excel, sheet.workbook.Worksheets
    gone = sheets.Add()
    gone.Name = "Replaced"
    gone.Range("A1").Value = "kept"
    sheet.sheet.Activate()
    excel.DisplayAlerts = False
    try:
        assert sheet.formula("G1", "=XrKeep(Replaced!A1)") == "[ranges.xlsx]Replaced!A1"
        sheet.sheet.Range("G1").ClearContents()
        gone.Delete()
        replacement = sheets.Add()
        replacement.Name = "Replaced"
        replacement.Range("A1").Value = "replacement"
        sheet.sheet.Activate()
        # Whether the new sheet has the old one's ID is Excel's choice: read only if it has.
        same = sheet.formula("G2", "=XrSameSheet(Replaced!A1)")
        refused = ("RuntimeError: pyinex.Range: Excel no longer has the sheet [ranges.xlsx]Replaced open; it was "
                   "deleted or renamed, or its workbook closed")
        expected = [repr("replacement")] * 2 if same else [refused] * 2
        assert sheet.formula("G3", "=XrKeptRead(0)") == repr(expected), f"same ID: {same}"
    finally:
        sheet.sheet.Range("G2:G3").ClearContents()
        sheets("Replaced").Delete()
        excel.DisplayAlerts = True


def test_a_range_returned_is_value_as_any_object_excel_cannot_take(sheet, cleared):
    assert sheet.formula("G1", "=XrReturned(B2)") == E("#VALUE!")
    message = "Pyinex: XrReturned: Range cannot be returned to Excel"
    window = cleared(holding=message)
    assert window is not None and message in window.text, window and window.text


def test_reading_outside_a_call_is_refused(sheet):
    assert sheet.formula("G1", "=XrReadLater(B2)") == "reading later"
    time.sleep(2)  # With no call meanwhile: the thread reads between calls.
    assert sheet.formula("G2", "=XrReadLaterResult(1)") == (
        "RuntimeError: pyinex.Range asks Excel, which allows that only while Excel is calculating a PyCall or a "
        "direct function")


def test_reading_on_a_scripts_thread_is_refused(sheet):
    assert sheet.formula("G1", "=XrOnThread(B2)") == (
        "RuntimeError: pyinex.Range asks Excel, which allows that only on Excel's own thread, not one a script started")


def test_a_function_recalculates_when_a_cell_it_refers_to_changes_and_only_then(sheet):
    sheet.put("H1", 1)
    sheet.put("H2", 2)
    assert sheet.formula("G1", "=XrCounted(H1:H2)") == "((1.0,), (2.0,))"
    count = sheet.excel.Evaluate("XrCalls(0)")
    sheet.put("H2", 5)  # Inside the reference: calculated again.
    assert sheet.value("G1") == "((1.0,), (5.0,))"
    assert sheet.excel.Evaluate("XrCalls(0)") == count + 1
    sheet.put("H9", 7)  # Outside it: not.
    sheet.put("Z1", 7)
    assert sheet.excel.Evaluate("XrCalls(0)") == count + 1
    sheet.excel.Calculate()  # Not volatile: an ordinary recalculation leaves it alone.
    assert sheet.excel.Evaluate("XrCalls(0)") == count + 1


def full_calculation(sheet, formulas):
    """With calculation manual, enters the cells L1:L3 that formulas read (=M1*2 and so on,
    M1:M3 = 1, 2, 3), then the formulas, in order, then calculates fully, which then meets
    the formulas before L1:L3; returns what the formulas' reads saw meanwhile."""
    app = sheet.excel
    app.Calculation = -4135  # xlCalculationManual
    try:
        for row in (1, 2, 3):
            sheet.put(f"M{row}", row)
            sheet.put(f"L{row}", f"=M{row}*2")
        for cell, formula in formulas:
            sheet.put(cell, formula)
        app.Evaluate("XrEvents(TRUE)")  # Cleared. (Evaluate may call a function twice.)
        app.CalculateFull()
        got = app.Evaluate("XrEvents(FALSE)")
        return eval(got[0] if isinstance(got, tuple) else got)  # Long text comes as a 1-tuple.
    finally:
        app.Calculation = -4105  # xlCalculationAutomatic
        for cell, _ in formulas:
            sheet.sheet.Range(cell).ClearContents()


@pytest.mark.parametrize(
    "formulas",
    [
        [("K1", "=XrSum(L1:L3)")],
        [("K1", '=XrSum(INDIRECT("L1:L3"))')],
        # K2 keeps the range and K1 reads it, so K1's formula names none of the cells it reads.
        [("K2", "=XrKeepSum(L1:L3)"), ("K1", "=XrSumKept(0)")],
    ],
    ids=["reference", "INDIRECT", "range kept by another cell"],
)
def test_a_cell_not_yet_calculated_stops_the_call_and_excel_calls_again(sheet, cleared, formulas):
    stop = ("pyinex.Range cannot ask Excel: a cell this call reads is not yet calculated, and Excel calls the "
            "function again once it is")
    events = full_calculation(sheet, formulas)
    assert "stopped: " + stop in events, events
    first = events.index("stopped: " + stop)
    # After the stop nothing more asks Excel, the function's own code and Pyinex's.
    assert events[first + 1:first + 3] == ["RuntimeError: " + stop,
                                           "RuntimeError: " + stop.replace("pyinex.Range", "pyinex.caller")]
    assert events[-1] == "read 12.0", events  # Called again once the cells had calculated.
    text = cleared().text
    assert "XrSum" not in text and "Traceback" not in text, text  # Nothing said of the stopped call.


# Reloads that change an argument from a value to a reference and back.


def test_a_value_becoming_a_reference_and_back_registers_again(app, book, xll, cleared):
    loaded(app, book, "xr_switch.py", ONE.format("XrSwitch", "a", "repr(a)"), "XrSwitch(0)")
    book.put("A1", 1)
    assert book.formula("B1", "=XrSwitch(A1)") == "1.0"
    before = Counter(registered_types(app, xll).values())
    book.script("xr_switch.py", ONE.format("XrSwitch", "a: pyinex.Range", "a.address()"))
    book.put("A1", 2)  # A call before the queue has run: its value never reaches the new code.
    assert book.value("B1") == E("#VALUE!")
    assert "Pyinex: XrSwitch is being registered again" in cleared(holding="being registered again").text
    assert when(lambda: shown(app.Evaluate("XrSwitch(A1)")) != E("#VALUE!"))
    after = Counter(registered_types(app, xll).values())
    assert (after - before, before - after) == (Counter(["QU"]), Counter(["QQ"]))
    book.put("A1", 3)
    assert book.value("B1") == "A1"
    book.script("xr_switch.py", ONE.format("XrSwitch", "a", "repr(a)"))
    book.put("A1", 4)
    assert book.value("B1") == E("#VALUE!")
    assert when(lambda: Counter(registered_types(app, xll).values()) == before)
    book.put("A1", 5)
    assert book.value("B1") == "5.0"
