"""The #NAME? repair in real Excel (docs/architecture.md, "The #NAME? repair"): the
direct-functions test group 8. Cells typed before their function registered, on two sheets,
in a formula that also uses other functions, in a spilling formula, inside IFERROR, and in
an IF branch first taken after registration, and in a legacy array formula, each right once
the function registers; a formula of a direct function registered before, text holding a
new name, and formulas holding it only in a text or within a longer name, untouched, as a
function counting its calls shows; a formula on a protected sheet reported, then repaired
by PyRepair.

Functions have names of their own, starting "Rp", since a registered name stays for the
session.
"""

import time

from conftest import E, shown
from test_pyload import when_registered

FUNCTIONS = """import pyinex

@pyinex.func
def RpTwice(x):
    return x * 2

@pyinex.func
def RpRows(n):
    return [[i] for i in range(1, int(n) + 1)]
"""

COUNTED = """import pyinex

calls = 0

@pyinex.func
def RpCounted():
    global calls
    calls += 1
    return calls

@pyinex.func
def RpXRpTwice(x):
    return 0
"""


def when(condition, timeout=15.0):
    """Whether condition() became true within timeout."""
    deadline = time.monotonic() + timeout
    while not condition():
        if time.monotonic() > deadline:
            return False
        time.sleep(0.1)
    return True


def test_cells_typed_before_registration_are_repaired(app, book):
    sheet, other = book.sheet, book.workbook.Worksheets.Add(After=book.sheet)
    book.script("rp_counted.py", COUNTED)
    assert book.formula("F1", '=PyLoad("rp_counted.py")') == "2 functions"
    assert when(lambda: isinstance(shown(app.Evaluate("RpCounted()")), float))
    assert when(lambda: shown(app.Evaluate("RpXRpTwice(1)")) == 0.0)
    assert book.formula("D1", "=RpCounted()") == (counted := book.value("D1"))  # Registered before: untouched.
    book.put("D2", "RpTwice(2) as text")  # Holds a new name, but is not a formula: untouched.
    # Formulas that hold the new name and "(", but do not call it: untouched. (The longer
    # name is registered: Excel calculates again any formula holding an unknown name each
    # time a name is registered.)
    untouched = {cell: book.formula(cell, formula) for cell, formula in (
        ("D3", '=RpCounted()+N("RpTwice(")'), ("D4", "=RpCounted()+RpXRpTwice(1)"))}
    book.put("A1", False)
    typed = [
        (sheet, "B1", "=RpTwice(2)", E("#NAME?"), 4.0),
        (other, "B1", "=RpTwice(3)", E("#NAME?"), 6.0),  # Another sheet.
        (sheet, "B2", "=ROUND(RpTwice(1.25),1)+SUM(1,2)", E("#NAME?"), 5.5),  # With other functions.
        (sheet, "C1", "=IFERROR(RpTwice(4),0)", 0.0, 8.0),  # Hidden behind the fallback.
        (sheet, "C2", "=IF(A1,RpTwice(5),-1)", -1.0, -1.0),  # A branch not taken.
        (sheet, "B3", "=RpRows(3)", E("#NAME?"), 1.0),  # Spills B3:B5 once it works.
    ]
    for where, cell, formula, before, _ in typed:
        where.Range(cell).Formula2 = formula
        assert shown(where.Range(cell).Value2) == before, cell
    # Ctrl+Shift+Enter: Excel changes it only whole.
    assert book.array("G1:G2", "=RpTwice(1)*{1;2}") == [[E("#NAME?")], [E("#NAME?")]]
    book.script("rp_functions.py", FUNCTIONS)
    book.put("F2", '=PyLoad("rp_functions.py")')
    assert when(lambda: book.value("B1") == 4.0), book.value("B1")
    for where, cell, formula, _, after in typed:
        assert shown(where.Range(cell).Value2) == after, cell
        assert where.Range(cell).Formula2 == formula, cell
    assert book.values("B3:B5") == [[1.0], [2.0], [3.0]]
    assert book.values("G1:G2") == [[2.0], [4.0]]
    assert sheet.Range("G1").HasArray and sheet.Range("G1").CurrentArray.Address == "$G$1:$G$2"
    assert sheet.Range("G1").FormulaArray == "=RpTwice(1)*{1;2}"
    book.put("A1", True)  # The branch, first taken now, has its call resolved.
    assert book.value("C2") == 10.0
    assert book.value("D1") == counted  # Not entered again, so not calculated again.
    assert {cell: book.value(cell) for cell in untouched} == untouched
    assert book.value("D2") == "RpTwice(2) as text" and not sheet.Range("D2").HasFormula


def test_a_protected_sheet_is_reported_and_pyrepair_repairs_it(app, book, cleared):
    loader = book.workbook.Worksheets.Add(After=book.sheet)
    book.script("rp_locked.py", "import pyinex\n\n@pyinex.func\ndef RpLocked(x):\n    return x + 1\n")
    assert book.formula("B1", "=RpLocked(1)") == E("#NAME?")
    book.sheet.Protect()  # Its cells are locked.
    try:
        loader.Range("A1").Formula2 = '=PyLoad("rp_locked.py")'
        text = cleared(holding="again (").text
        assert "Pyinex: cannot enter the formula in [book.xlsx]Sheet1!$B$1 again (" in text, text
        assert book.value("B1") == E("#NAME?")
    finally:
        book.sheet.Unprotect()
    app.Run("PyRepair")
    assert book.value("B1") == 2.0


def test_a_change_handler_that_runs_pyrepair_lets_the_repair_end(app, book):
    # A workbook's change handler running PyRepair, as a VBA Worksheet_Change handler with
    # Application.Run "PyRepair" would, as each formula the repair enters changes a cell. Here
    # the handler is this process's, connected through COM; it gives up after 10 runs, where
    # a repair that asked for itself would run it for ever.
    import win32com.client

    book.script("rp_rerun.py", "import pyinex\n\n@pyinex.func\ndef RpRerun(x):\n    return x * 2\n")
    assert book.formula("A1", '=PyLoad("rp_rerun.py")') == "1 function"
    assert when_registered(app, "RpRerun(1)") == 2.0
    assert book.formula("A2", "=RpRerun(1)") == 2.0
    runs = []

    class Handler:
        def OnChange(self, target):
            runs.append(target.Address)
            if len(runs) <= 10:
                app.Run("PyRepair")

    handler = win32com.client.WithEvents(book.sheet, Handler)
    try:
        app.Run("PyRepair")
    finally:
        handler.close()  # Disconnected, or Excel holds this process's handler and cannot quit.
    assert sorted(runs) == ["$A$1", "$A$2"]  # Each formula entered once.
