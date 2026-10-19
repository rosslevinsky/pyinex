"""The functions Excel has registered for the add-in, as Excel lists them
(docs/architecture.md, "Excel functions", "Direct functions" and "The queue"): the five and
the two commands, with their type strings, none marked thread-safe; and, for a copy of the
add-in whose pyinex.ini lists tests/excel/scripts/direct.py, those and direct.py's functions,
each in its slot. In the session's Excel, the direct functions other tests loaded with PyLoad
(PyFunc0 and on) are left out of the first comparison.
"""

from pathlib import Path

EXPECTED = {
    "PyCall": "Q" * 33,  # The result, file, function and 30 arguments.
    "PyConsole": "Q" * 7,
    "PyReloadOnChange": "QQ",
    "PyInfo": "QQ",
    "PyLoad": "QQ",
    "PyinexQueue": "J",  # The queue's command.
    "PyRepair": "J",  # A command.
}


def test_excel_lists_the_documented_functions(app, pytestconfig):
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    registered = {
        procedure: type_text
        for dll, procedure, type_text in app.RegisteredFunctions
        if Path(dll).resolve() == xll and not procedure.startswith("PyFunc")
    }
    assert registered == EXPECTED
    assert not any("$" in t for t in registered.values())  # Not thread-safe: the main thread only.


# direct.py's functions, in its order, with each one's type string: one Q for the result and
# one for each argument.
DIRECT = {
    "Describe": 1, "Pair": 2, "Length": 1, "EchoBack": 1, "Give": 1, "IsGiven": 2, "GiveError": 1, "Raises": 0,
    "Exits": 0, "Interrupts": 0, "Two": 2, "CallerOf": 2, "SheetOf": 0, "Attempt": 1, "OnAThread": 1,
    "PrintLater": 1, "StartTicking": 0, "Ticked": 0, "LoopThenCheck": 2, "Marked": 2, "MarkedToo": 0,
}  # fmt: skip


def test_excel_lists_the_direct_functions_pyinex_ini_names(direct_excel):
    registered = {
        procedure: type_text
        for dll, procedure, type_text in direct_excel.app.RegisteredFunctions
        if Path(dll).resolve() == direct_excel.xll.resolve()
    }
    slots = {f"PyFunc{i}": "Q" * (count + 1) for i, count in enumerate(DIRECT.values())}
    assert registered == EXPECTED | slots
    assert not any("$" in t for t in registered.values())
