"""The functions Excel has registered for the add-in, as Excel lists them
(docs/architecture.md, "Excel functions"): the four, with their type strings, none marked
thread-safe.
"""

from pathlib import Path

EXPECTED = {
    "PyCall": "Q" * 33,  # The result, file, function and 30 arguments.
    "PyConsole": "Q" * 7,
    "PyReloadOnChange": "QQ",
    "PyInfo": "QQ",
}


def test_excel_lists_the_documented_functions(app, pytestconfig):
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    registered = {
        procedure: type_text
        for dll, procedure, type_text in app.RegisteredFunctions
        if Path(dll).resolve() == xll
    }
    assert registered == EXPECTED
    assert not any("$" in t for t in registered.values())  # Not thread-safe: the main thread only.
