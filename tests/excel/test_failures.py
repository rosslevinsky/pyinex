"""Failures in real Excel (docs/architecture.md, "PyCall" and "Scripts"; docs/conversion.md,
"What an error means"): each gives #VALUE! and says why on the console, in a line starting
Pyinex:, and the next call works.

The console is cleared before each test, so the text a test finds was written by it.
"""

import pytest
from conftest import E, pycall


@pytest.fixture
def values(book):
    book.copy("values.py")
    return book


def fails_saying(book, cleared, formula, *said):
    """Enters the formula, checks it gives #VALUE!, and returns the console's text once it
    holds each of said, compared ignoring case (Excel may give a folder's name in another
    case than Windows does)."""
    assert book.formula("A1", formula) == E("#VALUE!")
    window = cleared(said[-1])
    text = window.text if window else ""
    for s in said:
        assert s.lower() in text.lower(), (s, text)
    assert book.formula("B1", pycall("values.py", "echo", '"still working"')) == "still working"
    return text


def test_a_path_that_is_not_text(values, cleared):
    fails_saying(values, cleared, '=PyCall(1, "f")', "Pyinex: PyCall needs the script's path and the function's name as text")


def test_a_script_that_is_not_there(values, cleared):
    fails_saying(values, cleared, pycall("missing.py", "f"), f"Pyinex: cannot find the script {values.folder / 'missing.py'}")


def test_a_function_that_is_not_there(values, cleared):
    fails_saying(values, cleared, pycall("values.py", "nope"), f"Pyinex: the script {values.folder / 'values.py'} has no function nope")


def test_a_name_that_is_not_a_function(values, cleared):
    fails_saying(
        values, cleared, pycall("values.py", "NOT_A_FUNCTION"), f"Pyinex: NOT_A_FUNCTION in the script {values.folder / 'values.py'} is not a function"
    )


@pytest.mark.parametrize(
    "function, shown",
    [
        ("raises", "ValueError: raised on purpose"),
        ("exits", "SystemExit: 3"),
        ("interrupts", "KeyboardInterrupt"),
        ("bad_error", "ValueError"),  # pyinex.ExcelError("#BAD!"): unknown text.
    ],
)
def test_an_exception_in_a_call(values, cleared, function, shown):
    fails_saying(
        values, cleared, pycall("values.py", function), f"Pyinex: {function} in the script", "raised an exception:", "Traceback (most recent call last)", shown
    )


def test_a_wrong_number_of_arguments(values, cleared):
    fails_saying(values, cleared, pycall("values.py", "two", "1"), "Pyinex: two in the script", "raised an exception:", "TypeError: two() missing 1 required positional argument: 'b'")


@pytest.mark.parametrize(
    "source, shown",
    [
        ("def f(:\n    pass\n", "SyntaxError"),
        ("raise RuntimeError('at import')\n", "RuntimeError: at import"),
        ("raise SystemExit(4)\n", "SystemExit: 4"),
        ("raise KeyboardInterrupt\n", "KeyboardInterrupt"),
    ],
)
def test_a_script_that_fails_to_load(values, cleared, source, shown):
    values.script("broken.py", source)
    fails_saying(values, cleared, pycall("broken.py", "f"), f"Pyinex: the script {values.folder / 'broken.py'} failed to load:", "Traceback (most recent call last)" if "def f(:" not in source else "SyntaxError", shown)


@pytest.mark.parametrize(
    "name, said",
    [
        ("dict", "Pyinex: values.py, give: dict cannot be returned to Excel"),
        ("set", "Pyinex: values.py, give: set cannot be returned to Excel"),
        ("complex", "Pyinex: values.py, give: complex cannot be returned to Excel"),
        ("too_large_int", "Pyinex: values.py, give: int too large to be a number in Excel"),
        ("long_text", "Pyinex: values.py, give: str longer than 32,767 UTF-16 code units, Excel's limit for a cell"),
        ("emoji_too_long", "Pyinex: values.py, give: str longer than 32,767 UTF-16 code units, Excel's limit for a cell"),
        ("too_deep", "Pyinex: values.py, give: row 1, column 2: a list or tuple inside a row is nested too deeply for Excel"),
        ("bad_in_block", "Pyinex: values.py, give: row 2, column 2: dict cannot be returned to Excel"),
    ],
)
def test_a_result_that_cannot_be_returned(values, cleared, name, said):
    fails_saying(values, cleared, pycall("values.py", "give", f'"{name}"'), said)

