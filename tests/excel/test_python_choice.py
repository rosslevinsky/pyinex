"""Python's start in real Excel (docs/architecture.md, "Choosing a Python", "Refusals" and
"Lifetime"): a refused Python, PyConsole before Python starts, and each PyInfo item
starting it. Each kind of run needs an Excel of its own, so the driver starts one for
each, and conftest.py collects only the tests marked for that run.

These tests take the session's Excel as it is (the excel fixture), without the identity
check, which would start Python first.
"""

import time
from pathlib import Path

import pytest
from conftest import EXCEL_PID, ITEMS, Console, E, python_dlls, shown


def evaluate(excel, formula):
    return shown(excel.Evaluate(formula))


@pytest.fixture
def pid(excel, pytestconfig):
    return pytestconfig.stash[EXCEL_PID]


# A pyinex.ini naming no Python.


@pytest.mark.kind("refused")
def test_a_refused_python_says_why_and_loads_nothing(excel, pid, expected, tmp_path):
    status = evaluate(excel, 'PyInfo("status")')
    assert status.startswith(f"cannot use Python {expected['executable']}: file not found"), status
    assert python_dlls(pid) == []
    script = tmp_path / "never.py"
    script.write_text("def f():\n    return 'ran'\n", encoding="utf-8")
    assert evaluate(excel, f'PyCall("{script}", "f")') == E("#VALUE!")
    assert python_dlls(pid) == []


@pytest.mark.kind("refused")
def test_after_a_refusal_pyinfo_gives_only_status_and_pyinex(excel, pytestconfig):
    for item in ITEMS:
        got = evaluate(excel, f'PyInfo("{item}")')
        if item == "status":
            assert "file not found" in got
        elif item == "pyinex":
            assert got.lower() == str(pytestconfig.getoption("--excel")).lower()
        else:
            assert got == E("#VALUE!"), item
    table = excel.Evaluate("PyInfo()")
    assert [row[0] for row in table] == ITEMS
    assert [shown(row[1]) == E("#VALUE!") for row in table] == [item not in ("status", "pyinex") for item in ITEMS]


@pytest.mark.kind("refused")
def test_after_a_refusal_the_console_and_reload_setting_work(excel, pid):
    assert excel.Evaluate("PyConsole(TRUE)") is True
    window = Console.find(pid)
    assert window is not None and window.shown
    text = window.text
    assert "Pyinex: cannot use Python" in text and "file not found" in text, text
    assert excel.Evaluate("PyConsole(FALSE)") is False
    assert excel.Evaluate("PyReloadOnChange(FALSE)") is False
    assert excel.Evaluate("PyReloadOnChange()") is False
    assert excel.Evaluate("PyReloadOnChange(TRUE)") is True
    assert python_dlls(pid) == []


@pytest.mark.kind("refused")
def test_a_refusal_holds_until_the_addin_is_reopened(excel, pid, expected, pytestconfig, tmp_path):
    # Last of the refused run's tests: it ends the refusal.
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    (xll.parent / "pyinex.ini").write_text(f"[python]\nexecutable = {expected['retry']}\n", encoding="utf-8")
    for _ in range(3):
        assert "file not found" in evaluate(excel, 'PyInfo("status")')
        assert evaluate(excel, 'PyInfo("version")') == E("#VALUE!")
    assert python_dlls(pid) == []
    # Reopened, Pyinex tries again, and finds the Python now named.
    assert excel.ExecuteExcel4Macro(f'UNREGISTER("{xll}")') is True
    assert excel.RegisterXLL(str(xll))
    assert evaluate(excel, 'PyInfo("status")') == "running"
    assert evaluate(excel, 'PyInfo("executable")').lower() == expected["retry"].lower()
    assert python_dlls(pid)
    # The refusal was said once, not on every call: counted once a line printed now has
    # arrived, after everything written before it.
    script = tmp_path / "marks.py"
    script.write_text("def f():\n    print('after the refusal')\n    return 1\n", encoding="utf-8")
    assert excel.Evaluate(f'PyCall("{script}", "f")') == 1.0
    deadline = time.monotonic() + 10
    while "after the refusal" not in Console.find(pid).text and time.monotonic() < deadline:
        time.sleep(0.05)
    text = Console.find(pid).text
    assert "after the refusal" in text
    assert text.count("Pyinex: cannot use Python") == 1, text


# PyConsole as the first call.


@pytest.mark.kind("console_first")
def test_pyconsole_works_before_python_starts(excel, pid, expected):
    assert python_dlls(pid) == []
    assert excel.Evaluate("PyConsole(TRUE)") is True
    window = Console.find(pid)
    assert window is not None and window.shown and window.title == "Pyinex console"
    assert excel.Evaluate("PyConsole(FALSE, TRUE)") is False
    assert python_dlls(pid) == [], "PyConsole started Python"
    # Python starts at the first call that needs it, and is the one chosen.
    assert evaluate(excel, 'PyInfo("status")') == "running"
    assert python_dlls(pid)
    assert evaluate(excel, 'PyInfo("executable")').lower() == expected["executable"].lower()


# Each PyInfo item as the first call.


@pytest.mark.kind("first_call")
def test_any_pyinfo_item_asked_first_starts_python(excel, pid, expected):
    item = expected["first"]
    assert python_dlls(pid) == []
    got = evaluate(excel, f'PyInfo("{item}")')
    assert not isinstance(got, E), f"PyInfo({item!r}) gave {got}"
    assert python_dlls(pid), f"PyInfo({item!r}) did not start Python"
    # What the first call gave is what the item is once Python runs.
    again = evaluate(excel, f'PyInfo("{item}")')
    assert (type(got), got) == (type(again), again)
    assert evaluate(excel, 'PyInfo("status")') == "running"
    assert evaluate(excel, 'PyInfo("executable")').lower() == expected["executable"].lower()
