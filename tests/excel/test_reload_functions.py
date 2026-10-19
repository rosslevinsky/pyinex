"""Reloading a script PyLoad loaded, in real Excel (docs/architecture.md, "Reloading a
script"): the direct-functions test group 6, each kind of change with the cells
checked before and after, and each with PyReloadOnChange(FALSE); a call made before the queue
has run; several edits while work is queued; and a script that now fails, recovered through
its PyLoad cell and through a call of one of its functions, with no restart.

A name registered stays registered for the rest of the session's Excel, so every test's
functions have names of their own, starting "Rf". Each change is a real save of the file;
the cell that calls the function depends on A1, so changing A1 recalculates it, as an
ordinary recalculation does.
"""

import time
from collections import Counter
from pathlib import Path

import pytest
from conftest import E, registered_types, shown
from test_pyload import when_registered

ONE = "import pyinex\n\n@pyinex.func\ndef {0}({1}):\n    return {2}\n"


@pytest.fixture
def xll(pytestconfig):
    return Path(pytestconfig.getoption("--excel")).resolve()


@pytest.fixture
def setting(app):
    """PyReloadOnChange for the test; turned back on after it."""

    def set_to(on):
        return app.Evaluate(f"PyReloadOnChange({'TRUE' if on else 'FALSE'})")

    try:
        yield set_to
    finally:
        set_to(True)


def loaded(app, book, name, text, probe):
    """A script written and loaded by PyLoad in A9, once probe (a formula) no longer gives
    #NAME?."""
    book.script(name, text)
    status = book.formula("A9", f'=PyLoad("{name}")')
    when_registered(app, probe)
    return status


def when(check, timeout=15.0):
    """Waits until check() is true, for the queue to run; returns its last answer."""
    deadline = time.monotonic() + timeout
    while not (done := check()) and time.monotonic() < deadline:
        time.sleep(0.1)
    return done


def test_a_changed_body_runs_at_the_next_recalculation(app, book, xll):
    loaded(app, book, "rf_body.py", ONE.format("RfBody", "a", "'old ' + str(a)"), "RfBody(0)")
    book.put("A1", 1)
    assert book.formula("B1", "=RfBody(A1)") == "old 1.0"
    before = registered_types(app, xll)
    book.script("rf_body.py", ONE.format("RfBody", "a", "'new body ' + str(a)"))
    book.put("A1", 2)
    assert book.value("B1") == "new body 2.0"
    time.sleep(1)
    assert registered_types(app, xll) == before  # Excel was told nothing.


def test_a_changed_signature_is_registered_again(app, book, xll, cleared):
    loaded(app, book, "rf_sig.py", ONE.format("RfSig", "a", "'one ' + str(a)"), "RfSig(0)")
    book.put("A1", 1)
    assert book.formula("B1", "=RfSig(A1)") == "one 1.0"
    before = Counter(registered_types(app, xll).values())
    book.script("rf_sig.py", ONE.format("RfSig", "a, b=''", "'two ' + str(a) + str(b)"))
    book.put("A1", 2)  # A call before the queue has run: the new code never sees its argument.
    assert book.value("B1") == E("#VALUE!")
    assert "Pyinex: RfSig is being registered again" in cleared(holding="being registered again").text
    assert when(lambda: shown(app.Evaluate("RfSig(1,2)")) == "two 1.02.0")
    after = Counter(registered_types(app, xll).values())
    assert (after - before, before - after) == (Counter(["QQQ"]), Counter(["QQ"]))  # Replaced, not added.
    # An ordinary recalculation of the cell reaches the new registration.
    book.put("A1", 3)
    assert book.value("B1") == "two 3.0"
    assert book.formula("B2", "=RfSig(4, 5)") == "two 4.05.0"


def test_a_removed_function_is_unregistered_and_its_name_cleared(app, book, xll, cleared):
    loaded(app, book, "rf_gone.py", ONE.format("RfKept", "a", "a") + "\n" + ONE.format("RfGone", "a", "'gone ' + str(a)"),
           "RfGone(0)")
    book.put("A1", 1)
    assert book.formula("B1", "=RfGone(A1)") == "gone 1.0"
    count = len(registered_types(app, xll))
    book.script("rf_gone.py", ONE.format("RfKept", "a", "a"))
    book.put("A1", 2)  # Its cell's call reloads the script.
    assert book.value("B1") == E("#VALUE!")
    assert "Pyinex: RfGone was removed from the script" in cleared(holding="was removed").text
    assert when(lambda: len(registered_types(app, xll)) == count - 1)
    assert book.formula("B2", "=RfGone(3)") == E("#VALUE!")  # Typed afresh.
    assert book.formula("B3", "=RfKept(3)") == 3.0


def test_a_renamed_function_is_removed_and_registered(app, book, xll):
    loaded(app, book, "rf_renamed.py", ONE.format("RfBefore", "a", "'before'"), "RfBefore(0)")
    book.put("A1", 1)
    assert book.formula("B1", "=RfBefore(A1)") == "before"
    count = len(registered_types(app, xll))
    book.script("rf_renamed.py", ONE.format("RfAfter", "a", "'after'"))
    book.put("A1", 2)
    assert book.value("B1") == E("#VALUE!")
    assert when_registered(app, "RfAfter(1)") == "after"
    assert when(lambda: len(registered_types(app, xll)) == count)  # One out, one in.
    assert book.formula("B2", "=RfBefore(1)") == E("#VALUE!")


def test_an_added_function_registers(app, book):
    loaded(app, book, "rf_added.py", ONE.format("RfFirst", "a", "a"), "RfFirst(0)")
    book.put("A1", 1)
    assert book.formula("B1", "=RfFirst(A1)") == 1.0
    book.script("rf_added.py", ONE.format("RfFirst", "a", "a") + "\n" + ONE.format("RfAdded", "a, b", "a * b"))
    book.put("A1", 2)
    assert book.value("B1") == 2.0
    assert when_registered(app, "RfAdded(2, 3)") == 6.0


CHANGES = {
    "body": ONE.format("{0}", "a", "'changed'"),
    "signature": ONE.format("{0}", "a, b", "'changed'"),
    "removed": "x = 1\n",
    "renamed": ONE.format("{0}Moved", "a", "'changed'"),
    "added": ONE.format("{0}", "a", "'steady'") + "\n" + ONE.format("{0}Extra", "a", "'extra'"),
    "fails": ONE.format("{0}", "a", "'changed'") + "\nraise ValueError('broken')\n",
}


@pytest.mark.parametrize("change", CHANGES)
def test_with_reloading_off_nothing_changes(app, book, xll, setting, change):
    name = "RfOff" + change.capitalize()
    loaded(app, book, f"rf_off_{change}.py", ONE.format(name, "a", "'steady'"), f"{name}(0)")
    book.put("A1", 1)
    assert book.formula("B1", f"={name}(A1)") == "steady"
    before = registered_types(app, xll)
    assert setting(False) is False
    book.script(f"rf_off_{change}.py", CHANGES[change].replace("{0}", name))
    book.put("A1", 2)
    assert book.value("B1") == "steady"
    assert book.formula("A9", f'=PyLoad("rf_off_{change}.py")') == "1 function"
    time.sleep(1)
    assert registered_types(app, xll) == before
    assert shown(app.Evaluate(f"{name}Extra(1)")) == E("#NAME?")


@pytest.mark.parametrize("statement, said", [("raise ValueError('now broken')", "ValueError: now broken"),
                                             ("raise SystemExit(8)", "SystemExit: 8"),
                                             ("raise KeyboardInterrupt", "KeyboardInterrupt")],
                         ids=["ValueError", "SystemExit", "KeyboardInterrupt"])
@pytest.mark.parametrize("recovery", ["call", "PyLoad"])
def test_a_script_that_now_fails_registers_nothing_and_recovers_once_fixed(app, book, cleared, statement, said, recovery):
    name = f"RfFails{said.split(':')[0]}{recovery}"
    script = f"rf_fails_{name.lower()}.py"
    working = ONE.format(name, "a", "'steady'")
    loaded(app, book, script, working, f"{name}(0)")
    book.put("A1", 1)
    assert book.formula("B1", f"={name}(A1)") == "steady"
    book.script(script, working + "\n" + ONE.format(f"{name}New", "a", "'new'") + f"\n{statement}\n")
    book.put("A1", 2)
    assert book.value("B1") == E("#VALUE!")
    text = cleared(holding=said).text
    assert "failed to load:" in text and said in text
    assert book.formula("A9", f'=PyLoad("{script}")') == E("#VALUE!")
    time.sleep(1)
    assert shown(app.Evaluate(f"{name}New(1)")) == E("#NAME?")  # Nothing new registered.
    assert shown(app.Evaluate(f"{name}(1)")) == E("#VALUE!")  # Nothing removed: it reaches Pyinex.
    book.script(script, working + "\n" + ONE.format(f"{name}New", "a", "'new'"))
    if recovery == "call":
        book.put("A1", 3)
        assert book.value("B1") == "steady"
    else:
        assert book.formula("A9", f'=PyLoad("{script}")') == "2 functions"
    assert when_registered(app, f"{name}New(1)") == "new"
    assert book.formula("B2", f"={name}(4)") == "steady"


EDITOR = """from pathlib import Path

def write(path, n, parameters, trigger):
    # Writes version n of rf_edits.py when trigger is 7, then returns trigger.
    if trigger == 7:
        Path(path).write_text(
            "import pyinex\\n\\n@pyinex.func\\ndef RfEdits(" + parameters + "):\\n    return 'v%d'\\n\\n" % n
            + "@pyinex.func\\ndef RfEditsOther(a):\\n    return a\\n", encoding="utf-8")
    return trigger
"""


def test_several_edits_in_one_calculation_while_work_is_queued(app, book, xll):
    # Excel cannot run the queue during a calculation. Each edit, written by a cell, is seen
    # by the next cell's call, so the second is made while the first change is still queued.
    loaded(app, book, "rf_edits.py",
           ONE.format("RfEdits", "a", "'v0'") + "\n" + ONE.format("RfEditsOther", "a", "a"), "RfEdits(0)")
    book.script("rf_editor.py", EDITOR)
    script = book.folder / "rf_edits.py"
    count = len(registered_types(app, xll))
    book.put("A1", 0)
    book.put("C1", f'=PyCall("rf_editor.py", "write", "{script}", 1, "a, b", A1)')
    book.put("C2", "=RfEditsOther(C1)")
    book.put("C3", f'=PyCall("rf_editor.py", "write", "{script}", 2, "a, b, c", C2)')
    book.put("C4", "=RfEditsOther(C3)")
    assert shown(app.Evaluate("RfEdits(1)")) == "v0"
    book.put("A1", 7)  # One calculation: C1 to C4 in order.
    assert book.values("C1:C4") == [[7.0], [7.0], [7.0], [7.0]]
    assert when(lambda: shown(app.Evaluate("RfEdits(1,2,3)")) == "v2")
    assert "QQQQ" in registered_types(app, xll).values()
    assert len(registered_types(app, xll)) == count  # The same functions, no slot left over.
    assert book.formula("B1", "=RfEdits(1,2,3)") == "v2"
