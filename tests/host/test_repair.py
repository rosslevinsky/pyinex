"""A workbook's PyLoad cells calculated as it opens, and the #NAME? repair
(docs/architecture.md, "When a workbook opens" and "The #NAME? repair"), as the fake Excel
shows them. Its stand-in for Excel's Application object gives the add-in the workbooks open,
their sheets and Excel's search of formulas, records each formula the add-in enters again,
and reports a workbook opened, or made from a template, to the handler the add-in connects
to Excel's application events. Which cells are chosen, which are left alone, the queued
work an open starts, and PyRepair.
"""

import pytest
from conftest import host_env
from hostdriver import Host
from test_direct_functions import wait_for
from test_pyload import direct_names


def script(folder, name, *functions):
    """A script marking a function of one argument, returning it doubled, under each name."""
    path = folder / name
    path.write_text("import pyinex\n" + "".join(f"\n@pyinex.func\ndef {f}(x):\n    return x * 2\n" for f in functions),
                    encoding="utf-8")
    return path


@pytest.fixture
def unopened(xllhost, pyinex_xll):
    """A fake Excel with the add-in not yet open, so that a test can fill workbooks first."""
    with Host(xllhost, pyinex_xll, env=host_env()) as h:
        yield h


def test_the_pyload_cells_of_workbooks_open_as_the_addin_opens_are_calculated(unopened, tmp_path):
    host = unopened
    script(tmp_path, "open.py", "Twice")
    host.set_folder(tmp_path)
    host.set_formula("[Book1]Sheet2", "B3", '=PyLoad("open.py")')
    host.set_formula("[Book1]Sheet1", "A1", "=Twice(2)")  # Typed before Twice was registered.
    host.set_formula("[Book1]Sheet1", "A2", "=Other(1)")  # Another name: left alone.
    host.set_cells("[Book1]Sheet1", "A3", "Twice(as text)")  # Found by the search, but not a formula.
    host.set_formula("[Book2]Sheet1", "C1", "=IFERROR(Twice(1),0)")  # The host does not calculate it.
    host.open()
    host.tick()
    assert direct_names(host) == ["Twice"]
    assert host.entered() == ([("[Book1]Sheet2", "$B$3", '=PyLoad("open.py")', "1 function"),
                               ("[Book1]Sheet1", "$A$1", "=Twice(2)", 4.0),
                               ("[Book2]Sheet1", "$C$1", "=IFERROR(Twice(1),0)", None)], 1)
    host.tick(3)
    assert host.entered() == ([], 0)  # Once only.


@pytest.mark.parametrize("new", [False, True], ids=["opened", "made from a template"])
def test_a_workbook_opened_later_is_calculated_when_excel_reports_it(host, tmp_path, new):
    script(tmp_path, "later.py", "Later")
    script(tmp_path, "typed.py", "Typed")
    host.set_folder(tmp_path)
    host.tick()  # The add-in connects to Excel's application events.
    assert host.entered() == ([], 0)
    host.set_formula("[Book1]Sheet1", "A1", '=PyLoad("typed.py")')  # Not in the workbook opened: left alone.
    host.set_formula("[Book3]Sheet1", "A1", '=PyLoad("later.py")')
    host.set_formula("[Book3]Sheet1", "A2", "=Later(5)")
    assert host.opened("Book3", new=new)
    host.tick()
    assert direct_names(host) == ["Later"]
    assert host.entered() == ([("[Book3]Sheet1", "$A$1", '=PyLoad("later.py")', "1 function"),
                               ("[Book3]Sheet1", "$A$2", "=Later(5)", 10.0)], 1)


def test_with_no_workbook_open_the_queue_waits_without_a_word(unopened):
    host = unopened
    host.set_workbook("none")
    host.open()
    assert host.tick(5) == (True, 0, 0)  # Waiting, but not through the running object table.
    assert "Pyinex:" not in host.stderr
    host.set_workbook("open")
    assert host.tick()[1] == 1
    assert host.opened("Book1")


def test_closing_disconnects_from_excels_events_and_reopening_connects_again(host):
    host.tick()
    assert host.opened("Book1")
    host.close()
    assert not host.opened("Book1")
    host.open()
    host.tick()
    assert host.opened("Book1")


def test_a_name_registered_again_is_left_alone(host, tmp_path):
    path = script(tmp_path, "again.py", "Again")
    host.set_formula("[Book1]Sheet1", "A1", "=Again(1)")
    assert host.call("PyLoad", str(path)) == "1 function"
    host.tick()
    assert host.entered() == ([("[Book1]Sheet1", "$A$1", "=Again(1)", 2.0)], 1)  # A new name.
    path.write_text("import pyinex\n\n@pyinex.func\ndef Again(x, y=3):\n    return x * y\n", encoding="utf-8")
    assert host.call("PyLoad", str(path)) == "1 function"  # The changed signature is queued.
    host.tick()
    assert [r.type for r in host.registrations() if r.name == "Again"] == ["QQQ"]
    assert host.entered() == ([], 0)  # Excel always knew the name: no repair.


def test_a_formula_that_cannot_be_entered_again_is_reported_and_pyrepair_enters_it(host, tmp_path):
    path = script(tmp_path, "locked.py", "Locked")
    host.set_formula("[Book1]Sheet1", "A1", "=Locked(1)")
    host.set_formula("[Book1]Sheet2", "A1", "=Locked(2)")
    host.lock("[Book1]Sheet1")
    assert host.call("PyLoad", str(path)) == "1 function"
    host.tick()
    assert host.entered() == ([("[Book1]Sheet2", "$A$1", "=Locked(2)", 4.0)], 1)  # The rest are entered.
    said = wait_for(host, "again (")
    assert "Pyinex: cannot enter the formula in [Book1]Sheet1!$A$1 again (Formula2 failed" in said
    assert "The cell is on a protected sheet." in said
    host.lock("[Book1]Sheet1", False)
    # Typed since the workbook opened, so calculated only by PyRepair, which also enters
    # every name a direct function holds.
    host.set_formula("[Book1]Sheet2", "B1", f'=PyLoad("{path}")')
    assert host.run("PyRepair") == 1
    assert host.entered() == ([("[Book1]Sheet2", "$B$1", f'=PyLoad("{path}")', "1 function"),
                               ("[Book1]Sheet1", "$A$1", "=Locked(1)", 2.0),
                               ("[Book1]Sheet2", "$A$1", "=Locked(2)", 4.0)], 1)


def test_only_formulas_that_call_the_name_are_entered_again(host, tmp_path):
    # The search finds the name followed by "(" anywhere in a formula; only a call of the
    # name itself is entered again, never one in a text, a quoted sheet name, a table's column
    # or a longer name.
    path = script(tmp_path, "twice.py", "RpTwice")
    sheet = "[Book1]Sheet1"
    host.set_formula(sheet, "A1", "=RpTwice(1)")
    host.set_formula(sheet, "A2", '=1+N("RpTwice(")')
    host.set_formula(sheet, "A3", "=MyRpTwice(1)")
    host.set_formula(sheet, "A4", "=SUM('RpTwice(x'!A1)")
    host.set_formula(sheet, "A5", "=Other.RpTwice(1)")
    host.set_formula(sheet, "A6", '=LEN("a""RpTwice(")+rptwice(2)')
    host.set_formula(sheet, "A7", "=SUM(Table1[[RpTwice(]])")  # A table's column so headed.
    host.set_formula(sheet, "A8", "=SUM(Table1['#OfItems])+RpTwice(3)")  # ' escapes # inside [].
    assert host.call("PyLoad", str(path)) == "1 function"
    host.tick()
    assert host.entered() == ([(sheet, "$A$1", "=RpTwice(1)", 2.0), (sheet, "$A$6", '=LEN("a""RpTwice(")+rptwice(2)', None),
                               (sheet, "$A$8", "=SUM(Table1['#OfItems])+RpTwice(3)", None)], 1)


def test_a_change_handler_that_searches_leaves_the_repair_its_own_cells(host, tmp_path):
    # A workbook's change handler may call Range.Find as each formula is entered, which
    # FindNext would then continue: every cell is found before any is entered.
    path = script(tmp_path, "handled.py", "Handled")
    sheet = "[Book1]Sheet1"
    for cell in ("A1", "A2", "A3"):
        host.set_formula(sheet, cell, f"=Handled({cell[1]})")
    host.set_formula(sheet, "B1", "=Unrelated(1)")
    host.set_formula(sheet, "B4", "=Unrelated(4)")
    host.set_handler("Unrelated(")
    assert host.call("PyLoad", str(path)) == "1 function"
    host.tick()
    assert host.entered() == ([(sheet, f"$A${n}", f"=Handled({n})", 2.0 * n) for n in (1, 2, 3)], 1)


def test_a_legacy_array_formula_is_entered_again_whole_once(host, tmp_path):
    # Excel refuses to change one cell of an array formula entered with Ctrl+Shift+Enter.
    path = script(tmp_path, "array.py", "InArray")
    host.set_array("[Book1]Sheet1", "B2:C3", "=InArray(1)")
    host.set_formula("[Book1]Sheet1", "A1", "=InArray(2)")
    assert host.call("PyLoad", str(path)) == "1 function"
    host.tick()
    assert host.entered() == ([("[Book1]Sheet1", "$A$1", "=InArray(2)", 4.0),
                               ("[Book1]Sheet1", "$B$2:$C$3", "=InArray(1)", None)], 1)
    assert "again (" not in host.stderr


def test_a_legacy_array_formula_too_long_to_enter_again_is_reported(host, tmp_path):
    # FormulaArray takes at most 255 characters: the cell is left alone and named.
    path = script(tmp_path, "long.py", "LongArray")
    long = "=LongArray(1)" + "+0" * 125
    host.set_array("[Book1]Sheet1", "B2:B3", long)
    assert host.call("PyLoad", str(path)) == "1 function"
    host.tick()
    assert host.entered() == ([], 1)
    assert ("Pyinex: cannot enter the formula in [Book1]Sheet1!$B$2:$B$3 again (a legacy array formula longer than 255 "
            "characters; enter it again with Ctrl+Shift+Enter)") in wait_for(host, "longer than 255")


def test_a_formula_that_cannot_be_read_is_reported_and_pyrepair_enters_it(host, tmp_path):
    path = script(tmp_path, "unread.py", "Unread")
    host.set_formula("[Book1]Sheet1", "A1", "=Unread(1)")
    host.set_formula("[Book1]Sheet2", "A1", "=Unread(2)")
    host.lock("[Book1]Sheet1", unreadable=True)
    assert host.call("PyLoad", str(path)) == "1 function"
    host.tick()
    assert host.entered() == ([("[Book1]Sheet2", "$A$1", "=Unread(2)", 4.0)], 1)
    said = wait_for(host, "again (")
    assert "Pyinex: cannot enter the formula in [Book1]Sheet1!$A$1 again (Formula failed" in said
    host.lock("[Book1]Sheet1", False)
    assert host.run("PyRepair") == 1
    assert host.entered() == ([("[Book1]Sheet1", "$A$1", "=Unread(1)", 2.0),
                               ("[Book1]Sheet2", "$A$1", "=Unread(2)", 4.0)], 1)


def test_pyrepair_refused_inside_a_worksheet_call_leaves_nothing_queued(host, tmp_path, pyinex_xll):
    # Run from inside a worksheet function, as a script could, PyRepair does nothing at all.
    path = script(tmp_path, "inside.py", "Inside")
    assert host.call("PyLoad", str(path)) == "1 function"
    host.tick()
    host.set_formula("[Book1]Sheet1", "A1", "=Inside(1)")
    host.set_formula("[Book1]Sheet1", "A2", f'=PyLoad("{path}")')
    assert host.entered() == ([], 1)
    caller = tmp_path / "caller.py"
    caller.write_text(f"import ctypes\n\ndef repair():\n    return ctypes.CDLL({str(pyinex_xll)!r}).PyRepair()\n",
                      encoding="utf-8")
    assert host.call("PyCall", str(caller), "repair") == 0.0
    said = "Pyinex: PyRepair cannot run inside a worksheet function"
    assert said in wait_for(host, said)
    assert host.tick(3)[1:] == (0, 0)
    assert host.entered() == ([], 0)


def test_a_cancelled_quit_with_nothing_registered_follows_workbooks_again(host, tmp_path):
    # Excel calls xlAutoClose before its save prompt; cancelled there, it carries on, and
    # with no direct function to register again the queue still connects to Excel's events.
    host.tick()
    assert host.close() == 1
    assert not host.opened("Book1")
    host.tick()
    script(tmp_path, "after.py", "After")
    host.set_folder(tmp_path)
    host.set_formula("[Book3]Sheet1", "A1", '=PyLoad("after.py")')
    assert host.opened("Book3")
    host.tick()
    assert direct_names(host) == ["After"]


def test_a_change_handler_that_runs_pyrepair_lets_the_repair_end(host, tmp_path):
    # PyRepair run again by a workbook's change handler, as each formula the repair enters
    # changes a cell, is part of the repair running: else each pass would ask for another.
    script(tmp_path, "rerun.py", "Rerun")
    host.set_folder(tmp_path)
    sheet = "[Book1]Sheet1"
    host.set_formula(sheet, "A1", '=PyLoad("rerun.py")')
    host.set_formula(sheet, "A2", "=Rerun(1)")
    assert host.call("PyLoad", str(tmp_path / "rerun.py")) == "1 function"
    host.tick()
    host.entered()
    host.set_change("run", "PyRepair")  # The host faults once the handler has run 25 times.
    assert host.run("PyRepair") == 1
    entered, calculations = host.entered()
    assert entered == [(sheet, "$A$1", '=PyLoad("rerun.py")', "1 function"), (sheet, "$A$2", "=Rerun(1)", 2.0)]
    assert calculations == 1


def test_pyrepair_excel_refuses_at_opening_is_reported(unopened):
    unopened.fail_name("PyRepair")
    unopened.open()
    assert "Pyinex: Excel refused to register the command PyRepair" in wait_for(unopened, "PyRepair")
