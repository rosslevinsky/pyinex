"""Direct functions' names in real Excel (docs/architecture.md, "Names"): the
direct-functions test group 4. Two names that differ only in case, one name in two
scripts, a clash with one of Excel's functions, with each of Pyinex's own and with another
add-in's, a name a workbook defines, names Excel cannot take, a Unicode name and the decorator's name=.

A name registered stays registered for the rest of the session's Excel, so every test's
functions have names of their own, starting "Nm".
"""

import time

import pytest
from conftest import E, shown
from test_pyload import when_registered

ONE = "import pyinex\n\n@pyinex.func(name={1!r})\ndef f{0}(a):\n    return {1!r} + ': ' + str(a)\n"


def named(*names):
    return "\n".join(ONE.format(i, n) for i, n in enumerate(names))


def load(app, book, script, *names, cell="A1"):
    """A script declaring a function under each name, loaded by PyLoad; waits for the last
    name to register."""
    book.script(script, named(*names))
    status = book.formula(cell, f'=PyLoad("{script}")')
    when_registered(app, f"{names[-1]}(1)")
    return status


def test_two_names_that_differ_only_in_case(app, book, cleared):
    assert load(app, book, "nm_twins.py", "NmTwin", "NMTWIN", "NmTwinAfter") == "2 functions"
    assert "Pyinex: NMTWIN in the script" in (text := cleared(holding="declares NmTwin already").text)
    assert "is not registered: the script declares NmTwin already" in text
    assert book.formula("B1", "=nmtwin(1)") == book.formula("B2", "=NMTWIN(1)") == "NmTwin: 1.0"


def test_one_name_in_two_scripts_the_first_wins(app, book, cleared):
    load(app, book, "nm_first.py", "NmShared")
    assert load(app, book, "nm_second.py", "nmshared", "NmSecondOwn", cell="A2") == "1 function"
    text = cleared(holding="registered that name first").text
    assert "Pyinex: nmshared in the script" in text and "nm_second.py" in text and "nm_first.py" in text
    assert book.formula("B1", "=NmShared(1)") == "NmShared: 1.0"


@pytest.mark.parametrize("name", ["SUM", "Concat", "XLOOKUP", "NOW", "Ünïcødé"])
def test_a_name_of_one_of_excels_functions_is_refused(app, book, cleared, name):
    # Excel would take the registration, and go on calling its own function. (It reads
    # Ünïcødé as its own UNICODE.)
    load(app, book, "nm_excels.py", name, "NmAfterExcels")
    text = cleared(holding="already knows that name").text
    assert f"Pyinex: cannot register {name} from the script nm_excels.py: Excel already knows that name" in text
    assert book.formula("B1", "=SUM(1,2)") == 3.0


def test_a_name_another_add_in_registered_is_refused(app, book, cleared):
    # Another add-in, as Excel sees one: a function of another DLL registered under a name.
    registered = app.ExecuteExcel4Macro('REGISTER("kernel32.dll","GetTickCount","J","NmTheirTicks",,1)')
    assert isinstance(registered, float), registered
    try:
        assert isinstance(shown(app.Evaluate("NmTheirTicks()")), float)
        load(app, book, "nm_theirs.py", "nmtheirticks", "NmAfterTheirs")
        assert "Pyinex: cannot register nmtheirticks from the script nm_theirs.py" in (
            cleared(holding="already knows that name").text)
        assert isinstance(shown(app.Evaluate("NmTheirTicks()")), float)  # Still theirs.
    finally:
        app.ExecuteExcel4Macro(f"UNREGISTER({registered})")


def test_a_name_pyinex_removed_and_another_add_in_took_is_refused(app, book, cleared):
    # A removed name still evaluates to an ID; another add-in's registration gives another.
    load(app, book, "nm_handed.py", "NmHanded", "NmHandedKept")
    # The queue removes before it registers, so once NmHandedNext registers, NmHanded is gone.
    book.script("nm_handed.py", named("NmHandedKept", "NmHandedNext"))
    assert book.formula("A1", '=PyLoad("nm_handed.py")') == "2 functions"
    assert when_registered(app, "NmHandedNext(1)") == "NmHandedNext: 1.0"
    registered = app.ExecuteExcel4Macro('REGISTER("kernel32.dll","GetTickCount","J","NmHanded",,1)')
    assert isinstance(registered, float), registered
    try:
        book.script("nm_handed.py", named("NmHanded", "NmHandedKept", "NmHandedNext"))
        book.formula("A1", '=PyLoad("nm_handed.py")')
        assert "Pyinex: cannot register NmHanded from the script nm_handed.py" in (
            cleared(holding="already knows that name").text)
        assert isinstance(shown(app.Evaluate("NmHanded()")), float)  # Still theirs.
    finally:
        app.ExecuteExcel4Macro(f"UNREGISTER({registered})")


@pytest.mark.parametrize("name", ["PyCall", "PyConsole", "PyReloadOnChange", "PyInfo", "PyLoad", "PyRepair"])
def test_pyinex_s_own_names_are_refused(app, book, cleared, name):
    load(app, book, "nm_own.py", name, "NmAfterOwn")
    assert f"Pyinex: {name} in the script" in (text := cleared(holding="Pyinex's own names").text)
    assert "is not registered: it is one of Pyinex's own names" in text
    assert app.Evaluate('PyInfo("status")') == "running"


@pytest.mark.parametrize(
    "name, why",
    [("A1", "it is a cell reference"), ("R1C1", "it is a cell reference"), ("R", "it is a cell reference"),
     ("TRUE", "it is a logical value"), ("x" * 256, "a name is at most 255 characters"),
     ("Nm Bad", "a name is letters, digits, underscores and periods, starting with a letter or underscore"),
     ("1Nm", "a name is letters, digits, underscores and periods, starting with a letter or underscore")],
    ids=["A1", "R1C1", "R", "TRUE", "too long", "a space", "digit first"],
)
def test_names_excel_cannot_take_are_refused(app, book, cleared, name, why):
    load(app, book, "nm_cannot.py", name, "NmAfterCannot")
    assert f"is not registered: Excel cannot take that name: {why}" in cleared(holding="cannot take that name").text


@pytest.mark.parametrize("name", ["NmGröße", "Nm.Dotted", "_nm_under", "XFE1", "Nm" + "x" * 253])
def test_names_excel_takes_are_registered(app, book, cleared, name):
    load(app, book, "nm_takes.py", name)
    # Evaluate takes at most 255 characters, so load's wait cannot see the longest name
    # register: the console's word is waited for before the formula is typed.
    assert f"Pyinex: registered {name} from" in cleared(holding=f"registered {name} from", timeout=15).text
    assert book.formula("B1", f"={name}(1)") == f"{name}: 1.0"


def test_a_name_typed_in_formulas_before_it_registers_is_free(app, book):
    # Excel remembers how a formula spelled an unknown name, but evaluates it alone to
    # #NAME? still, so it is not taken for a name Excel knows.
    assert book.formula("B1", "=NmTypedFirst(1)") == E("#NAME?")
    assert shown(app.Evaluate("NmTypedFirst(2)")) == E("#NAME?")
    load(app, book, "nm_typed.py", "NmTypedFirst")
    assert book.formula("B2", "=NmTypedFirst(3)") == "NmTypedFirst: 3.0"


def test_a_name_a_workbook_defines_is_free(app, book, cleared, workbook):
    # Excel evaluates a name alone as one of its own or an add-in's: a workbook's name is
    # not seen, so a function takes it, and other workbooks call the function.
    book.workbook.Names.Add("NmDefined", "=1")
    assert book.formula("B1", "=NmDefined") == 1.0
    load(app, book, "nm_defined.py", "NmDefined", "NmDefinedAfter")
    text = cleared(holding="registered NmDefinedAfter").text
    assert "registered NmDefined from the script" in text, text
    assert "already knows" not in text, text
    assert workbook().formula("A1", "=NmDefined(2)") == "NmDefined: 2.0"


def test_the_decorators_name_is_the_name_in_excel(app, book):
    book.script("nm_decorated.py", "import pyinex\n\n@pyinex.func(name='NmShown')\ndef hidden(a):\n    return a * 2\n")
    assert book.formula("A1", '=PyLoad("nm_decorated.py")') == "1 function"
    assert when_registered(app, "NmShown(2)") == 4.0
    time.sleep(0.5)
    assert shown(app.Evaluate("hidden(2)")) == E("#NAME?")
