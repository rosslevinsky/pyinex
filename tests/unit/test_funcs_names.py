"""The rules for direct functions' names (src/core/python/funcs.py; docs/architecture.md,
"Names"), in pure Python: compared without regard to case, the first registration wins, and
Pyinex's own names and names Excel cannot take are refused with a console message saying
why. (A name Excel or another add-in already uses is found by asking Excel, in the loader:
tests/host/test_names.py.)
"""

import pytest

ONE = "import pyinex\n\n@pyinex.func{1}\ndef f{0}(a):\n    return a\n"


def named(*names):
    """A script declaring a function under each name, in order."""
    return "\n".join(ONE.format(i, f"(name={n!r})") for i, n in enumerate(names))


def test_two_names_that_differ_only_in_case_in_one_script(load, capfd):
    rows = load(named("Twin", "TWIN"))
    assert [row[1] for row in rows] == ["Twin"]
    assert "Pyinex: TWIN in the script" in (err := capfd.readouterr().err)
    assert "the script declares Twin already" in err


def test_two_names_that_differ_only_in_case_in_two_scripts(core, script, capfd):
    first, second = script(named("Pair"), name="pair_one.py"), script(named("pAIR"), name="pair_two.py")
    assert [row[1] for row in core["load_functions"](str(first))] == ["Pair"]
    assert core["load_functions"](str(second)) == []
    assert f"pAIR in the script {second} is not registered: the script {first} registered that name first" in (
        capfd.readouterr().err)


@pytest.mark.parametrize("name", ["PyCall", "pyconsole", "PYRELOADONCHANGE", "PyInfo", "PyLoad", "PyRepair", "xlAutoRemove",
                                  "PyinexQueue140700000000000"])
def test_pyinex_s_own_names_are_refused(load, capfd, name):
    assert load(named(name)) == []
    assert f"{name} in the script" in (err := capfd.readouterr().err)
    assert "is not registered: it is one of Pyinex's own names" in err


@pytest.mark.parametrize(
    "name, why",
    [
        ("A1", "it is a cell reference"),
        ("xfd1048576", "it is a cell reference"),
        ("R1C1", "it is a cell reference"),
        ("R", "it is a cell reference"),
        ("c", "it is a cell reference"),
        ("RC5", "it is a cell reference"),
        ("TRUE", "it is a logical value"),
        ("Bad Name", "a name is letters, digits, underscores and periods, starting with a letter or underscore"),
        ("1digit", "a name is letters, digits, underscores and periods, starting with a letter or underscore"),
        ("x-y", "a name is letters, digits, underscores and periods, starting with a letter or underscore"),
        ("x" * 256, "a name is at most 255 characters"),
    ],
    ids=["A1", "last cell", "R1C1", "R", "C", "RC5", "TRUE", "space", "digit first", "hyphen", "too long"],
)
def test_names_excel_cannot_take_are_refused(load, capfd, name, why):
    rows = load(named(name, "AfterIt"))
    assert [row[1] for row in rows] == ["AfterIt"]
    assert f"is not registered: Excel cannot take that name: {why}" in capfd.readouterr().err


@pytest.mark.parametrize("name", ["XFE1", "AB", "Größe", "_under", "Disc.Rate", "x" * 255])
def test_names_excel_takes_are_registered(load, name):
    assert [row[1] for row in load(named(name))] == [name]
