"""Direct functions' names in the fake Excel (docs/architecture.md, "Names"): the
direct-functions test group 4. Names compare without regard to case; the first
registration wins; a name Excel or another add-in already uses, one of Pyinex's own, or one
Excel cannot take is refused, with a console message saying why, and the script's other
functions still register.

The fake Excel knows a few of Excel's own functions, SUM among them, and evaluates a name
alone as Excel does (tests/host/xllhost/xllhost.cpp); a question it cannot answer refuses
the name. One name in two scripts is
in test_pyload.py (test_a_name_another_script_registered_is_refused_naming_both).
"""

import pytest
from hostdriver import VALUE, Err
from test_direct_functions import wait_for
from test_pyload import direct_names
from test_reload_functions import edit

ONE = "import pyinex\n\n@pyinex.func(name={1!r})\ndef f{0}(a):\n    return {1!r} + ': ' + str(a)\n"


def named(*names):
    return "\n".join(ONE.format(i, n) for i, n in enumerate(names))


def load(host, path, *names):
    """A script declaring a function under each name, loaded by PyLoad; the queue run."""
    path.write_text(named(*names), encoding="utf-8")
    status = host.call("PyLoad", str(path))
    host.tick()
    return status


def test_two_names_that_differ_only_in_case(host, tmp_path):
    script = tmp_path / "twins.py"
    assert load(host, script, "Twin", "TWIN", "After") == "2 functions"
    assert direct_names(host) == ["Twin", "After"]
    assert f"Pyinex: TWIN in the script {script} is not registered: the script declares Twin already" in (
        wait_for(host, "declares Twin already"))
    # Excel compares names without regard to case, so either spelling calls the first.
    assert host.call("TWIN", 1.0) == host.call("twin", 1.0) == "Twin: 1.0"


def test_a_name_differing_only_in_case_in_another_script(host, tmp_path):
    first, second = tmp_path / "first.py", tmp_path / "second.py"
    load(host, first, "Pair")
    assert load(host, second, "pAIR") == "0 functions"
    assert f"Pyinex: pAIR in the script {second} is not registered: the script {first} registered that name first" in (
        wait_for(host, "registered that name first"))
    assert host.call("pair", 1.0) == "Pair: 1.0"


@pytest.mark.parametrize("name", ["SUM", "sum", "Concat"])
def test_a_name_of_one_of_excels_functions_is_refused(host, tmp_path, name):
    script = tmp_path / "excels.py"
    load(host, script, name, "AfterExcels")
    assert direct_names(host) == ["AfterExcels"]
    assert (f"Pyinex: cannot register {name} from the script {script}: Excel already knows that name, as one of its "
            "functions or another add-in's") in wait_for(host, "already knows that name")
    # Its name is free in Python again, so a later load says the same, not "registered".
    assert host.call("PyLoad", str(script)) == "2 functions"


def test_a_name_another_add_in_uses_is_refused(host, tmp_path):
    host.foreign("TheirFunction")
    script = tmp_path / "theirs.py"
    load(host, script, "theirfunction", "Ours")
    assert direct_names(host) == ["Ours"]
    assert "Pyinex: cannot register theirfunction from the script" in wait_for(host, "already knows that name")



def test_a_name_excel_cannot_answer_about_is_refused(host, tmp_path):
    host.fail_evaluates(1)  # The question about the first name.
    script = tmp_path / "unanswered.py"
    assert load(host, script, "Unanswered", "Answered") == "2 functions"
    assert direct_names(host) == ["Answered"]
    said = wait_for(host, "cannot register Unanswered")
    # Excel did not answer, which says nothing of whether the name is taken.
    assert f"Pyinex: cannot register Unanswered from the script {script}: Excel could not say whether it already knows that name" in said
    assert "already knows that name, as one of" not in said
    # Its name is free again, so the PyLoad cell's next calculation asks again.
    assert host.call("PyLoad", str(script)) == "2 functions"
    host.tick()
    assert direct_names(host) == ["Answered", "Unanswered"]

@pytest.mark.parametrize("name", ["PyCall", "PyConsole", "PyReloadOnChange", "PyInfo", "PyLoad", "pyload"])
def test_pyinex_s_own_names_are_refused(host, tmp_path, name):
    script = tmp_path / "own.py"
    load(host, script, name, "AfterOwn")
    assert direct_names(host) == ["AfterOwn"]
    assert f"Pyinex: {name} in the script {script} is not registered: it is one of Pyinex's own names" in (
        wait_for(host, "Pyinex's own names"))
    assert host.call("PyInfo", "status") == "running"  # Still Pyinex's.


@pytest.mark.parametrize(
    "name, why",
    [("A1", "it is a cell reference"), ("R1C1", "it is a cell reference"), ("x" * 256, "a name is at most 255 characters"),
     ("Bad Name", "a name is letters, digits, underscores and periods, starting with a letter or underscore")],
    ids=["A1", "R1C1", "too long", "a space"],
)
def test_names_excel_cannot_take_are_refused(host, tmp_path, name, why):
    script = tmp_path / "cannot.py"
    load(host, script, name, "AfterCannot")
    assert direct_names(host) == ["AfterCannot"]
    assert f"is not registered: Excel cannot take that name: {why}" in wait_for(host, "cannot take that name")


@pytest.mark.parametrize("name", ["Größe", "Disc.Rate", "_under", "x" * 255])
def test_unicode_and_other_names_excel_takes(host, tmp_path, name):
    load(host, tmp_path / "takes.py", name)
    assert direct_names(host) == [name]
    assert host.call(name, 1.0) == f"{name}: 1.0"


def test_a_name_a_reload_removed_is_free_for_another_script(host, tmp_path):
    first, second = tmp_path / "first.py", tmp_path / "second.py"
    load(host, first, "Handed")
    assert load(host, second, "Handed") == "0 functions"
    first.write_text("x = 1\n", encoding="utf-8")
    assert host.call("PyLoad", str(first)) == "0 functions"
    host.tick()
    assert host.call("Handed", 1.0) == Err(VALUE)
    assert load(host, second, "Handed") == "1 function"
    assert host.call("Handed", 2.0) == "Handed: 2.0"


@pytest.mark.parametrize("before", ["removed", "refused"])
def test_a_name_pyinex_no_longer_holds_is_asked_about_again(host, tmp_path, before):
    # Another add-in may take a name Pyinex removed, or one Excel refused it.
    script = tmp_path / "again.py"
    if before == "refused":
        host.fail_registrations(1)
    load(host, script, "Discount", "Other")
    if before == "removed":
        edit(script, named("Other"))
        assert host.call("PyLoad", str(script)) == "1 function"
        host.tick()
    assert direct_names(host) == ["Other"]
    host.foreign("Discount")
    edit(script, named("Discount", "Other"))
    assert host.call("PyLoad", str(script)) == "2 functions"
    host.tick()
    assert direct_names(host) == ["Other"]
    assert f"Pyinex: cannot register Discount from the script {script}: Excel already knows that name" in (
        wait_for(host, "already knows that name"))
