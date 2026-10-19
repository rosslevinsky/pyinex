"""What a reload of a script PyLoad loaded changes (src/core/python/funcs.py; docs/architecture.md,
"Reloading a script"), in pure Python: a changed body changes nothing in Excel; a changed
signature, help or category unregisters and registers again, and the old registration's
calls give #VALUE! meanwhile; a removed function is unregistered; a renamed one is both; an
added one registers; a script that now fails changes nothing; and nothing changes with
reloading off.

load_functions gives the rows the loader registers; what a reload decided outside it, in a
call, waits in changes() for the loader. A row of a key alone asks the loader to remove that
function; removals come before registrations, so a name is cleared before it is registered
again.
"""

import os

import pytest

ONE = "import pyinex\n\n@pyinex.func\ndef {0}({1}):\n    return {2}\n"


def edit(path, text):
    """Rewrites a script and moves its time on, so a reload sees the change at once."""
    stat = path.stat()
    path.write_text(text, encoding="utf-8")
    os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns + 2_000_000_000))


@pytest.fixture
def loaded(core, script):
    """loaded(text) -> (path, {name: key}): a script loaded as PyLoad loads it, its functions
    registered."""

    def run(text, name="reloaded.py"):
        path = script(text, name=name)
        rows = core["load_functions"](str(path))
        for row in rows:
            core["registered"](row[0], True)
        return path, {row[1]: row[0] for row in rows}

    return run


def test_a_changed_body_needs_no_change_in_excel(core, loaded):
    path, keys = loaded(ONE.format("body", "a", "a + 1"))
    assert core["call_function"](keys["body"], (1.0,), True) == 2.0
    edit(path, ONE.format("body", "a", "a + 100"))
    assert core["call_function"](keys["body"], (1.0,), True) == 101.0
    assert core["changes"]() == []


@pytest.mark.parametrize(
    "after",
    [
        ONE.format("changed", "a, b", "a"),  # Its parameters.
        ONE.format("changed", "b", "b"),  # A parameter's name.
        "import pyinex\n\n@pyinex.func(category='Other')\ndef changed(a):\n    return a\n",
        "import pyinex\n\n@pyinex.func\ndef changed(a):\n    '''New help.'''\n    return a\n",
        "import pyinex\n\n@pyinex.func(name='CHANGED')\ndef changed(a):\n    return a\n",  # The name's case.
    ],
    ids=["parameters", "parameter name", "category", "help", "case"],
)
def test_a_changed_signature_is_registered_again(core, loaded, capfd, after):
    path, keys = loaded(ONE.format("changed", "a", "a"))
    edit(path, after)
    # The old registration's call reads the old arguments: the new code never sees them.
    assert core["call_function"](keys["changed"], (1.0,), True) is core["FAILED"]
    assert "Pyinex: changed is being registered again" in capfd.readouterr().err
    removal, registration = core["changes"]()
    assert removal == [str(path), keys["changed"]]
    assert registration[0] == str(path) and registration[2].casefold() == "changed"
    assert registration[1] != keys["changed"]
    core["registered"](registration[1], True)
    assert core["call_function"](registration[1], (5.0,) * registration[3], True) == 5.0


def test_a_removed_function_is_unregistered_and_its_name_freed(core, loaded, script, capfd):
    path, keys = loaded(ONE.format("kept", "a", "a") + "\n" + ONE.format("removed", "a", "a"))
    edit(path, ONE.format("kept", "a", "a"))
    assert core["call_function"](keys["removed"], (1.0,), True) is core["FAILED"]
    assert f"Pyinex: removed was removed from the script {path}" in capfd.readouterr().err
    assert core["changes"]() == [[str(path), keys["removed"]]]
    assert core["call_function"](keys["kept"], (1.0,), True) == 1.0
    # Another script may take the name now.
    other = script(ONE.format("removed", "a", "-a"), name="other.py")
    assert [row[1] for row in core["load_functions"](str(other))] == ["removed"]


def test_a_renamed_function_is_removed_and_registered(core, loaded):
    path, keys = loaded(ONE.format("before", "a", "a"))
    edit(path, ONE.format("after", "a", "a"))
    assert core["call_function"](keys["before"], (1.0,), True) is core["FAILED"]
    removal, registration = core["changes"]()
    assert removal == [str(path), keys["before"]]
    assert registration[2:5] == ["after", 1, "a"]


def test_an_added_function_registers(core, loaded):
    path, keys = loaded(ONE.format("first", "a", "a"))
    edit(path, ONE.format("first", "a", "a") + "\n" + ONE.format("added", "a, b", "b"))
    assert core["call_function"](keys["first"], (1.0,), True) == 1.0
    ((script, key, name, count, *_),) = core["changes"]()
    assert (script, name, count) == (str(path), "added", 2) and key not in keys.values()


def test_pyload_calculated_again_gives_the_removals_first(core, loaded):
    path, keys = loaded(ONE.format("same", "a", "a") + "\n" + ONE.format("moved", "a", "a"))
    edit(path, ONE.format("same", "a", "a") + "\n" + ONE.format("moved", "a, b", "a"))
    rows = core["load_functions"](str(path))
    assert rows[0] == [keys["moved"]]  # The removal, before the name is registered again.
    assert rows[1] == [-1]  # same: already registered.
    assert rows[2][1:3] == ["moved", 2]
    assert core["changes"]() == []  # PyLoad's own rows are not given twice.


@pytest.mark.parametrize("statement, shown", [("raise ValueError('now broken')", "ValueError: now broken"),
                                              ("raise SystemExit(7)", "SystemExit: 7"),
                                              ("raise KeyboardInterrupt", "KeyboardInterrupt")])
def test_a_script_that_now_fails_changes_nothing_until_it_is_fixed(core, loaded, capfd, statement, shown):
    working = ONE.format("steady", "a", "a")
    path, keys = loaded(working)
    edit(path, working + "\n" + ONE.format("newcomer", "a", "a") + "\n" + statement + "\n")
    assert core["call_function"](keys["steady"], (1.0,), True) is core["FAILED"]
    err = capfd.readouterr().err
    assert "failed to load:" in err and shown in err
    assert core["load_functions"](str(path)) is core["FAILED"]
    assert core["changes"]() == []  # Nothing new registers, and nothing is removed.
    # Fixed: recovered through a call of one of its functions ...
    edit(path, working + "\n" + ONE.format("newcomer", "a", "-a"))
    assert core["call_function"](keys["steady"], (2.0,), True) == 2.0
    ((_, _, name, *_),) = core["changes"]()
    assert name == "newcomer"


def test_a_script_that_now_fails_recovers_through_pyload(core, loaded):
    working = ONE.format("steady", "a", "a")
    path, keys = loaded(working)
    edit(path, working + "\nraise ValueError('now broken')\n")
    assert core["load_functions"](str(path)) is core["FAILED"]
    edit(path, working)
    assert core["load_functions"](str(path)) == [[-1]]
    assert core["call_function"](keys["steady"], (3.0,), True) == 3.0


def test_with_reloading_off_nothing_changes(core, loaded):
    path, keys = loaded(ONE.format("fixed", "a", "a + 1"))
    edit(path, ONE.format("other", "a, b", "a"))
    assert core["call_function"](keys["fixed"], (1.0,), False) == 2.0  # The loaded version.
    assert core["load_functions"](str(path), None, False) == [[-1]]
    assert core["changes"]() == []


def test_several_edits_before_the_loader_asks(core, loaded):
    path, keys = loaded(ONE.format("edited", "a", "a"))
    edit(path, ONE.format("edited", "a, b", "a"))
    assert core["call_function"](keys["edited"], (1.0,), True) is core["FAILED"]
    first_removal, first = core["changes"]()
    # Edited again before the loader has registered the first change: a PyLoad calculation
    # removes the registration waiting before it registers the next; the last one wins.
    edit(path, ONE.format("edited", "a, b, c", "c"))
    second_removal, second = core["load_functions"](str(path))
    assert first_removal[1] == keys["edited"] and second_removal == [first[1]]
    assert second[1:4] == ["edited", 3, "a,b,c"]
    core["registered"](second[0], True)
    assert core["call_function"](second[0], (1.0, 2.0, 3.0), True) == 3.0
    assert core["call_function"](first[1], (1.0, 2.0), True) is core["FAILED"]


def test_a_failed_registration_again_frees_the_name_for_the_next_pyload(core, loaded):
    path, keys = loaded(ONE.format("retried", "a", "a"))
    edit(path, ONE.format("retried", "a, b", "a"))
    rows = core["load_functions"](str(path))
    assert rows[0] == [keys["retried"]]
    core["registered"](rows[1][0], False)  # Excel refused it.
    (again,) = core["load_functions"](str(path))
    assert again[1:3] == ["retried", 2]


def test_a_script_only_pycall_uses_is_not_followed(core, script):
    path = script(ONE.format("unloaded", "a", "a"), name="pycall_only.py")
    assert core["call"](str(path), "unloaded", (1.0,), True, None) == 1.0
    edit(path, ONE.format("renamed", "a", "a"))
    assert core["call"](str(path), "renamed", (1.0,), True, None) == 1.0
    assert core["changes"]() == []


def test_the_scripts_name_is_as_pyload_was_given_it(core, script):
    path = script(ONE.format("near", "a", "a"), name="near_named.py")
    (row,) = core["load_functions"]("near_named.py", str(path.parent))
    core["registered"](row[0], True)
    edit(path, ONE.format("near", "a, b", "a"))
    assert core["call_function"](row[0], (1.0,), True) is core["FAILED"]
    assert [change[0] for change in core["changes"]()] == ["near_named.py", "near_named.py"]


def test_a_load_during_a_call_s_reload_follows_that_reload_s_changes(core, loaded):
    # A call reloads version 2, whose changes wait for the loader; a PyLoad during the call
    # loads version 3, whose rows remove version 2's registration. So the loader queues
    # changes() before the load's rows (src/loader/functions.cpp, load), one stream in the
    # order the reloads happened.
    path, keys = loaded(ONE.format("nested", "a", "1"))
    edit(path, ONE.format("nested", "a, b", "2"))
    assert core["call_function"](keys["nested"], (1.0,), True) is core["FAILED"]
    edit(path, ONE.format("nested", "a, b, c", "3"))
    rows = core["load_functions"](str(path))
    (_, first), (_, second, *_) = core["changes"]()
    assert first == keys["nested"]
    assert rows[0] == [second] and rows[1][0] != second


@pytest.fixture
def helper(monkeypatch):
    """A module the scripts import and that no reload runs again, so an object a script marks
    from it is the same object after the reload."""
    import functools
    import sys
    import types

    module = types.ModuleType("reload_helper")

    def target(x, y=0):
        return [type(x).__name__, x, y]

    module.target = target
    module.rate = functools.partial(target, y=20)
    monkeypatch.setitem(sys.modules, "reload_helper", module)
    return module


def test_a_reused_partial_is_placed_by_its_signature_after_the_reload(core, loaded, helper):
    # (x, *, y=20), then (*, x=10, y=20): the same names and count, so nothing changes in
    # Excel, but x is keyword-only now.
    marked = "import pyinex, reload_helper\n{0}rate = pyinex.func(reload_helper.rate)\n"
    path, keys = loaded(marked.format(""))
    assert core["call_function"](keys["rate"], (1.0,), True) == ["float", 1.0, 20]
    edit(path, marked.format("reload_helper.rate.keywords['x'] = 10\n"))
    assert core["call_function"](keys["rate"], (1.0,), True) == ["float", 1.0, 20]
    assert core["call_function"](keys["rate"], (), True) == ["int", 10, 20]
    assert core["changes"]() == []


def test_a_reused_function_is_converted_by_its_annotation_after_the_reload(core, loaded, helper):
    marked = "import pyinex, reload_helper\nreload_helper.target.__annotations__['x'] = {0}\ntarget = pyinex.func(reload_helper.target)\n"
    path, keys = loaded(marked.format("int"))
    assert core["call_function"](keys["target"], (3.0,), True) == ["int", 3, 0]
    edit(path, marked.format("str"))
    assert core["call_function"](keys["target"], ("a",), True) == ["str", "a", 0]
    assert core["changes"]() == []


@pytest.mark.parametrize("refused", ["31 parameters", "an unreadable signature"])
def test_a_function_whose_new_signature_is_refused_says_why_not_that_it_was_removed(core, loaded, capfd, refused):
    path, keys = loaded(ONE.format("F", "x", "x"))
    if refused == "31 parameters":
        edit(path, ONE.format("F", ", ".join(f"x{i}" for i in range(31)), "0"))
        why = "it has 31 parameters, and Excel passes at most 30"
    else:
        edit(path, "import pyinex\n\nclass Odd:\n    @property\n    def __signature__(self):\n        raise ValueError('unreadable')\n"
                   "    def __call__(self, x):\n        return x\n\nF = pyinex.func(Odd())\n")
        why = "its signature cannot be read"
    capfd.readouterr()
    rows = core["load_functions"](str(path))
    assert rows == [[keys["F"]]]  # Removed from Excel, its registration being refused.
    err = capfd.readouterr().err
    assert why in err and "was removed from the script" not in err
    assert core["call_function"](keys["F"], (1.0,), False) is core["FAILED"]
    err = capfd.readouterr().err
    assert why in err and "was removed" not in err
