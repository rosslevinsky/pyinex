"""Direct functions in the fake Excel (docs/architecture.md, "Direct functions"): the
functions of the scripts pyinex.ini lists, registered when the add-in opens and called by
their own names, with PyCall's value rules, error rule and call context.

Each value case calls the function by its name and through PyCall, from the same script,
and asserts the two results equal as well as the documented one.
"""

import time
from pathlib import Path

import pytest
from conftest import HELPERS
from hostdriver import MISSING, NA, NUM, VALUE, Err, Int

DIRECT = Path(__file__).resolve().parent / "scripts" / "direct.py"


@pytest.fixture
def direct(startup):
    return startup(str(DIRECT))


def wait_for(host, text, timeout=10):
    """The host's standard error once it holds text; output arrives on a thread."""
    deadline = time.monotonic() + timeout
    while text not in host.stderr and time.monotonic() < deadline:
        time.sleep(0.05)
    return host.stderr


def both(host, function, *args):
    """The function's result called by its own name, asserted equal to PyCall's."""
    by_name = host.call(function, *args)
    pycall = host.call("PyCall", str(DIRECT), function, *args)
    assert repr(by_name) == repr(pycall), (function, args)
    return by_name


# Group 1: every value rule in docs/conversion.md.


@pytest.mark.parametrize(
    "value, arrives",
    [
        (1.5, "float: 1.5"),
        (3.0, "float: 3.0"),
        (-0.0, "float: -0.0"),
        (1e308, "float: 1e+308"),
        (Int(7), "float: 7.0"),
        ("text", "str: 'text'"),
        ("", "str: ''"),
        ("é ✓ 😀", "str: 'é ✓ 😀'"),
        ("a\ud800b", "str: 'a\\ud800b'"),
        (True, "bool: True"),
        (False, "bool: False"),
        (None, "NoneType: None"),
        (Err(NA), "ExcelError: ExcelError('#N/A')"),
        (Err(99), "ExcelError: ExcelError('#VALUE!')"),
        (((1.0, 2.0, 3.0),), "tuple: ((1.0, 2.0, 3.0),)"),
        (((1.0,), (2.0,)), "tuple: ((1.0,), (2.0,))"),
        (((1.0, "a"), (True, None)), "tuple: ((1.0, 'a'), (True, None))"),
        (((5.0,),), "float: 5.0"),
    ],
    ids=lambda v: "" if not isinstance(v, str) else v[:20],
)
def test_a_value_arrives_as_pycall_passes_it(direct, value, arrives):
    assert both(direct, "describe", value) == arrives


def test_a_value_pyinex_does_not_recognize_is_reported_naming_the_function(direct):
    assert both(direct, "describe", Err(99)) == "ExcelError: ExcelError('#VALUE!')"
    assert "Pyinex: Describe: argument 1: " in wait_for(direct, "Pyinex: Describe: argument 1: ")


@pytest.mark.parametrize("value", [1.5, "text", "", "a\ud800b", True, Err(VALUE), ((1.0, "two"), (True, ""), (Err(7), 6.0))], ids=repr)
def test_a_value_goes_back_as_it_came(direct, value):
    assert both(direct, "echo", value) == value


def test_text_at_excels_limit_goes_back_whole(direct):
    assert both(direct, "echo", "x" * 32767) == "x" * 32767


def test_arguments_left_out(direct):
    assert both(direct, "echo") == ""  # None: an empty cell's text.
    assert both(direct, "pair", 1.0, MISSING) == Err(VALUE)  # Dropped at the end: b is missing.
    assert both(direct, "pair", MISSING, 2.0) == "(None, 2.0)"


@pytest.mark.parametrize(
    "name, shows",
    [
        ("none", ""),
        ("true", True),
        ("false", False),
        ("text", "text"),
        ("float", 2.5),
        ("int", 666.0),
        ("huge_int", float(1234567891011121314)),
        ("nan", Err(NUM)),
        ("inf", Err(NUM)),
        ("longest_text", "x" * 32767),
        ("error", Err(NA)),
        ("row", ((1.0, 2.0, 3.0),)),
        ("tuple", ((1.0, "two", True),)),
        ("column", ((1.0,), (2.0,), (3.0,))),
        ("ragged", ((1.0, 2.0, 3.0), (4.0, "", ""), (5.0, "", ""))),
        ("empty", ""),
        ("empty_rows", (("",), ("",))),
        ("empty_row_in_block", ((1.0, 2.0), ("", ""))),
        ("nan_in_block", ((1.0, 2.0), (3.0, Err(NUM)))),
    ],
    ids=lambda v: v if isinstance(v, str) and len(v) < 20 else "",
)
def test_a_result_shows_as_pycalls_does(direct, name, shows):
    assert both(direct, "give", name) == shows


@pytest.mark.parametrize(
    "name, said",
    [
        ("too_large_int", "int too large to be a number in Excel"),
        ("long_text", "str longer than 32,767 UTF-16 code units, Excel's limit for a cell"),
        ("emoji_too_long", "str longer than 32,767 UTF-16 code units, Excel's limit for a cell"),
        ("too_deep", "row 1, column 2: a list or tuple inside a row is nested too deeply for Excel"),
        ("dict", "dict cannot be returned to Excel"),
        ("set", "set cannot be returned to Excel"),
        ("complex", "complex cannot be returned to Excel"),
        ("object", "object cannot be returned to Excel"),
        ("bad_in_block", "row 2, column 2: dict cannot be returned to Excel"),
    ],
)
def test_a_result_that_cannot_be_returned_is_value_naming_the_function(direct, name, said):
    assert both(direct, "give", name) == Err(VALUE)
    assert f"Pyinex: give: {said}" in wait_for(direct, f"Pyinex: give: {said}")


@pytest.mark.parametrize("function, shown", [("raises", "ValueError: raised on purpose"), ("exits", "SystemExit: 3"),
                                             ("interrupts", "KeyboardInterrupt")])
def test_an_exception_is_value_with_its_traceback(direct, function, shown):
    assert direct.call(function) == Err(VALUE)
    err = wait_for(direct, shown)
    assert f"Pyinex: {function} in the script {DIRECT} raised an exception:" in err
    assert "Traceback (most recent call last)" in err and "<pyinex/" not in err
    assert direct.call("echo", "still working") == "still working"


def test_a_wrong_number_of_arguments_is_python_s_error(direct):
    # Two are registered, so Excel passes two, missing ones as missing: dropped at the end.
    assert both(direct, "two", 1.0) == Err(VALUE)
    assert "TypeError: two() missing 1 required positional argument: 'b'" in wait_for(direct, "TypeError: two()")


def test_a_result_is_freed_by_excel(direct):
    _, dll_free, faults = direct.call_detail("give", "ragged")
    assert dll_free and not faults


# Group 12: the call context.


@pytest.mark.parametrize(
    "caller, options, address",
    [
        (("[Book.xlsx]Other", 12, 3), (), "C12"),
        (("[Book.xlsx]Other", 12, 3), (True, True), "[Book.xlsx]Other!R12C3"),
        (("[Book.xlsx]Sheet1", 12, 3, 3, 3), (), "C12:E14"),  # A legacy array formula's range.
        (("[Book.xlsx]Sheet1", 4, 27, 1, 1, True), (False, True), "[Book.xlsx]Sheet1!AA4"),  # No sheet given.
    ],
)
def test_caller_is_the_calling_cell(direct, caller, options, address):
    direct.set_caller(*caller)
    assert both(direct, "caller", *options) == address


def test_sheet_is_the_calling_sheet(direct):
    direct.set_caller("[Bücher ✓.xlsx]Лист 😀", 2, 2)
    assert both(direct, "sheet") == "[Bücher ✓.xlsx]Лист 😀"


def test_caller_and_sheet_refuse_outside_a_cell(direct):
    direct.set_caller(None)
    assert both(direct, "attempt", "caller") == "RuntimeError"
    assert both(direct, "attempt", "sheet") == "RuntimeError"


def test_esc_pressed_kept_and_cleared(direct):
    assert direct.call("esc") is False
    direct.set_escape(True)
    assert direct.call("esc") is True
    assert direct.call("esc", False) is True
    assert direct.call("esc", True) is True
    assert direct.call("esc") is False


@pytest.mark.parametrize("name", ["caller", "sheet", "esc_pressed"])
def test_asking_excel_from_a_scripts_thread_is_refused(direct, name):
    assert both(direct, "on_a_thread", name) == "RuntimeError"


def test_a_scripts_thread_prints_between_calls(direct):
    assert direct.call("print_later", "printed from a thread") == "started"
    time.sleep(0.5)
    direct.call("echo")  # Lets the host handle its messages.
    deadline = time.monotonic() + 10
    while "printed from a thread" not in direct.stdout and time.monotonic() < deadline:
        time.sleep(0.05)
    assert "printed from a thread" in direct.stdout


def test_the_gil_is_released_between_calls(direct):
    assert direct.call("start_ticking") == "ticking"
    first = direct.call("ticked")
    time.sleep(1)
    assert direct.call("ticked") - first > 20


# The Function Arguments dialog.


def test_in_the_function_wizard_a_direct_function_is_blank_and_runs_nothing(direct, tmp_path):
    marker = tmp_path / "ran.txt"
    direct.set_wizard(True)
    assert direct.call("mark", str(marker)) == ""
    direct.set_wizard(False)
    assert not marker.exists()
    assert direct.call("mark", str(marker)) == "ran" and marker.exists()


# pyinex.ini's list.

GOOD = "import pyinex\n\n@pyinex.func\ndef {0}(x):\n    return '{0}: ' + str(x)\n"


def direct_names(host):
    return [r.name for r in host.registrations() if r.procedure.startswith("PyFunc")]


def test_several_scripts_register_in_order_relative_and_absolute(startup, tmp_path):
    elsewhere = tmp_path / "elsewhere" / "second.py"
    elsewhere.parent.mkdir()
    elsewhere.write_text(GOOD.format("second"), encoding="utf-8")
    host = startup(f"first.py; {elsewhere}", **{"first.py": GOOD.format("first")})
    assert direct_names(host) == ["first", "second"]
    assert host.call("first", 1.0) == "first: 1.0"
    assert host.call("second", "x") == "second: x"


@pytest.mark.parametrize(
    "broken, said",
    [
        (None, "Pyinex: cannot find the script"),
        ("def f(:\n", "SyntaxError"),
        ("import pyinex\n\n@pyinex.func\ndef lost(x):\n    return x\n\nraise ValueError('at import')\n", "ValueError: at import"),
        ("import pyinex\n\n@pyinex.func\ndef lost(x):\n    return x\n\nraise SystemExit(4)\n", "SystemExit: 4"),
        ("import pyinex\n\n@pyinex.func\ndef lost(x):\n    return x\n\nraise KeyboardInterrupt\n", "KeyboardInterrupt"),
    ],
)
def test_a_script_that_fails_registers_nothing_and_the_others_still_do(startup, broken, said):
    files = {"first.py": GOOD.format("first"), "last.py": GOOD.format("last")}
    if broken is not None:
        files["broken.py"] = broken
    host = startup("first.py; broken.py; last.py", **files)
    assert direct_names(host) == ["first", "last"]
    err = wait_for(host, said)
    assert said in err and "broken.py" in err
    assert host.call("last", 2.0) == "last: 2.0"


def test_a_refused_signature_is_reported_and_the_rest_register(startup):
    # 31 parameters: *args, refused here once, is accepted now (test_signatures.py).
    over = ", ".join(f"p{i}" for i in range(31))
    text = GOOD.format("before") + f"\n@pyinex.func\ndef refused({over}):\n    pass\n\n" + GOOD.format("after")
    host = startup("mixed.py", **{"mixed.py": text})
    assert direct_names(host) == ["before", "after"]
    assert "Pyinex: refused in the script" in wait_for(host, "is not registered")


@pytest.mark.parametrize("listed", ["", " ; "])
def test_an_empty_list_leaves_python_unstarted(startup, listed):
    host = startup(listed)
    assert direct_names(host) == []
    host.close()
    # Before Python starts, nothing pins the add-in, and the host unloads it.
    assert host.unload() == (False, False)


def test_a_listed_scripts_top_level_runs_outside_any_call(startup):
    from test_pyload import TOP_LEVEL

    host = startup("top.py", **{"top.py": TOP_LEVEL.format("top_level")})
    assert host.call("top_level") == "RuntimeError, RuntimeError, RuntimeError"


def test_a_listed_script_starts_python_at_open(startup):
    host = startup("first.py", **{"first.py": GOOD.format("first")})
    host.close()
    assert host.unload() == (True, True)  # Pinned: Python started while it opened.


def test_reopening_registers_the_same_functions_again(startup):
    host = startup("first.py", **{"first.py": GOOD.format("first")})
    assert host.close() == 1 and host.open() == 1
    assert direct_names(host)[-1] == "first"
    (latest,) = {r.procedure for r in host.registrations() if r.name == "first"}
    assert latest == "PyFunc0"  # The slot table started again.
    assert host.call("first", 3.0) == "first: 3.0"


def test_reopened_with_another_list_an_old_name_never_reaches_the_new_function(startup, tmp_path):
    # Closing clears each name it registered (docs/architecture.md, "Lifetime"): without it,
    # Excel would still send old to the slot new now has.
    host = startup("old.py", **{"old.py": GOOD.format("old"), "new.py": GOOD.format("new")})
    assert host.call("old", 1.0) == "old: 1.0"
    assert host.close() == 1
    (tmp_path / "addin" / "pyinex.ini").write_text("[python]\nscripts = new.py\n", encoding="utf-8")
    assert host.open() == 1
    (slot,) = {r.procedure for r in host.registrations() if r.name == "new"}
    assert slot == "PyFunc0"  # The slot old had.
    assert host.call("new", 2.0) == "new: 2.0"
    assert host.call("old", 1.0) == Err(VALUE)  # old's cell: an error, never new's value.
    assert "old" not in direct_names(host)


def test_a_last_help_text_at_excels_limit_is_registered_whole(startup):
    # Excel shows the last help text it is given short, so Pyinex gives it one more, empty.
    long = "h" * 255
    text = (f"import pyinex\n\n@pyinex.func(help='{long}')\ndef bare():\n    pass\n\n"
            f"@pyinex.func\ndef one(a):\n    '''Help.\n\n    Args:\n        a: {long}\n    '''\n")
    host = startup("full.py", **{"full.py": text})
    registered = {r.name: r for r in host.registrations()}
    assert (registered["bare"].help, registered["bare"].argument_help) == (long, ("",))
    assert (registered["one"].help, registered["one"].argument_help) == ("Help.", (long, ""))


def test_pycall_still_works_beside_direct_functions(direct):
    assert direct.call("PyCall", str(HELPERS), "echo", "alive") == "alive"
