"""xlAutoOpen registers exactly the five documented functions and the two commands, PyRepair
and the queue's (docs/architecture.md, "Excel functions" and "The queue"), through the table-driven helper,
and then the direct functions of the scripts pyinex.ini lists (docs/architecture.md, "Direct
functions")."""

from test_direct_functions import DIRECT

ARGUMENTS = "file,function," + ",".join(f"arg{i}" for i in range(1, 31))


def test_registers_the_documented_functions(host):
    registered = {r.procedure if r.procedure == "PyinexQueue" else r.name: r for r in host.registrations()}
    assert set(registered) == {"PyCall", "PyConsole", "PyReloadOnChange", "PyInfo", "PyLoad", "PyinexQueue", "PyRepair"}
    expected = {
        "PyCall": ("Q" * 33, ARGUMENTS, 32),
        "PyConsole": ("QQQQQQQ", "show,clear,x,y,width,height", 6),
        "PyReloadOnChange": ("QQ", "on", 1),
        "PyInfo": ("QQ", "item", 1),
        "PyLoad": ("QQ", "file", 1),
    }
    # The queue's command, which no formula can use, under a name of its own (tests/host/test_queue.py).
    assert registered["PyinexQueue"].type == "J"
    # PyRepair, a command a user runs (tests/host/test_repair.py).
    assert (registered["PyRepair"].procedure, registered["PyRepair"].type) == ("PyRepair", "J")
    for name, (type_, arguments, help_count) in expected.items():
        r = registered[name]
        assert (r.procedure, r.type, r.arguments) == (name, type_, arguments), name
        assert r.category == "Pyinex", name
        assert r.help, name
        # Then one more, empty: Excel shows the last help text it is given short.
        assert len(r.argument_help) == help_count + 1 and all(r.argument_help[:-1]), name
        assert r.argument_help[-1] == "", name


def test_no_function_is_thread_safe(host):
    # Pyinex runs only on Excel's main thread (docs/architecture.md, "Threading").
    assert all("$" not in r.type for r in host.registrations())


def test_pycall_argument_text_fits_excels_limit(host):
    # Excel truncates a function's argument text beyond 255 characters.
    (pycall,) = [r for r in host.registrations() if r.name == "PyCall"]
    assert len(pycall.arguments) <= 255
    assert all(len(h) <= 255 for h in pycall.argument_help)


def test_registers_the_direct_functions_pyinex_ini_lists(startup):
    host = startup(str(DIRECT))
    registered = host.registrations()
    assert [r.procedure for r in registered[:7]] == ["PyCall", "PyConsole", "PyReloadOnChange", "PyInfo", "PyLoad",
                                                     "PyinexQueue", "PyRepair"]
    direct = {r.name: r for r in registered[7:]}
    # Every function direct.py marks, in its order, each in the next slot, and nothing else.
    assert list(direct) == ["echo", "Describe", "pair", "give", "raises", "exits", "interrupts", "two", "caller",
                            "sheet", "esc", "attempt", "on_a_thread", "print_later", "start_ticking", "ticked", "mark"]
    assert [r.procedure for r in direct.values()] == [f"PyFunc{i}" for i in range(len(direct))]
    echo, describe, caller = direct["echo"], direct["Describe"], direct["caller"]
    assert (echo.type, echo.arguments, echo.category, echo.help, echo.argument_help) == (
        "QQ", "value", "Pyinex", "Returns its argument.", ("Anything.", ""))
    assert (describe.type, describe.arguments, describe.category, describe.help, describe.argument_help) == (
        "QQ", "x", "Pyinex tests", "What arrived.", ("", ""))
    assert (caller.type, caller.arguments, caller.argument_help) == ("QQQ", "r1c1,full", ("", "", ""))
    # One help text more than the arguments, empty: Excel shows the last it is given short.
    assert (direct["sheet"].help, direct["sheet"].argument_help) == ("", ("",))
    assert (direct["mark"].help, direct["mark"].argument_help) == ("Writes the file named, so a test can tell the function ran.", ("", ""))
    assert direct["sheet"].type == "Q" and direct["sheet"].arguments == ""


def test_no_direct_function_is_thread_safe(startup):
    host = startup(str(DIRECT))
    assert len(host.registrations()) > 7
    assert all("$" not in r.type for r in host.registrations())


def test_a_pyinex_range_argument_is_registered_as_a_reference(startup):
    # docs/architecture.md, "pyinex.Range": U where an argument takes a reference, Q elsewhere.
    host = startup("ranged.py", **{"ranged.py": (
        "import pyinex\n\n"
        "@pyinex.func\ndef where(a, r: pyinex.Range, n: int = 0):\n    return 1\n\n"
        "@pyinex.func\ndef spread(first: pyinex.Range, *more: pyinex.Range):\n    return 1\n"
    )})
    direct = {r.name: r.type for r in host.registrations()[7:]}
    assert direct == {"where": "QQUQ", "spread": "Q" + "U" * 30}
