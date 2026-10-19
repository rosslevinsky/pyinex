"""Reloading a script PyLoad loaded, in the fake Excel (docs/architecture.md, "Reloading a
script"): the direct-functions test group 6, each kind of change with the cells
checked before and after, and each with PyReloadOnChange(FALSE); calls made before the queue
has run; a registration again that fails; several edits while work is queued; and a script
that now fails, recovered through PyLoad and through a call, with no restart.

A "cell" here is a call of the function by its name, as Excel makes when it calculates one.
"""

import os

import pytest
from hostdriver import VALUE, Err, Ref
from test_direct_functions import wait_for
from test_pyload import direct_names

ONE = "import pyinex\n\n@pyinex.func\ndef {0}({1}):\n    return {2}\n"


def edit(path, text):
    """Rewrites a script and moves its time on, so a reload sees the change at once."""
    stat = path.stat()
    path.write_text(text, encoding="utf-8")
    os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns + 2_000_000_000))


def loaded(host, path, text):
    """A script written and loaded by PyLoad, its functions registered."""
    path.write_text(text, encoding="utf-8")
    status = host.call("PyLoad", str(path))
    host.tick()
    return status


def types(host):
    """Each direct function's name and type string, as Excel has them registered."""
    return {r.name: r.type for r in host.registrations() if r.procedure.startswith("PyFunc")}


@pytest.fixture
def script(tmp_path):
    return tmp_path / "reloaded.py"


def test_a_changed_body_runs_at_the_next_call_with_no_call_into_excel(host, script):
    loaded(host, script, ONE.format("body", "a", "a + 1"))
    assert host.call("body", 1.0) == 2.0
    edit(script, ONE.format("body", "a", "a + 100"))
    host.excel_calls()
    assert host.call("body", 1.0) == 101.0
    assert host.excel_calls() == 0
    assert host.tick()[1:] == (0, 0)  # Nothing queued: Excel was not asked to run the queue.


def test_a_changed_signature_is_registered_again_and_meanwhile_gives_value(host, script):
    loaded(host, script, ONE.format("sig", "a", "'one: ' + str(a)"))
    assert types(host) == {"sig": "QQ"}
    host.set_busy(True)  # So the queue waits while the interval is checked.
    edit(script, ONE.format("sig", "a, b", "'two: ' + str(a) + str(b)"))
    # Called before the queue has run: Excel still passes one argument, and the new code
    # never sees it.
    assert host.call("sig", 1.0) == Err(VALUE)
    assert "Pyinex: sig is being registered again, since its script changed" in wait_for(host, "being registered again")
    assert types(host) == {"sig": "QQ"}
    host.set_busy(False)
    host.tick()
    assert types(host) == {"sig": "QQQ"}
    assert host.call("sig", 1.0, 2.0) == "two: 1.02.0"
    assert len([r for r in host.registrations() if r.procedure.startswith("PyFunc")]) == 1


def test_a_removed_function_is_unregistered_and_its_name_cleared(host, script):
    loaded(host, script, ONE.format("kept", "a", "a") + "\n" + ONE.format("gone", "a", "a"))
    assert host.call("gone", 1.0) == 1.0
    host.set_busy(True)
    edit(script, ONE.format("kept", "a", "a"))
    assert host.call("kept", 2.0) == 2.0  # The call that reloads the script.
    assert host.call("gone", 1.0) == Err(VALUE)
    assert f"Pyinex: gone was removed from the script {script}" in wait_for(host, "was removed")
    host.set_busy(False)
    host.tick()
    assert direct_names(host) == ["kept"]
    assert host.call("gone", 1.0) == Err(VALUE)  # Its name is cleared: it reaches nothing.


def test_a_renamed_function_is_removed_and_registered(host, script):
    loaded(host, script, ONE.format("before", "a", "'before'"))
    edit(script, ONE.format("after", "a", "'after'"))
    assert host.call("before", 1.0) == Err(VALUE)
    host.tick()
    assert direct_names(host) == ["after"]
    assert host.call("after", 1.0) == "after"
    assert host.call("before", 1.0) == Err(VALUE)


def test_an_added_function_registers(host, script):
    loaded(host, script, ONE.format("first", "a", "a"))
    edit(script, ONE.format("first", "a", "a") + "\n" + ONE.format("added", "a, b", "a * b"))
    assert host.call("first", 1.0) == 1.0
    host.tick()
    assert direct_names(host) == ["first", "added"]
    assert host.call("added", 2.0, 3.0) == 6.0


def test_pyload_calculated_again_does_the_same(host, script):
    loaded(host, script, ONE.format("viaload", "a", "a") + "\n" + ONE.format("dropped", "a", "a"))
    edit(script, ONE.format("viaload", "a, b", "b") + "\n" + ONE.format("new", "", "'new'"))
    assert host.call("PyLoad", str(script)) == "2 functions"
    host.tick()
    assert types(host) == {"viaload": "QQQ", "new": "Q"}
    assert host.call("viaload", 1.0, 2.0) == 2.0
    assert host.call("dropped", 1.0) == Err(VALUE)


CHANGES = {
    "body": ONE.format("steady", "a", "'changed'"),
    "signature": ONE.format("steady", "a, b", "'changed'"),
    "removed": "x = 1\n",
    "renamed": ONE.format("moved", "a", "'changed'"),
    "added": ONE.format("steady", "a", "'steady'") + "\n" + ONE.format("extra", "a", "'extra'"),
    "fails": ONE.format("steady", "a", "'changed'") + "\nraise ValueError('broken')\n",
}


@pytest.mark.parametrize("change", CHANGES)
def test_with_reloading_off_nothing_changes(host, script, change):
    loaded(host, script, ONE.format("steady", "a", "'steady'"))
    assert host.call("PyReloadOnChange", False) is False
    edit(script, CHANGES[change])
    assert host.call("steady", 1.0) == "steady"
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick()
    assert types(host) == {"steady": "QQ"}
    assert host.call("steady", 1.0) == "steady"


@pytest.mark.parametrize("statement, shown", [("raise ValueError('now broken')", "ValueError: now broken"),
                                              ("raise SystemExit(6)", "SystemExit: 6"),
                                              ("raise KeyboardInterrupt", "KeyboardInterrupt")],
                         ids=["ValueError", "SystemExit", "KeyboardInterrupt"])
@pytest.mark.parametrize("recovery", ["call", "PyLoad"])
def test_a_script_that_now_fails_registers_nothing_and_recovers_once_fixed(host, script, statement, shown, recovery):
    working = ONE.format("steady", "a", "'steady'")
    loaded(host, script, working)
    edit(script, working + "\n" + ONE.format("newcomer", "a", "'new'") + f"\n{statement}\n")
    assert host.call("steady", 1.0) == Err(VALUE)
    err = wait_for(host, shown)
    assert f"Pyinex: the script {script} failed to load:" in err and shown in err
    assert host.call("PyLoad", str(script)) == Err(VALUE)
    host.tick()
    assert direct_names(host) == ["steady"]  # Nothing new, and nothing removed.
    edit(script, working + "\n" + ONE.format("newcomer", "a", "'new'"))
    if recovery == "call":
        assert host.call("steady", 1.0) == "steady"
    else:
        assert host.call("PyLoad", str(script)) == "2 functions"
    host.tick()
    assert direct_names(host) == ["steady", "newcomer"]
    assert host.call("newcomer", 1.0) == "new" and host.call("steady", 1.0) == "steady"


def test_a_registration_again_that_fails_is_reported_and_tried_again_by_pyload(host, script):
    loaded(host, script, ONE.format("retried", "a", "'one'"))
    edit(script, ONE.format("retried", "a, b", "'two'"))
    host.fail_registrations(1)  # Before the work is queued: the queue may run at the next request.
    assert host.call("retried", 1.0) == Err(VALUE)
    host.tick()
    assert f"Pyinex: Excel refused to register retried from the script {script}" in wait_for(host, "refused")
    assert direct_names(host) == []
    assert host.call("retried", 1.0) == Err(VALUE)  # Its name was cleared before.
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick()
    assert host.call("retried", 1.0, 2.0) == "two"


def test_several_edits_while_work_is_queued(host, script):
    loaded(host, script, ONE.format("edited", "a", "'one'"))
    host.set_busy(True)  # Excel refuses to run the queue, so each change waits.
    edit(script, ONE.format("edited", "a, b", "'two'"))
    assert host.call("edited", 1.0) == Err(VALUE)
    edit(script, ONE.format("edited", "a, b, c", "'three'") + "\n" + ONE.format("also", "", "'also'"))
    assert host.call("PyLoad", str(script)) == "2 functions"
    edit(script, ONE.format("edited", "a, b, c", "'four'") + "\n" + ONE.format("also", "", "'also'"))
    host.tick(3)
    assert types(host) == {"edited": "QQ"}
    assert host.call("edited", 1.0) == Err(VALUE)
    host.set_busy(False)
    host.tick()
    assert types(host) == {"edited": "QQQQ", "also": "Q"}
    assert host.call("edited", 1.0, 2.0, 3.0) == "four"
    assert len([r for r in host.registrations() if r.procedure.startswith("PyFunc")]) == 2


# A function whose body calls PyLoad through the add-in's export, as a function calling into
# Excel could make Excel calculate a PyLoad cell during the call.
NESTING = """import ctypes, os, pyinex

@pyinex.func
def target({0}):
    return {1!r}

@pyinex.func
def nest(a):
    stat = os.stat({2!r})
    with open({2!r}, 'w', encoding='utf-8') as f:
        f.write({3!r})
    os.utime({2!r}, ns=(stat.st_atime_ns, stat.st_mtime_ns + 2_000_000_000))

    class XLOPER12(ctypes.Structure):
        _fields_ = [('str', ctypes.c_void_p), ('rest', ctypes.c_byte * 16), ('xltype', ctypes.c_uint32)]

    text = ctypes.create_unicode_buffer(chr(len({2!r})) + {2!r})
    ctypes.WinDLL({4!r}).PyLoad(ctypes.byref(XLOPER12(ctypes.addressof(text), xltype=2)))
    return 'nested'
"""


def test_a_pyload_during_a_call_that_reloaded_queues_after_that_reload(host, script, pyinex_xll):
    # The call reloads version 2, whose changes wait in Python until the call ends; its body
    # makes version 3 and calls PyLoad. Version 2's changes must be queued before version 3's,
    # or version 2's registration, already retired, ends up holding the name.
    third = NESTING.format("a, b, c", "three", str(script), "", str(pyinex_xll))
    loaded(host, script, NESTING.format("a", "one", str(script), "", str(pyinex_xll)))
    edit(script, NESTING.format("a, b", "two", str(script), third, str(pyinex_xll)))
    assert host.call("nest", 1.0) == "nested"
    host.tick()
    direct = [(r.name, r.type) for r in host.registrations() if r.procedure.startswith("PyFunc")]
    assert sorted(direct) == [("nest", "QQ"), ("target", "QQQQ")]
    assert host.call("target", 1.0, 2.0, 3.0) == "three"



# A function whose body calls PyLoad through the add-in's export for a script whose top level
# tries to ask Excel, which a script loaded by PyLoad may not, even within a call.
ASKING = """import ctypes, sys, pyinex

@pyinex.func
def Outer(a):
    class XLOPER12(ctypes.Structure):
        _fields_ = [('str', ctypes.c_void_p), ('rest', ctypes.c_byte * 16), ('xltype', ctypes.c_uint32)]

    text = ctypes.create_unicode_buffer(chr(len({0!r})) + {0!r})
    ctypes.WinDLL({1!r}).PyLoad(ctypes.byref(XLOPER12(ctypes.addressof(text), xltype=2)))
    return [m.ASKED for n, m in list(sys.modules.items()) if n.startswith('pyinex_inner_')][0]
"""


def test_a_pyload_within_a_call_runs_its_script_outside_any_call(host, script, pyinex_xll, tmp_path):
    inner = tmp_path / "inner.py"
    inner.write_text("import pyinex\ntry:\n    pyinex.caller()\n    ASKED = 'asked'\n"
                     "except RuntimeError:\n    ASKED = 'refused'\n", encoding="utf-8")
    loaded(host, script, ASKING.format(str(inner), str(pyinex_xll)))
    assert host.call("Outer", 1.0) == "refused"


# A function whose call Excel stops, then calls PyLoad, as ASKING does, for a script given by
# a path relative to the workbook's folder, which PyLoad would ask Excel for.
STOPPED = """import ctypes, pyinex

@pyinex.func
def Stopped(r: pyinex.Range):
    try:
        r.values()
    except RuntimeError:
        pass  # Excel stopped the call.

    class XLOPER12(ctypes.Structure):
        _fields_ = [('str', ctypes.c_void_p), ('rest', ctypes.c_byte * 16), ('xltype', ctypes.c_uint32)]

    text = ctypes.create_unicode_buffer(chr(len('inner.py')) + 'inner.py')
    ctypes.WinDLL({0!r}).PyLoad(ctypes.byref(XLOPER12(ctypes.addressof(text), xltype=2)))
    return 'loaded'
"""


def test_a_pyload_within_a_stopped_call_asks_excel_nothing(host, script, pyinex_xll):
    # Excel allows only xlFree once it has stopped a call (any other call is the host's
    # fault): the nested PyLoad leaves the stop in place, and finds no workbook folder.
    loaded(host, script, STOPPED.format(str(pyinex_xll)))
    host.stop("uncalced")
    assert host.call("Stopped", Ref("[Book1]Sheet1", "B2")) == Err(VALUE)
    host.stop("none")
    assert host.call("PyInfo", "status") == "running"
