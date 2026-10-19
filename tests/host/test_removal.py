"""Removing direct functions, and the pool's limits, in the fake Excel (docs/architecture.md,
"Removing a function" and "A full pool"): what removal leaves, removal with each step of the
name-clearing recipe failing (the slot is then kept out of use, and its old name never
reaches another function), the direct-functions test group 9 (the pool filled
exactly, one function more, many cycles of registering and removing), and group 11's
fake-Excel half: thousands of calls and hundreds of reloads with handles flat. (The fake
Excel checks every result's C runtime heap blocks as it frees it.)
"""

import pytest
from hostdriver import VALUE, Err
from test_direct_functions import wait_for
from test_pyload import direct_names
from test_reload_functions import ONE, edit, loaded, types

POOL = 1000


def slots(host):
    """The numbered exports registered, by name."""
    return {r.name: r.procedure for r in host.registrations() if r.procedure.startswith("PyFunc")}


def test_what_removal_leaves(host, tmp_path):
    script = tmp_path / "removed.py"
    loaded(host, script, ONE.format("leaving", "a", "'leaving'"))
    assert slots(host) == {"leaving": "PyFunc0"}
    script.write_text("x = 1\n", encoding="utf-8")
    assert host.call("PyLoad", str(script)) == "0 functions"
    host.tick()
    # Gone from the registration table; a cell that calls it, and one typed afresh, #VALUE!.
    assert slots(host) == {}
    assert host.call("leaving", 1.0) == Err(VALUE)
    # Its slot is free: the next function takes it, and the old name does not reach it.
    loaded(host, tmp_path / "next.py", ONE.format("arriving", "a", "'arriving'"))
    assert slots(host) == {"arriving": "PyFunc0"}
    assert host.call("leaving", 1.0) == Err(VALUE)
    assert host.call("arriving", 1.0) == "arriving"


@pytest.mark.parametrize(
    "fail",
    [
        lambda h: h.fail_unregisters(1),  # Unregistering the function.
        lambda h: h.fail_commands(1),  # Registering the stand-in command under its name.
        lambda h: h.fail_set_names(1),  # Deleting the name.
        lambda h: h.fail_unregisters(1, skip=1),  # Unregistering the stand-in command.
    ],
    ids=["unregister", "command", "set name", "unregister command"],
)
def test_a_removal_that_fails_at_any_step_keeps_the_slot_out_of_use(host, tmp_path, fail):
    script = tmp_path / "failing.py"
    loaded(host, script, ONE.format("stuck", "a", "'stuck'"))
    script.write_text("x = 1\n", encoding="utf-8")
    fail(host)  # Before the removal is queued: the queue's timer may run it at the next request.
    assert host.call("PyLoad", str(script)) == "0 functions"
    host.tick()
    assert "Pyinex: cannot remove stuck from Excel; until Excel restarts it gives #VALUE!" in wait_for(host, "cannot remove")
    assert host.call("stuck", 1.0) == Err(VALUE)
    # The next function takes another slot, and the old name never reaches it.
    loaded(host, tmp_path / "next.py", ONE.format("fresh", "a", "'fresh'"))
    assert slots(host)["fresh"] == "PyFunc1"
    assert host.call("stuck", 1.0) == Err(VALUE)
    assert host.call("fresh", 1.0) == "fresh"
    # Excel still makes of the name what it did when the removal failed, so it registers again.
    edit(script, ONE.format("stuck", "a", "'again'"))
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick()
    assert host.call("stuck", 1.0) == "again"
    # Removed again, this time cleanly, and then taken by another add-in: the slot kept out of
    # use does not let Pyinex register over that add-in's name.
    script.write_text("x = 1\n", encoding="utf-8")
    assert host.call("PyLoad", str(script)) == "0 functions"
    host.tick()
    host.foreign("stuck")
    edit(script, ONE.format("stuck", "a", "'over'"))
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick()
    assert f"Pyinex: cannot register stuck from the script {script}: Excel already knows that name" in (
        wait_for(host, "already knows that name"))
    # No slot of Pyinex's took the name: at most the one kept out of use still shows it.
    assert {r.procedure for r in host.registrations() if r.name == "stuck" and r.procedure.startswith("PyFunc")} <= {"PyFunc0"}


def many(count, prefix="Pool_"):
    """A script declaring count functions, named prefix0, prefix1, ..."""
    return ("import pyinex\n\n"
            f"for i in range({count}):\n"
            f"    globals()['f%d' % i] = pyinex.func(name='{prefix}%d' % i)(lambda x, i=i: i)\n")


def test_the_pool_filled_exactly(host, tmp_path):
    assert loaded(host, tmp_path / "full.py", many(POOL)) == f"{POOL} functions"
    assert len(slots(host)) == POOL
    assert host.call("Pool_0", 1.0) == 0.0 and host.call(f"Pool_{POOL - 1}", 1.0) == POOL - 1
    assert "slots are taken" not in host.stderr


def test_one_function_more_than_the_pool_is_reported_and_the_rest_work(host, tmp_path):
    script = tmp_path / "over.py"
    assert loaded(host, script, many(POOL + 1)) == f"{POOL + 1} functions"
    assert len(slots(host)) == POOL
    assert (f"Pyinex: cannot register Pool_{POOL} from the script {script}: all {POOL} of Pyinex's function slots are taken"
            in wait_for(host, "slots are taken"))
    assert host.stderr.count("slots are taken") == 1
    assert host.call("Pool_0", 1.0) == 0.0 and host.call(f"Pool_{POOL - 1}", 1.0) == POOL - 1


def test_many_cycles_of_registering_and_removing_reuse_one_slot(host, tmp_path):
    script = tmp_path / "cycled.py"
    loaded(host, script, ONE.format("Cycle0", "a", "'Cycle0'"))
    for n in range(1, 200):
        edit(script, ONE.format(f"Cycle{n}", "a", f"'Cycle{n}'"))
        assert host.call("PyLoad", str(script)) == "1 function"
        host.tick()
        assert slots(host) == {f"Cycle{n}": "PyFunc0"}
        assert host.call(f"Cycle{n}", 1.0) == f"Cycle{n}"
        assert host.call(f"Cycle{n - 1}", 1.0) == Err(VALUE)
    assert host.call("Cycle0", 1.0) == Err(VALUE)


def test_thousands_of_calls_and_hundreds_of_reloads_leave_handles_flat(host, tmp_path):
    script = tmp_path / "scaled.py"
    loaded(host, script, ONE.format("Scaled", "a", "a"))

    def cycles(n):
        for i in range(n):
            # Each reload changes the signature, so it is registered again.
            edit(script, ONE.format("Scaled", "a" if i % 2 else "a, b", "a"))
            assert host.call("Scaled", 1.0) == Err(VALUE)
            host.tick()
            for j in range(10):
                assert host.call("Scaled", *([float(j)] if i % 2 else [float(j), 0.0])) == float(j)

    cycles(20)  # Warm up.
    before = host.handles()
    cycles(300)  # 300 registrations again and 3,000 calls.
    assert host.handles() - before <= 3
    assert types(host) == {"Scaled": "QQ"}  # The last cycle's, one argument.



@pytest.mark.parametrize("skip", range(12))
def test_memory_running_out_after_a_failed_removal_keeps_the_slot_out_of_use(host, tmp_path, skip):
    # Memory runs out at each point after Excel refuses the removal in turn; whatever it
    # stopped, the next function never takes the slot Excel still sends the old name to.
    script = tmp_path / "failing.py"
    loaded(host, script, ONE.format("stuck", "a", "'stuck'"))
    script.write_text("x = 1\n", encoding="utf-8")
    host.fail_unregisters(1)
    host.fail_memory("unregister", skip)
    host.call("PyLoad", str(script))  # #VALUE! if memory ran out in it.
    host.tick(2)
    following = tmp_path / "next.py"
    following.write_text(ONE.format("fresh", "a", "'fresh'"), encoding="utf-8")
    for _ in range(3):  # Each PyLoad after one memory stopped tries again.
        host.call("PyLoad", str(following))
        host.tick(2)
        if "fresh" in direct_names(host):
            break
    assert host.call("fresh", 1.0) == "fresh", host.stderr
    assert host.call("stuck", 1.0) == Err(VALUE)  # Not the new function's result.


@pytest.mark.parametrize("skip", range(12))
def test_memory_running_out_during_a_removal_never_keeps_the_name_from_its_function(host, tmp_path, skip):
    # The removal succeeds, and memory runs out at each point after it in turn: what Excel
    # then makes of the name is still recorded, so the function registers under it again.
    script = tmp_path / "back.py"
    loaded(host, script, ONE.format("back", "a", "'first'"))
    script.write_text("x = 1\n", encoding="utf-8")
    host.fail_memory("unregister", skip)
    host.call("PyLoad", str(script))
    host.tick(2)
    edit(script, ONE.format("back", "a", "'again'"))
    for _ in range(3):
        host.call("PyLoad", str(script))
        host.tick(2)
        if "back" in direct_names(host):
            break
    assert host.call("back", 1.0) == "again", host.stderr
    assert "already knows that name" not in host.stderr


def test_a_name_whose_removal_failed_twice_registers_again(host, tmp_path):
    # Each failed removal keeps a slot with what Excel then made of the name; Excel's answer
    # now matches the second, and the name is free to take a third slot.
    script = tmp_path / "twice.py"
    loaded(host, script, ONE.format("twice", "a", "'one'"))
    for version in ("'two'", "'three'"):
        script.write_text("x = 1\n", encoding="utf-8")
        host.fail_unregisters(1)
        assert host.call("PyLoad", str(script)) == "0 functions"
        host.tick()
        edit(script, ONE.format("twice", "a", version))
        assert host.call("PyLoad", str(script)) == "1 function"
        host.tick()
        assert host.call("twice", 1.0) == version.strip("'"), host.stderr
    assert slots(host)["twice"] == "PyFunc2"


@pytest.mark.parametrize("skip", range(40))
def test_memory_running_out_anywhere_in_a_pyload_loses_no_new_function(host, tmp_path, skip):
    # In the loader or the core, as each allocation fails in turn: a function whose row did
    # not reach the queue is offered again by the next PyLoad.
    loaded(host, tmp_path / "first.py", ONE.format("first", "a", "'first'"))  # Python is running.
    script = tmp_path / "later.py"
    script.write_text(ONE.format("later", "a", "'later'"), encoding="utf-8")
    host.fail_memory("now", skip)
    host.call("PyLoad", str(script))  # #VALUE! if memory ran out in it.
    host.tick(2)
    for _ in range(3):
        if "later" in direct_names(host):
            break
        host.call("PyLoad", str(script))
        host.tick(2)
    assert host.call("later", 1.0) == "later", host.stderr


@pytest.mark.parametrize("skip", range(40))
def test_memory_running_out_anywhere_in_a_pyload_that_removes_a_function_still_frees_its_slot(host, tmp_path, skip):
    # Python has already let the function go, so the removal is made whatever the loader
    # failed to take: its slot is freed, and nothing is left registered.
    script = tmp_path / "leaving.py"
    loaded(host, script, ONE.format("leaving", "a", "'leaving'"))
    script.write_text("x = 1\n", encoding="utf-8")
    host.fail_memory("now", skip)
    host.call("PyLoad", str(script))
    host.tick(2)
    for _ in range(3):
        if not slots(host):
            break
        host.call("PyLoad", str(script))
        host.tick(2)
    assert slots(host) == {}, host.stderr


def test_the_add_in_closed_while_excel_removes_a_function_does_not_bring_it_back(host, tmp_path):
    # A close from within the removal's own call into Excel, as it registers the stand-in
    # command: the function being removed is not queued to be registered again.
    script = tmp_path / "going.py"
    loaded(host, script, ONE.format("going", "a", "'going'"))
    script.write_text("x = 1\n", encoding="utf-8")
    host.nest_at("xlAutoRemove", "close")
    assert host.call("PyLoad", str(script)) == "0 functions"
    host.tick(3)
    assert "going" not in direct_names(host), host.stderr
    assert slots(host) == {}
