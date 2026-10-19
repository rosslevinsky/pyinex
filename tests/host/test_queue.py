"""The queue that registers functions after start-up (docs/architecture.md, "The queue"), in
the fake Excel: the hidden command PyinexQueue<n>, run by the host's stand-in for
Application.Run when the queue's timer fires (Host.tick sends the timer's message), or by
the host's own request (Host.run); refusals, a failed xlfRegister, work queued while the
queue runs, the queue started again from inside itself, a queue window that could not be
made, a timer that could not be set, memory running out in the queue's command just
before Excel is told of a function and just after, and PyRepair run while a function
registers, in the queue's command and as the add-in opens.
"""

import re
import time

import pytest

from conftest import host_env
from test_direct_functions import wait_for

from hostdriver import _hex
from test_pyload import GOOD, direct_names, queue_command

WAITING = "Pyinex: Excel cannot register functions yet"


def two_functions(tmp_path, name="two.py", first="first", second="second"):
    script = tmp_path / name
    script.write_text(GOOD.format(first) + "\n" + GOOD.format(second), encoding="utf-8")
    return str(script)


def test_the_command_is_registered_as_a_command(host):
    (queue,) = [r for r in host.registrations() if r.procedure == "PyinexQueue"]
    assert queue.type == "J"
    # Its name is this copy's own, so that another copy of Pyinex has a command of its own.
    assert re.fullmatch(r"PyinexQueue\d+", queue.name)


def test_queued_work_is_run_by_the_command(host, tmp_path):
    host.set_workbook("none")  # So the queue cannot reach Excel's Application object itself.
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    assert host.tick(3) == (True, 0, 0)
    assert direct_names(host) == []
    assert host.run(queue_command(host)) == 1
    assert direct_names(host) == ["first", "second"]
    assert host.call("second", 1.0) == "second: 1.0"


def test_the_timer_runs_the_command_through_application_run(host, tmp_path):
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    host.tick()
    assert direct_names(host) == ["first", "second"]
    assert host.tick() == (True, 0, 0)  # Nothing left: no more attempts.


def test_a_refused_registration_is_reported_and_tried_again_by_pyload(host, tmp_path):
    script = two_functions(tmp_path)
    host.fail_registrations(1)
    assert host.call("PyLoad", script) == "2 functions"
    host.tick()
    assert direct_names(host) == ["second"]
    assert f"Pyinex: Excel refused to register first from the script {script}" in wait_for(host, "refused to register")
    # Its name is free again, so the PyLoad cell's next calculation registers it.
    assert host.call("PyLoad", script) == "2 functions"
    host.tick()
    assert direct_names(host) == ["second", "first"]
    assert host.call("first", 2.0) == "first: 2.0"


def test_work_queued_while_the_queue_runs_is_done_by_the_same_run(host, tmp_path):
    later = two_functions(tmp_path, "later.py", "third", "fourth")
    host.nest("call", "PyLoad", later)  # Calculated as the first function registers.
    host.set_workbook("none")  # So the timer's first try, whenever it fires, runs nothing.
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    assert host.tick() == (True, 0, 0)
    host.set_workbook("open")
    _, runs, refused = host.tick()
    assert (runs, refused) == (1, 0)
    assert direct_names(host) == ["first", "second", "third", "fourth"]
    assert [r.procedure for r in host.registrations() if r.procedure.startswith("PyFunc")] == [
        "PyFunc0", "PyFunc1", "PyFunc2", "PyFunc3"]
    assert host.call("fourth", 1.0) == "fourth: 1.0"


def test_the_queue_is_never_run_twice_at_once(host, tmp_path):
    host.nest("run", queue_command(host))  # Started again as the first function registers.
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    host.tick()
    assert "Pyinex: the queue is already running" in wait_for(host, "already running")
    # Each registered once, in order, by the first run.
    assert [(r.name, r.procedure) for r in host.registrations() if r.procedure.startswith("PyFunc")] == [
        ("first", "PyFunc0"), ("second", "PyFunc1")]


def test_thousands_of_refusals_then_the_work_is_done(host, tmp_path):
    host.set_busy(True)
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    assert host.tick(5000) == (True, 5000, 5000)  # Far past the first spike's 400.
    assert direct_names(host) == []
    host.set_workbook("none")  # Another cause: Excel's Application object cannot be found.
    assert host.tick(3) == (True, 0, 0)
    host.set_workbook("open")
    host.set_busy(False)
    assert host.tick() == (True, 1, 0)
    assert direct_names(host) == ["first", "second"]
    assert host.call("first", 5.0) == "first: 5.0"
    # Each cause is reported when it starts, and only then: 5,000 refusals are one report.
    assert wait_for(host, "registered second").count(WAITING) == 2


def test_with_no_workbook_window_the_work_waits_for_one(host, tmp_path):
    # The running object table holds no Excel of this process, so the work waits; it is
    # never dropped.
    host.set_workbook("none")
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    assert host.tick(20) == (True, 0, 0)
    assert direct_names(host) == []
    assert WAITING in wait_for(host, WAITING)
    host.set_workbook("open")
    host.tick()
    assert direct_names(host) == ["first", "second"]


def test_a_protected_view_window_is_passed_over(host, tmp_path):
    # Its object has no Application, so the queue looks further; here there is nothing more.
    host.set_workbook("protected")
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    assert host.tick(5) == (True, 0, 0)
    assert direct_names(host) == []
    host.set_workbook("open")
    host.tick()
    assert direct_names(host) == ["first", "second"]


def test_a_queue_window_not_made_is_made_when_pyload_calculates_again(host, tmp_path):
    script = two_functions(tmp_path)
    host.fail_windows(1)
    assert host.call("PyLoad", script) == "2 functions"
    assert "Pyinex: cannot start the queue's timer" in wait_for(host, "cannot start the queue's timer")
    assert host.tick() == (False, 0, 0)
    # Its functions are queued, so their names are taken and this calculation finds none
    # new; it starts the timer for the work waiting.
    assert host.call("PyLoad", script) == "2 functions"
    host.tick()
    assert direct_names(host) == ["first", "second"]


def test_a_timer_not_set_is_set_when_pyload_calculates_again(host, tmp_path):
    script = two_functions(tmp_path)
    host.fail_timers(1)
    assert host.call("PyLoad", script) == "2 functions"
    assert "Pyinex: cannot start the queue's timer (Windows error 8)" in wait_for(host, "cannot start the queue's timer")
    # The host handles Windows' messages before each request, so a timer set would have
    # fired by now: none was, and the work waits.
    time.sleep(0.5)
    assert direct_names(host) == []
    assert host.call("PyLoad", script) == "2 functions"
    deadline = time.monotonic() + 5  # The timer itself, with no tick, runs the command.
    while direct_names(host) != ["first", "second"] and time.monotonic() < deadline:
        time.sleep(0.05)
    assert direct_names(host) == ["first", "second"]


def pool(host):
    return [(r.name, r.procedure) for r in host.registrations() if r.procedure.startswith("PyFunc")]


def test_memory_running_out_before_excel_is_told_leaves_the_function_queued(host, tmp_path):
    host.set_workbook("none")  # So the queue's own timer cannot start the command.
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    host.fail_memory("evaluate")  # As the command asks Excel about first's name.
    assert host.run(queue_command(host)) == 0  # The command failed, and said so to Excel.
    assert pool(host) == []
    host.set_workbook("open")
    host.tick()
    assert pool(host) == [("first", "PyFunc0"), ("second", "PyFunc1")]  # Each once.
    assert host.call("first", 1.0) == "first: 1.0"


def test_memory_running_out_once_excel_has_registered_a_function_loses_nothing(host, tmp_path):
    host.set_workbook("none")
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    host.fail_memory("register")  # Once Excel has registered first.
    assert host.run(queue_command(host)) == 0
    host.set_workbook("open")
    host.tick()
    # first is not registered again, in another slot, and second is registered after it.
    assert pool(host) == [("first", "PyFunc0"), ("second", "PyFunc1")]
    assert host.call("first", 1.0) == "first: 1.0"


def test_pyrepair_run_while_excel_registers_a_function_then_memory_running_out_loses_nothing(host, tmp_path):
    # PyRepair, run from within xlfRegister, adds the names of the functions registered to
    # the repair's list, after the name of the one Excel is registering.
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    host.tick()
    host.set_workbook("none")
    assert host.call("PyLoad", two_functions(tmp_path, "more.py", "third", "fourth")) == "2 functions"
    host.nest("run", "PyRepair")
    host.fail_memory("register")  # Once Excel has registered third.
    assert host.run(queue_command(host)) == 0
    host.set_workbook("open")
    host.tick()
    assert pool(host) == [("first", "PyFunc0"), ("second", "PyFunc1"), ("third", "PyFunc2"), ("fourth", "PyFunc3")]
    assert host.call("third", 1.0) == "third: 1.0"


@pytest.mark.parametrize("refused", [False, True], ids=["registered", "refused"])
def test_pyrepair_run_while_a_listed_function_registers_waits_for_the_add_in_to_open(startup, refused):
    # A function pyinex.ini lists registers as the add-in opens, outside the queue's command;
    # PyRepair run then leaves its work to the queue's run after, so the formula calling the
    # function is entered once the function is registered, or, refused, left alone.
    def before(host):
        host.set_formula("[Book1]Sheet1", "A1", "=Listed(1)")
        host.nest("run", "PyRepair")
        if refused:
            host.fail_registrations(1)

    host = startup("listed.py", before=before, **{"listed.py": GOOD.format("Listed") + "\n" + GOOD.format("Other")})
    assert "Pyinex: the queue is already running" in wait_for(host, "already running")
    entered, _ = host.entered()
    if refused:
        assert direct_names(host) == ["Other"]
        assert "Pyinex: Excel refused to register Listed" in wait_for(host, "refused to register")
        assert [e for e in entered if e[1] == "$A$1"] == []
    else:
        assert direct_names(host) == ["Listed", "Other"]
        assert [e for e in entered if e[1] == "$A$1"] == [("[Book1]Sheet1", "$A$1", "=Listed(1)", "Listed: 1.0")]



def test_a_queue_command_excel_refuses_at_opening_is_reported_and_its_work_dropped(xllhost, pyinex_xll, tmp_path):
    # Application.Run cannot run a command never registered: rather than try for ever, the
    # work is dropped, and the console says why and what to do.
    from hostdriver import Host

    with Host(xllhost, pyinex_xll, env=host_env()) as host:
        host.fail_commands(1)  # The first command the add-in registers: the queue's.
        host.open()
        assert not [r for r in host.registrations() if r.procedure == "PyinexQueue"]
        assert "Excel refused to register Pyinex's queue command" in wait_for(host, "queue command")
        assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
        # The timer may already have fired, as the host handled its messages before this.
        _, runs, _ = host.tick()
        assert runs == 0  # Never Application.Run.
        assert "Pyinex: 2 functions cannot be registered: Excel refused Pyinex's queue command" in wait_for(host, "cannot be registered")
        assert host.tick() == (False, 0, 0)  # Its window and timer ended.
        assert direct_names(host) == []
        assert WAITING not in host.stderr


@pytest.mark.parametrize("skip", range(4))
def test_pyrepair_failing_inside_a_registration_leaves_the_queue_its_own(host, tmp_path, skip):
    # PyRepair, run from within xlfRegister, runs out of memory; the queue's command, run
    # there too, then still finds the queue running, and each function registers once.
    host.set_workbook("none")  # So only the run below starts the queue.
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    command = queue_command(host)
    host.nest_all(f"memory {skip}", f"run {_hex('PyRepair')}", f"run {_hex(command)}")
    host.run(command)
    host.set_workbook("open")
    host.tick(2)
    assert pool(host) == [("first", "PyFunc0"), ("second", "PyFunc1")]


def test_the_add_in_closed_while_excel_registers_a_function_loses_none(host, tmp_path):
    # A close from within xlfRegister puts what it removed ahead in the queue: the run takes
    # off the queue the function it registered, and leaves the rest to the next run.
    first = tmp_path / "first.py"
    first.write_text(GOOD.format("alpha"), encoding="utf-8")
    assert host.call("PyLoad", str(first)) == "1 function"
    host.tick()
    host.set_workbook("none")
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    host.nest_all("close")
    host.run(queue_command(host))
    host.set_workbook("open")
    host.tick(3)
    assert sorted(direct_names(host)) == ["alpha", "first", "second"]  # Each once.


@pytest.mark.parametrize("refused", [False, True], ids=["registered", "refused"])
def test_the_add_in_closed_while_excel_registers_a_new_name_is_safe(host, tmp_path, refused):
    host.set_workbook("none")
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    host.nest_all("close")
    if refused:
        host.fail_registrations(1)
    host.run(queue_command(host))
    host.set_workbook("open")
    host.tick(3)
    assert sorted(direct_names(host)) == (["second"] if refused else ["first", "second"])


def test_a_run_of_the_queue_that_stopped_is_reported_as_such(host, tmp_path):
    # Excel ran the command, which stopped before the end: the console says so, not that
    # there was no workbook to run it from.
    host.set_workbook("none")  # So the queue's own timer cannot start the command first.
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    host.fail_memory("register")  # Once Excel has registered first.
    host.set_workbook("open")
    host.tick()
    said = wait_for(host, "stopped before the end")
    assert f"{WAITING} (its last run of the queue stopped before the end); Pyinex keeps trying" in said
    host.tick()
    assert pool(host) == [("first", "PyFunc0"), ("second", "PyFunc1")]


def test_pyload_starts_again_a_timer_that_failed_with_only_a_repair_waiting(host, tmp_path):
    # A PyLoad cell's calculation starts the queue's timer again whatever waits, not only
    # functions to register. Here the timer's own message, which the host handles before its
    # next request, as Excel would, runs the repair; tick would send it whether set or not.
    script = two_functions(tmp_path)
    assert host.call("PyLoad", script) == "2 functions"
    host.tick()
    host.set_formula("[Book1]Sheet1", "A1", "=first(1)")
    host.set_workbook("none")  # PyRepair's run finds no workbook: the repair waits.
    host.fail_timers(1)
    host.run("PyRepair")
    host.set_workbook("open")
    host.entered()
    assert host.call("PyLoad", script) == "2 functions"  # Nothing new to register.
    time.sleep(0.5)
    entered, _ = host.entered()
    assert ("[Book1]Sheet1", "$A$1", "=first(1)", "first: 1.0") in entered


def test_the_add_in_closed_while_excel_registers_a_new_name_still_repairs_its_formulas(host, tmp_path):
    # A formula typed before its function was registered is entered again once it is, even
    # when a close, from within that registration, emptied the list of names to repair.
    host.set_formula("[Book1]Sheet1", "A1", "=first(1)")
    host.set_workbook("none")
    assert host.call("PyLoad", two_functions(tmp_path)) == "2 functions"
    host.nest_all("close")
    host.run(queue_command(host))
    host.set_workbook("open")
    host.entered()
    host.tick(3)
    entered, _ = host.entered()
    assert ("[Book1]Sheet1", "$A$1", "=first(1)", "first: 1.0") in entered, host.stderr


@pytest.mark.parametrize("skip", range(12))
def test_memory_running_out_after_a_new_name_registers_still_repairs_its_formulas(host, tmp_path, skip):
    # As each allocation after the registration fails in turn, the names to repair are kept
    # until the repair has what it needs, so the next run enters the formula again.
    host.set_formula("[Book1]Sheet1", "A1", "=first(1)")
    host.set_workbook("none")
    script = tmp_path / "one.py"
    script.write_text(GOOD.format("first"), encoding="utf-8")
    assert host.call("PyLoad", str(script)) == "1 function"
    host.fail_memory("register", skip)
    host.set_workbook("open")
    host.entered()
    host.tick(4)
    entered, _ = host.entered()
    assert ("[Book1]Sheet1", "$A$1", "=first(1)", "first: 1.0") in entered, host.stderr
