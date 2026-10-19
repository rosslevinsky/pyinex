"""Opening, closing, reopening and unloading the add-in, and when the pyinex module may ask
Excel for something (docs/architecture.md, "Lifetime" and "When they may be called"); with
work queued, and a quit cancelled at the save prompt (docs/architecture.md, "The queue")."""

import shutil
import time

from conftest import attempt, evaluate, host_env
from hostdriver import VALUE, Err, _hex
from test_direct_functions import DIRECT, wait_for
from test_pyload import GOOD, direct_names
from test_reload_functions import edit


def test_open_call_close_reopen_call(host):
    assert evaluate(host, "41 + 1", "kept = 'from before closing'") == "42"
    assert host.close() == 1
    assert host.open() == 1
    # The same interpreter: what the scripts set up is still there.
    assert evaluate(host, "kept") == "from before closing"


def test_once_python_starts_the_addin_stays_loaded(host):
    evaluate(host, "0", "kept = 'still here'")
    host.close()
    # Excel unloads the add-in after closing it; Python holds pointers into both DLLs.
    assert host.unload() == (True, True)
    assert host.open() == 1
    assert evaluate(host, "kept") == "still here"


def test_before_python_starts_the_addin_unloads(host):
    host.close()
    assert host.unload() == (False, False)
    assert host.open() == 1
    assert host.call("PyInfo", "status") == "running"


def test_no_python_runs_while_the_addin_unloads_or_excel_exits(host):
    setup = """
import atexit, os
class Noisy:
    def __del__(self):
        os.write(2, b'finalizer ran')
noisy = Noisy()
atexit.register(os.write, 2, b'atexit ran')
"""
    evaluate(host, "0", setup)
    host.close()
    host.unload()
    assert host.quit() == 0  # Exited, not killed after hanging.
    assert "finalizer ran" not in host.stderr
    assert "atexit ran" not in host.stderr


# What a list of results is when each of the three was refused.
REFUSED = "(['RuntimeError'], 3)"
SUMMARY = "sorted({{r.split(':')[0] for r in {0}}}), len({0})"


def test_a_thread_a_script_started_is_refused(host):
    host.set_caller("[Book1]Sheet1", row=2, column=3)
    setup = """
import pyinex, threading
results = []
def ask():
    for f in (pyinex.caller, pyinex.sheet, pyinex.esc_pressed):
        try:
            results.append(f())
        except RuntimeError as e:
            results.append('RuntimeError: ' + str(e))
"""
    evaluate(host, "0", setup)
    host.excel_calls()
    # During a call, from another thread.
    evaluate(host, "0", "t = threading.Thread(target=ask)\nt.start()\nt.join()")
    assert host.excel_calls() == 0
    # Between calls.
    evaluate(host, "0", "results.clear()\nt = threading.Thread(target=ask)\nt.start()")
    time.sleep(0.5)
    assert host.excel_calls() == 0
    assert evaluate(host, SUMMARY.format("results"), "t.join()") == REFUSED, evaluate(host, "results")
    # And Excel's own thread still may.
    assert evaluate(host, "pyinex.caller()") == "C2"


def test_code_run_while_the_addin_closes_is_refused(host):
    # Closing flushes the output streams, so a replaced sys.stdout's flush runs then, on
    # Excel's thread but outside any PyCall.
    setup = """
import pyinex, sys
class Asking:
    def __init__(self, stream):
        self.stream, self.results = stream, []
    def write(self, text):
        return self.stream.write(text)
    def flush(self):
        # Each flush replaces the last one's results: PyCall flushes after every call too.
        results = []
        for f in (pyinex.caller, pyinex.sheet, pyinex.esc_pressed):
            try:
                results.append(f())
            except RuntimeError as e:
                results.append('RuntimeError: ' + str(e))
        self.results = results
asking = sys.stdout = Asking(sys.stdout)
"""
    evaluate(host, "0", setup)
    assert evaluate(host, "asking.results") == "['A1', '[Book1]Sheet1', False]"  # In a call, allowed.
    host.excel_calls()
    host.close()
    assert host.excel_calls() == 0
    host.open()
    summary = attempt(host, SUMMARY.format("asking.results"), "sys.stdout = asking.stream")
    assert summary == REFUSED, evaluate(host, "asking.results")


def test_handles_and_reference_counts_stay_flat(host, tmp_path):
    path = tmp_path / "counted.py"

    def edit(n):
        # A different size each time, so the script reloads.
        path.write_bytes(b"def f(x):\n    return x" + b" " * (n % 7) + b"\n")

    def calls(n):
        for i in range(n):
            if i % 10 == 0:
                edit(i // 10)
            assert host.call("PyCall", str(path), "f", float(i)) == float(i)

    calls(30)  # Warm up: Python opens what it opens once.
    before = host.handles()
    calls(300)  # 300 calls, 30 reloads.
    assert host.handles() - before <= 3

    # A module kept in the cache gains no references from being called.
    find = "next(m for n, m in sys.modules.items() if n.startswith('pyinex_counted_'))"
    first = evaluate(host, f"sys.getrefcount({find})", "import gc, sys, weakref")
    for i in range(100):
        host.call("PyCall", str(path), "f", 1.0)
    assert evaluate(host, f"sys.getrefcount({find})") == first
    # And a reloaded one is freed.
    evaluate(host, "0", f"old = weakref.ref({find})")
    edit(1000)
    host.call("PyCall", str(path), "f", 1.0)
    assert evaluate(host, "old()", "gc.collect()") == "None"


# Closing with work queued (docs/architecture.md, "Lifetime").


def test_unloaded_with_work_queued_then_reopened_nothing_is_left_from_before(host, tmp_path):
    script = tmp_path / "queued.py"
    script.write_text(GOOD.format("queued"), encoding="utf-8")
    host.set_busy(True)  # Excel refuses to run the queue, so the work stays queued.
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick(3)
    assert host.close() == 1
    assert host.unload() == (True, True)  # Pinned: Python is running.
    host.set_busy(False)
    # Excel has unregistered the add-in's functions, so the queue sees it was unloaded,
    # says so, and ends its window and timer without running the command.
    assert host.tick() == (True, 0, 0)
    assert "Pyinex: the add-in was unloaded, so 1 function will not be registered again" in wait_for(host, "was unloaded")
    assert host.tick() == (False, 0, 0)
    assert host.open() == 1
    assert host.tick() == (False, 0, 0)  # No timer, window or work from before.
    assert direct_names(host) == []
    # The name is free again, and its PyLoad cell's next calculation registers it, in the
    # first slot: none was kept from before.
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick()
    assert [(r.name, r.procedure) for r in host.registrations() if r.name == "queued"] == [("queued", "PyFunc0")]
    assert host.call("queued", 1.0) == "queued: 1.0"


def test_closed_with_work_queued_and_reopened_nothing_is_left_from_before(host, tmp_path):
    script = tmp_path / "queued.py"
    script.write_text(GOOD.format("queued"), encoding="utf-8")
    host.set_busy(True)
    assert host.call("PyLoad", str(script)) == "1 function"
    assert host.close() == 1 and host.open() == 1
    assert "Pyinex: 1 function queued before the add-in closed will not be registered" in wait_for(host, "before the add-in closed")
    host.set_busy(False)
    assert host.tick() == (False, 0, 0)
    assert direct_names(host) == []
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick()
    assert host.call("queued", 2.0) == "queued: 2.0"


def test_a_quit_cancelled_at_the_save_prompt_brings_the_functions_back(host, tmp_path):
    # Excel calls xlAutoClose before its save prompt; if the user cancels, Excel carries on
    # with the add-in loaded, as the host does after "close".
    first, other = tmp_path / "first.py", tmp_path / "other.py"
    first.write_text(GOOD.format("alpha") + "\n" + GOOD.format("beta"), encoding="utf-8")
    other.write_text(GOOD.format("gamma"), encoding="utf-8")
    assert host.call("PyLoad", str(first)) == "2 functions"
    host.tick()
    assert host.call("alpha", 1.0) == "alpha: 1.0"
    host.set_busy(True)  # So the queue waits while the state between is checked.
    assert host.close() == 1
    assert direct_names(host) == []
    assert host.call("alpha", 1.0) == Err(VALUE)  # Removed, and its name cleared.
    # Work queued after closing waits behind the functions closing removed.
    assert host.call("PyLoad", str(other)) == "1 function"
    assert host.call("alpha", 1.0) == Err(VALUE)
    host.set_busy(False)
    host.tick()
    assert direct_names(host) == ["alpha", "beta", "gamma"]
    # Each name reaches its own function, and no other.
    assert [host.call(n, 2.0) for n in ("alpha", "beta", "gamma")] == ["alpha: 2.0", "beta: 2.0", "gamma: 2.0"]
    assert len({r.procedure for r in host.registrations() if r.procedure.startswith("PyFunc")}) == 3
    # Opening again later still starts from what is registered.
    assert host.close() == 1 and host.open() == 1
    assert direct_names(host) == []


def test_a_cancelled_quit_brings_back_the_functions_pyinex_ini_lists(startup):
    host = startup(str(DIRECT))
    names = direct_names(host)
    assert host.close() == 1
    assert direct_names(host) == []
    host.tick()
    assert direct_names(host) == names
    assert host.call("echo", "back") == "back"


def test_protected_view_after_a_cancelled_quit_is_not_taken_for_an_unload(host, tmp_path):
    # With a Protected View window active, Excel's Evaluate gives #N/A: not a sign the
    # add-in was unloaded, so nothing queued is dropped.
    first, other = tmp_path / "first.py", tmp_path / "other.py"
    first.write_text(GOOD.format("alpha"), encoding="utf-8")
    other.write_text(GOOD.format("gamma"), encoding="utf-8")
    assert host.call("PyLoad", str(first)) == "1 function"
    host.tick()
    assert host.close() == 1  # A quit cancelled at the save prompt: Excel carries on.
    host.set_workbook("both")
    host.tick()
    assert host.call("alpha", 1.0) == "alpha: 1.0"
    # Later work, with the Protected View window still active.
    assert host.call("PyLoad", str(other)) == "1 function"
    host.tick()
    assert host.call("gamma", 1.0) == "gamma: 1.0"
    assert "was unloaded" not in host.stderr


def test_unloaded_while_another_copy_is_loaded_the_work_waits_and_runs_no_other_queue(host, tmp_path, pyinex_xll):
    # Another copy of Pyinex, loaded after this one, has taken the name PyInfo, and has a
    # queue command of its own. Its path does not prove this copy was unloaded, so the work
    # waits rather than being dropped; and the other copy's command never runs in place of
    # this copy's.
    script = tmp_path / "queued.py"
    script.write_text(GOOD.format("alpha") + "\n" + GOOD.format("beta"), encoding="utf-8")
    copy = tmp_path / "copy" / "pyinex.xll"
    copy.parent.mkdir()
    shutil.copy(pyinex_xll, copy)
    assert host.call("PyLoad", str(script)) == "2 functions"
    host.tick()
    host.set_busy(True)  # So the work waits until the copy is loaded.
    assert host.close() == 1
    assert host.unload() == (True, True)  # Pinned: Python is running.
    assert host.open(copy) == 1
    assert host.call("PyInfo", "pyinex") == str(copy)
    host.set_busy(False)
    # This copy keeps trying to run its own command, which is gone, and never the other
    # copy's. The work is not dropped either, since nothing proves this copy was unloaded
    # (docs/architecture.md, "Lifetime").
    for _ in range(5):
        assert host.tick()[0]  # The window and its timer remain.
    (theirs,) = [r.name for r in host.registrations() if r.procedure == "PyinexQueue"]
    tried = wait_for(host, "cannot register functions yet")
    assert "Cannot run the macro 'PyinexQueue" in tried and f"'{theirs}'" not in tried
    assert "was unloaded" not in host.stderr
    assert direct_names(host) == []



def test_pyrepair_run_while_the_add_in_opens_again_loses_no_workbook_function(host, tmp_path):
    # xlAutoOpen with no xlAutoClose before it removes what was registered; PyRepair, run as
    # the first removal registers its stand-in command, waits for opening to end.
    (tmp_path / "book.py").write_text(GOOD.format("Booked"), encoding="utf-8")
    host.set_folder(tmp_path)
    host.set_formula("[Book1]Sheet1", "A1", '=PyLoad("book.py")')
    assert host.call("PyLoad", str(tmp_path / "book.py")) == "1 function"
    host.tick()
    host.nest_at("xlAutoRemove", f"run {_hex('PyRepair')}")
    host.open()
    host.tick(3)
    assert direct_names(host) == ["Booked"]
    assert host.call("Booked", 1.0) == "Booked: 1.0"


def test_pyrepair_run_while_the_add_in_closes_leaves_it_closed(host, tmp_path):
    script = tmp_path / "closing.py"
    script.write_text(GOOD.format("closing"), encoding="utf-8")
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick()
    host.nest_at("xlAutoRemove", f"run {_hex('PyRepair')}")
    assert host.close() == 1
    assert host.unload() == (True, True)
    # Excel has unregistered the add-in's functions: the queue sees it was unloaded and ends,
    # never running the command.
    _, runs, _ = host.tick(3)
    assert runs == 0
    assert host.tick() == (False, 0, 0)


def test_a_worksheet_function_excel_refuses_at_opening_is_reported(xllhost, pyinex_xll):
    from hostdriver import Host

    with Host(xllhost, pyinex_xll, env=host_env()) as host:
        host.fail_name("PyLoad")
        host.open()
        assert "Pyinex: Excel refused to register PyLoad; close and open the add-in again" in wait_for(host, "refused")
        names = {r.name for r in host.registrations()}
        assert "PyLoad" not in names and {"PyCall", "PyInfo"} <= names  # The others are there.


def test_a_cancelled_quit_in_a_session_that_never_used_pyinex_starts_no_python(host):
    # Asking Excel whether the add-in was unloaded calculates PyInfo("pyinex"), which needs
    # no Python: the add-in is still not pinned, and unloads.
    assert host.close() == 1
    host.tick()
    assert host.unload() == (False, False)


def test_reopening_with_reloading_off_keeps_a_listed_script_as_it_loaded(startup, tmp_path):
    # PyReloadOnChange(FALSE) loads each script once, those pyinex.ini lists too, however
    # often the add-in opens.
    host = startup("kept.py", **{"kept.py": GOOD.format("Kept")})
    assert host.call("Kept", 1.0) == "Kept: 1.0"
    assert host.call("PyReloadOnChange", False) is False
    edit(tmp_path / "addin" / "kept.py", GOOD.format("Kept").replace("'Kept: '", "'Changed: '"))
    assert host.close() == 1
    assert host.open() == 1
    host.tick()
    assert host.call("Kept", 1.0) == "Kept: 1.0"
