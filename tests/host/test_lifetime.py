"""Opening, closing, reopening and unloading the add-in, and when the pyinex module may ask
Excel for something (docs/architecture.md, "Lifetime" and "When they may be called")."""

import time

from conftest import attempt, evaluate


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
