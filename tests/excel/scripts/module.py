"""Functions the real-Excel tests of the pyinex module call (test_pyinex_module.py,
test_reload.py), copied beside each test's workbook first.

The 2010 functions are carried over from 0.83's Examples/PythonExtensionTest.py and
Examples/PyinexTest.py, with the 2010 caller functions' new names.
"""

import threading
import time

import pyinex


def caller(r1c1=False, full=False):
    return pyinex.caller(r1c1=bool(r1c1), full=bool(full))


def sheet():
    return pyinex.sheet()


def refused(name):
    """What pyinex.<name>() raises here, called from no cell (through Application.Run)."""
    try:
        value = getattr(pyinex, name)()
    except Exception as e:
        return type(e).__name__
    return f"nothing raised: {value!r}"


def refused_on_a_thread(name):
    """What pyinex.<name>() raises on a thread the script started."""
    found = []

    def run():
        try:
            getattr(pyinex, name)()
            found.append("nothing raised")
        except Exception as e:
            found.append(type(e).__name__)

    thread = threading.Thread(target=run)
    thread.start()
    thread.join()
    return found[0]


def refused_between_calls(go, done):
    """Starts a thread that, once the file go exists (written after this call has returned),
    asks each of caller, sheet and esc_pressed, and writes what each did to the file done."""

    def ask():
        import os

        while not os.path.exists(go):
            time.sleep(0.05)
        seen = []
        for f in (pyinex.caller, pyinex.sheet, pyinex.esc_pressed):
            try:
                f()
                seen.append(f.__name__ + ": answered")
            except RuntimeError:
                seen.append(f.__name__ + ": RuntimeError")
        with open(done + ".partial", "w", encoding="utf-8") as out:
            out.write("\n".join(seen))
        os.replace(done + ".partial", done)  # Whole, or not there at all.

    threading.Thread(target=ask, daemon=True).start()
    return "started"


# 2010: Examples/PythonExtensionTest.py, CallerA1() and the rest now pyinex.caller() and
# pyinex.sheet().


def CallerA1Test():
    return pyinex.caller()


def CallerA1FullTest():
    return pyinex.caller(full=True)


def CallerR1C1Test():
    return pyinex.caller(r1c1=True)


def CallerR1C1FullTest():
    return pyinex.caller(r1c1=True, full=True)


def CallerSheetTest():
    return pyinex.sheet()


# 2010: Examples/PyinexTest.py.

watchMeGetIncremented = 4  # this is set when module is imported


def PlayWithGlobalVariable():
    global watchMeGetIncremented
    watchMeGetIncremented = watchMeGetIncremented + 1
    return watchMeGetIncremented
