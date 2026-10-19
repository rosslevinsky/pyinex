"""Direct functions the real-Excel tests register through pyinex.ini (test_direct_functions.py,
test_registration.py). conftest.py's direct_excel copies this file, and values.py, whose
values it returns, beside a copy of the add-in that lists it. The tests call each function
by its name and through PyCall, from this same file, and compare."""

import pathlib
import threading
import time

import pyinex
from values import RETURNS


@pyinex.func
def Describe(x):
    """What arrived: its type's name and its repr, as text."""
    return f"{type(x).__name__}: {x!r}"


@pyinex.func
def Pair(a, b):
    """Describe() of each argument, in one row."""
    return [Describe(a), Describe(b)]


@pyinex.func
def Length(text):
    return len(text)


@pyinex.func
def EchoBack(x=None):  # Not Echo: ECHO is one of Excel's own macro functions.
    return x


@pyinex.func(name="Give")
def give(name):
    return RETURNS[name]


@pyinex.func
def IsGiven(name, x):
    """Whether x, read back from a cell, is exactly what Give(name) returned."""
    return x == RETURNS[name]


@pyinex.func
def GiveError(text):
    return pyinex.ExcelError(text)


@pyinex.func
def Raises():
    raise ValueError("raised on purpose")


@pyinex.func
def Exits():
    raise SystemExit(3)


@pyinex.func
def Interrupts():
    raise KeyboardInterrupt


@pyinex.func
def Two(a, b):
    return a + b


# The call context.


@pyinex.func
def CallerOf(r1c1=False, full=False):
    return pyinex.caller(r1c1=bool(r1c1), full=bool(full))


@pyinex.func
def SheetOf():
    return pyinex.sheet()


@pyinex.func
def Attempt(name):
    """What pyinex.<name>() raises, or what it returns, as text."""
    try:
        return repr(getattr(pyinex, name)())
    except Exception as e:
        return type(e).__name__


@pyinex.func
def OnAThread(name):
    """What pyinex.<name>() raises on a thread this function starts."""
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


@pyinex.func
def PrintLater(text):
    """Prints text from a thread of its own, after this call has returned."""
    threading.Thread(target=lambda: (time.sleep(0.2), print(text))).start()
    return "started"


ticks = [0]


@pyinex.func
def StartTicking():
    def tick():
        while True:
            ticks[0] += 1
            time.sleep(0.01)

    threading.Thread(target=tick, daemon=True).start()
    return "ticking"


@pyinex.func
def Ticked():
    return ticks[0]


@pyinex.func
def LoopThenCheck(marker, clear):
    """Loops until Esc, then says whether the press is still there for the next check."""
    pathlib.Path(marker).write_text("started", encoding="utf-8")
    for _ in range(600):  # 30 seconds.
        time.sleep(0.05)
        if pyinex.esc_pressed(clear=bool(clear)):
            return f"stopped; still pressed: {pyinex.esc_pressed()}"
    return "not stopped"


# Insert Function and the Function Arguments dialog.


@pyinex.func(category="Pyinex tests")
def Marked(path, note=None):
    """Writes the file named, so a test can tell it ran.

    Args:
        path: The file to write.
        note: Not used.
    """
    pathlib.Path(path).write_text("ran", encoding="utf-8")
    return "ran"


@pyinex.func(name="MarkedToo", category="Pyinex tests", help="Help given to the decorator.")
def marked_too():
    return "ran too"
