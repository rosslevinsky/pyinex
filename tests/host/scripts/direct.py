"""Direct functions the fake-Excel tests register through pyinex.ini and call by name
(test_direct_functions.py, test_registration.py). The tests call each through PyCall too,
from this same file, and compare."""

import math
import threading
import time

import pyinex

RETURNS = {
    "none": None,
    "true": True,
    "false": False,
    "text": "text",
    "float": 2.5,
    "int": 666,
    "huge_int": 1234567891011121314,
    "too_large_int": 10**400,
    "nan": math.nan,
    "inf": math.inf,
    "longest_text": "x" * 32767,
    "long_text": "x" * 32768,
    "emoji_too_long": "\U0001f600" * 16384,
    "row": [1, 2, 3],
    "column": [[1], [2], [3]],
    "ragged": [[1, 2, 3], [4], 5],
    "empty": [],
    "empty_rows": [[], []],
    "empty_row_in_block": [[1, 2], []],
    "nan_in_block": [[1, 2], [3, math.nan]],
    "too_deep": [[1, [2]]],
    "dict": {"a": 1},
    "set": {1},
    "complex": 1 + 2j,
    "object": object(),
    "bad_in_block": [[1, 2], [3, {}]],
    "error": pyinex.ExcelError("#N/A"),
    "tuple": (1, "two", True),
}


@pyinex.func
def echo(value=None):
    """Returns its argument.

    Args:
        value: Anything.
    """
    return value


@pyinex.func(name="Describe", category="Pyinex tests", help="What arrived.")
def describe(x):
    return f"{type(x).__name__}: {x!r}"


@pyinex.func
def pair(a, b):
    return repr((a, b))


@pyinex.func
def give(name):
    return RETURNS[name]


@pyinex.func
def raises():
    raise ValueError("raised on purpose")


@pyinex.func
def exits():
    raise SystemExit(3)


@pyinex.func
def interrupts():
    raise KeyboardInterrupt


@pyinex.func
def two(a, b):
    return a + b


@pyinex.func
def caller(r1c1=False, full=False):
    return pyinex.caller(r1c1=bool(r1c1), full=bool(full))


@pyinex.func
def sheet():
    return pyinex.sheet()


@pyinex.func
def esc(clear=False):
    return pyinex.esc_pressed(clear=bool(clear))


@pyinex.func
def attempt(name):
    """What pyinex.<name>() returns or raises, as text."""
    try:
        return repr(getattr(pyinex, name)())
    except Exception as e:
        return type(e).__name__


@pyinex.func
def on_a_thread(name):
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
def print_later(text):
    """Prints text from a thread of its own, after this call has returned."""
    threading.Thread(target=lambda: (time.sleep(0.2), print(text))).start()
    return "started"


ticks = [0]


@pyinex.func
def start_ticking():
    def tick():
        while True:
            ticks[0] += 1
            time.sleep(0.01)

    threading.Thread(target=tick, daemon=True).start()
    return "ticking"


@pyinex.func
def ticked():
    return ticks[0]


@pyinex.func
def mark(path):
    """Writes the file named, so a test can tell the function ran."""
    open(path, "w").close()
    return "ran"


def not_marked(x):
    return x
