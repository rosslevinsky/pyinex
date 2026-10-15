"""Functions the real-Excel Esc tests call (test_esc.py), copied beside each test's
workbook first. Each writes the marker file named when its loop starts, which is when the
test presses Esc."""

import pathlib
import time

import pyinex


def loop_then_check(marker, clear):
    """Loops until Esc, then says whether the press is still there for the next check. One
    value, not a block: a kept press stops Excel's calculation before a block could spill."""
    pathlib.Path(marker).write_text("started", encoding="utf-8")
    for _ in range(600):  # 30 seconds.
        time.sleep(0.05)
        if pyinex.esc_pressed(clear=bool(clear)):
            return f"stopped; still pressed: {pyinex.esc_pressed()}"
    return "not stopped"


def loop_default(marker):
    """loop_then_check with esc_pressed's default, which keeps the press: checked twice."""
    pathlib.Path(marker).write_text("started", encoding="utf-8")
    for _ in range(600):
        time.sleep(0.05)
        if pyinex.esc_pressed():
            return f"stopped; still pressed: {pyinex.esc_pressed()}, {pyinex.esc_pressed()}"
    return "not stopped"


def loop_then_block(marker, clear):
    """Loops until Esc, then returns a block of two cells across."""
    pathlib.Path(marker).write_text("started", encoding="utf-8")
    for _ in range(600):
        time.sleep(0.05)
        if pyinex.esc_pressed(clear=bool(clear)):
            return ["stopped", "spilled"]
    return ["not stopped", "spilled"]


def long_loop(marker, id, count, clearBreak):
    pathlib.Path(marker).write_text("started", encoding="utf-8")
    return LongLoop(id, count, clearBreak)


# 2010: Examples/PythonBreakTest.py, with pyinex.Break(clearBreak) now
# pyinex.esc_pressed(clear=clearBreak).


def LongLoop(id, count, clearBreak):
    for i in range(0, int(count)):
        print(id + " " + str(i))
        time.sleep(1)
        if pyinex.esc_pressed(clear=bool(clearBreak)):
            print("Excel escape key pressed")
            return id + " aborted"

    return id + " completed"
