"""The pyinex module: which cell is calling, and whether the user pressed Esc.

pyinex_module.xlsx calls these functions with PyCall. The module is built into the add-in,
so this script runs only inside Excel. docs/architecture.md, "The pyinex module", has the
details.
"""

import time

import pyinex


def where():
    return pyinex.caller()


def where_r1c1():
    return pyinex.caller(r1c1=True)


def where_full():
    """The workbook and sheet in front, without quotes even when the sheet name has spaces."""
    return pyinex.caller(full=True)


def which_sheet():
    return pyinex.sheet()


def where_block():
    """In a formula entered into a range of cells, caller() gives that range."""
    here = pyinex.caller()
    return [[here, here], [here, here]]


def slow_count(seconds):
    """Counts for a while, stopping early if the user presses Esc."""
    for _ in range(int(seconds * 10)):
        if pyinex.esc_pressed():
            print("slow_count: stopped by Esc")
            return "stopped by Esc"
        time.sleep(0.1)
    return "finished"
