"""A real Esc press in real Excel (docs/architecture.md, "The pyinex module":
pyinex.esc_pressed), with the 2010 PythonBreakTest.

Entering the formula through COM does not return until its calculation ends, so a thread
started first presses the key: it waits for the marker file the script writes as its loop
starts, checks that Excel is the foreground window (made so before, sending no key), and
sends Esc with SendInput, as a keyboard would.
"""

import ctypes
import threading
import time
from ctypes import wintypes

import pytest
from conftest import foreground, pycall, typed

VK_ESCAPE, KEYEVENTF_KEYUP, INPUT_KEYBOARD = 0x1B, 0x0002, 1


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_size_t)]


class INPUT(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT), ("padding", ctypes.c_byte * 32)]

    _anonymous_ = ("u",)
    _fields_ = [("type", wintypes.DWORD), ("u", _U)]


def press(key):
    """Presses and releases one key, as a keyboard would."""
    keys = (INPUT * 2)()
    for i, flags in enumerate((0, KEYEVENTF_KEYUP)):
        keys[i].type, keys[i].ki.wVk, keys[i].ki.dwFlags = INPUT_KEYBOARD, key, flags
    return ctypes.windll.user32.SendInput(2, keys, ctypes.sizeof(INPUT)) == 2


def press_esc_once(marker, excel, cancel, report, timeout=30.0):
    """The sender, on a thread of its own: once the script's loop has started, and only if
    Excel is the foreground window, presses Esc. Records what it did, or why it did not, in
    report; does nothing once cancel is set."""
    deadline = time.monotonic() + timeout
    while not marker.exists():
        if cancel.is_set():
            return
        if time.monotonic() > deadline:
            report["error"] = "the script's loop never started"
            return
        time.sleep(0.05)
    if foreground() != excel:
        report["error"] = "Excel was not the foreground window, so no key was sent"
        return
    if not cancel.is_set():
        report["sent"] = press(VK_ESCAPE)


@pytest.fixture
def esc(app, book, excel_in_front):
    """A function entering a formula whose script loops until Esc and pressing Esc during
    it; returns the workbook. Excel is in front throughout, so the key goes to it and
    nowhere else. The test fails if the key could not be sent."""
    book.copy("esc.py")

    def enter(cell, function, *args):
        marker = book.folder / f"started {cell}"
        report, cancel = {}, threading.Event()
        sender = threading.Thread(target=press_esc_once, args=(marker, excel_in_front, cancel, report), daemon=True)
        sender.start()
        try:
            book.formula(cell, pycall("esc.py", function, f'"{marker}"', *args))
        finally:
            cancel.set()
            sender.join(35)
        assert not sender.is_alive(), "the sender is still running"
        assert "error" not in report, report["error"]
        assert report.get("sent"), "SendInput did not send Esc"
        return book

    return enter


XL_DONE, XL_PENDING = 0, 2  # Application.CalculationState.


@pytest.mark.parametrize("clear, kept", [("FALSE", True), ("TRUE", False)])
def test_esc_stops_a_loop_and_is_kept_or_cleared(esc, clear, kept):
    book = esc("A1", "loop_then_check", clear)
    assert book.value("A1") == f"stopped; still pressed: {kept}"


@pytest.mark.parametrize("clear, spilled, state", [("FALSE", None, XL_PENDING), ("TRUE", "spilled", XL_DONE)])
def test_a_kept_press_stops_excels_calculation_too(esc, app, clear, spilled, state):
    # Excel sees a press left unclear: it stops calculating once the function returns,
    # before the block it returned spills, and leaves the rest pending.
    book = esc("A1", "loop_then_block", clear)
    assert book.values("A1:B1") == [["stopped", spilled]]
    assert app.CalculationState == state


def test_the_next_call_runs_normally(esc):
    book = esc("A1", "loop_then_check", "TRUE")
    assert book.value("A1") == "stopped; still pressed: False"
    book.script("after.py", "import pyinex\n\ndef f():\n    return ['ran', pyinex.esc_pressed()]\n")
    book.formula("A3", pycall("after.py", "f"))
    assert typed(book.values("A3:B3")) == typed([["ran", False]])


def test_esc_pressed_keeps_the_press_by_default(esc):
    book = esc("A1", "loop_default")
    assert book.value("A1") == "stopped; still pressed: True, True"


@pytest.mark.parametrize("clear", ["FALSE", "TRUE"])
def test_long_loop_2010(esc, clear):
    book = esc("A1", "long_loop", '"first"', "30", clear)
    assert book.value("A1") == "first aborted"
