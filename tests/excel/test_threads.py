"""PyCall's thread with multi-threaded recalculation on, in real Excel (docs/architecture.md,
"Threading"): Excel may calculate on several threads, but every PyCall runs on its main
thread, the one that started Python.
"""

import win32process
from conftest import pycall

XL_THREAD_MODE_AUTOMATIC = 0  # xlThreadModeAutomatic
CELLS = 500


def test_every_call_runs_on_excels_main_thread(book):
    book.script(
        "thread.py",
        "import ctypes, threading\n\n"
        "def f(row):\n"
        "    return [ctypes.windll.kernel32.GetCurrentThreadId(), threading.main_thread().ident]\n",
    )
    main_thread = win32process.GetWindowThreadProcessId(book.excel.Hwnd)[0]  # The thread owning Excel's window.
    setting = book.excel.MultiThreadedCalculation
    before = setting.Enabled, setting.ThreadMode
    try:
        setting.Enabled = True
        setting.ThreadMode = XL_THREAD_MODE_AUTOMATIC
        assert setting.ThreadCount > 1, "Excel calculates on one thread only here"
        # Thread-safe work of Excel's own beside each call, so the calculation is spread.
        book.sheet.Range(f"C1:C{CELLS}").Formula2 = "=SUMPRODUCT(SEQUENCE(2000)*ROW())"
        book.sheet.Range(f"A1:A{CELLS}").Formula2 = pycall("thread.py", "f", "ROW()")
        book.excel.CalculateFull()
        threads = book.values(f"A1:B{CELLS}")
    finally:
        setting.Enabled, setting.ThreadMode = before
    assert {(t, p) for t, p in threads} == {(float(main_thread), float(main_thread))}
