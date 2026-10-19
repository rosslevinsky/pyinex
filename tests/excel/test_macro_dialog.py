"""The Macro dialog (Alt+F8) in real Excel, opened as a user opens it, with the keys: it
lists no direct function and not PyRepair, which Excel hides as an add-in's command, though
it lists a command a workbook defines; PyRepair typed in it runs; a direct function's name
typed in it calls the function with no arguments, as a formula giving none would.

The dialog's controls are drawn by Excel, not made of windows of their own, so the test
reads them as a screen reader does, through Microsoft Active Accessibility (MSAA): Excel
answers for each item of the list and for the macro name's box. The keys go to whichever
window has the keyboard, so Excel is brought to the front first.

Functions have names of their own, starting "Md", since a registered name stays for the
session.
"""

import ctypes
import time

from conftest import E
from test_pyload import when_registered

DIALOG = "bosa_sdm_XL9"  # The class of Excel's own dialogs.
# MSAA's dispatch IDs (oleacc.h) and roles.
ACC_CHILD, ACC_CHILD_COUNT, ACC_NAME, ACC_VALUE, ACC_ROLE = -5002, -5001, -5003, -5004, -5006
ROLE_LISTITEM, ROLE_TEXT = 34, 42
VK_RETURN, VK_SHIFT, VK_ESCAPE, VK_MENU, VK_F8 = 0x0D, 0x10, 0x1B, 0x12, 0x77

CALLED = """import pyinex

@pyinex.func
def MdNeeds(x):
    return x

@pyinex.func
def MdSeen(*xs):
    print('MdSeen called with', xs)
    return len(xs)
"""


def tap(*keys):
    """Presses the keys together, as a user would, and lets them go in reverse."""
    import win32api

    for vk in keys:
        win32api.keybd_event(vk, 0, 0, 0)
    for vk in reversed(keys):
        win32api.keybd_event(vk, 0, 2, 0)  # KEYEVENTF_KEYUP


def type_text(text):
    import win32api

    for ch in text:
        code = win32api.VkKeyScan(ch)
        tap(*([VK_SHIFT] if code & 0x100 else []), code & 0xFF)


def dialogs(pid):
    """Excel's dialogs shown now."""
    import win32gui
    import win32process

    found = []

    def visit(hwnd, _):
        if (win32process.GetWindowThreadProcessId(hwnd)[1] == pid and win32gui.GetClassName(hwnd) == DIALOG
                and win32gui.IsWindowVisible(hwnd)):
            found.append(hwnd)

    win32gui.EnumWindows(visit, None)
    return found


def when(condition, timeout=15.0):
    """condition()'s value once true within timeout, else None."""
    deadline = time.monotonic() + timeout
    while not (value := condition()):
        if time.monotonic() > deadline:
            return None
        time.sleep(0.1)
    return value


def accessible(hwnd):
    """The window's MSAA object, as an IDispatch."""
    import pythoncom
    import pywintypes

    iid = (ctypes.c_byte * 16).from_buffer_copy(bytes(pywintypes.IID("{618736E0-3C3D-11CF-810C-00AA00389B71}")))
    pointer = ctypes.c_void_p()
    ctypes.oledll.oleacc.AccessibleObjectFromWindow(ctypes.c_void_p(hwnd), ctypes.c_long(-4),  # OBJID_CLIENT
                                                    ctypes.byref(iid), ctypes.byref(pointer))
    try:
        return pythoncom.ObjectFromAddress(pointer.value, pythoncom.IID_IDispatch)  # A reference of its own.
    finally:  # The reference AccessibleObjectFromWindow gave, released: IUnknown::Release.
        table = ctypes.cast(pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(table[2])(pointer)


def get(acc, dispid, *args):
    import pythoncom

    return acc.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, 1, *args)


def controls(acc):
    """(role, name, value) of each control, and of the controls within it."""
    import pythoncom

    found = []
    for i in range(1, get(acc, ACC_CHILD_COUNT) + 1):
        try:
            child = get(acc, ACC_CHILD, i)
        except pythoncom.com_error:
            child = None
        if child is not None:  # A control holding others: the list holds its items.
            child = child.QueryInterface(pythoncom.IID_IDispatch)
            found.append((get(child, ACC_ROLE, 0), get(child, ACC_NAME, 0), None))
            found += controls(child)
        else:
            try:
                value = get(acc, ACC_VALUE, i)
            except pythoncom.com_error:
                value = None
            found.append((get(acc, ACC_ROLE, i), get(acc, ACC_NAME, i), value))
    return found


def to_front(hwnd):
    """Whether the window came to the front. Windows lets a process bring a window to the
    front while its thread shares the input of the window in front, so it does."""
    import win32api
    import win32gui
    import win32process

    user32 = ctypes.windll.user32
    front = win32gui.GetForegroundWindow()
    ours = win32api.GetCurrentThreadId()
    theirs = win32process.GetWindowThreadProcessId(front)[0] if front else 0
    attach = theirs not in (0, ours)
    if attach:
        user32.AttachThreadInput(ours, theirs, True)
    try:
        return bool(user32.SetForegroundWindow(ctypes.c_void_p(hwnd)))
    finally:
        if attach:
            user32.AttachThreadInput(ours, theirs, False)


class MacroDialog:
    """The Macro dialog, opened with Alt+F8 as a user opens it."""

    def __init__(self, app):
        import win32process

        self.app = app
        self.pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]

    def __enter__(self):
        import win32gui

        assert not dialogs(self.pid), "a dialog is already open"
        hwnd = self.app.Hwnd
        win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
        if not to_front(hwnd):
            # Windows may refuse, as when no window is in front to share input with; a key
            # pressed first is the usual way round that. Alt is the one used, and once let go
            # shows Excel's key tips, which Escape hides.
            import win32api

            win32api.keybd_event(VK_MENU, 0, 0, 0)
            try:
                to_front(hwnd)
            finally:
                win32api.keybd_event(VK_MENU, 0, 2, 0)
            tap(VK_ESCAPE)
        assert when(lambda: win32gui.GetForegroundWindow() == hwnd, 5), "Excel did not come to the front"
        tap(VK_MENU, VK_F8)
        try:
            found = when(lambda: dialogs(self.pid))
            assert found, "Alt+F8 opened no dialog"
            self.hwnd = found[0]
            assert get(accessible(self.hwnd), ACC_NAME, 0) == "Macro"
        except BaseException:
            self.__exit__()
            raise
        return self

    def __exit__(self, *_):
        if dialogs(self.pid):  # Left open by a failure: closed, so the next test can run.
            tap(VK_ESCAPE)
            when(lambda: not dialogs(self.pid), 5)

    def macros(self):
        """The names the dialog's list shows."""
        return [name for role, name, _ in controls(accessible(self.hwnd)) if role == ROLE_LISTITEM
                and name not in ("All Open Workbooks", "This Workbook") and not name.endswith(".xlsx")]

    def run(self, name):
        """Types the name in the macro name's box, which has the keyboard when the dialog
        opens, and presses Enter, which presses Run; waits for the dialog to close."""
        type_text(name)
        box = [value for role, _, value in controls(accessible(self.hwnd)) if role == ROLE_TEXT]
        assert box == [name]
        tap(VK_RETURN)
        assert when(lambda: not dialogs(self.pid)), f"the dialog stayed open after running {name}"


def test_the_macro_dialog_lists_no_direct_function_and_not_pyrepair(app, book):
    book.script("md_listed.py", "import pyinex\n\n@pyinex.func\ndef MdListed():\n    return 1\n")
    assert book.formula("A1", '=PyLoad("md_listed.py")') == "1 function"
    assert when_registered(app, "MdListed()") == 1.0
    # A command the workbook defines on an Excel 4 macro sheet, which the dialog lists: so
    # an empty list would be one read and found empty, not one that could not be read.
    sheet = book.workbook.Excel4MacroSheets.Add()
    sheet.Range("A1").Formula = "=RETURN()"
    book.workbook.Names.Add("MdCommand", f"='{sheet.Name}'!$A$1", True, 2)  # 2: xlCommand
    with MacroDialog(app) as dialog:
        assert dialog.macros() == ["MdCommand"]


def test_pyrepair_typed_in_the_macro_dialog_runs(app, book, cleared):
    # As test_repair.py's protected sheet: a formula that could not be entered again is
    # repaired only by PyRepair.
    loader = book.workbook.Worksheets.Add(After=book.sheet)
    book.script("md_locked.py", "import pyinex\n\n@pyinex.func\ndef MdLocked(x):\n    return x + 1\n")
    assert book.formula("B1", "=MdLocked(1)") == E("#NAME?")
    book.sheet.Protect()
    try:
        loader.Range("A1").Formula2 = '=PyLoad("md_locked.py")'
        assert "again (" in cleared(holding="again (").text
        assert book.value("B1") == E("#NAME?")
    finally:
        book.sheet.Unprotect()
    with MacroDialog(app) as dialog:
        dialog.run("PyRepair")
    assert when(lambda: book.value("B1") == 2.0, 5)


def test_a_direct_function_typed_in_the_macro_dialog_is_called_with_no_arguments(app, book, cleared):
    book.script("md_called.py", CALLED)
    assert book.formula("A1", '=PyLoad("md_called.py")') == "2 functions"
    assert when_registered(app, "MdSeen(1)") == 1.0
    assert book.formula("B1", "=MdNeeds(3)") == 3.0
    with MacroDialog(app) as dialog:
        dialog.run("MdSeen")
    assert cleared(holding="MdSeen called with ()").text.count("MdSeen called with ()") == 1
    with MacroDialog(app) as dialog:
        dialog.run("MdNeeds")  # As =MdNeeds() would: the error reported, Excel unharmed.
    text = cleared(holding="missing 1").text
    assert "MdNeeds() missing 1 required positional argument: 'x'" in text, text
    assert book.formula("B2", "=MdNeeds(4)") == 4.0
