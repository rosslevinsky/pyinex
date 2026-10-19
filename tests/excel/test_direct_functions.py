"""Direct functions in real Excel (docs/architecture.md, "Direct functions"): the functions of
tests/excel/scripts/direct.py, which a copy of the add-in lists in its pyinex.ini, called
by their own names in an Excel of their own (conftest.py, direct_excel).

Group 1: every value rule a worksheet can produce, each called by name and through PyCall
from the same script, the two results asserted equal as well as the documented one. Group
12: the call context. Group 10: what Insert Function and the Function Arguments dialog
show, read as a screen reader reads them, and that typing a formula there runs no Python.
Last, closing the add-in and opening it again with another list (docs/architecture.md,
"Lifetime").
"""

import ctypes
import re
import threading
import time
from ctypes import wintypes

import pytest
from conftest import E, bring_to_front, foreground, typed
from test_esc import INPUT, KEYEVENTF_KEYUP, INPUT_KEYBOARD, press_esc_once


@pytest.fixture
def book(direct_excel, tmp_path):
    """A workbook in the direct functions' Excel, closed after the test."""
    made = direct_excel.book(tmp_path)
    try:
        yield made
    finally:
        made.workbook.Close(SaveChanges=False)


def both(direct, book, function, *args, cells=("C1", "D1")):
    """The function's value in the first cell, called by its name, asserted equal to PyCall's
    in the second."""
    by_name = book.formula(cells[0], f"={function}({','.join(args)})")
    pycall = book.formula(cells[1], direct.pycall(function if function != "Give" else "give", *args))
    assert typed([[by_name]]) == typed([[pycall]]), (function, args)
    return by_name


# Group 1: values from Excel to Python (docs/conversion.md, "Excel to Python").


@pytest.mark.parametrize(
    "constant, arrives",
    [
        ("3", "float: 3.0"),
        ("3.5", "float: 3.5"),
        ("-2.25", "float: -2.25"),
        ("1E+300", "float: 1e+300"),
        ("TRUE", "bool: True"),
        ("FALSE", "bool: False"),
        ("text", "str: 'text'"),
        ("é ✓ 😀", "str: 'é ✓ 😀'"),
        ('=""', "str: ''"),
        ("=1/4", "float: 0.25"),
        ("=DATE(2026,1,2)", "float: 46024.0"),
        ("#NULL!", "ExcelError: ExcelError('#NULL!')"),
        ("=1/0", "ExcelError: ExcelError('#DIV/0!')"),
        ("=#VALUE!", "ExcelError: ExcelError('#VALUE!')"),
        ("=#REF!", "ExcelError: ExcelError('#REF!')"),
        ("=#NAME?", "ExcelError: ExcelError('#NAME?')"),
        ("=#NUM!", "ExcelError: ExcelError('#NUM!')"),
        ("=NA()", "ExcelError: ExcelError('#N/A')"),
        ("=FILTER({1,2},{FALSE,FALSE})", "ExcelError: ExcelError('#VALUE!')"),  # #CALC!, newer than the SDK.
    ],
)
def test_a_cell_arrives_as_pycall_passes_it(direct_excel, book, constant, arrives):
    book.put("A1", constant)
    assert both(direct_excel, book, "Describe", "A1") == arrives


def test_an_empty_cell_arrives_as_none(direct_excel, book):
    assert both(direct_excel, book, "Describe", "A1") == "NoneType: None"


@pytest.mark.parametrize(
    "argument, arrives",
    [
        ("A1:A1", "float: 1.0"),
        ("A1:C1", "tuple: ((1.0, 2.0, 3.0),)"),
        ("A1:A3", "tuple: ((1.0,), (4.0,), (7.0,))"),
        ("A1:C3", "tuple: ((1.0, 2.0, 3.0), (4.0, 5.0, 6.0), (7.0, None, 'nine'))"),
        ("{1}", "float: 1.0"),
        ("{1,2,3}", "tuple: ((1.0, 2.0, 3.0),)"),
        ("{1;2;3}", "tuple: ((1.0,), (2.0,), (3.0,))"),
        ('{1,"a";TRUE,#N/A}', "tuple: ((1.0, 'a'), (True, ExcelError('#N/A')))"),
        ("(A1:B2,C3)", "ExcelError: ExcelError('#VALUE!')"),
    ],
)
def test_a_range_or_array_arrives_as_rows_of_rows(direct_excel, book, argument, arrives):
    for cell, value in zip(["A1", "B1", "C1", "A2", "B2", "C2", "A3", "C3"], ["1", "2", "3", "4", "5", "6", "7", "nine"]):
        book.put(cell, value)
    assert both(direct_excel, book, "Describe", argument, cells=("E1", "F1")) == arrives


def test_text_at_excels_limit_arrives_whole(direct_excel, book):
    book.put("A1", '=REPT("😀",16383)')
    assert both(direct_excel, book, "Length", "A1") == 16383.0


def test_an_argument_left_out_in_the_middle_is_none(direct_excel, book):
    book.formula("A1", "=Pair(,3)")
    book.formula("A2", direct_excel.pycall("Pair", "", "3"))
    assert book.values("A1:B2") == [["NoneType: None", "float: 3.0"]] * 2


def test_an_argument_left_out_at_the_end_is_dropped(direct_excel, book):
    assert both(direct_excel, book, "EchoBack") == ""  # EchoBack(): its default, None.


# Group 1: values from Python to Excel (docs/conversion.md, "Python to Excel").


@pytest.mark.parametrize(
    "name, shows",
    [
        ("true", True),
        ("false", False),
        ("text", "text"),
        ("float", 2.5),
        ("int", 666.0),
        ("huge_int", float(1234567891011121314)),
        ("intenum", 7.0),
        ("float64", 1.25),
        ("unicode", "é ✓ 😀"),
        ("str_subclass", "text"),
        ("none", ""),
        ("nan", E("#NUM!")),
        ("inf", E("#NUM!")),
        ("too_large_int", E("#VALUE!")),
        ("long_text", E("#VALUE!")),
        ("emoji_too_long", E("#VALUE!")),
        ("dict", E("#VALUE!")),
        ("set", E("#VALUE!")),
        ("complex", E("#VALUE!")),
        ("object", E("#VALUE!")),
    ],
)
def test_a_return_value_shows_as_pycalls_does(direct_excel, book, name, shows):
    got = both(direct_excel, book, "Give", f'"{name}"')
    assert (type(got), got) == (type(shows), shows)


def test_none_shows_nothing_but_is_not_blank(direct_excel, book):
    assert book.formula("A1", '=Give("none")') == ""
    assert book.formula("B1", '=A1=""') is True
    assert book.formula("B2", "=ISBLANK(A1)") is False


@pytest.mark.parametrize("name", ["longest_text", "longest_emoji"])
def test_text_at_excels_limit_is_returned_whole(direct_excel, book, name):
    book.formula("A1", f'=Give("{name}")')
    assert len(book.value("A1").encode("utf-16-le")) // 2 == 32767
    assert book.formula("B2", f'=IsGiven("{name}", A1)') is True


@pytest.mark.parametrize("error", ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"])
def test_a_returned_excel_error_shows_that_error(direct_excel, book, error):
    assert both(direct_excel, book, "GiveError", f'"{error}"') == E(error)


@pytest.mark.parametrize(
    "name, cells, shows",
    [
        ("row", "A1:C1", [[1.0, 2.0, 3.0]]),
        ("column", "A1:A3", [[1.0], [2.0], [3.0]]),
        ("ragged", "A1:C3", [[1.0, 2.0, 3.0], [4.0, "", ""], [5.0, "", ""]]),
        ("empty", "A1:A1", [[""]]),
        ("empty_rows", "A1:A2", [[""], [""]]),
        ("empty_row_in_block", "A1:B2", [[1.0, 2.0], ["", ""]]),
        ("nan_in_block", "A1:B2", [[1.0, 2.0], [3.0, E("#NUM!")]]),
        ("list_subclass", "A1:C1", [[1.0, True, "x"]]),
        ("bools_in_block", "A1:B2", [[True, False], [1.0, 0.0]]),
        ("too_deep", "A1:B2", [[E("#VALUE!"), None], [None, None]]),
        ("bad_in_block", "A1:B2", [[E("#VALUE!"), None], [None, None]]),
    ],
)
def test_a_list_fills_a_block_as_pycalls_does(direct_excel, book, name, cells, shows):
    book.formula("A1", f'=Give("{name}")')
    book.formula("A11", direct_excel.pycall("give", f'"{name}"'))
    pycall_cells = ":".join(c[0] + str(int(c[1:]) + 10) for c in cells.split(":"))
    assert typed(book.values(cells)) == typed(shows)
    assert typed(book.values(pycall_cells)) == typed(shows)


@pytest.mark.parametrize(
    "formula, said",
    [
        ('=Give("dict")', "Pyinex: Give: dict cannot be returned to Excel"),
        ('=Give("bad_in_block")', "Pyinex: Give: row 2, column 2: dict cannot be returned to Excel"),
        ("=Raises()", "Pyinex: Raises in the script"),
        ("=Exits()", "SystemExit: 3"),
        ("=Interrupts()", "KeyboardInterrupt"),
        ("=Two(1)", "TypeError: Two() missing 1 required positional argument: 'b'"),
    ],
)
def test_a_failure_is_value_and_said_on_the_console(direct_excel, book, formula, said):
    direct_excel.cleared()
    assert book.formula("A1", formula) == E("#VALUE!")
    window = direct_excel.console(holding=said)
    assert window is not None and said in window.text, window and window.text
    assert book.formula("B1", '=EchoBack("still working")') == "still working"


def test_an_exception_shows_its_traceback(direct_excel, book):
    direct_excel.cleared()
    assert book.formula("A1", "=Raises()") == E("#VALUE!")
    window = direct_excel.console(holding="ValueError: raised on purpose")
    text = window.text if window else ""
    assert f"Pyinex: Raises in the script {direct_excel.script} raised an exception:" in text, text
    assert "Traceback (most recent call last)" in text and "<pyinex/" not in text


# Group 12: the call context (docs/architecture.md, "The pyinex module").


@pytest.mark.parametrize(
    "options, address",
    [((), "C12"), (("TRUE",), "R12C3"), (("FALSE", "TRUE"), "[book.xlsx]Sheet1!C12"), (("TRUE", "TRUE"), "[book.xlsx]Sheet1!R12C3")],
)
def test_caller_in_both_address_styles(direct_excel, book, options, address):
    assert book.formula("C12", f"=CallerOf({','.join(options)})") == address
    assert book.formula("C12", direct_excel.pycall("CallerOf", *options)) == address


def test_sheet_and_caller_are_the_calling_ones_not_the_active_ones(direct_excel, book, tmp_path):
    other = book.workbook.Worksheets.Add(After=book.sheet)
    other.Name = "Other"
    other.Activate()
    assert book.formula("C12", "=CallerOf(FALSE, TRUE)") == "[book.xlsx]Sheet1!C12"
    (tmp_path / "elsewhere").mkdir()
    elsewhere = direct_excel.book(tmp_path / "elsewhere", name="elsewhere.xlsx")
    try:
        elsewhere.sheet.Activate()
        assert book.formula("D4", "=CallerOf(FALSE, TRUE)") == "[book.xlsx]Sheet1!D4"
        assert book.formula("E5", "=SheetOf()") == "[book.xlsx]Sheet1"
    finally:
        elsewhere.workbook.Close(SaveChanges=False)


def test_caller_of_a_legacy_array_formula_is_its_range(direct_excel, book):
    assert book.array("C12:E14", "=CallerOf()") == [["C12:E14"] * 3] * 3


def test_caller_of_a_spilling_formula_is_its_cell(direct_excel, book):
    book.formula("B3", '=CallerOf() & {"", ""}')
    assert book.values("B3:C3") == [["B3", "B3"]]


@pytest.mark.parametrize("name", ["caller", "sheet"])
def test_caller_and_sheet_refuse_outside_a_cell(direct_excel, book, name):
    # Application.Run calls the function from no cell.
    assert direct_excel.app.Run("Attempt", name) == "RuntimeError"


def test_evaluate_calls_as_if_from_the_active_cell(direct_excel, book):
    book.sheet.Range("E5").Select()
    assert direct_excel.app.Evaluate("CallerOf()") == "E5"


@pytest.mark.parametrize("name", ["caller", "sheet", "esc_pressed"])
def test_asking_excel_from_a_scripts_thread_is_refused(direct_excel, book, name):
    assert both(direct_excel, book, "OnAThread", f'"{name}"') == "RuntimeError"


def test_a_scripts_thread_prints_between_calls(direct_excel, book):
    direct_excel.cleared()
    assert book.formula("A1", '=PrintLater("printed from a thread")') == "started"
    window = direct_excel.console(holding="printed from a thread", timeout=10)
    assert window is not None and "printed from a thread" in window.text


def test_the_gil_is_released_between_calls(direct_excel, book):
    assert book.formula("A1", "=StartTicking()") == "ticking"
    first = book.formula("A2", "=Ticked()")
    time.sleep(1)
    assert book.formula("A3", "=Ticked()") - first > 20


@pytest.mark.parametrize("clear, kept", [("FALSE", True), ("TRUE", False)])
def test_esc_pressed_kept_and_cleared(direct_excel, book, clear, kept):
    hwnd = direct_excel.app.Hwnd
    before = foreground()
    assert bring_to_front(hwnd), "Excel could not be made the foreground window"
    marker = book.folder / f"started {clear}"
    report, cancel = {}, threading.Event()
    sender = threading.Thread(target=press_esc_once, args=(marker, hwnd, cancel, report), daemon=True)
    sender.start()
    try:
        result = book.formula("A1", f'=LoopThenCheck("{marker}", {clear})')
    finally:
        cancel.set()
        sender.join(35)
        if before and before != hwnd:
            bring_to_front(before)
    assert "error" not in report, report["error"]
    assert report.get("sent"), "SendInput did not send Esc"
    assert result == f"stopped; still pressed: {kept}"


# Group 10: Insert Function and the Function Arguments dialog.

VK_SHIFT, VK_CONTROL, VK_MENU, VK_ESCAPE, VK_F3 = 0x10, 0x11, 0x12, 0x1B, 0x72
KEYEVENTF_UNICODE = 0x0004


def send(*events):
    """Sends keys as one SendInput: each event a (virtual key, up) pair, or a str typed as
    Unicode characters."""
    flat = []
    for e in events:
        if isinstance(e, str):
            for unit in memoryview(e.encode("utf-16-le")).cast("H"):  # UTF-16 code units.
                flat += [(0, unit, KEYEVENTF_UNICODE), (0, unit, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP)]
        else:
            key, up = e
            flat.append((key, 0, KEYEVENTF_KEYUP if up else 0))
    keys = (INPUT * len(flat))()
    for i, (vk, scan, flags) in enumerate(flat):
        keys[i].type, keys[i].ki.wVk, keys[i].ki.wScan, keys[i].ki.dwFlags = INPUT_KEYBOARD, vk, scan, flags
    assert ctypes.windll.user32.SendInput(len(flat), keys, ctypes.sizeof(INPUT)) == len(flat), "SendInput failed"


def chord(*keys):
    """Keys held down in order, then released in reverse: chord(VK_CONTROL, ord("A"))."""
    return [(k, False) for k in keys] + [(k, True) for k in reversed(keys)]


def dialog(pid, title, timeout=10.0, present=True):
    """The window of Excel's dialog with this title (class bosa_sdm_XL...), or None; waits
    up to timeout for it to appear, or, with present=False, to go."""
    import win32gui
    import win32process

    deadline = time.monotonic() + timeout
    while True:
        found = []

        def visit(hwnd, _):
            if (win32process.GetWindowThreadProcessId(hwnd)[1] == pid and win32gui.GetClassName(hwnd).startswith("bosa_sdm_XL")
                    and win32gui.GetWindowText(hwnd) == title and win32gui.IsWindowVisible(hwnd)):
                found.append(hwnd)

        win32gui.EnumWindows(visit, None)
        if bool(found) == present or time.monotonic() > deadline:
            return found[0] if found else None
        time.sleep(0.1)


class GUID(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD), ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]


IID_IDISPATCH = GUID(0x00020400, 0, 0, (ctypes.c_ubyte * 8)(0xC0, 0, 0, 0, 0, 0, 0, 0x46))
OBJID_CLIENT = 0xFFFFFFFC
ACC_CHILD_COUNT, ACC_CHILD, ACC_NAME, ACC_VALUE, ACC_DESCRIPTION = -5001, -5002, -5003, -5004, -5005


def accessible_texts(hwnd):
    """Every name, value and description the window's accessible objects give, as a screen
    reader reads them (Microsoft Active Accessibility): how a test reads a dialog of Excel's,
    which draws its own controls."""
    import pythoncom

    pointer = ctypes.c_void_p()
    ctypes.oledll.oleacc.AccessibleObjectFromWindow(
        wintypes.HWND(hwnd), wintypes.DWORD(OBJID_CLIENT), ctypes.byref(IID_IDISPATCH), ctypes.byref(pointer))
    root = pythoncom.ObjectFromAddress(pointer.value, pythoncom.IID_IDispatch)

    def get(obj, dispid, *args):
        try:
            return obj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True, *args)
        except pythoncom.com_error:
            return None

    texts = []

    def read(obj, child):
        texts.extend(t for t in (get(obj, d, child) for d in (ACC_NAME, ACC_VALUE, ACC_DESCRIPTION)) if t)

    def walk(obj, depth):
        for i in range(1, (get(obj, ACC_CHILD_COUNT) or 0) + 1):
            child = get(obj, ACC_CHILD, i)
            if child is None:
                read(obj, i)
            elif depth < 20:
                read(child, 0)
                walk(child, depth + 1)

    read(root, 0)
    walk(root, 0)
    return texts


@pytest.fixture
def in_front(direct_excel, book):
    """The direct functions' Excel in front, with A1 of a new workbook selected, for keys
    sent as a keyboard sends them; what was in front before is put back after."""
    before = foreground()
    book.sheet.Range("A1").Select()
    assert bring_to_front(direct_excel.app.Hwnd), "Excel could not be made the foreground window"
    try:
        yield direct_excel
    finally:
        if before and before != direct_excel.app.Hwnd:
            bring_to_front(before)


def test_the_function_arguments_dialog_shows_the_help_and_runs_no_python(in_front, book):
    marker = book.folder / "typed in the dialog.txt"
    send("=Marked(", *chord(VK_CONTROL, ord("A")))
    window = dialog(in_front.pid, "Function Arguments")
    assert window, "the Function Arguments dialog did not open"
    try:
        send(f'"{marker}"')
        time.sleep(1.5)  # Excel calculates the formula in the dialog as it is typed.
        texts = accessible_texts(window)
    finally:
        send(*chord(VK_ESCAPE))
        dialog(in_front.pid, "Function Arguments", present=False)
        send(*chord(VK_ESCAPE))
    shown = "\n".join(texts)
    assert "Marked" in shown, shown
    assert "Writes the file named, so a test can tell it ran." in shown, shown
    assert "The file to write." in shown, shown
    # A blank: the dialog shows the function's value as "" and no formula result, and no
    # other text in quotes but the argument typed.
    assert '""' in texts and "Formula result =" in texts, shown
    assert not [t for t in texts if t.startswith('"') and t not in ('""', f'"{marker}"')], shown
    assert not marker.exists(), "Python ran in the Function Arguments dialog"
    assert book.formula("B1", f'=Marked("{marker}")') == "ran" and marker.exists()


def test_insert_function_lists_the_category_and_its_help(in_front, book):
    send(*chord(VK_SHIFT, VK_F3))
    window = dialog(in_front.pid, "Insert Function")
    assert window, "the Insert Function dialog did not open"
    seen = []
    try:
        send(*chord(VK_MENU, ord("C")))  # Or select a category.
        for _ in range(10):  # Each P moves to the next category starting with P.
            send(*chord(ord("P")))
            time.sleep(0.3)
            seen = accessible_texts(window)
            if any(t.startswith("MarkedToo,") for t in seen):
                break
    finally:
        for _ in range(3):  # The first may only close the category's list.
            send(*chord(VK_ESCAPE))
            if not dialog(in_front.pid, "Insert Function", timeout=2, present=False):
                break
    shown = "\n".join(seen)
    # The category's functions, each with its help, and the selected one's arguments.
    # (The list's own value, its selected item cut short, is among the texts too.)
    assert "Marked, Writes the file named, so a test can tell it ran." in seen, shown
    assert "MarkedToo, Help given to the decorator." in seen, shown
    assert not [t for t in seen if t.startswith(("Describe,", "EchoBack,", "PyCall,"))], shown  # Other categories'.
    assert "Marked(path,note)" in seen, shown


# Closing (docs/architecture.md, "Lifetime"): the add-in is closed with UNREGISTER and
# opened with RegisterXLL, as tests/excel/test_lifetime.py does.


def reopened(direct, listed):
    """Closes the add-in, lists other scripts in its pyinex.ini, and opens it again."""
    ini = direct.folder / "pyinex.ini"
    ini.write_text(re.sub(r"(?m)^scripts = .*$", f"scripts = {listed}", ini.read_text("utf-8")), "utf-8")
    assert direct.app.ExecuteExcel4Macro(f'UNREGISTER("{direct.xll}")') is True
    assert direct.app.RegisterXLL(str(direct.xll))


def test_reopened_with_another_list_an_old_name_never_reaches_the_new_function(direct_excel, book):
    for name in ("OldOne", "NewOne"):
        text = f"import pyinex\n\n@pyinex.func\ndef {name}(x):\n    return '{name}: ' + str(x)\n"
        (direct_excel.folder / f"{name}.py").write_text(text, "utf-8")
    try:
        reopened(direct_excel, "OldOne.py")
        assert book.formula("A1", "=OldOne(1)") == "OldOne: 1.0"
        reopened(direct_excel, "NewOne.py")
        slots = {p for dll, p, _ in direct_excel.app.RegisteredFunctions if p.startswith("PyFunc")}
        assert slots == {"PyFunc0"}  # NewOne in the slot OldOne had; direct.py's all removed.
        assert book.formula("A2", "=NewOne(2)") == "NewOne: 2.0"
        direct_excel.app.CalculateFull()
        assert book.value("A1") == E("#VALUE!")  # OldOne's cell: an error, never NewOne's value.
        assert book.formula("A3", "=OldOne(3)") == E("#VALUE!")  # Typed again.
    finally:
        reopened(direct_excel, "direct.py")
    assert book.formula("A4", "=EchoBack(4)") == book.formula("B4", direct_excel.pycall("EchoBack", "4"))
