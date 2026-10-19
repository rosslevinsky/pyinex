"""Fixtures for the real-Excel tests: one Excel for the session, a workbook per test.

These tests drive Microsoft Excel itself, through COM, with pywin32. They run only when
asked, with the add-in to test:

    py -3.14 -m pytest tests/excel --excel C:\\path\\to\\pyinex.xll

and are normally run by scripts/excel_tests.py, which runs them once per Python and
leaves the machine as it found it (docs/testing.md). Without --excel nothing here is
collected, so `pytest tests` and CI never start Excel. They need Windows, 64-bit Excel,
pywin32 and the signed-in desktop: Excel cannot be automated from an SSH or service
session. Lacking any of those, they skip, saying why, and the driver counts a skip as a
failure.

Excel embeds whichever Python the add-in chooses, not the one running pytest. The driver
picks it with a pyinex.ini beside the add-in and passes --excel-run, a folder holding
expected.json (the Python it should be); the session checks Excel embedded that one before
any test runs, and writes excel.json there (Excel's process id and start time, version and
build) for the driver, which ends that Excel if it has not quit once pytest has exited. Each
other Excel a test starts is reported the same way, in an "excel <name>.json" file
(report_excel).
"""

import ctypes
import gc
import itertools
import json
import math
import re
import shutil
import sys
import time
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
EXCEL_PID = pytest.StashKey[int]()
ROOT = HERE.parents[1]
SCRIPTS = HERE / "scripts"  # Scripts the tests call, copied beside a test's workbook.
ITEMS = ["status", "python", "pyinex", "version", "executable", "prefix", "venv", "addin_version"]  # PyInfo's.

# The numbers Excel's COM interface gives for error values: -2146826288 (#NULL!, error
# 2000) plus the error's number less 2000.
ERRORS = {
    -2146826288: "#NULL!",
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826245: "#GETTING_DATA",
    -2146826243: "#SPILL!",
    -2146826238: "#CALC!",
}


def process_created(pid):
    """When the process started, in FILETIME units, or None: with the pid, the driver's
    proof that a process is still this session's Excel (scripts/excel_tests.py, created)."""
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.GetProcessTimes.argtypes = [ctypes.c_void_p, *[ctypes.POINTER(ctypes.c_ulonglong)] * 4]
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    handle = kernel32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not handle:
        return None
    try:
        times = [ctypes.c_ulonglong() for _ in range(4)]
        return times[0].value if kernel32.GetProcessTimes(handle, *[ctypes.byref(t) for t in times]) else None
    finally:
        kernel32.CloseHandle(handle)

def pytest_addoption(parser):
    group = parser.getgroup("excel", "real Excel")
    group.addoption("--excel", metavar="XLL", help="run the real-Excel tests on this pyinex.xll")
    group.addoption("--excel-run", metavar="FOLDER", help="the driver's folder for this run (expected.json in, excel.json out)")


def pytest_ignore_collect(collection_path, config):
    """Nothing here is collected without --excel; driver_checks/ only when named."""
    path = Path(collection_path)
    if path == HERE / "conftest.py":
        return None
    if not config.getoption("--excel", default=None):
        return True
    checks = HERE / "driver_checks"
    if path == checks or checks in path.parents:
        named = [Path(a.split("::")[0]).resolve() for a in config.args]
        if not any(n == checks or checks in n.parents for n in named):
            return True
    return None


def pytest_configure(config):
    config.addinivalue_line(
        "markers",
        "kind(name): run only in the driver's run of that kind (refused, console_first, first_call, "
        "lifetime), each in an Excel of its own; unmarked tests run in the per-Python runs",
    )


def run_kind(config):
    """The kind of run the driver started: "normal" unless its expected.json says else."""
    folder = config.getoption("--excel-run", default=None)
    if not folder:
        return "normal"
    return json.loads((Path(folder) / "expected.json").read_text("utf-8")).get("kind", "normal")


def pytest_collection_modifyitems(config, items):
    """A file named on the command line is collected whatever pytest_ignore_collect says, so
    without --excel its tests are deselected here; with it, so are the tests of another
    kind of run than this one."""
    kind = run_kind(config) if config.getoption("--excel", default=None) else None

    def keep(item):
        if HERE not in Path(item.path).parents:
            return True
        if kind is None:
            return False
        marker = item.get_closest_marker("kind")
        return (marker.args[0] if marker else "normal") == kind

    dropped = [item for item in items if not keep(item)]
    if dropped:
        config.hook.pytest_deselected(items=dropped)
        items[:] = [item for item in items if keep(item)]


def pycall(script, function, *args):
    """A PyCall formula; each argument is formula text, so "A1", '"text"' or "1/0"."""
    return "=PyCall(" + ",".join([f'"{script}"', f'"{function}"', *args]) + ")"


class E:
    """An Excel error value read from a cell, by its name: E("#N/A"). It equals only an E
    of the same name, never text, so a cell holding the text "#N/A" does not pass as the
    error."""

    def __init__(self, name):
        self.name = name

    def __eq__(self, other):
        return isinstance(other, E) and other.name == self.name

    def __hash__(self):
        return hash(self.name)

    def __repr__(self):
        return f"E({self.name!r})"


def python_dlls(pid):
    """The Python DLLs (python3*.dll) loaded in a process, by their paths."""
    return modules(pid, r"python3\d*\.dll")


def modules(pid, name):
    """The modules loaded in a process whose file names match the pattern name, by path."""
    from ctypes import wintypes

    class MODULEENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD), ("th32ModuleID", wintypes.DWORD), ("th32ProcessID", wintypes.DWORD),
            ("GlblcntUsage", wintypes.DWORD), ("ProccntUsage", wintypes.DWORD), ("modBaseAddr", ctypes.c_void_p),
            ("modBaseSize", wintypes.DWORD), ("hModule", ctypes.c_void_p), ("szModule", ctypes.c_wchar * 256),
            ("szExePath", ctypes.c_wchar * 260),
        ]  # fmt: skip

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateToolhelp32Snapshot.restype = ctypes.c_void_p
    kernel32.Module32FirstW.argtypes = kernel32.Module32NextW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    for _ in range(20):  # Windows may fail the snapshot while the process loads a module.
        snapshot = kernel32.CreateToolhelp32Snapshot(0x18, pid)  # TH32CS_SNAPMODULE | TH32CS_SNAPMODULE32
        if snapshot not in (None, ctypes.c_void_p(-1).value):
            break
        time.sleep(0.05)
    else:
        raise AssertionError("cannot list Excel's modules")
    found, entry = [], MODULEENTRY32W()
    entry.dwSize = ctypes.sizeof(entry)
    try:
        # A process always has modules, so a list that will not start is a failure, and one
        # that stops early must stop for want of more (ERROR_NO_MORE_FILES), not an error.
        assert kernel32.Module32FirstW(snapshot, ctypes.byref(entry)), "cannot list Excel's modules"
        while True:
            if re.fullmatch(name, entry.szModule, re.IGNORECASE):
                found.append(entry.szExePath)
            if not kernel32.Module32NextW(snapshot, ctypes.byref(entry)):
                assert ctypes.get_last_error() == 18, "the list of Excel's modules stopped early"
                break
    finally:
        kernel32.CloseHandle(snapshot)
    return found


def shown(value):
    """A value read from Excel, with an error value made an E."""
    if isinstance(value, int) and not isinstance(value, bool) and value in ERRORS:
        return E(ERRORS[value])
    return value


def typed(rows):
    """Rows of cells as (type, value) pairs, so 1.0 and TRUE, equal in Python, differ."""
    return [[(type(v).__name__, v) for v in row] for row in rows]


def rows_of(value):
    """A range's Value2 as rows, whatever its size."""
    return [list(r) for r in value] if isinstance(value, tuple) else [[value]]


def registered_types(app, xll):
    """The numbered exports Excel has registered for the add-in xll, each with its type
    string, read from Excel's registration table (Application.RegisteredFunctions: DLL,
    procedure, type string). The reader is checked first against a registration known to be
    there, PyCall's, so that an export it does not find is absent from the table, not missed
    by the reader."""
    rows = [(Path(dll).resolve(), procedure, type_) for dll, procedure, type_ in (app.RegisteredFunctions or ())]
    mine = {procedure: type_ for dll, procedure, type_ in rows if dll == Path(xll).resolve()}
    assert mine.get("PyCall") == "Q" * 33, f"the registration table's reader did not find PyCall's registration: {rows}"
    return {procedure: type_ for procedure, type_ in mine.items() if procedure.startswith("PyFunc")}


def same(got, expected):
    """A cell's value equals what was expected; floats to 12 significant figures."""
    if expected is None:
        return got in (None, "")  # Pyinex returns empty text for a blank cell.
    if isinstance(expected, float) and isinstance(got, float):
        return math.isclose(got, expected, rel_tol=1e-12)
    return type(got) is type(expected) and got == expected


class Book:
    """A saved workbook of one sheet, in a test's own folder."""

    def __init__(self, excel, workbook):
        self.excel = excel
        self.workbook = workbook
        self.sheet = workbook.Worksheets(1)
        self.folder = Path(workbook.Path)

    def formula(self, cell, formula):
        """Enters a formula (spilling, as Microsoft 365 does) and returns its value."""
        self.sheet.Range(cell).Formula2 = formula
        return self.value(cell)

    def array(self, cells, formula):
        """Enters a legacy array formula over the cells and returns their values, as rows."""
        self.sheet.Range(cells).FormulaArray = formula
        return self.values(cells)

    def value(self, cell):
        return shown(self.sheet.Range(cell).Value2)

    def values(self, cells):
        return [[shown(v) for v in row] for row in rows_of(self.sheet.Range(cells).Value2)]

    def script(self, name, text):
        """Writes a script beside the workbook and returns its name."""
        (self.folder / name).write_text(text, encoding="utf-8")
        return name

    def copy(self, name):
        """Copies a script from tests/excel/scripts beside the workbook; returns its name."""
        shutil.copyfile(SCRIPTS / name, self.folder / name)
        return name

    def put(self, cell, value):
        """Puts a constant, or a formula's text, in a cell, as typing it would."""
        self.sheet.Range(cell).Formula2 = value


def _win32():
    if sys.platform != "win32":
        pytest.skip("real-Excel tests need Windows")
    try:
        import pywintypes  # noqa: F401
        import win32api  # noqa: F401
        from win32com.client import DispatchEx  # noqa: F401
    except ImportError:
        pytest.skip("real-Excel tests need pywin32 (py -m pip install pywin32)")


# Excel rejects a call from another process while it is busy, as it is while it runs a
# macro such as Pyinex's queue command just after a workbook opens or a PyLoad cell
# calculates (docs/architecture.md, "When a workbook opens"). With no message filter, COM
# fails such a call at once (RPC_E_CALL_REJECTED), so the session registers one on its
# thread, as Microsoft advises a program automating Excel to do: COM then tries a call
# Excel asked to retry later every 100 ms, for up to a minute. pywin32 offers no
# IMessageFilter, so it is made here with ctypes. Kept for the session, as COM holds it;
# the filter the thread had before is put back when the session ends.
_message_filter = []


def _install_message_filter(pytestconfig):
    from ctypes import POINTER, WINFUNCTYPE, byref, c_long, c_ulong, c_void_p
    from ctypes.wintypes import DWORD

    import pythoncom  # noqa: F401  Initializes COM on this thread, as a single-threaded apartment.

    interfaces = {bytes.fromhex(f"{n:02x}000000" "0000" "0000" "c000000000000046") for n in (0x00, 0x16)}  # IUnknown, IMessageFilter
    retry_later, handled, wait_default = 2, 0, 2  # SERVERCALL_RETRYLATER, SERVERCALL_ISHANDLED, PENDINGMSG_WAITDEFPROCESS

    @WINFUNCTYPE(c_long, c_void_p, c_void_p, POINTER(c_void_p))
    def query_interface(this, iid, out):
        found = ctypes.string_at(iid, 16) in interfaces
        out[0] = this if found else None
        return 0 if found else -2147467262  # S_OK, E_NOINTERFACE

    @WINFUNCTYPE(c_ulong, c_void_p)
    def count(this):  # AddRef and Release: the filter lives as long as the process.
        return 1

    @WINFUNCTYPE(DWORD, c_void_p, DWORD, c_void_p, DWORD, c_void_p)
    def handle_incoming_call(this, kind, caller, ticks, info):
        return handled

    @WINFUNCTYPE(DWORD, c_void_p, c_void_p, DWORD, DWORD)
    def retry_rejected_call(this, callee, ticks, reject):
        return 100 if reject == retry_later and ticks < 60_000 else 0xFFFFFFFF  # Wait 100 ms, or cancel.

    @WINFUNCTYPE(DWORD, c_void_p, c_void_p, DWORD, DWORD)
    def message_pending(this, callee, ticks, kind):
        return wait_default

    functions = (query_interface, count, count, handle_incoming_call, retry_rejected_call, message_pending)
    vtable = (c_void_p * len(functions))(*(ctypes.cast(f, c_void_p) for f in functions))
    instance = (c_void_p * 1)(ctypes.addressof(vtable))
    register = ctypes.windll.ole32.CoRegisterMessageFilter
    register.argtypes = [c_void_p, POINTER(c_void_p)]
    previous = c_void_p()
    if register(ctypes.addressof(instance), byref(previous)) != 0:
        pytest.exit("cannot register a COM message filter", returncode=2)
    _message_filter.extend([functions, vtable, instance])

    def restore():
        # COM takes its own reference to the filter put back, and releases this one's.
        register(previous, None)
        if previous:
            release = WINFUNCTYPE(c_ulong, c_void_p)(ctypes.cast(previous, POINTER(POINTER(c_void_p)))[0][2])
            release(previous)  # The reference CoRegisterMessageFilter handed back.

    pytestconfig.add_cleanup(restore)  # After every session fixture, Excel's included.


@pytest.fixture(scope="session")
def run_folder(pytestconfig):
    """The driver's folder for this run, or None when pytest is run by hand."""
    folder = pytestconfig.getoption("--excel-run", default=None)
    return Path(folder) if folder else None


def report_excel(run_folder, name, pid):
    """Writes an Excel's process id and start time to "excel <name>.json" in the driver's
    folder, so the driver waits for that Excel too once pytest has exited, and fails the run
    if it has to end it (scripts/excel_tests.py). Nothing without the driver."""
    if run_folder is not None:
        report = {"pid": pid, "created": process_created(pid)}
        (run_folder / f"excel {name}.json").write_text(json.dumps(report), "utf-8")


@pytest.fixture(scope="session")
def expected(run_folder):
    """What the driver says Excel should embed: executable, version, venv."""
    if run_folder is None:
        return None
    return json.loads((run_folder / "expected.json").read_text("utf-8"))


@pytest.fixture(scope="session")
def excel(pytestconfig, run_folder, expected):
    """A new Excel with the add-in loaded and Python started, ended after the session."""
    _win32()
    _install_message_filter(pytestconfig)
    import pywintypes
    import win32process
    from win32com.client import DispatchEx

    # Window positions and sizes in real pixels, as Excel and Pyinex use them, not scaled.
    ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # PER_MONITOR_AWARE_V2
    xll = Path(pytestconfig.getoption("--excel")).resolve()
    if not xll.is_file() or not (xll.parent / "pyinex_core.dll").is_file():
        pytest.exit(f"{xll} or pyinex_core.dll beside it not found", returncode=2)
    report = {}

    def write_report():
        if run_folder is not None:
            (run_folder / "excel.json").write_text(json.dumps(report, indent=2), "utf-8")

    try:
        app = DispatchEx("Excel.Application")  # A new Excel, not one the user has open.
    except pywintypes.com_error as e:
        pytest.skip(f"cannot start Excel: {e}")
    report["pid"] = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    report["created"] = process_created(report["pid"])
    write_report()  # At once, so the driver can end this Excel if the run hangs.
    try:
        app.Visible = True
        app.DisplayAlerts = False
        report.update(version=app.Version, build=app.Build)
        write_report()
        if not app.RegisterXLL(str(xll)):
            pytest.exit(f"Excel would not load {xll}", returncode=2)
        pytestconfig.stash[EXCEL_PID] = report["pid"]
        yield app
    finally:
        try:
            for workbook in list(app.Workbooks):
                workbook.Close(SaveChanges=False)
            app.Quit()
        except pywintypes.com_error:
            pass
        app = None
        gc.collect()
        # Excel ends only once nothing refers to it, and pytest holds this fixture's value
        # until the session is over: the driver waits for Excel after pytest has exited.


@pytest.fixture(scope="session")
def embedded(excel, expected):
    """The Python Excel embedded, checked against the driver's expectation before any test
    that uses Excel; a mismatch ends the session, since every result would be another
    Python's."""
    found = {
        "executable": excel.Evaluate('PyInfo("executable")'),
        "version": excel.Evaluate('PyInfo("version")'),
        "venv": excel.Evaluate('PyInfo("venv")'),
    }
    if expected is not None:
        wrong = [
            f"{k}: expected {expected[k]!r}, Excel embedded {found[k]!r}"
            for k in ("executable", "version", "venv")
            if (str(found[k]).lower() if k == "executable" else found[k])
            != (str(expected[k]).lower() if k == "executable" else expected[k])
        ]
        if wrong:
            pytest.exit("wrong Python in Excel: " + "; ".join(wrong), returncode=3)
    return found


@pytest.fixture
def app(excel, embedded):
    """The session's Excel, once its Python is known to be the one intended."""
    return excel


@pytest.fixture
def book(app, tmp_path):
    """A new workbook saved as book.xlsx in the test's own folder, closed afterwards."""
    workbook = app.Workbooks.Add()
    while workbook.Worksheets.Count > 1:
        workbook.Worksheets(workbook.Worksheets.Count).Delete()
    workbook.SaveAs(str(tmp_path / "book.xlsx"), FileFormat=51)  # 51: .xlsx
    try:
        yield Book(app, workbook)
    finally:
        workbook.Close(SaveChanges=False)


@pytest.fixture
def workbook(app):
    """A function making a workbook of one sheet, named sheet if given, saved as path (an
    .xlsx) or, with no path, left unsaved; each is closed after the test. Excel will not
    save into a folder whose name holds brackets, but opens a workbook there, so such a
    workbook is saved beside its folder, copied in, and opened."""
    made = []

    def make(path=None, sheet=None):
        workbook = app.Workbooks.Add()
        made.append(workbook)
        while workbook.Worksheets.Count > 1:
            workbook.Worksheets(workbook.Worksheets.Count).Delete()
        if sheet:
            workbook.Worksheets(1).Name = sheet
        if path:
            path = Path(path)
            path.parent.mkdir(parents=True, exist_ok=True)
            if "[" not in str(path.parent) and "]" not in str(path.parent):
                workbook.SaveAs(str(path), FileFormat=51)  # 51: .xlsx
            else:
                clean = next(p for p in path.parents if "[" not in str(p) and "]" not in str(p))
                staged = clean / f"staged {path.name}"
                workbook.SaveAs(str(staged), FileFormat=51)
                made.remove(workbook)
                workbook.Close(SaveChanges=False)
                shutil.copyfile(staged, path)
                workbook = app.Workbooks.Open(str(path))
                made.append(workbook)
        return Book(app, workbook)

    try:
        yield make
    finally:
        for workbook in made:
            workbook.Close(SaveChanges=False)


@pytest.fixture
def examples(tmp_path):
    """A copy of examples/, so the scripts' compiled files are written there, not in the
    repository."""
    folder = tmp_path / "examples"
    shutil.copytree(ROOT / "examples", folder, ignore=shutil.ignore_patterns("__pycache__"))
    return folder


# Pyinex's console window.

CONSOLE_CLASS = "PyinexConsole"


class Console:
    """Pyinex's console window in the session's Excel, read from outside its process."""

    def __init__(self, hwnd):
        self.hwnd = hwnd

    @staticmethod
    def find(pid):
        import win32gui
        import win32process

        found = []

        def visit(hwnd, _):
            if win32process.GetWindowThreadProcessId(hwnd)[1] == pid and win32gui.GetClassName(hwnd) == CONSOLE_CLASS:
                found.append(hwnd)

        win32gui.EnumWindows(visit, None)
        return Console(found[0]) if found else None

    @property
    def title(self):
        import win32gui

        return win32gui.GetWindowText(self.hwnd)

    @property
    def shown(self):
        import win32gui

        return bool(win32gui.IsWindowVisible(self.hwnd))

    @property
    def rect(self):
        """(left, top, width, height) in pixels."""
        import win32gui

        left, top, right, bottom = win32gui.GetWindowRect(self.hwnd)
        return left, top, right - left, bottom - top

    def close(self):
        """What the close button does."""
        import win32con
        import win32gui

        win32gui.SendMessage(self.hwnd, win32con.WM_CLOSE, 0, 0)

    def colors(self):
        """How many pixels of each color the window's text box shows, as {COLORREF: count},
        drawn by the window itself (PrintWindow), so a window over it does not matter."""
        import win32gui

        box = win32gui.FindWindowEx(self.hwnd, 0, "EDIT", None)
        left, top, right, bottom = win32gui.GetClientRect(box)
        width, height = right - left, bottom - top
        user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
        for f, args, result in [
            (user32.GetDC, [ctypes.c_void_p], ctypes.c_void_p),
            (user32.ReleaseDC, [ctypes.c_void_p, ctypes.c_void_p], ctypes.c_int),
            (user32.PrintWindow, [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint], ctypes.c_int),
            (gdi32.CreateCompatibleDC, [ctypes.c_void_p], ctypes.c_void_p),
            (gdi32.CreateCompatibleBitmap, [ctypes.c_void_p, ctypes.c_int, ctypes.c_int], ctypes.c_void_p),
            (gdi32.SelectObject, [ctypes.c_void_p, ctypes.c_void_p], ctypes.c_void_p),
            (gdi32.GetBitmapBits, [ctypes.c_void_p, ctypes.c_long, ctypes.c_void_p], ctypes.c_long),
            (gdi32.DeleteObject, [ctypes.c_void_p], ctypes.c_int),
            (gdi32.DeleteDC, [ctypes.c_void_p], ctypes.c_int),
        ]:
            f.argtypes, f.restype = args, result
        screen = user32.GetDC(None)
        memory = gdi32.CreateCompatibleDC(screen)
        bitmap = gdi32.CreateCompatibleBitmap(screen, width, height)
        old = gdi32.SelectObject(memory, bitmap)
        try:
            assert user32.PrintWindow(box, memory, 2), "PrintWindow failed"  # PW_RENDERFULLCONTENT
            pixels = ctypes.create_string_buffer(width * height * 4)
            gdi32.GetBitmapBits(bitmap, len(pixels), pixels)
        finally:
            gdi32.SelectObject(memory, old)
            gdi32.DeleteObject(bitmap)
            gdi32.DeleteDC(memory)
            user32.ReleaseDC(None, screen)
        counts = {}
        raw = pixels.raw
        for i in range(0, len(raw), 4):
            b, g, r = raw[i], raw[i + 1], raw[i + 2]  # Blue, green, red, unused.
            color = r | g << 8 | b << 16  # A COLORREF.
            counts[color] = counts.get(color, 0) + 1
        return counts

    @property
    def text(self):
        """The text box's text. GetWindowText gives only a caption across processes; a text
        box's text must be asked for."""
        import win32con
        import win32gui

        box = win32gui.FindWindowEx(self.hwnd, 0, "EDIT", None)
        if not box:
            return ""
        send = ctypes.windll.user32.SendMessageW
        send.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_void_p]
        send.restype = ctypes.c_ssize_t
        buffer = ctypes.create_unicode_buffer(send(box, win32con.WM_GETTEXTLENGTH, 0, None) + 1)
        send(box, win32con.WM_GETTEXT, len(buffer), ctypes.cast(buffer, ctypes.c_void_p))
        return buffer.value


@pytest.fixture
def console(app, pytestconfig):
    """A function returning Pyinex's console window, waiting up to five seconds for it and
    for any text given to appear in it (the add-in writes output from another thread)."""

    def get(holding=None, timeout=5.0):
        deadline = time.monotonic() + timeout
        while True:
            window = Console.find(pytestconfig.stash[EXCEL_PID])
            if window is not None and (holding is None or holding in window.text):
                return window
            if time.monotonic() > deadline:
                return window
            time.sleep(0.05)

    return get


@pytest.fixture
def mark(app, console, tmp_path_factory):
    """A function that prints a new marker through Python and returns the console's text
    once the marker is in it. Output reaches the console in the order it was written, so
    the text then holds everything written before the call."""
    script = tmp_path_factory.mktemp("mark") / "mark.py"
    script.write_text("def mark(text):\n    print(text)\n", encoding="utf-8")
    count = itertools.count()

    def mark():
        text = f"marker {next(count)} {time.monotonic_ns()}"
        assert app.Evaluate(f'PyCall("{script}", "mark", "{text}")') == ""
        window = console(holding=text, timeout=10)
        assert window is not None and text in window.text, f"{text!r} never reached the console"
        return window.text

    return mark


def foreground():
    user32 = ctypes.windll.user32
    user32.GetForegroundWindow.restype = ctypes.c_void_p
    return user32.GetForegroundWindow() or 0


def bring_to_front(hwnd):
    """Makes hwnd the foreground window without sending any key: this thread shares input
    with the current foreground window's for a moment, which lets it hand the foreground
    on (AttachThreadInput). True if hwnd is then in front."""
    user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
    user32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
    user32.BringWindowToTop.argtypes = [ctypes.c_void_p]
    current = foreground()
    mine, theirs = kernel32.GetCurrentThreadId(), user32.GetWindowThreadProcessId(current, None) if current else 0
    attached = bool(theirs) and theirs != mine and user32.AttachThreadInput(mine, theirs, True)
    try:
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
    finally:
        if attached:
            user32.AttachThreadInput(mine, theirs, False)
    deadline = time.monotonic() + 5
    while foreground() != hwnd and time.monotonic() < deadline:
        time.sleep(0.05)
    return foreground() == hwnd


@pytest.fixture
def excel_in_front(app):
    """Excel as the foreground window for the test; the window that was in front before is
    put back after."""
    before = foreground()
    assert bring_to_front(app.Hwnd), "Excel could not be made the foreground window"
    try:
        yield app.Hwnd
    finally:
        if before and before != app.Hwnd:
            bring_to_front(before)


@pytest.fixture
def desktop():
    """The desktop's left, top, width and height, spanning every monitor."""
    metric = ctypes.windll.user32.GetSystemMetrics
    return metric(76), metric(77), metric(78), metric(79)  # SM_XVIRTUALSCREEN and so on.


PERSONALIZE = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"


@pytest.fixture
def app_mode():
    """A function setting Windows' app mode to "light" or "dark", as Settings does: the
    registry value, then the message telling open windows. It is put back after the test,
    and the driver checks that it was."""
    import winreg

    def read():
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, PERSONALIZE) as key:
            try:
                return winreg.QueryValueEx(key, "AppsUseLightTheme")[0]
            except FileNotFoundError:
                return None

    def announce():
        result = ctypes.c_ulong()
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "ImmersiveColorSet", 0x0002, 5000, ctypes.byref(result))

    def set_to(light):
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, PERSONALIZE, 0, winreg.KEY_SET_VALUE) as key:
            if light is None:
                winreg.DeleteValue(key, "AppsUseLightTheme")
            else:
                winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, light)
        announce()

    before = read()
    try:
        yield lambda mode: set_to(1 if mode == "light" else 0)
    finally:
        set_to(before)


@pytest.fixture
def cleared(app, console, mark):
    """The console function, with the console cleared first, so text found in it was
    written by this test: earlier output is waited for before clearing."""
    mark()
    app.Evaluate("PyConsole(FALSE, TRUE)")
    return console


# An Excel of a test's own, for direct functions.


class DirectExcel:
    """An Excel of its own whose add-in, a copy in a folder of its own, lists
    scripts/direct.py in its pyinex.ini, so the add-in registers direct.py's functions as it
    opens (docs/architecture.md, "Direct functions")."""

    def __init__(self, app, pid, xll):
        self.app = app
        self.pid = pid
        self.xll = xll
        self.folder = xll.parent
        self.script = xll.parent / "direct.py"

    def book(self, folder, sheet=None, name="book.xlsx"):
        """A new workbook of one sheet, saved in folder."""
        workbook = self.app.Workbooks.Add()
        while workbook.Worksheets.Count > 1:
            workbook.Worksheets(workbook.Worksheets.Count).Delete()
        if sheet:
            workbook.Worksheets(1).Name = sheet
        workbook.SaveAs(str(folder / name), FileFormat=51)  # 51: .xlsx
        return Book(self.app, workbook)

    def pycall(self, function, *args):
        """A PyCall formula calling the function in this add-in's direct.py, by full path."""
        return pycall(str(self.script), function, *args)

    def console(self, holding=None, timeout=5.0):
        """Pyinex's console window in this Excel, waiting up to timeout for holding."""
        deadline = time.monotonic() + timeout
        while True:
            window = Console.find(self.pid)
            if window is not None and (holding is None or holding in window.text):
                return window
            if time.monotonic() > deadline:
                return window
            time.sleep(0.05)

    def cleared(self):
        """Clears the console once what was written before has arrived; returns this."""
        marker = f"cleared {time.monotonic_ns()}"
        self.app.Evaluate(f'PyCall("{self.script}", "PrintLater", "{marker}")')
        window = self.console(holding=marker, timeout=10)
        assert window is not None and marker in window.text, "the console never showed the marker"
        self.app.Evaluate("PyConsole(FALSE, TRUE)")
        return self


@pytest.fixture(scope="session")
def direct_excel(excel, embedded, pytestconfig, run_folder, tmp_path_factory):
    """The DirectExcel of the session, started at first use and ended after the session.
    Its pyinex.ini names the Python the session's Excel embedded."""
    import pywintypes
    import win32process
    from win32com.client import DispatchEx

    xll = Path(pytestconfig.getoption("--excel")).resolve()
    folder = tmp_path_factory.mktemp("direct addin")
    for source in (xll, xll.parent / "pyinex_core.dll", SCRIPTS / "direct.py", SCRIPTS / "values.py"):
        shutil.copyfile(source, folder / source.name)
    (folder / "pyinex.ini").write_text(f"[python]\nexecutable = {embedded['executable']}\nscripts = direct.py\n", "utf-8")
    app = DispatchEx("Excel.Application")
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    report_excel(run_folder, "direct", pid)  # At once, so the driver can end it if the run hangs.
    direct = DirectExcel(app, pid, folder / xll.name)
    try:
        app.Visible = True
        app.DisplayAlerts = False
        assert app.RegisterXLL(str(direct.xll)), f"Excel would not load {direct.xll}"
        yield direct
    finally:
        try:
            for workbook in list(app.Workbooks):
                workbook.Close(SaveChanges=False)
            app.Quit()
        except pywintypes.com_error:
            pass
        direct.app = app = None
        gc.collect()
        # As for the session's Excel, the driver waits for this one after pytest has exited.
