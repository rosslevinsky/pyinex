"""pyinex.Range in the fake Excel (docs/architecture.md, "pyinex.Range"; docs/conversion.md,
"References"): each kind of reference as Excel passes it, with its sheet; several areas; values
that are not references refused; a whole column read without reading the column; each stop
code, with no call into Excel after it but xlFree, from caught exceptions and from cleanup,
everything Excel handed over freed, and nothing on the console; a range kept after its call;
the refusals outside a call and on a script's thread (Range's part of group 12); a kept range
whose sheet is gone; a range returned; a stop met as the result is freed, by a direct function
and a PyCall; the queue waiting until a stopped call returns; and reloads that change an
argument from a value to a reference and back.

The fake Excel's sheets hold what set_cells puts there; a reference's xlCoerce and COUNTA read
them, and a whole column or row read cell by cell is a fault, as is any call into Excel after
a stop but xlFree. One script holds the functions, loaded once by PyLoad in a host of the
module's own.
"""

import pytest
from conftest import host_env
from hostdriver import REF, VALUE, Err, Host, Ref, SRef
from test_direct_functions import wait_for
from test_reload_functions import edit, loaded, types

SCRIPT = """import ctypes, os, sys, threading, time
from ctypes import wintypes
import pyinex

def attempt(f):
    try:
        return repr(f())
    except Exception as e:
        return f'{type(e).__name__}: {e}'

@pyinex.func
def Shown(r: pyinex.Range):
    return repr([r.workbook, r.sheet, r.address(), r.address(r1c1=True), r.address(full=True), r.rows, r.columns])

@pyinex.func
def Values(r: pyinex.Range):
    return repr([r.values(), r.empty()])

@pyinex.func
def Areas(r: pyinex.Range):
    return repr([r.address(), [(a.address(), a.values()) for a in r.areas], attempt(r.values), attempt(lambda: r.rows)])

@pyinex.func
def Both(plain, r: pyinex.Range):
    return repr([plain, r.address()])

kept = []

@pyinex.func
def Keep(r: pyinex.Range):
    kept.append(r)
    return r.address()

@pyinex.func
def Returned(r: pyinex.Range):
    return r

@pyinex.func
def Kept():
    r = kept[-1]
    return repr([r.address(full=True), r.rows, r.values()])

@pyinex.func
def KeptRead():
    r = kept[-1]
    return repr([attempt(r.values), attempt(r.empty), attempt(r.areas[0].values)])

later = []

@pyinex.func
def ReadLater(r: pyinex.Range):
    def read():
        time.sleep(0.3)
        later.append(attempt(r.values))
        print('read later', file=sys.stderr)
    threading.Thread(target=read).start()
    return 'reading later'

@pyinex.func
def ReadLaterResult():
    return later[-1] if later else 'not yet'

@pyinex.func
def OnThread(r: pyinex.Range):
    out = []
    t = threading.Thread(target=lambda: out.append(attempt(r.values)))
    t.start()
    t.join()
    return out[0]

log = []

@pyinex.func
def Stops(first: pyinex.Range, second: pyinex.Range):
    log[:] = [attempt(first.values)]
    log.append(attempt(second.values))
    for ask in (second.values, second.empty, pyinex.caller, pyinex.sheet, pyinex.esc_pressed):
        log.append(attempt(ask))
    return 'returned'

@pyinex.func
def StopsRaising(r: pyinex.Range):
    try:
        return repr(r.values())
    finally:
        log[:] = [attempt(r.empty)]

class Late(float):
    # Reads a range only as Pyinex lets the result go, once it has converted it.
    def __del__(self):
        log[:] = [attempt(self.range.values)]

def late(value, r):
    result = Late(value)
    result.range = r
    return result

@pyinex.func
def ReadsWhenFreed(r: pyinex.Range):
    return late(1.0, r)

def read_kept_when_freed():  # By PyCall.
    return late(2.0, kept[-1])

user32 = ctypes.WinDLL('user32')
user32.FindWindowExW.restype = wintypes.HWND
user32.FindWindowExW.argtypes = (wintypes.HWND, wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR)
user32.GetWindowThreadProcessId.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.DWORD))
user32.PostMessageW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
user32.PeekMessageW.argtypes = (ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT, wintypes.UINT)
user32.TranslateMessage.argtypes = user32.DispatchMessageW.argtypes = (ctypes.POINTER(wintypes.MSG),)

def queue_window():
    # This process's: another host's would not answer.
    window = None
    while True:
        window = user32.FindWindowExW(-3, window, 'PyinexQueueWindow', None)  # -3: HWND_MESSAGE.
        process = wintypes.DWORD()
        if not window or (user32.GetWindowThreadProcessId(window, ctypes.byref(process)) and process.value == os.getpid()):
            return window

@pyinex.func
def StopsThenTimes(r: pyinex.Range, pump):
    # Meets a stop, then posts the queue's timer message (WM_TIMER, ID 1), as a timer due
    # would; with pump, handles Windows messages, as a message box would.
    log[:] = [attempt(r.values)]
    log.append(bool(user32.PostMessageW(queue_window(), 0x0113, 1, 0)))
    message = wintypes.MSG()
    while pump and user32.PeekMessageW(ctypes.byref(message), None, 0, 0, 1):  # 1: PM_REMOVE.
        user32.TranslateMessage(ctypes.byref(message))
        user32.DispatchMessageW(ctypes.byref(message))
    return 'returned'

@pyinex.func
def StopsThenRunsTheQueue(r: pyinex.Range):
    # Meets a stop, then runs the queue's command itself, as only something else could.
    log[:] = [attempt(r.values)]
    xll = ctypes.WinDLL(os.environ['PYINEX_TEST_XLL'])
    log.append(xll.PyinexQueue())
    return 'returned'

@pyinex.func
def Log():
    return repr(log)

@pyinex.func
def Mark(text):
    print(text, file=sys.stderr)
    return text
"""


@pytest.fixture(scope="module")
def ranges(xllhost, pyinex_xll, tmp_path_factory):
    path = tmp_path_factory.mktemp("ranges") / "ranges.py"
    path.write_text(SCRIPT, encoding="utf-8")
    with Host(xllhost, pyinex_xll, env=host_env(PYINEX_TEST_XLL=str(pyinex_xll))) as h:
        h.open()
        assert h.call("PyLoad", str(path)) == "18 functions"
        h.tick()
        h.set_cells("[Book1]Sheet1", "B2", ((1.0, ""), (None, "x")))
        h.set_cells("[Book1]Sheet1", "D1", 10.0)
        h.set_cells("[Book1]Sheet1", "D3", 30.0)
        h.set_cells("[Book1]Sheet1", "E5", True)
        h.set_cells("[Book1]Data", "A1", (("data", 2.0),))
        h.set_cells("[Other.xlsx]Sheet1", "A1", "other book")
        h.script = path
        yield h


SHEET1 = "[Book1]Sheet1"


@pytest.mark.parametrize(
    "reference, shown",
    [
        (Ref(SHEET1, "B2"), ["Book1", "Sheet1", "B2", "R2C2", "[Book1]Sheet1!B2", 1, 1]),
        (Ref(SHEET1, "B2:C3"), ["Book1", "Sheet1", "B2:C3", "R2C2:R3C3", "[Book1]Sheet1!B2:C3", 2, 2]),
        (Ref(SHEET1, "D:D"), ["Book1", "Sheet1", "D1:D1048576", "R1C4:R1048576C4", "[Book1]Sheet1!D1:D1048576", 1048576, 1]),
        (Ref(SHEET1, "5:5"), ["Book1", "Sheet1", "A5:XFD5", "R5C1:R5C16384", "[Book1]Sheet1!A5:XFD5", 1, 16384]),
        (Ref("[Book1]Data", "A1:B2"), ["Book1", "Data", "A1:B2", "R1C1:R2C2", "[Book1]Data!A1:B2", 2, 2]),
        (Ref("[Other.xlsx]Sheet1", "A1"), ["Other.xlsx", "Sheet1", "A1", "R1C1", "[Other.xlsx]Sheet1!A1", 1, 1]),
    ],
    ids=["cell", "range", "whole column", "whole row", "another sheet", "another workbook"],
)
def test_each_reference_gives_its_workbook_sheet_address_and_size(ranges, reference, shown):
    assert ranges.call("Shown", reference) == repr(shown)


def test_a_reference_with_no_sheet_is_on_the_sheet_calculating(ranges):
    ranges.set_caller("[Book1]Data", 5, 5)
    try:
        assert ranges.call("Shown", SRef("A1:B1")) == repr(["Book1", "Data", "A1:B1", "R1C1:R1C2", "[Book1]Data!A1:B1", 1, 2])
        assert ranges.call("Values", SRef("A1:B1")) == repr([(("data", 2.0),), ((False, False),)])
    finally:
        ranges.set_caller(SHEET1)


@pytest.mark.parametrize(
    "reference, values",
    [
        (Ref(SHEET1, "B2"), [1.0, False]),
        (Ref(SHEET1, "B2:C3"), [((1.0, ""), (None, "x")), ((False, False), (True, False))]),
        (Ref(SHEET1, "Z99"), [None, True]),
        (Ref("[Book1]Data", "A1:B1"), [(("data", 2.0),), ((False, False),)]),
        (Ref("[Other.xlsx]Sheet1", "A1"), ["other book", False]),
    ],
    ids=["cell", "range", "empty cell", "another sheet", "another workbook"],
)
def test_values_and_empty_cells(ranges, reference, values):
    # A defined name reaches the add-in as the reference it names: one of these.
    assert ranges.call("Values", reference) == repr(values)


def test_a_whole_column_reads_only_its_used_part(ranges):
    # Reading every cell of the column would be a fault in the fake Excel.
    assert ranges.call("Values", Ref(SHEET1, "D:D")) == repr([((10.0,), (None,), (30.0,)), ((False,), (True,), (False,))])
    assert ranges.call("Values", Ref(SHEET1, "5:5")) == repr([((None, None, None, None, True),), ((True,) * 4 + (False,),)])
    assert ranges.call("Values", Ref(SHEET1, "Y:Y")) == repr([None, True])


def test_several_areas(ranges):
    assert ranges.call("Areas", Ref(SHEET1, "B2:C3", "E5")) == repr([
        "B2:C3,E5",
        [("B2:C3", ((1.0, ""), (None, "x"))), ("E5", True)],
        "ValueError: B2:C3,E5 has 2 areas: use its areas",
        "ValueError: B2:C3,E5 has 2 areas: use its areas",
    ])


@pytest.mark.parametrize(
    "value, shown",
    [
        (5.0, "5.0"),
        ("B2", "the text 'B2'"),
        (((1.0, 2.0),), "a range of 1 × 2 cells"),  # An array constant, or a workbook not open.
        (Err(REF), "the error #REF!"),  # A reference to a sheet since deleted.
    ],
    ids=["number", "text", "array", "deleted sheet"],
)
def test_a_value_that_is_not_a_reference_is_refused(ranges, value, shown):
    assert ranges.call("Shown", value) == Err(VALUE)
    assert f"Pyinex: Shown: argument 1 (r): pyinex.Range takes a reference, not {shown}" in wait_for(ranges, shown)


def test_values_and_a_reference_in_one_call(ranges):
    assert ranges.call("Both", 7.0, Ref(SHEET1, "C3")) == repr([7.0, "C3"])


def test_a_range_kept_after_its_call(ranges):
    assert ranges.call("Keep", Ref("[Book1]Data", "A1:B1")) == "A1:B1"
    # Its address and size need no call; a later call reads it.
    assert ranges.call("Kept") == repr(["[Book1]Data!A1:B1", 1, (("data", 2.0),)])


@pytest.mark.parametrize("gone, after", [("[Book1]Deleted", "[Book1]Added"), ("[Closed.xlsx]Sheet1", "[Opened.xlsx]Sheet1")],
                         ids=["sheet deleted", "workbook closed"])
def test_a_kept_range_whose_sheet_is_gone_is_not_read(ranges, gone, after):
    ranges.set_cells(gone, "A1", "kept")
    assert ranges.call("Keep", Ref(gone, "A1")) == "A1"
    ranges.forget(gone)
    ranges.set_cells(after, "A1", "another sheet")  # Given the ID Excel freed.
    refused = (f"RuntimeError: pyinex.Range: Excel no longer has the sheet {gone} open; it was deleted or renamed, "
               "or its workbook closed")
    assert ranges.call("KeptRead") == repr([refused] * 3)


def test_a_range_returned_is_value_as_any_object_excel_cannot_take(ranges):
    assert ranges.call("Returned", Ref(SHEET1, "B2")) == Err(VALUE)
    assert "Pyinex: Returned: Range cannot be returned to Excel" in wait_for(ranges, "Range cannot be returned")


def test_reading_outside_a_call_is_refused(ranges):
    assert ranges.call("ReadLater", Ref(SHEET1, "B2")) == "reading later"
    wait_for(ranges, "read later")  # With no call meanwhile: the thread reads between calls.
    assert ranges.call("ReadLaterResult") == ("RuntimeError: pyinex.Range asks Excel, which allows that only while Excel is calculating a "
                      "PyCall or a direct function")


def test_reading_on_a_scripts_thread_is_refused(ranges):
    assert ranges.call("OnThread", Ref(SHEET1, "B2")) == (
        "RuntimeError: pyinex.Range asks Excel, which allows that only on Excel's own thread, not one a script started")


# A call Excel stops: xlretUncalced or xlretAbort.

UNCALCED = "a cell this call reads is not yet calculated, and Excel calls the function again once it is"
ABORT = "the calculation was interrupted, and Excel calls the function again in the next calculation"


def said_during(host, request):
    """What the console says while request() runs, between two markers printed by calls."""
    import uuid

    start, end = f"<{uuid.uuid4()}>", f"<{uuid.uuid4()}>"
    host.call("Mark", start)
    result = request()
    host.call("Mark", end)
    return result, wait_for(host, end).split(start)[1].split(end)[0].strip()


@pytest.mark.parametrize("code, why", [("uncalced", UNCALCED), ("abort", ABORT)])
def test_after_a_stop_nothing_asks_excel_and_the_call_gives_value_silently(ranges, code, why):
    # The first reading succeeds; the second stops. Any call into Excel after the stop but
    # xlFree is a fault, which fails the request, and a value Excel handed over and never
    # freed is one too.
    ranges.stop(code, skip=1)
    assert said_during(ranges, lambda: ranges.call("Stops", Ref(SHEET1, "B2"), Ref(SHEET1, "C3"))) == (Err(VALUE), "")
    log = eval(ranges.call("Log"))
    assert log[0] == "1.0"
    assert log[1:] == [f"RuntimeError: pyinex.{name} cannot ask Excel: {why}"
                       for name in ("Range", "Range", "Range", "caller", "sheet", "esc_pressed")]
    # The next call is not stopped.
    assert ranges.call("Values", Ref(SHEET1, "C3")) == repr(["x", False])


def test_a_stop_in_a_finally_block_and_raised_says_nothing(ranges):
    ranges.stop("uncalced")
    assert said_during(ranges, lambda: ranges.call("StopsRaising", Ref(SHEET1, "B2"))) == (Err(VALUE), "")
    assert eval(ranges.call("Log")) == [f"RuntimeError: pyinex.Range cannot ask Excel: {UNCALCED}"]


@pytest.mark.parametrize("skip", [0, 1, 5])
def test_a_stop_while_finding_a_whole_column_s_used_part(ranges, skip):
    ranges.stop("uncalced", skip=skip)
    assert said_during(ranges, lambda: ranges.call("Values", Ref(SHEET1, "D:D"))) == (Err(VALUE), "")
    assert ranges.call("Values", Ref(SHEET1, "D:D")).startswith("[((10.0,),")


# Reloads that change an argument from a value to a reference and back.

ONE = "import pyinex\n\n@pyinex.func\ndef {0}({1}):\n    return {2}\n"


@pytest.fixture
def script(tmp_path):
    return tmp_path / "switched.py"


def test_a_value_becoming_a_reference_registers_again(host, script):
    loaded(host, script, ONE.format("Switch", "a", "repr(a)"))
    assert types(host) == {"Switch": "QQ"}
    host.set_busy(True)  # So the queue waits while the interval is checked.
    edit(script, ONE.format("Switch", "a: pyinex.Range", "a.address()"))
    # Called before the queue has run: Excel still passes values, which never reach the new code.
    assert host.call("Switch", 1.0) == Err(VALUE)
    assert "Pyinex: Switch is being registered again, since its script changed" in wait_for(host, "being registered again")
    assert types(host) == {"Switch": "QQ"}
    host.set_busy(False)
    host.tick()
    assert types(host) == {"Switch": "QU"}
    assert host.call("Switch", Ref(SHEET1, "B2")) == "B2"


def test_a_reference_becoming_a_value_registers_again(host, script):
    loaded(host, script, ONE.format("Switch", "a: pyinex.Range", "a.address()"))
    assert types(host) == {"Switch": "QU"}
    host.set_busy(True)
    edit(script, ONE.format("Switch", "a", "repr(a)"))
    # Excel still passes a reference, which never reaches the new code, and nothing reads it.
    host.excel_calls()
    assert host.call("Switch", Ref(SHEET1, "B2")) == Err(VALUE)
    assert host.excel_calls() == 2  # Naming its sheet (xlSheetNm, then xlFree); no cell is read.
    host.set_busy(False)
    host.tick()
    assert types(host) == {"Switch": "QQ"}
    assert host.call("Switch", 1.0) == "1.0"


def test_a_registration_as_a_reference_that_fails_is_tried_again_by_pyload(host, script):
    loaded(host, script, ONE.format("Retried", "a", "'value'"))
    edit(script, ONE.format("Retried", "a: pyinex.Range", "'reference'"))
    host.fail_registrations(1)
    assert host.call("Retried", 1.0) == Err(VALUE)
    host.tick()
    assert f"Pyinex: Excel refused to register Retried from the script {script}" in wait_for(host, "refused")
    assert types(host) == {}
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick()
    assert types(host) == {"Retried": "QU"}
    assert host.call("Retried", Ref(SHEET1, "B2")) == "reference"


def test_several_edits_between_value_and_reference_while_queued(host, script):
    loaded(host, script, ONE.format("Edited", "a", "'one'"))
    host.set_busy(True)
    edit(script, ONE.format("Edited", "a: pyinex.Range", "'two'"))
    assert host.call("Edited", 1.0) == Err(VALUE)
    edit(script, ONE.format("Edited", "a, b: pyinex.Range", "'three'"))
    assert host.call("PyLoad", str(script)) == "1 function"
    edit(script, ONE.format("Edited", "a: pyinex.Range, b", "'four'"))
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick(3)
    assert types(host) == {"Edited": "QQ"}
    host.set_busy(False)
    host.tick()
    assert types(host) == {"Edited": "QUQ"}
    assert host.call("Edited", Ref(SHEET1, "B2"), 2.0) == "four"


# A stop met as the result is freed.


def test_a_stop_while_the_result_is_freed_gives_value_silently(ranges):
    # The result converts to 1.0 first; its finalizer then meets the stop.
    ranges.stop("uncalced")
    assert said_during(ranges, lambda: ranges.call("ReadsWhenFreed", Ref(SHEET1, "B2"))) == (Err(VALUE), "")
    assert eval(ranges.call("Log")) == [f"RuntimeError: pyinex.Range cannot ask Excel: {UNCALCED}"]


def test_a_stop_while_the_result_is_freed_in_a_pycall_gives_value_silently(ranges):
    assert ranges.call("Keep", Ref(SHEET1, "B2")) == "B2"
    ranges.stop("abort")
    assert said_during(ranges, lambda: ranges.call("PyCall", str(ranges.script), "read_kept_when_freed")) == (Err(VALUE), "")
    assert eval(ranges.call("Log")) == [f"RuntimeError: pyinex.Range cannot ask Excel: {ABORT}"]
    assert ranges.call("PyCall", str(ranges.script), "read_kept_when_freed") == 2.0


# The queue waits until a stopped call returns.


def test_the_queue_s_command_run_inside_a_stopped_call_waits(ranges, tmp_path):
    queued = tmp_path / "queued.py"
    queued.write_text("import pyinex\n\n@pyinex.func\ndef QueuedRunDirectly():\n    return 1.0\n", encoding="utf-8")
    ranges.set_busy(True)  # So the work stays queued.
    try:
        assert ranges.call("PyLoad", str(queued)) == "1 function"
        ranges.tick()
    finally:
        ranges.set_busy(False)
    # The command, run inside the stopped call, registers nothing there (the fake Excel
    # faults any call into Excel but xlFree after the stop) and leaves the work for later.
    ranges.stop("uncalced")
    assert ranges.call("StopsThenRunsTheQueue", Ref(SHEET1, "B2")) == Err(VALUE)
    assert eval(ranges.call("Log")) == [f"RuntimeError: pyinex.Range cannot ask Excel: {UNCALCED}", 0]
    ranges.tick()
    assert ranges.call("QueuedRunDirectly") == 1.0


@pytest.mark.parametrize("pump", [True, False], ids=["messages handled in the call", "timer due after it"])
def test_the_queue_waits_until_a_stopped_call_returns(ranges, tmp_path, pump):
    name = "QueuedPumped" if pump else "QueuedPosted"
    queued = tmp_path / "queued.py"
    queued.write_text(f"import pyinex\n\n@pyinex.func\ndef {name}():\n    return 1.0\n", encoding="utf-8")
    ranges.set_busy(True)  # So the work stays queued.
    try:
        assert ranges.call("PyLoad", str(queued)) == "1 function"
        ranges.tick()
    finally:
        ranges.set_busy(False)
    # Any call into Excel after the stop but xlFree, before the function returns, is a fault;
    # once it has returned, none is.
    ranges.stop("uncalced")
    assert ranges.call("StopsThenTimes", Ref(SHEET1, "B2"), pump) == Err(VALUE)
    assert eval(ranges.call("Log")) == [f"RuntimeError: pyinex.Range cannot ask Excel: {UNCALCED}", True]
    ranges.tick()
    assert ranges.call(name) == 1.0



KEEPER = """import pyinex

kept = []

@pyinex.func
def KeepIt(r: pyinex.Range):
    kept[:] = [r]
    return 1

def read():
    return kept[0].values()

def read_quietly():
    try:
        return kept[0].values()
    except RuntimeError:
        return 'caught'
"""


@pytest.mark.parametrize("function", ["read", "read_quietly"])
def test_a_pycall_excel_stopped_says_nothing(host, tmp_path, function):
    # Excel answers that a cell the PyCall reads is not yet calculated: it calls the function
    # again, so a failure, raised or a result, is not reported.
    script = tmp_path / "keeper.py"
    loaded(host, script, KEEPER)
    assert host.call("KeepIt", Ref("[Book1]Sheet1", "A1")) == 1.0
    host.stop("uncalced")
    assert host.call("PyCall", str(script), function) == Err(VALUE)
    host.stop("none")
    assert host.call("PyCall", str(script), "read") == ""  # Asked again, it reads the empty cell.
    assert "raised an exception" not in host.stderr and "cannot be returned" not in host.stderr


def test_a_function_called_with_no_cell_says_so(host, tmp_path):
    script = tmp_path / "who.py"
    loaded(host, script, "import pyinex\n\n@pyinex.func\ndef Who():\n    try:\n        return pyinex.caller()\n"
                         "    except RuntimeError as e:\n        return str(e)\n")
    host.set_caller(None)
    assert host.call("Who") == "the function running was not called from a cell"


def test_a_value_pyinex_does_not_know_is_read_as_value_and_reported(host, tmp_path):
    script = tmp_path / "unknown.py"
    loaded(host, script, "import pyinex\n\n@pyinex.func\ndef Read(r: pyinex.Range):\n    return repr(r.values())\n")
    host.set_cells("[Book1]Sheet1", "A1", Err(99))
    assert host.call("Read", Ref("[Book1]Sheet1", "A1")) == "ExcelError('#VALUE!')"
    assert "pyinex.Range: Excel error code 99 is not one Pyinex knows" in wait_for(host, "error code 99")
