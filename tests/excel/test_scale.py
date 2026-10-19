"""Many calls in real Excel leak nothing (docs/architecture.md, "Lifetime"): Excel's handle
count and private memory after 1,000 warm-up calls and after 10,000 more, and every
result right; for PyCall and for a direct function. And many reloads leak nothing: a
script's, and a direct function's registration made again (the direct-functions test
group 11).
"""

import ctypes
import time
from ctypes import wintypes

from conftest import EXCEL_PID, E, pycall, shown

CELLS = 1000
ROUNDS = 10
# Measured in testing after the 10,000 calls: from 11 fewer to 2 more handles, and
# from 4.5 MB less to 8.6 MB more private memory, Excel's own variation from run to run.
MOST_HANDLES = 20
MOST_MEMORY = 32 * 2**20


class PROCESS_MEMORY_COUNTERS_EX(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD), ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t), ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t), ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t), ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t), ("PrivateUsage", ctypes.c_size_t),
    ]  # fmt: skip


def handles_and_memory(pid):
    """Excel's open handles and its private memory in bytes."""
    kernel32, psapi = ctypes.windll.kernel32, ctypes.windll.psapi
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.GetProcessHandleCount.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.DWORD)]
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    psapi.GetProcessMemoryInfo.argtypes = [ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD]
    process = kernel32.OpenProcess(0x1000 | 0x0010, False, pid)  # QUERY_LIMITED_INFORMATION, VM_READ
    assert process, "cannot open Excel's process"
    try:
        count, memory = wintypes.DWORD(), PROCESS_MEMORY_COUNTERS_EX()
        memory.cb = ctypes.sizeof(memory)
        assert kernel32.GetProcessHandleCount(process, ctypes.byref(count))
        assert psapi.GetProcessMemoryInfo(process, ctypes.byref(memory), memory.cb)
        return count.value, memory.PrivateUsage
    finally:
        kernel32.CloseHandle(process)


RELOADS = 600
# Excel's handle count drifts by up to about 70 on its own (measured re-entering an unchanged
# formula 200 times, and by +70 across 400 reloads), so the bound for reloads is wider; a
# leak of a handle every four reloads or more still fails it.
MOST_RELOAD_HANDLES = 150


def test_reloads_leak_nothing(book, pytestconfig):
    # Each reload's replaced module must be freed, and nothing else kept: a helper, which is
    # never reloaded, holds a weak reference to each version of the script.
    pid = pytestconfig.stash[EXCEL_PID]
    (book.folder / "helper_versions.py").write_text("import weakref\n\nversions = []\n", encoding="utf-8")
    source = (
        "import gc, sys\n\nimport helper_versions\n\n"
        "helper_versions.versions.append(__import__('weakref').ref(sys.modules[__name__]))\n\n"
        "VERSION = {n}\n\n"
        "def alive():\n    gc.collect()\n"
        "    return f'{{VERSION}} of {{len(helper_versions.versions)}}, alive {{sum(r() is not None for r in helper_versions.versions)}}'\n"
        "# {pad}\n"
    )
    book.script("versions.py", source.format(n=0, pad=""))
    assert book.formula("A1", pycall("versions.py", "alive")) == "0 of 1, alive 1"
    first = handles_and_memory(pid)
    for n in range(1, RELOADS + 1):
        book.script("versions.py", source.format(n=n, pad=" " * n))  # Each larger: the size alone says it changed.
        # This version is loaded, every one before it was, and only it is alive.
        assert book.formula("A1", pycall("versions.py", "alive")) == f"{n} of {n + 1}, alive 1"
    last = handles_and_memory(pid)
    print(f"after {RELOADS} reloads: {last[0] - first[0]:+} handles, {(last[1] - first[1]) / 2**20:+.1f} MB")
    assert last[0] - first[0] <= MOST_RELOAD_HANDLES
    assert last[1] - first[1] <= MOST_MEMORY


def test_ten_thousand_calls_leak_nothing(book, pytestconfig):
    pid = pytestconfig.stash[EXCEL_PID]
    # Each call converts text and a block both ways.
    book.script(
        "scale.py",
        "import sys\n\ndef f(row, block):\n    return f'row {int(row)}: ' + ','.join(str(int(v)) for r in block for v in r)\n\n"
        "def refs():\n    return sys.getrefcount(sys.modules[__name__])\n",
    )
    book.sheet.Range("C1:D2").Value2 = ((1, 2), (3, 4))
    cells = book.sheet.Range(f"A1:A{CELLS}")
    cells.Formula = pycall("scale.py", "f", "ROW()", "$C$1:$D$2")  # The 1,000 warm-up calls.
    expected = [f"row {r}: 1,2,3,4" for r in range(1, CELLS + 1)]
    assert [row[0] for row in book.values(f"A1:A{CELLS}")] == expected
    first = handles_and_memory(pid)
    refs = book.formula("E1", pycall("scale.py", "refs"))
    for round in range(ROUNDS):
        book.excel.CalculateFull()
        assert [row[0] for row in book.values(f"A1:A{CELLS}")] == expected, f"round {round + 1}"
    last = handles_and_memory(pid)
    # A call that found the script in the cache kept no reference to it.
    assert book.formula("E1", pycall("scale.py", "refs")) == refs
    print(f"after {CELLS} calls: {first[0]} handles, {first[1] / 2**20:.1f} MB private; "
          f"after {CELLS * ROUNDS} more: {last[0]} handles, {last[1] / 2**20:.1f} MB")  # fmt: skip
    assert last[0] - first[0] <= MOST_HANDLES, f"{last[0] - first[0]} more handles"
    assert last[1] - first[1] <= MOST_MEMORY, f"{(last[1] - first[1]) / 2**20:.1f} MB more private memory"


def test_ten_thousand_direct_function_calls_leak_nothing(app, book, pytestconfig):
    pid = pytestconfig.stash[EXCEL_PID]
    book.script(
        "sc_direct.py",
        "import pyinex\n\n@pyinex.func\ndef ScDirect(row, block):\n"
        "    return f'row {int(row)}: ' + ','.join(str(int(v)) for r in block for v in r)\n",
    )
    assert book.formula("F1", '=PyLoad("sc_direct.py")') == "1 function"
    deadline = time.monotonic() + 15
    while shown(app.Evaluate("ScDirect(1, 1)")) == E("#NAME?") and time.monotonic() < deadline:
        time.sleep(0.1)
    book.sheet.Range("C1:D2").Value2 = ((1, 2), (3, 4))
    book.sheet.Range(f"A1:A{CELLS}").Formula = "=ScDirect(ROW(), $C$1:$D$2)"  # The 1,000 warm-up calls.
    expected = [f"row {r}: 1,2,3,4" for r in range(1, CELLS + 1)]
    assert [row[0] for row in book.values(f"A1:A{CELLS}")] == expected
    first = handles_and_memory(pid)
    for round in range(ROUNDS):
        book.excel.CalculateFull()
        assert [row[0] for row in book.values(f"A1:A{CELLS}")] == expected, f"round {round + 1}"
    last = handles_and_memory(pid)
    print(f"after {CELLS * ROUNDS} direct calls: {last[0] - first[0]:+} handles, {(last[1] - first[1]) / 2**20:+.1f} MB")
    assert last[0] - first[0] <= MOST_HANDLES, f"{last[0] - first[0]} more handles"
    assert last[1] - first[1] <= MOST_MEMORY, f"{(last[1] - first[1]) / 2**20:.1f} MB more private memory"


REGISTRATIONS = 200


def test_registering_again_hundreds_of_times_leaks_nothing(app, book, pytestconfig):
    # Each reload changes the signature, so the function is removed, its name cleared, and
    # registered again through the queue.
    pid = pytestconfig.stash[EXCEL_PID]
    source = "import pyinex\n\n@pyinex.func\ndef ScAgain({0}):\n    return 'v{1}'\n"

    def version(n):
        return source.format("a" if n % 2 else "a, b", n)

    book.script("sc_again.py", version(0))
    assert book.formula("F1", '=PyLoad("sc_again.py")') == "1 function"
    first = None
    for n in range(1, REGISTRATIONS + 1):
        book.script("sc_again.py", version(n))
        assert book.formula("F1", '=PyLoad("sc_again.py")') == "1 function"
        call = "ScAgain(1)" if n % 2 else "ScAgain(1, 2)"
        deadline = time.monotonic() + 15
        while shown(app.Evaluate(call)) != f"v{n}" and time.monotonic() < deadline:
            time.sleep(0.05)
        assert shown(app.Evaluate(call)) == f"v{n}", n
        if n == 20:
            first = handles_and_memory(pid)  # After a warm-up.
    last = handles_and_memory(pid)
    print(f"after {REGISTRATIONS - 20} registrations again: {last[0] - first[0]:+} handles, "
          f"{(last[1] - first[1]) / 2**20:+.1f} MB")
    assert last[0] - first[0] <= MOST_RELOAD_HANDLES
    assert last[1] - first[1] <= MOST_MEMORY
