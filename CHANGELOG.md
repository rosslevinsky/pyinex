# Changelog

Versions from 2026 on are calendar versions, `YEAR.MONTH.N`: `2026.10.0` is the first
release of October 2026, and `2026.10.1` the next one that month. Each month starts again
at `.0`. A version heading with no date has not been released yet.

## 2026.10.0 — October 15, 2026

A rewrite for current Excel and Python, and the first release since 0.83 alpha. Each
worksheet function keeps its purpose, but several changed name or behavior. Existing
workbooks and scripts need the changes below.

**Requirements**

- 64-bit Excel on Windows (tested with Microsoft 365) and 64-bit CPython 3.10 or later
  (tested with 3.10 to 3.15). One add-in works with every Python from 3.10 on. 32-bit
  Excel and Python 2 are no longer supported.
- The add-in is two files, `pyinex.xll` and `pyinex_core.dll`, which must be in the same
  folder.

**Functions**

- Four functions instead of six:
  - `PyConsole` gains a `clear` argument, which replaces `PyClearConsole`.
  - `PyConsole` loses `preserveLocation`. A position you give is applied each time the
    console is shown; with no position, the window stays where it was.
  - `PyModuleFreshnessCheck` is now `PyReloadOnChange`.
  - `PyLoadedLibrary` is now `PyInfo`, which reports more.
  - `PyVerbose` is removed. Set `PYTHONVERBOSE`, or `verbose` in `pyinex.ini`, before
    Excel starts.
- `PyCall` takes up to 30 arguments instead of 15. It no longer inspects the function, so
  any callable works, including classes and built-in functions, and Python itself applies
  default values and `*args`.
- A relative script path is relative to the workbook's folder, not to Excel's current
  folder.

**Values** ([docs/conversion.md](https://github.com/rosslevinsky/pyinex/blob/main/docs/conversion.md))

- Every range of two or more cells arrives as a tuple of rows. Before, a single row
  arrived as a flat tuple.
- Excel errors arrive as `pyinex.ExcelError` objects, not as text. Returning one puts that
  error in the cell.
- Every failure gives `#VALUE!` and writes the reason to the console. Before, the cell
  showed message text.
- `None` shows as nothing, and NaN and infinity show as `#NUM!`.
- Text longer than Excel's limit of 32,767 characters is an error. It is never cut short.

**Scripts**

- Scripts are loaded by their full path. Two scripts with the same name in different
  folders no longer collide, and paths with non-ASCII characters work without Windows'
  short file names.
- Every exception, including `SystemExit`, becomes `#VALUE!` and a traceback on the
  console. No exception a script raises can end Excel.
- Threads that a script starts keep running between calls, because Pyinex holds Python's
  global lock only during a call.

**The pyinex module**

- Scripts must `import pyinex`. Pyinex no longer adds the name to every script.
- `CallerA1`, `CallerA1Full`, `CallerR1C1` and `CallerR1C1Full` become
  `pyinex.caller(r1c1=False, full=False)`. `CallerSheet` becomes `pyinex.sheet()`.
  `Break` becomes `pyinex.esc_pressed(clear=False)`.
- Sheet and workbook names keep every Unicode character.

**Choosing a Python**

- Pyinex uses `PYINEX_PYTHON`, then a `pyinex.ini` beside the add-in, then the first
  `python.exe` on `PATH`, then the newest Python recorded in the registry by python.org's
  installer.
- A virtual environment's `python.exe` works as that environment.
- If Pyinex cannot use a Python, it refuses it and gives the reason on the console and in
  `PyInfo("status")`. It does not crash Excel.
- `pyinex.ini` may be saved as UTF-8, with or without a byte-order mark, or as UTF-16, so
  the path may contain any character. Windows' own reader for such files reads UTF-8
  without a byte-order mark as the 8-bit code page, so Pyinex reads UTF-8 files itself.

**The console**

- The console is a window of Pyinex's own, not a Windows console, which Microsoft 365
  Excel refused to create in testing.
- It is hidden until `PyConsole` shows it. When shown, it comes to the front without
  taking the keyboard from Excel. Closing it only hides it.
- It uses Windows' light or dark mode.
- Output is UTF-8. `print()` from a thread that a script started appears at once.

**The project**

- Built with CMake and MSVC instead of Visual C++ 2008 solutions.
- Built directly on Excel's C API instead of the xlw library. The add-in contains no
  Microsoft code: it is compiled against the SDK's `XLCALL.H` only.
- Tested without Excel on every push, in GitHub Actions, under each supported Python; and
  tested in real Excel under each supported Python with `scripts/excel_tests.py`.
- New example workbooks (`.xlsx`). Each sheet shows every result beside its expected
  value.

## 0.83 alpha — January 18, 2010

- The first version to include its source code.
- Python files whose path or name has non-ASCII characters can be loaded, by using
  Windows' short path names. This can fail, because short names can be turned off in
  Windows.
- Functions with a `*args` parameter can be called. All 15 arguments are passed: the
  named parameters first, and the rest in the tuple. Keyword arguments are not supported,
  because a cell holding `name=value` could also be ordinary text.
- Folders already on `sys.path` are not added again.
- `PythonExtensionTest.py` shows how a script that uses the `pyinex` module can run
  outside Excel, with a stand-in `pyinex` class.
- `PyModuleFreshnessCheck` can turn off the check for changed scripts, which makes calls
  faster.

## 0.82 — July 9, 2009

The first version fit for production use. It supports Excel 2002 and 2007 (and probably
2003), and the python.org builds of Python 2.5, 2.6 and 3.1, with Unicode strings.
