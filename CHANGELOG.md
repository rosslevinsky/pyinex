# Changelog

Versions from 2026 on are calendar versions, `YEAR.MONTH.N`: `2026.10.0` is the first
release of October 2026, and `2026.10.1` the next one that month. Each month starts again
at `.0`. A version heading with no date has not been released yet.

## 2026.10.1 — October 19, 2026

Functions a cell calls by their own names. Existing workbooks and scripts work unchanged:
`PyCall` and the other functions behave as before.

**Direct functions**

- A function a script marks `@pyinex.func` becomes a worksheet function under its own
  name: `=Discount(A1, B1)` instead of `=PyCall("model.py", "Discount", A1, B1)`. It
  follows PyCall's rules for values, errors and threads.
- Insert Function and the Function Arguments dialog show its argument names, from the
  signature, and its help and each argument's help, from the docstring.
  `@pyinex.func(name=..., category=..., help=...)` sets the name Excel uses, the category
  (by default "Pyinex") and the help.
- Any signature that fits 30 Excel arguments: positional parameters, then keyword-only
  ones, then `*args`. An argument left out takes its default.
- Annotations convert arguments: `int`, `float`, `str`, `bool`, a list of rows such as
  `list[list[float]]`, `numpy.ndarray`, and any of these or `None`. A value that does not
  fit gives `#VALUE!` and a console message naming the function, the argument and the
  value. Text is never read as a number or a logical value.
- A parameter annotated `pyinex.Range` receives a reference instead of its values, written
  with Excel's own reference syntax. It gives its workbook, sheet, address, size and areas,
  and during a call its values and which cells are empty. A whole column reads only its
  used part.
- Names are compared without regard to case. The first script to register a name keeps
  it; a name another script, Excel or another add-in already uses is refused with a reason
  on the console, as is one of Pyinex's own.
- With `PyReloadOnChange` on, a changed function body takes effect at the next call, as
  does any change Excel is not told of, such as a default value. A change to what Excel is
  told (the number of arguments, their names or types, the help), a removed, renamed or
  added function is registered or removed through Excel shortly after, and a removed name
  no longer reaches another function.

**Loading scripts**

- `PyLoad(file)` loads a script at any time and registers its functions, in every open
  workbook. It returns the number of functions it accepted, such as `2 functions`. A
  script that fails to load registers nothing, and the next calculation of the cell loads
  the fixed file.
- `pyinex.ini` takes a `scripts` list, whose functions register when the add-in opens.
  Python then starts when the add-in opens; otherwise it still starts at first use.
- When a workbook opens, even in a new Excel, Pyinex calculates its `PyLoad` cells, on
  hidden sheets too, so its functions register before anything changes. Entering those
  cells again marks the workbook as changed.
- A formula typed before its function was registered shows `#NAME?` until entered again.
  Pyinex now enters such formulas again once the function registers, including calls
  inside `IFERROR` and in `IF` branches not yet taken.
- A new command, `PyRepair`, run from the Macro dialog or `Application.Run`, does both on
  demand. It loads a workbook opened while a macro had turned Excel's events off, which
  Pyinex cannot see opening.

**Other changes**

- Every function's last argument help shows whole in the Function Arguments dialog,
  PyCall's included: Pyinex adds an empty one after it, as Microsoft advises.
- When Excel answers that a cell a `pyinex.Range` reads is not yet calculated, the rest of
  that call, a PyCall's included, makes no further call into Excel and gives `#VALUE!`, as
  Microsoft requires; Excel calls the function again once the cell is calculated.
- A new example workbook, `direct_functions.xlsx`, uses `PyLoad`, direct functions,
  annotations and `pyinex.Range`.
- The Python inside the add-in's core is kept as `.py` files and compiled into
  `pyinex_core.dll` when it is built. The add-in is still two files.
- A traceback no longer shows any of Pyinex's own frames, including one after the script's
  own (as when `pyinex.ExcelError` refuses its text) and those of chained exceptions.
- A relative script path in a workbook with no folder says why: no calling cell, as from a
  macro, or a workbook opened from OneDrive or SharePoint, whose location is a web address.
  It no longer claims the workbook was never saved.
- `pyinex.ini` saved as UTF-16 big-endian is read; one saved as UTF-16 with no byte-order
  mark is reported, rather than silently ignored.
- A Python that fails to start is refused with Python's own reason, not the last line of
  its thread dump; Python 2 is refused as too old. If Pyinex cannot set up a Python that
  started, the console says which step failed and why, and that Excel must be restarted.
- `PyConsole` refuses a position no monitor shows, as between monitors of different sizes.
- The console no longer holds back the end of a line after a byte that is not UTF-8, and
  keeps Pyinex's messages when Excel has standard output but no standard error.
- A call Excel stopped, to calculate a cell it reads first, no longer prints a failure.
- A result too large for memory gives `#VALUE!` with a console message.
- Two script paths whose short hashes collide get modules of their own.
- `PyInfo("pyinex")` does not start Python.

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
