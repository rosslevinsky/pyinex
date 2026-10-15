# Architecture

This document describes how Pyinex works, and why it works that way, for anyone who changes
it. It starts with what a user sees: the functions, the argument rule, scripts and the
`pyinex` module. It then covers what happens inside: choosing a Python, the add-in's
lifetime, threads and the console. The last sections,
[How a call moves through the code](#how-a-call-moves-through-the-code) and
[Not yet built](#not-yet-built), connect these to the source files and to later work.
Value conversion has its own document, [conversion.md](./conversion.md).

Each decision gives its reason and names the test file that checks it. Where this version
differs from Pyinex 0.83 (2010) for a reason worth knowing, the section says so.

## Excel functions

Pyinex has four worksheet functions, in the function category "Pyinex".

| Function | Purpose |
|---|---|
| `PyCall(file, function, arg1, …, arg30)` | Call a Python function. |
| `PyConsole(show, [clear], [x], [y], [width], [height])` | Show, hide, clear and place the console window. |
| `PyReloadOnChange([on])` | Turn reloading of changed scripts on or off, or report the setting. |
| `PyInfo([item])` | Report which Python and which add-in are loaded. |

None of them is registered as thread-safe (the `$` flag in the type string), so Excel
calls them only on its main thread. See [Threading](#threading).

None of them is volatile. A cell recalculates when its inputs change, or on a full
recalculation (Ctrl+Alt+F9).

**In the Function Arguments dialog.** This dialog, which Excel opens from Insert Function,
calculates a formula while the user types it. There, each function returns a blank and does
nothing else: PyCall runs no Python, PyConsole changes no window, PyReloadOnChange changes
no setting, and PyInfo does not start Python. Otherwise every keystroke would run the
function.

Pyinex recognizes the dialog as Microsoft's sample code does: a window in Excel's process
whose class name starts with `bosa_sdm_XL` and whose title is "Function Arguments".
Microsoft's sample also checks that Excel's main window owns the dialog. Pyinex does not,
because since Excel 2013 each workbook has its own window, which may own the dialog
instead. The title is the English one. In Excel in another language, the functions
therefore run in the dialog as they do in a cell: typing is slower, but no result is wrong.

Tests: `tests/host/test_registration.py` (names, argument counts, type strings, no `$`),
`tests/host/test_calls.py` (behavior, including each function in the dialog).

### PyCall

`PyCall(file, function, arg1, …, arg30)` loads the Python script `file` (see
[Scripts](#scripts)), finds `function` in it, calls it with the arguments, and returns the
result to the cell, converted as [conversion.md](./conversion.md) describes.

- **`file`** is the path of a `.py` file. A relative path is relative to the folder of the
  workbook that contains the calling cell, so a workbook and its scripts can move together.
  Pyinex finds that folder by calling `CELL("filename")` on the calling cell and taking the
  text up to the last backslash. Excel returns `C:\folder\[Book.xlsx]Sheet1`. When the
  workbook's only sheet has the workbook's name, Excel returns only `C:\folder\Book.xlsx`
  (`tests/host/test_calls.py` checks both forms). The folder name may
  contain brackets: Excel will not save a workbook into such a folder, but will open one
  from it. An unsaved workbook has no folder, so in one a relative path is an error. (In
  0.83 a relative path was relative to Excel's current folder, which the user does not
  control.)
- **`function`** is the name of a callable object at the top level of the script: a
  function, a class, or any other callable. Pyinex does not split the name at dots.
- **Up to 30 arguments** follow `file` and `function`. Excel allows 255. Thirty is enough for real use and
  keeps the list in the Function Arguments dialog short. Changing it means changing
  `max_arguments` in `src/loader/addin.cpp` and, to match, PyCall's registered type string,
  its argument names and help, and the parameters of the exported `PyCall` function.
- **Every failure gives `#VALUE!`** and writes the reason to the console. Failures include:
  no usable Python, file not found, function not found, an exception, and a result that
  cannot be returned (see conversion.md). A function that returns a `pyinex.ExcelError` gets that error instead.

Tests: `tests/host/test_calls.py`, `tests/unit/test_call.py`; in real Excel,
`tests/excel/test_failures.py`.

### PyConsole

`PyConsole(show, [clear], [x], [y], [width], [height])`:

- `show` TRUE shows the console, and FALSE hides it. The function returns `show`.
- `clear` TRUE clears the console's text each time the formula calculates. The user should
  set it back to FALSE afterwards. Otherwise the console is cleared on every
  recalculation, at whatever point in the calculation Excel reaches this cell.
- `x` and `y` set the position of the window's upper-left corner, and `width` and `height`
  set its size, in pixels. They are applied each time the console changes from hidden to
  shown. While the console is shown, the user can move it.
- Each of the four is rounded to the nearest whole pixel, with halves rounded away from
  zero (`0.5` → 1, `-0.5` → -1, `2.5` → 3), before it is checked. So a width of `0.4`
  rounds to 0 and is refused, and `0.5` rounds to 1 and is accepted.
- `x` must be given with `y`, and `width` with `height`. `x` and `y` must lie on the
  desktop, which covers every monitor, so they can be negative for a monitor to the left
  of or above the main one. `width` and `height` must be more than 0. A value that breaks
  any of these rules gives `#VALUE!` and a console message.
- With no position or size, the window stays where it was. Giving a position means "put it
  here", and leaving it out means "leave it where it is".
- PyConsole does not start Python.

Tests: `tests/host/test_console.py`; in real Excel, `tests/excel/test_console.py` (every
argument, each refusal and its message, placement, rounding, the close button, clearing).

### PyReloadOnChange

`PyReloadOnChange([on])`:

- TRUE: before each PyCall, Pyinex checks whether the script's file has changed, and
  reloads the script if it has (see [Scripts](#scripts)).
- FALSE: Pyinex loads each script once and keeps it.
- With no argument, it changes nothing.

It always returns the current setting. The default is TRUE, which suits editing scripts
while Excel is open. Changing the setting does not recalculate anything.

Tests: `tests/host/test_calls.py`; in real Excel, `tests/excel/test_reload.py`.

### PyInfo

`PyInfo([item])` reports:

| `item` | Returns |
|---|---|
| `"status"` | `"running"`, or the reason Python could not start (see [Refusals](#refusals)) |
| `"python"` | the full path of the Python DLL in use, for example `C:\…\Python313\python313.dll` |
| `"pyinex"` | the full path of the add-in |
| `"version"` | `sys.version` |
| `"executable"` | the `python.exe` Pyinex chose |
| `"prefix"` | `sys.prefix` |
| `"venv"` | TRUE when the chosen Python is a virtual environment |
| `"addin_version"` | the add-in's version |

Item names are compared without regard to case. With no `item`, it returns all of them as a two-column
table of names and values. Asking for any item starts Python if it has not started. An
unknown item gives `#VALUE!`.

Tests: `tests/host/test_calls.py`; in real Excel, `tests/excel/test_pyinex_module.py`.

### Removed: PyVerbose

0.83's `PyVerbose(n)` changed Python's verbosity while Python was running. Current Python
has no supported way to do that after it has started. Verbosity can still be set before
start-up: set `PYTHONVERBOSE` before starting Excel, or `verbose` in `pyinex.ini` (see
[Choosing a Python](#choosing-a-python)).

## Argument rule

Excel always passes PyCall all 30 argument slots, and marks the ones the user did not type
as missing. Pyinex:

1. **Drops missing arguments at the end.** `=PyCall("s.py", "f", A1, B1)` calls `f(a, b)`.
2. **Passes a missing argument in the middle as `None`**, the same as an empty cell.
   `=PyCall("s.py", "f", A1, , C1)` calls `f(a, None, c)`.
3. **Leaves everything else to Python.** It does not inspect the function or count its
   parameters. Python applies default values, fills `*args`, and raises `TypeError` for a
   wrong number of arguments, which becomes `#VALUE!` with the message on the console.

Pyinex 0.83 read the function's parameter count from its code object. That failed
for anything that is not a plain Python function, such as a class, a built-in function or
a `functools.partial`, and every call paid the cost of the inspection.

Arguments are positional only. Pyinex does not turn a cell such as `rate=0.1`
into a keyword argument, because that cell could equally be ordinary text.

Tests: `tests/unit/test_call.py` (plain functions, `*args`, default values, missing
arguments at the end and in the middle, a class, a `functools.partial`, an object with
`__call__`, and built-in functions with no signature to inspect, such as `max` and `int`);
in real Excel, `tests/excel/test_values_in.py` (missing arguments, 30 arguments, `*args`).

## Scripts

This section describes how the `file` named in a PyCall becomes a Python module.

- **Loaded by path.** Pyinex loads the file with `importlib`, from its full path. Its cache
  of loaded scripts is keyed by that full path, compared without regard to case, as Windows
  compares paths. Two scripts named `model.py` in different folders are two different
  modules, and no path needs Windows' short file names.
- **Module name.** Each script is registered in `sys.modules` under a name made from its
  file name and a short hash of its full path, for example `pyinex_model_3f9a1c2e`. The
  name is unique, and `__name__ == "__main__"` is false, as for any imported module.
- **Other files beside a script.** The first time a script from a folder loads, Pyinex puts that folder
  at the front of `sys.path`, where it stays, as `python script.py` does for the script's
  folder. So `import helper` finds a `helper.py` beside the script, including inside
  functions. These files are ordinary imports: Python caches them in `sys.modules` under
  their plain names, every script that imports one shares it, and Pyinex does **not**
  reload them when they change. Only scripts named in PyCall are reloaded. To use a changed
  helper, restart Excel or call `importlib.reload`. As with `python script.py`, a file in
  the script's folder with the name of a standard module, such as `random.py`, hides that
  module.
- **Compiled from source on every load.** Pyinex reads the script's source and compiles it
  each time it loads the script. It never reads or writes a cached `.pyc` file. Python's
  cache records the source file's modification time only to the second, so an edit that
  kept the file's size and was saved within the same second could otherwise run the old
  code. Scripts are small, so compiling costs little. (Helper files are ordinary imports
  and use Python's cache as usual.)
- **Files a script opens** are read and written as the chosen `python.exe` would read and
  write them, because the embedded Python starts with the same UTF-8 mode and default
  encoding. From Python 3.15, UTF-8 mode is on by default, so `open()` without an
  `encoding` uses UTF-8. Python 3.14 and earlier use the Windows code page, such as
  `cp1252`. Pyinex sets neither. A script that must read a file the same way under every
  Python should name the encoding. (A script's own source is not affected: Python reads
  source as UTF-8 in every version.) Tests:
  `tests/host/test_embedded_python.py::test_starts_as_its_python_exe_would`; in real
  Excel, `tests/excel/test_identity.py::test_files_open_with_the_encoding_python_exe_uses`.
- **Reloading.** With `PyReloadOnChange` on, each PyCall reads the file's last-modified
  time, at the file system's full precision (100 nanoseconds on NTFS), and its size, and
  compares them with the values recorded when the script was loaded. If either changed,
  Pyinex loads the script again into a new module object, so module-level variables start
  again from their initial values. If both are unchanged, Pyinex uses the
  cached module.
- **Why Pyinex checks on each call.** The check is one `os.stat` per call. Windows can
  instead report file changes as they happen, but Pyinex 0.83 found those reports
  unreliable for files on network drives, and receiving them would need a thread of its
  own. A check on each call is simpler, and it sees every change to a file's time or size.
- **Failures.** A script that fails to load leaves no entry in the cache. Causes include a
  missing file, a syntax error, and an exception in its top-level code. Every cell that
  uses the script shows `#VALUE!`, and the console shows why. If a changed script fails to
  reload, Pyinex drops its old version too, so cells never run code that no longer
  matches the file. The next PyCall tries to load it again.
- **Exceptions of every kind** while loading or calling, including `SystemExit` and
  `KeyboardInterrupt`, become `#VALUE!` and a traceback on the console, formatted by
  Python's `traceback` module. No exception a script raises can end Excel. A script can
  still end the process on purpose, for example with `os._exit`, as any code could. Pyinex
  does not print errors with `PyErr_Print`, which ends the process on `SystemExit`.

Tests: `tests/unit/test_module_cache.py` (including an edit of the same size within the
same second), `tests/host/test_exceptions.py`; in real Excel, `tests/excel/test_paths.py`
(absolute and relative paths, two workbooks sharing a script, scripts with the same name,
helper files, spaces and Unicode in every name) and `tests/excel/test_reload.py` (each
reload rule above, with files saved as an editor saves them).

## The pyinex module

`pyinex` is a Python module built into the add-in. Scripts `import pyinex` to use it. It
exists only inside Excel. The add-in's core puts it in `sys.modules` when Python starts,
so `import pyinex` finds it before any file of that name. The extension module used by the
unit tests contains only its `ExcelError` type.

### Functions

- **`pyinex.caller(r1c1=False, full=False)`** returns the address of the cell whose PyCall
  is running, in A1 style (`C12`) or, with `r1c1=True`, in R1C1 style (`R12C3`). With
  `full=True` it puts the workbook and sheet in front, as in `[Book.xlsx]Sheet1!C12`,
  without quotes even when the sheet name has spaces. For a formula entered
  into a range of cells (a legacy array formula), it returns that range, for example
  `C12:E14`. Columns run from `A` to `XFD`. `tests/unit/test_colnames.py` checks the column
  letters at every boundary.
- **`pyinex.sheet()`** returns the calling cell's workbook and sheet, for example
  `[Book.xlsx]Sheet1`. Workbook and sheet names keep every character, including Unicode.
- **`pyinex.esc_pressed(clear=False)`** returns True if the user has pressed Esc during
  the current calculation. A long-running function should check it and return early. With
  `clear=True`, Pyinex tells Excel to forget the key press, so other cells that check see
  False until Esc is pressed again. The default keeps the key press, so one press of Esc
  stops every function that checks. Excel also sees a key press that is kept: after the
  function returns, Excel stops the calculation and leaves the rest of it pending
  (`Application.CalculationState` is `xlPending`). A block that the cell returns then does
  not spill until the next calculation, which runs the formula again
  (`tests/excel/test_esc.py`).
- **`pyinex.ExcelError`** is the type that Excel errors arrive as and that a function can
  return; see [conversion.md](./conversion.md#excel-errors).

### When they may be called

`caller`, `sheet` and `esc_pressed` ask Excel for information. Microsoft allows that only
on Excel's calculation thread, while Excel is calling the add-in. So each of them raises
`RuntimeError`, without asking Excel:

- when called from any other thread, such as one that a script started;
- when called outside a PyCall, for example by code that runs while the add-in closes.

`caller` and `sheet` also raise `RuntimeError` when the running PyCall was not called from
a cell, as happens through `Application.Run` in VBA or COM. Excel's
`Application.Evaluate`, however, calls the function as if from the active cell, so there
they return that cell. `ExcelError` works anywhere.

Tests: in real Excel, `tests/excel/test_pyinex_module.py` (every form against the calling
cell, legacy array and spilling formulas, the far corner of the grid, and the refusals); `tests/host/test_pyinex_module.py` (both address styles,
full and short forms, a Unicode sheet name, `esc_pressed` clearing and keeping the key
press), `tests/host/test_lifetime.py` (the refusals).

## Choosing a Python

One add-in works with any installed x64 CPython 3.10 or later from python.org, because it
uses only Python's stable ABI: the part of Python's C interface that does not change
between versions. The add-in chooses a `python.exe` the first time it needs Python: at the
first PyCall, or the first PyInfo item.

### Two DLLs: the loader and the core

The add-in is two files, in the same folder:

- **`pyinex.xll`, the loader**, is the file Excel loads. It does not link to Python. It
  holds the Excel entry points and the table of functions to register. It chooses and
  checks a Python (below), loads that Python's DLLs, and then loads the core.
- **`pyinex_core.dll`, the core**, holds all the code that uses Python. It links to
  `python3.dll`, the stable ABI's DLL. It is compiled against Python 3.15's headers with
  `Py_LIMITED_API=0x030A0000`, so it can use nothing newer than 3.10's stable ABI, and
  nothing that 3.15 deprecates.

The loader calls the core through a few exported C functions, listed in `src/core/api.h`.
Both DLLs are built from the same `src/excel` code against the same C runtime, so the
loader's `xlAutoFree12` can free a value that the core returned.

**Why two DLLs.** Windows will not load a DLL linked to `python3.dll` until it has found a
`python3.dll`, and that happens before the add-in can choose one. Microsoft's linker can
delay loading a DLL until its first use (`/DELAYLOAD`), but not a DLL whose data the code
uses, and Python's C interface includes data: `Py_None`, the exception types, and the type
objects used by checks such as `PyFloat_Check`. Linked that way, the linker refuses:
`LNK1194: cannot delay-load 'python3.dll' due to import of data symbol
'__imp_PyBool_Type'`. So the code that uses Python is in a DLL of its own, which the loader
loads once it has chosen a Python.

**Loading the chosen Python's DLLs.** The loader loads `python3NN.dll` and `python3.dll`
by full path from the chosen installation, and then loads the core. The core's link to
`python3.dll` then uses the copy already loaded. Both are loaded by full path: with only
`python3NN.dll` loaded, Windows can find another installation's `python3.dll` on `PATH`,
and the core then runs with that Python.

### Where Pyinex looks, in order

1. **`PYINEX_PYTHON`**, an environment variable holding the path of a `python.exe`, or of
   the folder that contains one.
2. **`pyinex.ini`** in the add-in's folder. A relative `executable` is relative to that
   folder, so the two can move together.

   ```ini
   [python]
   executable = C:\Users\me\venvs\finance\Scripts\python.exe
   ; optional: Python's start-up verbosity, as PYTHONVERBOSE
   verbose = 0
   ```

   The file may be saved as UTF-8, with or without a byte-order mark, or as UTF-16. Pyinex
   reads UTF-8 files itself, with Windows' rules: names are compared without regard to
   case, `;` starts a comment, and quotes around a value are removed. Windows' own reader,
   `GetPrivateProfileStringW`, reads a file with no byte-order mark as the 8-bit code page,
   but editors save UTF-8, so a path such as `C:\Users\Zoë\…` would be read as a different
   path. A UTF-16 file is still read by Windows' reader. Tests:
   `tests/host/test_python_discovery.py::test_ini_names_a_path_beyond_the_ansi_code_page`.

3. **The first `python.exe` on `PATH`**, skipping the Microsoft Store's placeholder
   `python.exe` under `…\Microsoft\WindowsApps`. That placeholder opens the Store instead
   of running Python.
4. **The registry**, when `PATH` has none. python.org's installer records each
   installation under `Software\Python\PythonCore\<version>`
   ([PEP 514](https://peps.python.org/pep-0514/)): in `HKEY_CURRENT_USER` when installed
   for one user, its default, or in `HKEY_LOCAL_MACHINE` for all users. Pyinex takes the
   newest version whose name is a plain `3.N`. Names such as `3.14-32`, `3.14-arm64` and
   `3.14t` are 32-bit, ARM64 and free-threaded builds, which it would refuse. Of two
   entries for the same version, the one for the user comes first. The program is the
   entry's `InstallPath` `ExecutablePath` value, or else `python.exe` in the folder that
   `InstallPath`'s default value names.

   The installer leaves `PATH` alone unless asked, so without this step a default
   installation could be used only through `PYINEX_PYTHON` or `pyinex.ini`. It comes
   last because `PATH` and the other two are choices someone made; the registry only
   records what is installed. Tests: `tests/host/test_python_discovery.py`, the tests
   named `test_registry_*`. The fake Excel gives each test a registry of its own, so no
   test sees or changes the machine's.

### Virtual environments

A `python.exe` in a virtual environment (a folder with a `pyvenv.cfg` file) is used as
that environment, the same as running it would be:

- `sys.prefix` and `sys.exec_prefix` are the environment's folder. `sys.base_prefix` is
  the Python it was made from (the `home` line in `pyvenv.cfg`).
- The environment's packages import. The base Python's packages do not, unless the
  environment was made with `--system-site-packages`.
- The user's own site-packages folder (`%APPDATA%\Python`) is included or excluded as it
  would be for that `python.exe`: included for a plain installation and for an environment
  made with `--system-site-packages`, excluded for an ordinary environment.
- `sys.executable` is the chosen `python.exe`, so `subprocess` and `multiprocessing` start
  the right Python.

**How.** Python's configuration API, which could set this directly, is not in the stable
ABI. `Py_SetProgramName` is, but Python 3.11 deprecated it, and testing showed that Python
3.10 ignores it: the embedded `sys.executable` remained Excel. Instead, the core sets the
environment variable `__PYVENV_LAUNCHER__` to the chosen `python.exe` just before starting
Python. A virtual environment's own `python.exe` on Windows uses the same variable to
start its base Python. Python takes `sys.executable` from it, finds the `pyvenv.cfg` beside
it, sets the prefixes and site-packages from that file, and then removes the variable, so
processes that a script starts do not inherit it. This works on each Python from 3.10, for
both kinds of environment, with no deprecated call.

Python still takes two values from the host program (Excel), and the core corrects them
immediately after start-up: it sets `sys._base_executable` to the value the chosen
`python.exe` reports, and removes the `sys.path` entry for Excel's folder.
`tests/host/test_embedded_python.py` and `test_python_discovery.py` compare the embedded
interpreter's version, executables, prefixes and `sys.path` with what the `python.exe`
itself reports, for a plain installation and both kinds of environment.

Pyinex sets every variable that affects Python's start-up (`__PYVENV_LAUNCHER__`, and
`PYTHONHOME` and `PYTHONVERBOSE` below) in two places: Windows' copy of the environment
and the release C runtime's copy (`ucrtbase.dll`), which is where Python reads them. A
debug build of the add-in uses its own C runtime, and a change made only there would not
reach Python.

### Inherited environment variables

The add-in gets Excel's environment variables.

- **`PYTHONHOME`** is removed before Python starts, with a console note
  (`ignoring PYTHONHOME (…)`). Pyinex finds the standard library from the Python it chose,
  and a `PYTHONHOME` that names another installation would make start-up fail.
- **`PYTHONPATH`** is used, as `python.exe` uses it: its folders go on `sys.path`. The
  console notes it at start-up, because a `PYTHONPATH` meant for another Python version
  can cause confusing import errors.
- **Other `PYTHON…` variables** (`PYTHONVERBOSE`, `PYTHONUTF8`,
  `PYTHONDONTWRITEBYTECODE` and the rest) work as they do for `python.exe`. `verbose` in
  `pyinex.ini` sets `PYTHONVERBOSE` for Excel's process only when it is not already set.

### Refusals

Pyinex refuses a Python that it cannot run safely, so that Excel does not crash. The
console shows one line, for example:

```text
Pyinex: cannot use Python C:\Python39\python.exe: version 3.9 is too old; Pyinex needs 3.10 or later. Choose another with PYINEX_PYTHON or pyinex.ini.
```

| Case | Reason given |
|---|---|
| The named file does not exist | `file not found` |
| Nothing found on `PATH` or in the registry | `no python.exe on PATH, and none registered` (adding `skipped the Microsoft Store placeholder at …` when it skipped one) |
| Older than 3.10 | `version 3.N is too old; Pyinex needs 3.10 or later` |
| Not an x64 build | `it is a 32-bit (x86) build` or `it is an ARM64 build`, followed by `; Excel needs the x64 build`. Excel on Windows on Arm loads only x64 add-ins, and an x64 add-in can load only an x64 Python. |
| The Store placeholder, named explicitly | `it is the Microsoft Store placeholder, not an installed Python; install Python from python.org` |
| A free-threaded build (`python3.13t.exe`) | `free-threaded builds are not supported` |
| A virtual environment whose base Python is missing | `its pyvenv.cfg names base Python <path>, which does not exist` (or `names no base Python`) |
| Not a Windows program | `it is not an x64 Windows program` |
| It cannot start (see below) | `it failed to start (exit code N): <its last line of output>`, or `it did not finish starting within 60 seconds` |
| `pyinex_core.dll` is missing beside the add-in, or blocked | `cannot load <path>\pyinex_core.dll (Windows error N)` |
| Another add-in has already loaded a Python into Excel | `another add-in has already loaded Python from <path>; Pyinex cannot share Excel with it` |

The last check is needed because one process can hold only one Python. If another add-in
has loaded a Python DLL, Windows could give Pyinex that copy instead of the one it chose,
and Python's start-up call would do nothing, without an error, if the other add-in's
Python were already running. So Pyinex refuses whenever any Python DLL is already in
Excel's process, from any installation.

All these checks run before Python is loaded, so each refusal is safe. The checks on the
file come first: whether it exists, its name, and its CPU type, read from its program
header. Then, before loading anything, the loader **runs the chosen `python.exe` once**,
with the same environment the embedded Python will have, and has it print four things: its
version, whether it is free-threaded, the path of its `python3NN.dll`, and its
`sys._base_executable`. This is how Pyinex finds the DLL to load.

The same run also detects an installation that is damaged, for example one whose standard
library was deleted. With such an installation, the embedded start-up call would fail, and
the stable ABI's start-up call ends the process when it fails, which would end Excel. Run
as a separate process, it fails without harm, and the refusal gives its last line of
output. The run takes a fraction of a second, once per Excel session. It cannot detect an
installation that changes between the run and the start-up a moment later, or a start-up
setting that only Excel's process sees, such as an `EXCEL._pth` file (Python's path
override file) beside `EXCEL.EXE`.

A Python newer than the newest tested version (3.15) is used, with a console note that it
is untested.

After a refusal:

- every PyCall gives `#VALUE!`;
- `PyInfo("status")` returns the refusal message, and `PyInfo("pyinex")` the add-in's
  path, since neither needs Python; every other item gives `#VALUE!`, and `PyInfo()` gives
  the table with `#VALUE!` beside those items;
- `PyConsole` and `PyReloadOnChange` work as usual, since they do not use Python.

`PyInfo("status")` and `PyInfo("pyinex")` work after a refusal so that the user can read
the reason in a cell without opening the console. Pyinex tries again only when the add-in
is reopened (removed and added again in Excel's Add-ins dialog), not on every call. Excel's
environment cannot change in the meantime, and a new attempt on every call would repeat
the same message on every recalculation.

Tests: `tests/host/test_python_discovery.py` (`PATH`, `PYINEX_PYTHON`, `pyinex.ini`, a
virtual environment with and without `--system-site-packages`, a package only in the
user's site-packages, `PYTHONHOME`, each refusal including a Python DLL loaded before the
add-in, and each function's result after a refusal), `tests/host/test_embedded_python.py`
(the running interpreter is the Python chosen). In real Excel,
`tests/excel/test_python_choice.py` (a `pyinex.ini` naming no Python, and each function
after it; each `PyInfo` item starting Python; `PyConsole` before Python starts) and the
driver's runs under every Python, a virtual environment, and a virtual environment in a
folder whose name contains spaces and Unicode.

## Lifetime

What happens when Excel opens, closes and reopens the add-in:

- **Opening** (`xlAutoOpen`): Pyinex registers its functions. Python does not start yet,
  so an Excel session that never calls Pyinex does not load Python.
- **First use:** Pyinex chooses and starts Python (above). It then pins the add-in in
  memory (`GetModuleHandleExW` with `GET_MODULE_HANDLE_EX_FLAG_PIN`), so Windows never
  unloads it. This is needed because Python and the `pyinex` module now hold pointers to
  the add-in's code.
- **Closing** (`xlAutoClose`, which Excel also calls when the user only unticks the add-in
  in the Add-ins dialog): Pyinex flushes Python's console output. Code that runs then, such
  as the `flush` method of a replaced `sys.stdout`, runs outside any PyCall, so a `pyinex`
  function that would ask Excel for information raises `RuntimeError` (see
  [When they may be called](#when-they-may-be-called)). The interpreter and the script cache
  are left as they are, and Python is **never finalized**. After Python has started, the
  add-in stays loaded because it is pinned. It is also pinned once it has made the console
  window, because the window's code and the thread that sends text to it are in the
  add-in. If neither has happened, the add-in unloads normally.
- **Reopening** (`xlAutoOpen` again): Pyinex registers its functions again and uses the
  running interpreter and its script cache. Changed scripts reload as usual. A Python that
  was refused is tried again.
- **Excel exits:** the process ends without finalizing Python, so `atexit` handlers and
  object finalizers do not run. Scripts that write files should close them.

Why Python is never finalized: CPython cannot reliably start again in the same process
after it has been finalized, and extension modules such as NumPy cannot be initialized a
second time at all. Finalizing at exit can also hang while waiting for threads a script
started. One interpreter for the life of the process is simpler and safe. The cost is that
replacing the add-in's files needs an Excel restart once Python has started.

No Python code runs from `DllMain`, or from a C++ static object's constructor or
destructor. Windows holds its loader lock while those run, and calling Python then can
deadlock.

Tests: `tests/host/test_lifetime.py` (open, call, close, reopen, call; the add-in still
in memory after the host unloads it, which shows the pin works; refusals after closing; no
Python during unloading; handle and reference counts unchanged across many calls and
reloads). In real Excel, `tests/excel/test_lifetime.py` (unloading the add-in with Excel's
`UNREGISTER`, whose `xlAutoClose` flushes output outside any call, and loading it again
with the same interpreter and module state) and `tests/excel/test_scale.py` (handles and
private memory across 10,000 calls).

## Threading

- PyCall and the other functions run only on Excel's main calculation thread, because they
  are not registered as thread-safe.
- Pyinex holds the GIL (Python's global interpreter lock, which lets one thread at a time
  run Python code) only while it runs a call. After start-up it releases the GIL, and it
  takes the GIL again for each call (`PyGILState_Ensure` and `PyGILState_Release`). So
  threads that a script starts keep running between Excel's calls.
- `pyinex` functions that ask Excel for information refuse to run on any other thread
  (see [When they may be called](#when-they-may-be-called)), because Microsoft forbids
  calls to Excel's C API from threads on which Excel did not call the add-in.
- Pyinex has one thread of its own. It sends output to the console window and never calls
  Python (see [Console](#console)).
- Esc does not interrupt running Python. A long calculation should check
  `pyinex.esc_pressed()`.
- A call that runs a window's message loop, such as `matplotlib.pyplot.show()` with an
  interactive backend or Tk's `mainloop()`, does not return until that window closes, and
  Excel waits for it.

Tests: `tests/host/test_threads.py` (a thread that a script starts keeps running between
calls), `tests/host/test_lifetime.py` (the refusal); in real Excel,
`tests/excel/test_threads.py` (with multi-threaded recalculation on, every call runs on
Excel's main thread, the one that started Python), `tests/excel/test_output.py` (a
script's thread printing between calls) and `tests/excel/test_pyinex_module.py` (the
refusal).

## Console

Pyinex's console is a window of its own: a read-only text box titled "Pyinex console".
`print()`, tracebacks and Pyinex's own messages are written to it. It is not a Windows
console, because Microsoft 365 Excel can refuse to create one: on Windows 11 on Arm,
`AllocConsole` fails with "invalid parameter" (error 87), whether a user or automation
started Excel, while another program on the same machine can create one. A window of Pyinex's own also cannot end Excel when it is closed, as closing
a Windows console would, so nothing has to be disabled.

- **The window.** Pyinex creates it, hidden, when it is first needed: when Python starts,
  or when `PyConsole` first shows it. It is a separate window with its own taskbar button. A window owned by Excel would close with the workbook window that
  owned it, since each workbook has its own window. When shown, it comes to the front
  without taking the keyboard from Excel. Closing it hides it, and `PyConsole(TRUE)` shows
  it again. It keeps about the last million characters and removes the oldest. Once the
  window exists, the add-in stays loaded, because the window's code is in the add-in.
- **Light or dark.** The window uses Windows' app mode (Settings, Personalization,
  Colors), and changes when the mode changes while the window is open. In dark mode it
  shows light text on dark gray, with a dark title bar and scroll bar. The title bar's
  color comes from `dwmapi.dll` and the scroll bar's from `uxtheme.dll`, two system DLLs
  that the add-in uses for nothing else. Windows documents no dark scroll bar. Pyinex uses
  `DarkMode_Explorer`, the theme of File Explorer's own scroll bars; if that theme is
  missing, the scroll bar stays light.
- **Where output goes.** Excel starts with no standard output or standard error. In that
  case, before Python starts, Pyinex creates a pipe and connects standard output and
  standard error to it. Python's `sys.stdout` and `sys.stderr`, which Python opens on
  them at start-up, then write into the pipe, and Python needs no change. Python takes them
  from the C runtime's `stdout` and `stderr` streams. In a program with a window, such as
  Excel, those streams start marked as having no output, so Pyinex first reopens them and
  then connects them, and file descriptors 1 and 2, to the pipe. It does this in
  `ucrtbase.dll`, the Universal C runtime that Python 3 and C extensions use, so their
  `printf` output also arrives. It names that DLL instead of using its own runtime,
  because a debug build of the add-in uses the debug runtime.
- **A thread of Pyinex's own reads the pipe** and passes the text to the window. The
  window adds the text when Excel's thread next handles its window messages: between
  calculations, or during a long calculation when it lets Windows run. The thread never
  calls Python and never waits for the window, so a script that prints a lot of text
  cannot wait on a busy Excel. Text that arrives faster than Excel shows it is kept, up to
  the window's limit. The thread starts before any output goes to the pipe, because a pipe
  that nobody reads would stop Excel once it was full.
- **UTF-8, one line at a time.** At start-up, Pyinex sets `sys.stdout` and `sys.stderr`
  to UTF-8 and to write each line as soon as it ends. So any character can be printed, and
  `print()` from a thread that a script started appears without waiting for a buffer to
  fill. Pyinex's own messages are also written in UTF-8. After each PyCall, Pyinex also
  flushes both streams.
- **A host that already has output keeps it.** When the process already has standard
  output and standard error, as when Excel is started from a terminal, or in the fake Excel
  used by the tests, which captures the output, Pyinex leaves them alone and Python writes
  there. The window still exists for `PyConsole`, but receives no text.
- Pyinex's own messages (refusals, conversion errors, start-up notes) go to the same place.
  Each line starts with `Pyinex:`.

Tests: `tests/host/test_console.py`:

- In the fake Excel with output of its own, `print()` in a function and the traceback of a
  failing call reach the host's captured output.
- In a host started with no output, as Excel is, these reach the window instead, along
  with Unicode text split across reads, C code's `puts`, and `print()` from a script's own
  thread.
- Showing, hiding, closing, clearing, placement and rounding.
- Light and dark colors, and a change of mode while the window is open.

The fake Excel handles its windows' messages before each request, as Excel does between
calculations. In real Excel, `tests/excel/test_examples.py` shows the console through
`examples/addin.xlsx` and checks that the window is visible and holds the text the
example's `print()` wrote; `tests/excel/test_output.py` checks each way output arrives
(`print()`, tracebacks, a script's thread, C code's `puts`, any character, a Unicode
traceback); and `tests/excel/test_console.py` checks the colors the window draws in light
and dark mode and when the mode changes, read from its pixels.

## Calling Excel

The add-in calls Excel through one function, `call_excel` (`src/excel/excel.cpp`). Excel
exports a function named `MdCallBack12` from its own program file. `call_excel` finds it
once, with `GetProcAddress` on the process's main module, and passes every call to it. In
a process that exports no such function, every call fails with `xlretFailed`.

Microsoft's Excel SDK does the same in `Excel12v`, in `XLCALL.CPP`. Pyinex does not use
it. The SDK's license forbids distributing the SDK's files, and
does not say whether a program compiled from them may be distributed. With a function of
its own, the add-in contains no code of Microsoft's. It uses only the type and
constant definitions in `XLCALL.H`, at compile time ([sdk/README.md](../sdk/README.md)).
Tests: all of `tests/host`, whose fake Excel exports `MdCallBack12` as Excel does, and
all of `tests/excel`.

## Test dependencies

The add-in depends on nothing but Windows, the C runtime and the Python it chooses. Both
DLLs use Microsoft's Visual C++ runtime (`MSVCP140.dll`, `VCRUNTIME140.dll` and
`VCRUNTIME140_1.dll`). Microsoft 365 installs a copy in Excel's own folder, which Windows
searches first; with another Excel, Microsoft's Visual C++ Redistributable (x64) supplies
it. The tests use four Python packages, none of which the add-in loads or includes
([testing.md](./testing.md)):

- **pytest** runs all three kinds of tests.
- **openpyxl** writes the example workbooks (`scripts/make_examples.py`) and checks them
  without Excel (`tests/unit/test_examples.py`), on any operating system.
- **pywin32** controls real Excel through COM, and reads Pyinex's console window from
  outside Excel's process (`tests/excel`). Only the Python that runs the real-Excel tests
  needs it.
- **NumPy** is installed in each Python that the real-Excel tests embed. The tests use it
  for eigenvalues and matrix multiplication, and a NumPy array's
  `.tolist()` is how a script returns a large block, which the tests check at the full size
  of Excel's grid. Pyinex converts no NumPy array itself; only NumPy's `float64`, a
  subclass of `float`, is accepted as a number ([conversion.md](./conversion.md)).

## How a call moves through the code

This section lists what happens, file by file, when a cell calculates
`=PyCall("model.py", "f", A1)`. The other three functions take steps 1 and 2 (PyInfo also
takes step 3) and then do their own small amount of work in `src/loader/addin.cpp`.

1. **Opening.** Excel loads `pyinex.xll` and calls `xlAutoOpen` (`src/loader/addin.cpp`),
   which registers the four functions from the table at the top of that file through
   `register_functions` (`src/excel/excel.cpp`).
2. **The call arrives.** Excel calls the exported `PyCall` with 32 pointers to
   `XLOPER12`, Excel's value type: the file, the function and 30 argument slots. Every
   entry point runs its body inside `guarded` (`src/excel/excel.h`), which turns a C++
   exception into `#VALUE!` so that the exception does not reach Excel. In the Function
   Arguments dialog, the function returns a blank at once (`in_function_wizard`,
   `src/excel/excel.cpp`).
3. **Python, the first time.** `python()` (`src/loader/python.cpp`) runs once per Excel
   session. It:
   - creates the console window and sends output there, if the process has no output
     (`prepare_output`, `src/loader/console_window.cpp`);
   - sets the environment that Python will see;
   - chooses a `python.exe`, and runs it once to check it and to learn its DLL's path;
   - loads that `python3NN.dll`, the same folder's `python3.dll`, and `pyinex_core.dll`,
     each by full path;
   - finds the core's four exported functions (`src/core/api.h`);
   - pins the add-in in memory;
   - calls the core's `pyinex_start` (`src/core/core.cpp`).

   `pyinex_start` starts Python and runs a short Python program that corrects
   `sys._base_executable` and `sys.path` and sets the output streams to UTF-8. It then
   creates the `ExcelError` type (`src/core/convert.cpp`), the `pyinex` module
   (`src/core/module.cpp`) and PyCall's Python code (`src/core/call.cpp`), and releases
   the GIL. If Python was refused, the refusal is kept, and every later call returns
   `#VALUE!` at once.
4. **Into the core.** `pyinex_call` (`src/core/core.cpp`) takes the GIL, records that Excel
   is calling (so that the `pyinex` module may ask Excel for information), and calls
   `core::call` (`src/core/call.cpp`).
5. **Arguments.** `core::call` drops the missing arguments at the end and converts the rest
   with `to_python` (`src/core/convert.cpp`). It asks Excel for the workbook's folder only
   when the path is relative (`caller_folder`, `src/excel/excel.cpp`: `xlfCaller`, then
   `CELL("filename")`).
6. **The script and the function.** The rest of PyCall is written in Python, kept as a
   string in `src/core/call.cpp` and compiled at start-up, because loading files and
   catching every exception take less code in Python than through the C API. Its `call`
   function finds or loads the script (`_load`: the cache keyed by path, the size and time
   check, compiling from source), finds the function, and calls it. If an exception
   occurs, it prints a traceback without Pyinex's own frames, and returns a value that
   means "failed, and the reason has been printed".
7. **The result.** Back in C++, `to_excel` (`src/core/convert.cpp`) converts the Python
   result into an `excel::Value`, or into `#VALUE!` with a console message that names the
   problem. Pyinex flushes `sys.stdout` and `sys.stderr`, releases the GIL, and returns the
   value as an `XLOPER12` on the heap, marked `xlbitDLLFree`.
8. **Freeing.** Excel copies the value into the cell and passes it back to `xlAutoFree12`
   (`src/loader/addin.cpp`), which frees it with `free_returned` (`src/excel/excel.cpp`).
   The core and the loader use the same C runtime, so the loader can free what the core
   allocated.

**Closing.** `xlAutoClose` calls the core's `pyinex_close`, which flushes Python's output.
Python itself keeps running (see [Lifetime](#lifetime)).

The tests divide the code at the same points: `tests/unit` calls the code of steps 5 to 7
directly, through an extension module built from the same source files, and `tests/host`
runs steps 1 to 8 in a fake Excel (see [testing.md](./testing.md)).

## Not yet built

### Direct functions

This feature is not in this release. It would let a cell call a Python function by its own
name, `=Discount(A1, B1)` instead of `=PyCall("model.py", "Discount", A1, B1)`. The
function would be declared in Python, for example with `@pyinex.func`, and would appear in
the Function Arguments dialog with its argument names and help text. This release only
avoids decisions that would prevent it.

**How it would work.** Excel registers only functions that the XLL exports by name, and a
Python function written after the add-in was built cannot be one of those names. So the
XLL would contain a fixed number of numbered spare exports. At registration, each Python
function would take one of them, and Excel would be given the Python name for it. Each
spare export would do what `PyCall` does, with the file and function fixed at
registration. The value rules, the error rule, script loading and reloading, choosing a
Python, and the main-thread rule would all apply unchanged.

**What this release does for it.**

- The registration code takes a table of entries of any length and knows nothing about
  particular functions. The four functions above are registered through it. Tests:
  `tests/host/test_registration.py`.
- The spare exports would go in the loader, `pyinex.xll`, which is the DLL that Excel
  loads, and each would call the core (see
  [Two DLLs](#two-dlls-the-loader-and-the-core)). The loader already registers its
  functions from a table and passes calls to the core, so the split needs no change.

**Questions left open.**

- Python would have to start when the add-in opens, to learn which functions exist. That
  should be a setting, so that `PyCall` users keep the start at first use.
- Which scripts to read at start-up; probably a list in `pyinex.ini`.
- A workbook that calculates before its functions are registered shows `#NAME?` until the
  next recalculation.
- When a reloaded script adds, removes or renames a function, the old name must be
  removed. Excel's `xlfUnregister` leaves removed names visible in the Function Arguments
  dialog.

### Other ideas

These are possible later work, not yet built:

- **Keyword arguments.** A cell still cannot show whether `name=value` is a keyword
  argument or text (see [Argument rule](#argument-rule)).
- **Running Python text typed in a cell**, instead of a function in a file.
- **Restarting Python** without restarting Excel. CPython cannot do this reliably (see
  [Lifetime](#lifetime)).
- **Writing output and errors to a log file.** Output goes only to the console window
  ([Console](#console)).
- **A command to reload a script**, and **change reports from Windows** in place of the
  check on each call. Neither is needed: see [Scripts](#scripts).
- **Calls on several calculation threads, or several interpreters.** Pyinex runs every
  call on Excel's main thread, in one interpreter ([Threading](#threading)).
