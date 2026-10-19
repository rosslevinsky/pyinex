# Testing

This document explains how to build Pyinex, what each kind of test checks and why, and how
to run them.

Pyinex is an Excel add-in, so it builds and runs only on Windows. GitHub Actions
(`.github/workflows/ci.yml`) builds it and runs the unit and fake-Excel tests on every
push. The real-Excel tests need a Windows desktop with Excel, so they are run by hand.

There are three kinds of tests. Each finds problems that the others cannot:

| Kind | What it runs | Where it runs | What only it can find |
|---|---|---|---|
| Unit tests, `tests/unit` | the core's own code, with no Excel | Windows and CI; the examples test on any system | a wrong rule, tested case by case and quickly |
| Fake Excel, `tests/host` | the real add-in, in a small program that acts as Excel | Windows and CI | a broken Excel memory rule, and cases real Excel cannot be made to produce |
| Real Excel, `tests/excel` | the add-in in Excel itself, under every Python | a Windows desktop with Excel | a case where Excel behaves differently from what the fake Excel's authors expected |

The fake Excel answers the add-in only as its authors expected Excel to answer. Real Excel
has found defects that every fake-Excel test passed, which is why the real-Excel tests
exist.

## Building

On Windows you need Visual Studio 2022 Build Tools (C++), CMake 3.26 or later, and an x64
Python 3.15 from python.org. The core compiles against that Python's headers and
`python3.lib`. With several Pythons installed, set `Python_ROOT_DIR` to choose one.

```
cmake --preset x64-release
cmake --build --preset x64-release
```

The build writes everything to `build/x64-release/bin/`:

- the add-in: `pyinex.xll`, with `pyinex_core.dll` beside it (see
  [architecture.md](./architecture.md#two-dlls-the-loader-and-the-core));
- the unit tests' extension module, `_pyinex_testing.pyd`;
- the test programs `xllhost.exe`, `xllhost_gui.exe`, `badxll.xll` and `fakepython.exe`.

The `x64-debug` preset builds the same files into `build/x64-debug/bin/`.

The first configure downloads Microsoft's Excel SDK installer (4.6 MB) and unpacks it in
the build folder. The SDK's license forbids keeping its files in the repository. See
`sdk/README.md`, which also explains how to build offline.

`python scripts/count_lines.py` counts Pyinex's own C++ under `src/`, and fails if the
count is above 3,000 lines. CI runs it on every push.

The unit and fake-Excel tests below need pytest, openpyxl and NumPy in each Python that
runs them, and CI installs all three:

```
python -m pip install pytest openpyxl numpy
```

NumPy is for the tests of the `numpy.ndarray` annotation
([conversion.md](./conversion.md#type-annotations-direct-functions)), which fail without it
rather than skip, so a run without NumPy cannot pass by leaving them out. The rows of that
table for a Python without NumPy are tested by making `import numpy` fail inside the test
(`sys.modules["numpy"] = None`, in `tests/unit` and in the fake Excel's own process),
since every real-Excel run has NumPy.

## Unit tests (`tests/unit`)

These tests run the add-in's own code with no Excel at all. `_pyinex_testing.pyd` is a
Python extension module compiled from the same source files as `pyinex_core.dll`. pytest
loads it by its full path from the build's `bin` folder, and the tests call the add-in's
code through it:

- `test_conversion.py` checks every rule in [conversion.md](./conversion.md), in both
  directions (`src/core/convert.cpp`). A test describes an Excel value as a tuple, such as
  `("num", 1.5)` or `("multi", 2, 2, [...])`; `tests/unit/conftest.py` lists the forms. So
  a test can build any value Excel could pass, including values Pyinex does not recognize.
- `test_colnames.py` checks the column letters for every column from `A` to `XFD`, in both
  directions, against a list built by simple counting (`src/excel/address.cpp`).
- `test_call.py` and `test_module_cache.py` run PyCall's own code (`src/core/call.cpp` and
  `src/core/python/call.py`):
  the argument rule, every kind of exception, and loading, reloading and failing to load
  scripts that the tests write into temporary folders. The module's `call` takes the
  workbook's folder as an argument, since there is no Excel to ask.
- `test_funcs.py` runs the Python half of direct functions (`src/core/python/funcs.py`,
  in `call.py`'s namespace, as the core runs it) under plain Python, with a stand-in
  `pyinex` module: what `@pyinex.func` collects from a script, the help texts from
  docstrings, the signatures refused, and nothing from a script that raised. It needs no
  extension module, so it also runs on macOS and Linux. `test_funcs_reload.py` runs, the
  same way, what a reload of a script PyLoad loaded changes, `test_funcs_names.py` the
  rules for names, `test_funcs_signatures.py` the Excel positions of each kind of
  parameter, each case of a position left out and what the decorator can mark, and
  `test_funcs_annotations.py` every row of the annotation table, NumPy's arrays included,
  with and without NumPy. `test_funcs_range.py` runs `pyinex.Range`: its addresses, sizes
  and areas, the used part of a whole column or row, found by counting, with a stand-in
  sheet answering the questions the core would put to Excel. It needs the extension
  module, so it runs on Windows only. (A unit test's file name must differ from every fake-Excel
  test's: both folders are collected in one run.)
- `test_core_python.py` compiles each of the core's Python files (`src/core/python/`)
  under the Python running the tests, so a syntax only newer Pythons accept fails under
  3.10; checks that the build compiled each into `_pyinex_testing.pyd` whole; and runs
  `bootstrap.py` from that text: its output streams, `info` and `fix_up`.
- `test_examples.py` checks the example workbooks without Excel (see
  [Real Excel](#real-excel-testsexcel)). It needs only openpyxl, so it also runs on macOS
  and Linux, where it and the `test_funcs*.py` files other than `test_funcs_range.py` are
  the only unit tests collected.
- `test_excel_driver.py` checks how the real-Excel driver saves and restores registry keys,
  using a scratch key and no Excel.

The extension module uses the stable ABI, as the core does, so one build works with every
Python. On Windows, after building:

```
python -m pytest tests/unit
```

To test another build, set `PYINEX_BUILD_DIR` to its `bin` folder.

## Fake Excel (`tests/host`)

These tests run the real add-in without Excel. `xllhost.exe` is a small Windows program
that acts as Excel does towards an add-in:

- It loads the add-in, calls `xlAutoOpen`, and records what the add-in registers.
- It answers the add-in's calls into Excel: `xlGetName`, `xlFree`, `xlCoerce`,
  `xlfCaller`, `xlSheetNm`, `xlAbort`, `xlGetHwnd`, and `CELL("filename")` for the calling
  workbook's folder.
- It calls the add-in's functions as Excel does when it calculates a cell.
- It counts the add-in's calls into Excel, so a test can show that a refused call made
  none.
- It can unload the add-in with `FreeLibrary`, as Excel does after closing it, and report
  whether the add-in stayed loaded.
- It runs a registered command as Excel runs one, in a command context, on request, and
  allows `xlfRegister`, `xlfUnregister` and `xlfSetName` only in a command context
  (`xlAutoOpen`, `xlAutoClose` or a command), as Excel does: anywhere else they fail, and the
  host reports a fault.
- It stands in for Excel's COM `Application` object, as a workbook window gives it to
  `AccessibleObjectFromWindow`: `Run` runs a command by its name, as Excel's
  `Application.Run` does, and `Evaluate` answers `PyInfo("pyinex")` with the path of the
  add-in whose `PyInfo` the name reaches. A test can close the workbook window, make it a
  Protected View window, whose object has no `Application`, add an active Protected View
  window, while which `Evaluate` gives `#N/A`, or make `Run` and `Evaluate` refuse, as Excel
  does while a cell is edited. Through `Workbooks` it gives the workbooks open, each
  named in its sheets' names (`[Book1]Sheet1`), their worksheets, and Excel's search of
  formulas (`Find`, `FindNext`), which also finds text, with each cell's `Address`,
  `HasFormula` and `Formula2`; a formula put in a cell (`formula`) and entered again is
  recorded, and calculated when it calls a function with plain arguments; a sheet can be
  locked, as a protected sheet is, so entering a formula there fails. A handler the add-in
  connects to its application events hears a workbook reported opened (`opened`). It keys
  each registration by the add-in's path, as Excel does, so a test can load a second copy
  of the add-in. It can send the add-in's queue
  window its timer's message as many times as it likes, so thousands of attempts take a
  moment, and count the calls of `Run` and the refusals.
- It can make the add-in's next registrations fail, as Excel's refusal does, or the
  creation of its queue window, and carry out a request inside the add-in's next
  registration from a command, as if Excel had done it then: running the queue's command
  again, or calculating `PyLoad`.

It also enforces Excel's memory rules, and reports each broken rule as a fault:

- a value Excel gave the add-in that was freed twice, or never freed;
- a returned value in heap memory without `xlbitDLLFree`, which Excel would never free;
- a returned value whose text or arrays `xlAutoFree12` does not free (the host counts the
  heap's blocks in use before and after);
- an array inside an array, or an empty array.

A test can set:

- the calling cell, with or without its sheet, since Excel may give either;
- the calling workbook's folder, or none, for an unsaved workbook;
- whether Esc was pressed;
- whether the Function Arguments dialog is open. The host creates a window with the
  dialog's class and title, owned by a workbook's window;
- what the add-in reads from the registry: Windows' light or dark app mode, and the
  Pythons that python.org's installer records. From its start, the host's
  `HKEY_CURRENT_USER` and `HKEY_LOCAL_MACHINE` are empty keys of its own, with only a
  copy of Windows' cryptography settings, which Python's start-up reads. A test adds what
  it needs, and the host removes the keys when it exits. So no test sees or changes the
  machine's registry.

A test can also read the process's handle count, and read Pyinex's console window: its
title, position, size, text and colors. It can close the window as its close button does.
The host handles its windows' messages before each request, as Excel does between
calculations, so text that the add-in sends to the window arrives. Tests that need Python
code inside the add-in call the functions in `scripts/helpers.py` through PyCall. Tests of
direct functions (`test_direct_functions.py`, `test_registration.py`) open a copy of the
add-in, in a folder of its own, whose `pyinex.ini` lists the scripts to load as it opens
(`conftest.py`, `startup`), such as `scripts/direct.py`; the host calls each function by
the name it was registered under, through its numbered export. As in Excel, registration
IDs may be negative (the host's always are, so an add-in that takes an ID for positive
fails here), registering an export again gives it the same ID, and a name keeps its ID when the function is
unregistered (`xlfUnregister`) until it is deleted (`xlfSetName`); a name whose function is
gone gives `#VALUE!`. `test_pyload.py`, `test_queue.py` and `test_lifetime.py` test
`PyLoad`, the queue that registers its functions, and closing with work queued, through the
requests above. Two failures reach the add-in's own calls into Windows: as the host loads
the add-in it points the add-in's imports of `SetTimer` and `malloc` at its own, so that
`fail timer` makes the queue's timer fail and `fail memory` makes one allocation fail just
after the add-in asks Excel about a name, Excel registers a function or Excel unregisters
one, after a given number more succeed, as when memory runs out inside the queue's command,
or at once (`fail memory now`), when the core's allocations fail too, once Python has started
and the host has pointed `pyinex_core.dll`'s import of `malloc` at its own as well; tests
that do not know which allocation matters try each in turn. `nest` carries out
requests as the add-in calls `xlfRegister`, as Excel might within that call: running a
command (`PyRepair`, the queue's), closing the add-in, or running out of memory; `nestat`
does so at a chosen registration, such as the stand-in command a removal registers. Nobody
has seen real Excel do any of these within `xlfRegister`: they test that Pyinex is safe if
it does. `fail name` refuses one registration by name, as `PyLoad`'s or the queue
command's at opening; and `test_pyload.py` gives the loader a malformed
description by replacing Pyinex's own describer, which a script reaches through
`pyinex.func`'s globals. `test_reload_functions.py` tests what a changed script changes in Excel,
`test_names.py` the rules for names, with the host's own `xlfEvaluate` of a name, which
knows a few of Excel's functions and fails on request (`fail evaluate`), and a name
registered for another add-in (`foreign`), and
`test_removal.py` what removal leaves, removal with each step failing on request (`fail
unregister`, `fail command`, `fail setname`), and the pool filled. `test_signatures.py`
and `test_annotations.py` test how a call's arguments are placed by the function's
signature and converted by its annotations, each loading one script by `PyLoad` in a host
of the module's own; `test_annotations.py` also makes `import numpy` fail in the host's
Python, for the rows without NumPy. `test_range.py` tests `pyinex.Range` with references
the host passes (`Ref`, `SRef`) to cells it holds (`cells`), reading each with `xlCoerce`
and `COUNTA`, which fault on reading a whole column cell by cell; and a call Excel stops
(`stop`), after which any call into Excel but `xlFree` is a fault until the function
that met it returns; and a sheet deleted or its workbook closed (`forget`), its ID passing
to the next sheet named. `test_repair.py` tests the `PyLoad` cells calculated as a workbook
opens and the `#NAME?` repair: the cells chosen and those left alone, read from what the
host recorded (`entered`), with `PyRepair`; a legacy array formula (`array`), a sheet whose
formulas cannot be entered or read (`lock`), a change handler's own search (`handler`), and
a change handler that runs a command (`change`), which the host stops, with a fault, at 25
runs. The host's stand-in for `Application.Evaluate` calculates `PyInfo("pyinex")` by
calling the add-in's `PyInfo`, as Excel does. (`tick` sends its timer
message only to a queue window of the host's own process, so hosts open side by side do not
reach each other's.)

`badxll.xll` breaks the memory rules on purpose, so that `test_memory_rules.py` can check
that the host notices each one. `hostdriver.py` controls the host from pytest, and every
request fails the test if it caused a fault. Requests go over a pipe of their own, so the
add-in's printed output arrives separately, in `Host.stdout` and `Host.stderr`.

The host starts without a console, as Excel does. `test_console.py` also starts it as
Excel starts, with no standard output or standard error at all, so that the add-in's output
goes to the console window it creates. For that it uses `xllhost_gui.exe`, the same
program built as a program with a window, as Excel is. Only in such a program does the C
runtime start with its output streams marked as having no output, and the host checks that
they are before it runs.

**Which Python the add-in embeds.** `conftest.py` starts every host with an environment of
its own: no `PYTHON…` or `PYINEX…` variables, a `PATH` of only Windows' folders, and
`PYINEX_PYTHON` naming the Python that runs pytest. So the add-in embeds the same Python
as the tests, and the first check in each run, made from inside the add-in, confirms it.
The discovery tests (`test_python_discovery.py`) change that environment on purpose. They
create installations that a machine cannot have, such as one that is too old, 32-bit or
broken, in temporary folders. Some use `fakepython.exe` in place of `python.exe`; it prints
whatever the test asks for.

On Windows, after building:

```
python -m pytest tests/host
```

As for the unit tests, `PYINEX_BUILD_DIR` chooses another build, for example
`build\x64-debug\bin`.

One build must work with every supported Python, so test it under each, without
rebuilding. In `cmd.exe`:

```
for %v in (3.10 3.11 3.12 3.13 3.14 3.15) do py -%v -m pytest tests -q
```

CI does the same. One job builds the add-in and uploads it as the `pyinex` artifact, with
the test programs and `_pyinex_testing.pyd` as the `pyinex-test-harness` artifact. Then a
job for each of Python 3.10 to 3.15 downloads both artifacts and runs `tests/unit`, then
`tests/host`.

A second workflow, `.github/workflows/release.yml`, runs when a version tag such as
`v2026.10.0` is pushed. It builds the add-in, runs `tests/unit` and `tests/host` under
Python 3.15, and publishes a GitHub release: a zip of `pyinex.xll`, `pyinex_core.dll`,
`LICENSE` and `NOTICE`, with that version's `CHANGELOG.md` entry as its notes. It also
records, signed, that it built the zip and the two DLLs (`actions/attest`), which
`gh attestation verify` checks.

## Real Excel (`tests/excel`)

These tests run the add-in in Microsoft Excel itself. They make the checks that no other
test can make, because they do not depend on what anyone expected Excel to do.

`tests/excel` is a pytest suite that controls Excel through COM, with pywin32. pytest
collects it only when it is given the add-in to test (`--excel`), so `pytest tests` and
CI never start Excel. Each test works in a new, saved workbook in a temporary folder of
its own, and writes there any script it calls. The session starts one new Excel, leaving
any open Excel alone, loads the add-in, and checks that Excel embedded the intended Python
before any test runs. Excel rejects a call from another process while it is busy, as it is
while Pyinex's queue runs its command after a calculation or a workbook's opening, and
pywin32 then fails the call at once. So the session registers a COM message filter on its
thread (`IMessageFilter`, made with ctypes, since pywin32 has none), and COM tries such a
call again every 100 ms, for up to a minute. One script, `scripts/excel_tests.py` (the
driver, below), runs the suite under every Python.

### What each file checks

- `test_examples.py` runs the example workbooks. Every formula on every sheet of
  `examples/*.xlsx` is a check, and the cells just to the right of its result hold the
  value it should show. For a formula entered into a range of cells, the expected block,
  of the same size, is just to its right. Because `addin.xlsx` shows the console and
  prints to it, this file also checks that Pyinex's console window is shown and holds that
  text. `direct_functions.xlsx` loads its script with `PyLoad`, so its direct functions
  register a moment after the workbook opens; its cells are read once none shows
  `#NAME?`.

  `scripts/make_examples.py` writes the workbooks from tables in the script, with
  openpyxl, on any system. To change a workbook, change a table and run the script again;
  never edit a workbook in Excel. `tests/unit/test_examples.py` checks, without Excel, that
  every formula has an expected value beside it, uses only the documented functions, or
  the direct functions that the workbook's own script marks `@pyinex.func` when the
  workbook loads that script with `PyLoad`, and names a script and function that exist,
  and that the committed workbooks match what the script writes now.
- `test_identity.py` records which Python Excel embedded, checks that a virtual
  environment sees only its own packages, and checks that files open with the same
  encoding as under that `python.exe`.
- `test_values_in.py`, `test_values_out.py`, `test_failures.py` and `test_blocks.py` check
  the value rules in [conversion.md](./conversion.md) that a worksheet can produce, in both
  directions; each failure's `#VALUE!` and its `Pyinex:` line on the console; and blocks up
  to the full height and width of the grid. They include the 2010 value checks
  (`TypesOfReturnValues`, `SortList`, `CellWordcount`, `HasVarargs`, `UnicodeTest`). The
  functions they call are in `tests/excel/scripts/values.py`, copied beside each test's
  workbook. Returning a NumPy array is tested in `test_numpy.py`.
- `test_paths.py` and `test_reload.py` check how PyCall finds, loads and reloads scripts:
  relative paths from a workbook whose only sheet has its name and from a folder with
  brackets in its name, an unsaved workbook, two workbooks sharing a script, spaces and
  Unicode in every name, and each reload rule, with files saved as an editor saves them.
- `test_pyinex_module.py` checks `pyinex.caller` and `pyinex.sheet` against the calling
  cell, and every `PyInfo` item. `test_numpy.py` runs the 2010 NumPy demo. Their scripts
  are in `tests/excel/scripts/`. Python caches helper modules by their plain names for the
  whole session, so each test gives its helpers names that no other test uses.
- `test_console.py` and `test_output.py` check Pyinex's console window and each way output
  reaches it. Colors are read from the pixels the window draws (`PrintWindow`). The tests
  print solid block characters, so that some pixels have exactly the text's color. A test
  that changes Windows' app mode changes it back, and the driver checks that it did.
- `test_esc.py` presses a real Esc key. The COM call that enters the formula does not
  return until the calculation ends, so a separate thread in the test waits for the
  script's loop to start, brings Excel to the front, and sends the key with `SendInput`.
- `test_python_choice.py` (a refused Python, `PyConsole` before Python starts, each
  `PyInfo` item starting Python) and `test_lifetime.py` (unloading and loading the add-in
  again; unloading it with direct functions queued to come back, and loading it again; and a
  quit canceled at the save prompt, from a thread whose `Application.Quit` waits on the
  prompt while the test presses Esc in it, with direct functions registered and with none) run only in the driver's runs of their kind,
  each in an Excel of its own.
- `test_threads.py` checks that every call runs on Excel's main thread when multi-threaded
  recalculation is on.
- `test_scale.py` checks that 10,000 calls, each with the right result, add at most 20
  handles and 32 MB of private memory to Excel. In testing, the change ranged from 11
  fewer to 2 more handles, and from 4.5 MB less to 8.6 MB more memory, which is Excel's own
  variation from run to run. It checks 10,000 calls of a direct function the same way, and
  600 reloads of a script and 200 registrations again of a direct function with the wider
  bound on handles that reloads need.
- `test_reload_functions.py` checks each kind of change to a script PyLoad loaded, with the
  cells before and after, and each with `PyReloadOnChange(FALSE)`; a script that now fails,
  recovered through a call and through its PyLoad cell; and several edits made within one
  calculation, while Excel cannot run the queue. `test_names.py` checks the rules for names,
  with another add-in's function registered by Excel 4's `REGISTER` and a name a workbook
  defines. `test_removal.py`
  checks what removal leaves, in Excel's registration table (read by `conftest.py`'s
  `registered_types`, which first checks it finds PyCall's registration), in Insert Function
  and in cells; many cycles of registering and removing; and, in an Excel of its own whose
  `pyinex.ini` lists a script of 1,000 functions, the pool filled exactly and one function
  more.
- `test_signatures.py` checks each count of arguments from 0 to 30, each position by its
  value; defaults, `*args`, keyword-only parameters and the examples in
  [architecture.md](./architecture.md#direct-functions); a formula with more arguments than
  the function's positions, which gives `#VALUE!` without a call; each refused signature
  among functions that still work; a class and a `functools.partial`; and reloads that change a function's
  positions. `test_annotations.py` checks each annotation converting and refusing from
  cells, with its console message; annotated and unannotated parameters in one function;
  an annotation Pyinex does not convert; and a changed annotation, which Excel is not told
  of. `test_range.py` checks `pyinex.Range` with every kind of reference, values that are
  not references, the refusals, a full calculation that reaches a function before the cells
  it reads, and the recalculation [architecture.md](./architecture.md#pyinexrange) states.
  Their functions have names of their own (`Xs…`, `Xa…`, `Xr…`), since a registered name stays
  for the session, and each module loads its script once, in a workbook of its own.
- `test_registration.py` checks that Excel lists the five functions and the two commands,
  with their type strings, none marked thread-safe, and, for an add-in whose `pyinex.ini`
  lists `scripts/direct.py`, those and each of its direct functions in its slot.
- `test_pyload.py` checks `PyLoad` and the queue: relative and absolute paths, each way a
  script fails to load and the fixed file after it, a script loaded twice and from two
  workbooks, one script through `PyLoad` and PyCall, registration queued during a full
  recalculation, a long cell edit and a modal dialog (keys sent by the script's own top
  level as `PyLoad` calculates, so that Excel is busy before the queue's first try; only the
  console is read until the test presses Esc), no workbook open, only a Protected View
  window open, and two Excels, each registering its own functions only. A name registered
  stays registered in the session's Excel, so each test's functions have names of their
  own.
- `test_workbook_open.py` checks that a workbook saved in the session's Excel and opened in
  a new Excel loads its scripts, shown by a file the script writes, before anything is
  changed, and that its cells then follow their inputs: with a literal path, a path in a
  cell, the `PyLoad` cell on a hidden sheet, a volatile function, and the workbook open
  before the add-in loads; that one opened with Excel's events off is missed until
  Ctrl+Alt+F9 or `PyRepair`; and that an Excel with no workbook open says nothing until one
  opens, then loads it. `test_repair.py` checks the `#NAME?` repair (the plan's test
  group 8): cells typed before registration on two sheets, with other functions, spilling,
  inside `IFERROR`, in an untaken `IF` branch and in a legacy array formula; cells left
  untouched, among them formulas holding the name only in a text or a longer name; and a protected
  sheet's cell reported, then repaired by `PyRepair`. `test_macro_dialog.py` opens the Macro
  dialog as a user does, pressing Alt+F8 with Excel in front, and reads it as a screen reader
  would, through Microsoft Active Accessibility, since Excel draws its controls itself: the
  dialog lists a command an Excel 4 macro sheet defines but no direct function and not
  `PyRepair`; `PyRepair` typed in it runs; and a direct function's name typed in it calls the
  function with no arguments. It needs the desktop unlocked and nothing else taking the
  keyboard.
- `test_direct_functions.py` checks direct functions in an Excel of their own, started by
  `conftest.py`'s `direct_excel` with a copy of the add-in whose `pyinex.ini` lists
  `scripts/direct.py`: every value rule a worksheet can produce, each called by the
  function's name and through PyCall and the two compared; the call context (`caller`,
  `sheet`, a real Esc, a script's thread); and what Insert Function and the Function
  Arguments dialog show, read through Windows' accessibility interface, as a screen reader
  reads them, with keys sent by `SendInput`, and that a formula typed in the dialog shows
  a blank and runs no Python; and, last, the add-in closed and opened again with another list, after which
  a cell using a name from before gives `#VALUE!`, never the function that took its slot.
- `driver_checks/` holds tests that are wrong on purpose, which check the driver itself
  (see [Checking the driver](#checking-the-driver)). pytest collects them only when they
  are named.

### The driver

`scripts/excel_tests.py` runs the suite several times, each time in a new Excel, because
Python starts once per Excel process and never restarts:

- once for each of Python 3.10 to 3.15;
- once for a virtual environment made from 3.15 (`venv`);
- once for a virtual environment made from 3.15 in a folder whose name contains spaces
  and Unicode, `venv Bücher ✓ 😀` (`unicode-venv`);
- then, under 3.15, the runs that each need an Excel of their own. Each of these runs only
  the tests marked for it (`@pytest.mark.kind`; `tests/excel/conftest.py` reads the run's
  kind from its `expected.json`):
  - `refused`, whose `pyinex.ini` names no Python;
  - `console-first`, where `PyConsole` is the first call;
  - `first-status` to `first-addin_version`, eight runs, where that `PyInfo` item is the
    first call;
  - `lifetime`, which unloads the add-in and loads it again.

The driver needs Windows, and pytest and pywin32 for the Python that runs it. The suite
is written for Microsoft 365 Excel, 64-bit, in English: it enters formulas with
`Formula2`, its block tests expect spilling, and its dialog tests look for the English
titles "Function Arguments", "Insert Function" and "Macro". Another Excel, or Excel in
another language, fails those tests for that reason alone.
Each Python under test needs NumPy; the driver installs NumPy in each virtual environment
it makes. The Python the virtual environments are made from (3.15) also needs pytest:
`test_identity.py` checks that pytest is installed in that base Python and that the
virtual environment cannot import it. Run it at the
signed-in, unlocked desktop, because Microsoft does not support automating Excel from an
SSH or service session. Close Excel first, and run it from a checkout of the commit the
add-in was built from:

```
py -3.14 scripts\excel_tests.py C:\path\to\pyinex.xll
```

**How it chooses the Python.** For each run, the driver copies `pyinex.xll` and
`pyinex_core.dll` into a new temporary folder, removes from the copies the mark Windows
puts on downloaded files (as `Unblock-File` does), and writes a `pyinex.ini` beside them
that names that run's `python.exe`. It cannot use `PYINEX_PYTHON`, because an Excel started
through COM gets the user's environment, not the driver's (see
[architecture.md](./architecture.md#choosing-a-python)). The suite then compares
`PyInfo("executable")`, `PyInfo("version")` and `PyInfo("venv")` with the Python the run
chose, and stops if Excel embedded a different one.

**It refuses to start**, and says why:

- at a locked screen, or outside the desktop session;
- while any Excel is running. The driver never ends an Excel it did not start, and it
  could not restore settings around one;
- when the environment Excel will get sets `PYINEX_PYTHON`, `PYTHONHOME`, `PYTHONPATH`,
  `PYTHONSTARTUP` or `PYTHONUSERBASE`, because any of these would choose or change Excel's
  Python.

**A run passes only if every test passed.** Each of these fails the run:

- a failed test, an error, a skipped test, or no tests at all;
- a run that reaches its time limit (`--timeout`, one hour by default), after which the
  driver ends that run's Excel;
- an Excel that does not quit within 90 seconds and has to be ended. In testing, Excel
  sometimes took a minute to quit, with or without an add-in. This holds for the session's
  Excel and for each other Excel a test starts, such as `direct_excel`'s: `conftest.py`'s
  `report_excel` records its process for the driver, which waits for it in the same way;
- a setting the driver could not restore.

If an Excel is still running after a run and the driver could not end it, the run fails
and the driver stops, because that Excel may change settings as it quits. To end an Excel,
the driver must have its process handle and the start time that the test session
recorded, so it cannot end another process that reused the same process ID.

**It stops early when the rest would fail the same way.** A run whose Excel had to be
ended, that reached its time limit, or in which no tests ran stops the driver after it:
such a fault lies in the add-in or the machine, not in one Python, so each run left would
almost always meet it, at a quarter of an hour each. Failed tests do not stop it, because
which Pythons a test fails under helps tell why. `--keep-going` runs the rest whatever
happens. `scripts/check_excel_driver.py` checks both.

The driver exits with 0 only if every run passed; 1 otherwise; 2 when it refuses to start;
and 130 when stopped with Ctrl+C.

**It leaves the machine as it found it.** Before each run, it saves the registry keys that
the tests may change:

- Windows' app mode: `HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize`;
- Excel's options: `HKCU\Software\Microsoft\Office\16.0\Excel\Options`, which also lists
  the add-ins Excel opens at start-up;
- Excel's list of add-ins: `…\Excel\Add-in Manager`.

After the run, once its Excel has ended, the driver restores any key that changed and
reads it again to check. It does this even when a test failed, the run reached its time
limit, or the driver was stopped with Ctrl+C. `tests/unit/test_excel_driver.py` checks this
against a scratch key.

Everything else the driver writes goes in a temporary folder, which it removes at the end,
or in the results folder. pytest runs without its cache, without writing compiled files,
and without any `PYTEST_` variable, since a setting such as `PYTEST_ADDOPTS=-k …` would run
only some tests and could pass. The tests copy any script they call out of the repository
first.

**Results** go to the folder named by `--out`, by default a new folder under `%TEMP%`:

- a JUnit XML file for each run (`3.10.xml` … `venv.xml`);
- `summary.json`: the SHA-256 of each file of the build, the versions of Excel and
  Windows, and each run's Python, test counts and problems;
- `log.txt`: everything the driver printed.

`--python 3.12` runs only that run; it can be given more than once, and `venv`,
`unicode-venv`, `refused` and the other kinds name the other runs. `--tests` runs a
different selection than `tests/excel`. Give it with `--python` naming the runs it suits:
a run of a kind runs only the tests marked for that kind, so one given an ordinary test
file finds no tests, fails, and stops the driver. To check one file under one Python:

```
py -3.14 scripts\excel_tests.py C:\path\to\pyinex.xll --python 3.14 --tests tests\excel\test_range.py
```

### Checking the driver

`scripts/check_excel_driver.py` checks that the driver fails, refuses and cleans up as it
should. It needs no one once started. Run it at the desktop, with Excel closed:

```
py -3.14 scripts\check_excel_driver.py C:\path\to\pyinex.xll
```

It runs the driver under the newest Python, 3.15, only, on each test in `tests/excel/driver_checks/`:

| Check | Expected exit code |
|---|---|
| a test that fails | 1 |
| a test that is skipped | 1 |
| a script that never returns, with a 60-second time limit | 1 |
| a test that starts a second Excel, reports it as `direct_excel` does, and leaves it running | 1 |
| a test that changes the app mode and then never returns, with a 60-second time limit | 1 |
| the same, stopped with Ctrl+C after the app mode has changed | 130 |
| the driver started with `PYTHONPATH` set | 2 |
| the driver started while an Excel of the check's own is open | 2 |

After each check, it confirms that no Excel is left running and that the settings the
driver saves are as they were. It exits with 0 only if every check passed. The refusal at a
locked screen is checked by reading the code, since nobody can start the check at a locked
screen.

### What it covers

The two tables below list every check the 2010 version made, and every fake-Excel test.
Beside each is the real-Excel test that makes the same check now, or the reason no
real-Excel test does. File names are in `tests/excel/` unless another folder is given.

**The 2010 examples.** These are the scripts in `Examples/` and `Scripts/TestSheets.py`
of Pyinex 0.83, the repository's first commit. Their workbooks ran each function in Excel.

| 2010 | Now |
|---|---|
| `PyinexTest.py`: `PlayWithGlobalVariable` | `test_reload.py::test_a_module_variable_persists_between_calls_2010` |
| `TypesOfReturnValues` | `test_values_out.py::test_types_of_return_values_2010` |
| `SortList`, `stringize` | `test_values_out.py::test_sort_list_2010` |
| `CellWordcount` | `test_values_in.py::test_cell_wordcount_2010` |
| `HasVarargs`, `Columnize` | `test_values_in.py::test_varargs_2010` |
| `USFederalDebt`, `DebtGetter` | Dropped: it read a third party's web page (a US debt clock) whose layout it depended on, so a test of it would fail for reasons not Pyinex's. What it showed of Pyinex, a column returned as rows, is `test_values_out.py::test_a_list_fills_a_block`. |
| `TestHarnessFunc` | Dropped: only 2010's C++ test harness called it; the fake Excel (`tests/host`) replaced that harness, and returning a row is `test_values_out.py::test_a_list_fills_a_block`. |
| `PythonBreakTest.py`: `LongLoop` | `test_esc.py::test_long_loop_2010`, with a real Esc press |
| `PythonExtensionTest.py`: `CallerA1Test`, `CallerA1FullTest`, `CallerR1C1Test`, `CallerR1C1FullTest`, `CallerSheetTest` | `test_pyinex_module.py::test_python_extension_test_2010` |
| `TestAll` | Dropped: it printed the five above and returned one, each checked on its own. |
| `UnicodeTest.py`: `ReturnConstString`, `ReturnString` | `test_values_out.py::test_unicode_2010` |
| `NumPyDemo.py`: `ComputeEigenstructure`, `MatrixMultiply`, `HammerIntoPositiveDefinite` | `test_numpy.py::test_compute_eigenstructure_2010`, `::test_matrix_multiply_2010`, `::test_hammer_into_positive_definite_2010` |
| `PyVersion` | `test_numpy.py::test_the_python_numpy_runs_in_is_excels` |
| `CheckVersion` | Dropped: it refused every Python but 2.5 and 2.6; NumPy runs under every supported Python. |
| `Scripts/TestSheets.py`: `main`, `RunOneExcelVersion` | `scripts/excel_tests.py`, which runs every test under every Python, each in a fresh Excel, as `TestSheets.py` opened each example workbook under each build. |
| `HackExcelCLSID` | Dropped: it chose among several installed Excel versions; the tests use one. |

**The fake Excel's tests** (`tests/host/`). The second column uses three labels:

- "Fake only": the test cannot run in real Excel, and the entry says why.
- "Left to the fake": real Excel could make the check, but this suite does not, and the
  entry says why.
- "Partly": the real-Excel test makes part of the check, and the entry says which part it
  leaves out.

| `tests/host/` | Real Excel |
|---|---|
| `test_calls.py::test_pycall_round_trips_values` | `test_values_in.py::test_a_cell_arrives_as_its_python_value`, `test_values_out.py::test_a_return_value_shows_as_documented`; an unpaired surrogate and an unknown error code are fake only, since a worksheet cannot hold one (Excel passes newer errors as `#VALUE!`) |
| `::test_pycall_round_trips_text_at_excels_limit` | `test_values_in.py::test_text_at_excels_limit_arrives_whole`, `test_values_out.py::test_text_at_excels_limit_is_returned_whole` |
| `::test_pycall_round_trips_ranges` | `test_values_in.py::test_a_range_goes_back_out_the_same_shape` |
| `::test_pycall_none_is_a_blank_cell` | `test_values_out.py::test_none_shows_nothing_but_is_not_blank` |
| `::test_pycall_result_is_freed_by_excel` | Fake only: how Excel frees a result (`xlAutoFree12`) can be watched only inside the host. Real Excel shows a mistake as a crash or a leak: `test_scale.py` |
| `::test_pycall_exception_is_value_error_with_traceback` | `test_failures.py::test_an_exception_in_a_call`, `test_output.py::test_a_traceback_reaches_the_console` |
| `::test_pycall_returned_excel_error` | `test_values_out.py::test_a_returned_excel_error_shows_that_error`, `::test_a_list_fills_a_block[nan_in_block]` |
| `::test_pycall_relative_path_uses_the_workbook_folder` | `test_paths.py::test_a_relative_path_is_from_the_workbooks_folder` |
| `::test_pycall_relative_path_when_excel_gives_the_caller_without_its_sheet` | Fake only: a caller given as a reference with no sheet (`xltypeSRef`), which this Excel never gave; the forms it gives are in `test_paths.py` |
| `::test_pycall_relative_path_with_brackets_in_the_folder` | `test_paths.py::test_a_folder_with_brackets` |
| `::test_pycall_relative_path_when_excel_gives_the_workbook_without_its_sheet` | `test_paths.py::test_a_workbook_whose_only_sheet_has_its_name`, `::test_a_folder_with_brackets_and_a_sheet_named_like_the_workbook` |
| `::test_pycall_relative_path_in_an_unsaved_workbook` | `test_paths.py::test_an_unsaved_workbook_refuses_a_relative_path` |
| `::test_pycall_file_not_text` | `test_failures.py::test_a_path_that_is_not_text` |
| `::test_pycall_reloads_a_changed_script` | `test_reload.py::test_a_changed_script_is_reloaded` |
| `::test_reload_on_change_setting` | `test_reload.py::test_reload_on_change_off_keeps_the_loaded_script` |
| `::test_reload_on_change_refuses_text` | `test_reload.py::test_reload_on_change_refuses_text` |
| `::test_console_returns_show` | `test_console.py::test_show_and_hide` |
| `::test_info_items`, `::test_info_names_ignore_case`, `::test_info_table`, `::test_info_unknown_item` | `test_pyinex_module.py::test_every_pyinfo_item`, `::test_pyinfo_names_ignore_case`, `::test_pyinfo_with_no_item_is_the_whole_table`, `::test_an_unknown_pyinfo_item` |
| `::test_every_function_is_blank_in_the_function_wizard`, `::test_functions_run_once_the_wizard_closes` | Left to the fake: Excel's wizard is a modal dialog (`Range.FunctionWizard`), and driving one from a test, with keys sent to it while it calculates, is fragile in a way the rest of this suite is not; Pyinex recognizes the wizard by its window's class and title, which the fake makes as Excel does |
| `test_console.py::test_print_and_tracebacks_reach_the_hosts_output`, `::test_a_host_with_output_keeps_it_after_pyconsole_makes_a_window` | Fake only: a host with output of its own, which Excel never has; Excel's case is `test_output.py` |
| `::test_with_no_output_python_writes_to_the_console` | `test_output.py::test_print_reaches_the_console`, `::test_a_traceback_reaches_the_console`, `::test_c_codes_output_reaches_the_console` |
| `::test_the_console_shows_any_character` | `test_output.py::test_any_character_reaches_the_console` |
| `::test_print_from_a_scripts_thread_reaches_the_console_between_calls` | `test_output.py::test_a_scripts_thread_prints_between_calls` |
| `::test_with_no_output_pyconsoles_complaints_reach_the_console`, `::test_pyconsole_refuses_bad_arguments` | `test_console.py::test_a_bad_argument_is_refused` |
| `::test_pyconsole_makes_the_window`, `::test_pyconsole_does_not_start_python` | `test_python_choice.py::test_pyconsole_works_before_python_starts` |
| `::test_the_console_follows_windows_light_or_dark_mode` | `test_console.py::test_its_colors_follow_the_app_mode` |
| `::test_the_console_follows_a_change_of_mode_while_open` | `test_console.py::test_its_colors_follow_a_change_of_mode_while_open` |
| `::test_closing_the_console_hides_it` | `test_console.py::test_closing_hides_it_and_showing_places_it_again` |
| `::test_the_console_stays_pyinexs_across_unloading` | `test_lifetime.py::test_the_console_alone_keeps_the_addin_loaded` (made before Python starts), `::test_unloading_and_loading_again_keeps_python_and_its_scripts` (after) |
| `::test_pyconsole_places_the_window_as_it_is_shown` | `test_console.py::test_placing_applies_as_it_is_shown` |
| `::test_pyconsole_rounds_halves_away_from_zero` | `test_console.py::test_placing_rounds_halves_away_from_zero` |
| `::test_pyconsole_refuses_a_place_off_the_desktop` | `test_console.py::test_a_place_off_the_desktop_is_refused` |
| `::test_pyconsole_clears_the_text` | `test_console.py::test_clearing_empties_it`, `::test_clearing_while_shown` |
| `test_embedded_python.py::test_reports_the_chosen_python`, `::test_python_is_running` | `test_identity.py::test_excel_embedded_the_python_this_run_chose`, `test_pyinex_module.py::test_every_pyinfo_item` |
| `::test_starts_as_its_python_exe_would` | Partly: `test_identity.py` (executable, version, a virtual environment's packages) and `test_pyinex_module.py::test_every_pyinfo_item` (prefix, DLL), and `test_identity.py::test_files_open_with_the_encoding_python_exe_uses` (UTF-8 mode and the default encoding, which 3.15 changed). The rest of the comparison with `python.exe` (`sys._base_executable`, `sys.exec_prefix`, `sys.base_prefix`, `sys.path`) is left to the fake: a script in Excel could run `python.exe` and compare, but the start-up code is the same in any host, and the identity checks above show Excel took the Python chosen |
| `::test_eval_returns_text_and_errors` | Fake only: it tests the fake's own helper for evaluating Python |
| `test_exceptions.py::test_every_kind_in_every_place_in_one_process` | Each with its traceback: `test_failures.py::test_an_exception_in_a_call`, `::test_a_script_that_fails_to_load`, `test_reload.py::test_a_reload_that_raises_drops_the_old_version`, `::test_a_change_to_a_syntax_error_drops_the_old_version` |
| `test_lifetime.py::test_open_call_close_reopen_call`, `::test_once_python_starts_the_addin_stays_loaded` | `test_lifetime.py::test_unloading_and_loading_again_keeps_python_and_its_scripts` |
| `::test_before_python_starts_the_addin_unloads` | `test_lifetime.py::test_before_python_starts_unloading_takes_the_addin_out` |
| `::test_no_python_runs_while_the_addin_unloads_or_excel_exits` | `test_lifetime.py::test_unloading_and_loading_again_keeps_python_and_its_scripts`: the flush at `xlAutoClose` is outside any call, and the `atexit` handler and finalizer it leaves write files the driver looks for once Excel has exited |
| `::test_a_thread_a_script_started_is_refused` | `test_pyinex_module.py::test_caller_sheet_and_esc_pressed_refuse_on_another_thread` (during a call), `::test_caller_sheet_and_esc_pressed_refuse_between_calls` |
| `::test_code_run_while_the_addin_closes_is_refused` | `test_lifetime.py::test_unloading_and_loading_again_keeps_python_and_its_scripts` (`caller`, `sheet` and `esc_pressed` in the flush at `xlAutoClose`) |
| `::test_handles_and_reference_counts_stay_flat` | `test_scale.py::test_ten_thousand_calls_leak_nothing` (handles, memory, and the cached script's reference count across calls), `::test_reloads_leak_nothing` (each replaced module freed, through weak references, and handles and memory across 600 reloads, within 150 handles, twice Excel's own drift) |
| `test_memory_rules.py` (all ten) | Fake only: they check that the fake notices a broken Excel memory rule, with `badxll.xll`'s deliberate faults; real Excel would crash or leak without saying which. `::test_pyinex_passes` shows Pyinex keeps the rules; in real Excel, `test_scale.py` |
| `test_pyinex_module.py::test_caller_in_both_address_styles` | `test_pyinex_module.py::test_caller_in_both_address_styles`, `::test_the_calling_sheet_not_the_active_one` |
| `::test_caller_given_as_a_reference_with_no_sheet` | Fake only: `xltypeSRef`, as above |
| `::test_caller_of_a_formula_in_a_range_is_the_range` | `test_pyinex_module.py::test_caller_of_a_legacy_array_formula_is_its_range` |
| `::test_caller_at_the_grids_far_corner` | `test_pyinex_module.py::test_caller_at_the_grids_far_corner` |
| `::test_caller_outside_a_cell_raises` | `test_pyinex_module.py::test_caller_and_sheet_refuse_outside_a_cell` |
| `::test_sheet_keeps_every_character` | `test_paths.py::test_unicode_and_spaces_in_every_name` (`sheet()` and `caller(full=True)`) |
| `::test_esc_pressed_keeps_the_key_press_by_default`, `::test_esc_pressed_can_clear_the_key_press` | `test_esc.py::test_esc_pressed_keeps_the_press_by_default`, `::test_esc_stops_a_loop_and_is_kept_or_cleared` |
| `::test_excel_error_is_the_type_values_arrive_as` | `test_values_in.py::test_an_excel_error_behaves_as_documented` |
| `test_python_discovery.py::test_first_python_on_path`, `::test_path_skips_the_store_placeholder`, `::test_pyinex_python_names_a_folder`, `::test_pyinex_python_comes_before_the_ini_and_path`, `::test_the_ini_comes_before_path`, `::test_a_relative_folder_on_path_gives_an_absolute_executable`, `::test_user_site_packages`, `::test_pythonhome_is_removed`, `::test_pythonpath_is_honored`, `::test_pythonpath_naming_the_host_folder_keeps_it`, `::test_refusal_nothing_on_path`, `::test_refusal_only_the_store_placeholder_on_path` | Fake only: each sets the environment Excel starts with (`PATH`, `PYINEX_PYTHON`, `PYTHONHOME` and so on). An Excel started through COM takes the user's own, which a test must not change, and the driver refuses to start when it sets any variable that would steer Python |
| `::test_registry_python_when_path_has_none`, `::test_registry_install_folder_without_an_executable_path`, `::test_registry_comes_after_path`, `::test_registry_newest_version_by_number`, `::test_registry_skips_other_builds_and_missing_files`, `::test_registry_skips_the_store_placeholder`, `::test_registry_user_before_machine_at_the_same_version`, `::test_registry_newer_machine_entry_before_older_user_entry` | Fake only: each writes the registry Python entries it needs, which a test must not do to the machine's own registry |
| `::test_ini_path_relative_to_the_addin` | The driver's `unicode-venv` run, whose `pyinex.ini` names its Python by a relative path |
| `::test_ini_names_a_path_beyond_the_ansi_code_page` | Partly: the driver's `unicode-venv` run, which found the defect it tests, writes UTF-8; UTF-8 with a byte-order mark and UTF-16 are left to the fake, being the same reader |
| `::test_ini_is_read_by_windows_rules`, `::test_ini_reads_only_the_first_python_section`, `::test_ini_verbose_sets_pythonverbose` | Left to the fake: the add-in's own reading of the file, the same in any host, so an Excel run per case would add only time |
| `::test_venv_starts_as_its_python_exe_would` | Partly: the driver's `venv` and `unicode-venv` runs (`test_identity.py`), as for `test_starts_as_its_python_exe_would` above |
| `::test_venv_packages_and_isolation` | `test_identity.py::test_a_virtual_environment_sees_only_its_own_packages` |
| `::test_system_site_packages_venv_sees_base_packages` | Left to the fake: one more kind of virtual environment, which Python itself, not Pyinex, sets up from `pyvenv.cfg` once started as the driver's are |
| `::test_newer_than_tested_python_is_used_with_a_note` | Fake only: it needs a Python newer than 3.15, which `fakepython.exe` pretends to be |
| `::test_refusal` | Partly: `test_python_choice.py` (the `refused` run: a `pyinex.ini` naming no Python, and every function after it). The other reasons, each from `fakepython.exe` or a crafted file, are left to the fake: the same checks before loading, in any host, whose effects on every function the refused run shows |
| `::test_refusal_python_already_loaded` | Left to the fake: it needs another add-in that loads a Python into Excel first |
| `::test_refusal_is_retried_when_the_addin_reopens` | `test_python_choice.py::test_a_refusal_holds_until_the_addin_is_reopened` |
| `::test_the_check_run_ignores_modules_in_the_current_folder`, `::test_the_check_run_resolves_relative_paths_as_python_will`, `::test_a_check_run_with_much_output_is_read_while_it_runs` | Left to the fake: they set the current folder and environment the check run inherits, which in Excel are the user's own, or need a `python.exe` that misbehaves (`fakepython.exe`) |
| `::test_a_missing_core_is_retried_when_the_addin_reopens` | Left to the fake: the add-in's own handling of a missing `pyinex_core.dll`, the same in any host |
| `test_registration.py::test_registers_the_documented_functions`, `::test_no_function_is_thread_safe` | Partly: `test_registration.py::test_excel_lists_the_documented_functions` (the five functions and the two commands, with their type strings, none thread-safe). Excel's list gives no names of arguments, category or help, which only the fake checks |
| `::test_pycall_argument_text_fits_excels_limit` | Fake only: it checks the text the add-in registers against Excel's 255-character limit, a check of the text itself; `test_registration.py` shows real Excel registered every function |
| `test_threads.py::test_a_thread_keeps_running_between_calls` | `test_output.py::test_a_scripts_thread_prints_between_calls`; and with multi-threaded recalculation, `test_threads.py::test_every_call_runs_on_excels_main_thread` |

## Direct functions, PyLoad and pyinex.Range, rule by rule

Each rule that [architecture.md](./architecture.md) and [conversion.md](./conversion.md)
state for direct functions, `PyLoad` and `pyinex.Range`, with the test that checks it in
each layer: `tests/unit`, `tests/host` (the fake Excel) and `tests/excel` (real Excel).
Where a layer has none, the entry says why. "Partly" says which part of the rule the tests
named check, and which part is left to another layer or is a gap: a part no test checks.
Each table names its rules' section.

**Direct functions** ([architecture.md](./architecture.md#direct-functions))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| A marked function registers under its own name, or the decorator's `name=` | `test_funcs.py::test_collects_the_scripts_marked_functions_in_order`, `test_funcs.py::test_name_category_and_help_from_the_decorator` | `test_registration.py::test_registers_the_direct_functions_pyinex_ini_lists` | `test_registration.py::test_excel_lists_the_direct_functions_pyinex_ini_names`, `test_names.py::test_the_decorators_name_is_the_name_in_excel` |
| The decorator returns the function unchanged | `test_funcs.py::test_the_decorator_returns_the_function_unchanged` | Left to the unit test: plain Python, the same in any host | As for the fake |
| Argument names from the signature; help and argument help from the docstring's `Args:`; category, by default "Pyinex" | `test_funcs.py::test_help_and_argument_help_from_the_docstring`, `test_funcs.py::test_a_docstring_without_args_gives_function_help_only`, `test_funcs.py::test_no_docstring_gives_empty_help` | `test_registration.py::test_registers_the_direct_functions_pyinex_ini_lists` | `test_direct_functions.py::test_insert_function_lists_the_category_and_its_help`, `test_direct_functions.py::test_the_function_arguments_dialog_shows_the_help_and_runs_no_python` |
| Texts cut at 255 characters, argument names only whole; one more, empty, argument help | `test_funcs.py::test_texts_longer_than_excel_takes_are_cut`, `test_funcs.py::test_argument_names_are_given_only_whole` | Partly: `test_direct_functions.py::test_a_last_help_text_at_excels_limit_is_registered_whole` (a last help text of exactly 255 characters registered whole, and the extra empty help; the cutting is left to the unit tests) | Partly: `test_direct_functions.py::test_insert_function_lists_the_category_and_its_help` shows the help of a function with no arguments whole; the cutting is a check of the registered text, which only the fake reads |
| `pyinex.ini`'s `scripts` load at open, in order, relative to the add-in's folder; one that fails is reported and the others load | Left to the fake: the loader's C++ reads the list | `test_direct_functions.py::test_several_scripts_register_in_order_relative_and_absolute`, `test_direct_functions.py::test_a_script_that_fails_registers_nothing_and_the_others_still_do` | Partly: `test_registration.py::test_excel_lists_the_direct_functions_pyinex_ini_names` and `test_removal.py::test_the_pool_filled_exactly_then_one_function_more` load listed scripts; order and a failing script are left to the fake, the add-in's own reading, the same in any host |
| Python starts at open only when the list names a script | Left to the fake, as above | `test_direct_functions.py::test_a_listed_script_starts_python_at_open`, `test_direct_functions.py::test_an_empty_list_leaves_python_unstarted` | `test_python_choice.py::test_pyconsole_works_before_python_starts` and `test_python_choice.py::test_any_pyinfo_item_asked_first_starts_python` (no list: no Python until the first call); with a list, `test_registration.py::test_excel_lists_the_direct_functions_pyinex_ini_names` |
| A script loaded by `PyLoad` or the list runs outside any call: at its top level `caller`, `sheet` and `esc_pressed` raise `RuntimeError` | Left to the fake: the refusal is the core's C++ | `test_pyload.py::test_a_scripts_top_level_runs_outside_any_call`, `test_direct_functions.py::test_a_listed_scripts_top_level_runs_outside_any_call` | Left to the fake: the core's own check, the same in any host. (`test_pyinex_module.py::test_caller_and_sheet_refuse_outside_a_cell` checks another refusal, a call from no cell, so it is not this rule's test) |
| Registered only after the whole script has run: a decorated function then `ValueError`, `SystemExit` or `KeyboardInterrupt` registers nothing, first time and once fixed | `test_funcs.py::test_a_script_that_raises_after_a_marked_function_registers_nothing`, `test_funcs.py::test_a_script_that_raised_registers_once_it_is_fixed` | `test_pyload.py::test_a_script_that_fails_registers_nothing_until_it_is_fixed` | `test_pyload.py::test_a_script_that_raises_registers_nothing_until_it_is_fixed` |
| Only the functions the script marks itself, not ones it imports | `test_funcs.py::test_functions_the_script_imports_are_not_its_own` | Left to the unit test: decided in Python, the same in any host | As for the fake |
| Positions: positional parameters, then keyword-only, then `*args` to the last given; the plan's examples | `test_funcs_signatures.py::test_positions_are_positional_then_keyword_only_then_args`, `test_funcs_signatures.py::test_the_plans_examples` | `test_signatures.py::test_keyword_only_parameters`, `test_signatures.py::test_args_and_keyword_only_together_the_plans_examples` | `test_signatures.py::test_keyword_only_parameters`, `test_signatures.py::test_args_and_keyword_only_together_the_plans_examples` |
| 0 to 30 arguments, each position by its value, across the fourth and fifth | Partly: `test_funcs_signatures.py::test_thirty_positions_are_accepted` (a description for 30 parameters; passing each position its value is left to the other layers) | `test_signatures.py::test_each_count_from_0_to_30_passes_each_position_its_value`, `test_signatures.py::test_the_fourth_and_fifth_positions_each_have_their_own_value` | `test_signatures.py::test_each_count_from_0_to_30_passes_each_position_its_value`, `test_signatures.py::test_the_fourth_and_fifth_positions_each_have_their_own_value` |
| `*args` with none, one or many values; its positions numbered, each with its help; `**kwargs` takes no position | `test_funcs_signatures.py::test_args_take_none_one_or_many`, `test_funcs_signatures.py::test_args_positions_each_have_its_help`, `test_funcs_signatures.py::test_thirty_positions_are_accepted` | `test_signatures.py::test_args_with_none_one_and_many` | `test_signatures.py::test_args_with_none_one_and_many` |
| A position left out takes its default; with none, `None` before the last given, and Python's `TypeError` after it | `test_funcs_signatures.py::test_a_position_left_out_takes_its_default`, `test_funcs_signatures.py::test_one_left_out_with_no_default_before_the_last_given_is_none`, `test_funcs_signatures.py::test_one_left_out_with_no_default_after_the_last_given_is_pythons_error` | `test_signatures.py::test_defaults`, `test_signatures.py::test_a_missing_argument_in_the_middle_is_none`, `test_direct_functions.py::test_a_wrong_number_of_arguments_is_python_s_error` | `test_signatures.py::test_defaults`, `test_signatures.py::test_a_missing_argument_in_the_middle_is_none`, `test_direct_functions.py::test_a_failure_is_value_and_said_on_the_console` (`=Two(1)`: Python's `TypeError`) |
| More arguments than positions: `#VALUE!` from Excel, which calls nothing | No test: Excel refuses such a formula before any call, so the core never meets one (`test_funcs_signatures.py::test_more_arguments_than_positions_are_never_passed` checks only that a function of no positions registers a count of 0) | `test_signatures.py::test_more_arguments_than_registered_never_reach_the_function` (the fake refuses the call, as Excel does; Excel's own `#VALUE!` is real Excel's) | `test_signatures.py::test_more_arguments_than_registered_give_value_without_a_call` |
| Refused, naming the function: more than 30 positions, a signature that cannot be read or raises; the script's other functions register | `test_funcs_signatures.py::test_more_than_thirty_positions_are_refused_and_the_rest_register`, `test_funcs.py::test_an_unreadable_signature_is_refused`, `test_funcs.py::test_a_signature_that_raises_as_it_is_read_is_refused_and_the_rest_register`, `test_funcs.py::test_any_other_failure_describing_a_function_refuses_only_that_function` | Partly: `test_signatures.py::test_a_refused_signature_leaves_the_rest_registered_and_working`, `test_direct_functions.py::test_a_refused_signature_is_reported_and_the_rest_register` (more than 30 positions, a signature that cannot be read); one that raises as it is read is left to the unit tests, being decided in Python, the same in any host | Partly: `test_signatures.py::test_each_refusal_leaves_the_rest_registered_and_working` (the same cases as the fake); one that raises, as for the fake |
| A name imported only for type checkers (`if TYPE_CHECKING:`): from Python 3.14, unquoted, it does not stop registering, and its argument arrives unconverted; before 3.14 it must be quoted, as Python evaluates it as the script runs | `test_funcs.py::test_an_annotation_only_a_type_checker_imports_does_not_stop_registering` and `test_funcs_annotations.py::test_a_forward_reference_that_cannot_be_evaluated_leaves_the_value_and_says_so` (each from 3.14, skipped before); quoted, in every Python, `test_funcs_annotations.py::test_without_numpy_a_string_annotation_naming_it_leaves_the_value` and `test_funcs_annotations.py::test_an_annotation_that_cannot_be_evaluated_leaves_the_value_and_says_so`. Before 3.14 the unquoted name fails the script as any exception at its top level does (`test_funcs.py::test_a_script_that_raises_after_a_marked_function_registers_nothing`) | Left to the unit test: reading a signature is plain Python, the same in any host | As for the fake |
| A class, a `functools.partial` and a built-in function given to the decorator | `test_funcs_signatures.py::test_a_class_is_registered_with_its_constructors_parameters`, `test_funcs_signatures.py::test_only_the_class_marked_is_registered_not_its_subclasses_or_instances`, `test_funcs_signatures.py::test_a_partial_has_the_parameters_it_leaves`, `test_funcs_signatures.py::test_a_built_in_function_cannot_be_marked_and_the_script_loads` | `test_signatures.py::test_a_class`, `test_signatures.py::test_a_partial`, `test_signatures.py::test_a_refused_signature_leaves_the_rest_registered_and_working` (the built-in) | `test_signatures.py::test_a_class`, `test_signatures.py::test_a_partial`, `test_signatures.py::test_each_refusal_leaves_the_rest_registered_and_working` (the built-in) |
| A call follows PyCall's value rules, each case compared with PyCall | Partly: `test_funcs.py::test_a_collected_function_is_called_as_pycall_calls_it` (numbers and an exception through PyCall's calling code; the value cases, each compared with PyCall, are in the other layers) | `test_direct_functions.py::test_a_value_arrives_as_pycall_passes_it`, `test_direct_functions.py::test_a_value_goes_back_as_it_came`, `test_direct_functions.py::test_text_at_excels_limit_goes_back_whole`, `test_direct_functions.py::test_a_result_shows_as_pycalls_does`, `test_direct_functions.py::test_arguments_left_out` | `test_direct_functions.py::test_a_cell_arrives_as_pycall_passes_it`, `test_direct_functions.py::test_an_empty_cell_arrives_as_none`, `test_direct_functions.py::test_a_range_or_array_arrives_as_rows_of_rows`, `test_direct_functions.py::test_text_at_excels_limit_arrives_whole`, `test_direct_functions.py::test_an_argument_left_out_in_the_middle_is_none`, `test_direct_functions.py::test_a_return_value_shows_as_pycalls_does`, `test_direct_functions.py::test_text_at_excels_limit_is_returned_whole`, `test_direct_functions.py::test_a_list_fills_a_block_as_pycalls_does` |
| Every failure is `#VALUE!` with the reason, and any traceback, on the console; a conversion message starts with the function's name | Left to the fake: the result's conversion is the core's C++ | `test_direct_functions.py::test_an_exception_is_value_with_its_traceback`, `test_direct_functions.py::test_a_result_that_cannot_be_returned_is_value_naming_the_function`, `test_direct_functions.py::test_a_value_pyinex_does_not_recognize_is_reported_naming_the_function` | `test_direct_functions.py::test_a_failure_is_value_and_said_on_the_console`, `test_direct_functions.py::test_an_exception_shows_its_traceback` |
| The call context: `caller`, `sheet`, `esc_pressed`, a script's thread refused, the GIL released between calls | Left to the fake: it needs a call from Excel | `test_direct_functions.py::test_caller_is_the_calling_cell`, `test_direct_functions.py::test_sheet_is_the_calling_sheet`, `test_direct_functions.py::test_esc_pressed_kept_and_cleared`, `test_direct_functions.py::test_asking_excel_from_a_scripts_thread_is_refused`, `test_direct_functions.py::test_the_gil_is_released_between_calls` | `test_direct_functions.py::test_sheet_and_caller_are_the_calling_ones_not_the_active_ones`, `test_direct_functions.py::test_caller_of_a_legacy_array_formula_is_its_range`, `test_direct_functions.py::test_caller_of_a_spilling_formula_is_its_cell`, `test_direct_functions.py::test_caller_and_sheet_refuse_outside_a_cell`, `test_direct_functions.py::test_esc_pressed_kept_and_cleared`, `test_direct_functions.py::test_asking_excel_from_a_scripts_thread_is_refused`, `test_direct_functions.py::test_the_gil_is_released_between_calls` |
| Blank in the Function Arguments dialog, running no Python | Left to the fake: it needs the dialog's window | `test_direct_functions.py::test_in_the_function_wizard_a_direct_function_is_blank_and_runs_nothing` | `test_direct_functions.py::test_the_function_arguments_dialog_shows_the_help_and_runs_no_python` (the function's value shown as `""` and no formula result, and no file written) |
| Not thread-safe (no `$`) and not volatile (no `!`): recalculated when a cell it refers to changes, and only then | Left to the fake: the type string is the loader's | `test_registration.py::test_no_direct_function_is_thread_safe`, `test_registration.py::test_registers_the_direct_functions_pyinex_ini_lists` (several type strings compared whole) | `test_registration.py::test_excel_lists_the_direct_functions_pyinex_ini_names` (type strings compared whole); when it recalculates, `test_range.py::test_a_function_recalculates_when_a_cell_it_refers_to_changes_and_only_then` (a function taking a reference; for one taking values, Excel's own rule, with no test of its own) |
| A full pool reports each function that did not fit, and the rest keep working | Left to the fake: the slot table is the loader's | `test_removal.py::test_the_pool_filled_exactly`, `test_removal.py::test_one_function_more_than_the_pool_is_reported_and_the_rest_work` | `test_removal.py::test_the_pool_filled_exactly_then_one_function_more` |
| A slot is reused at once, after a removal whose every step succeeded, with no old name reaching the new function | Left to the fake, as above | `test_removal.py::test_many_cycles_of_registering_and_removing_reuse_one_slot` | Partly: `test_removal.py::test_many_cycles_of_registering_and_removing` (the number of registrations stays flat across cycles, not which slot; reusing the same slot is checked in the fake) |
| A description whose count is outside 0 to 30, or that lacks its count's help, is refused before Excel is told | Python never hands one over: `test_funcs_signatures.py::test_more_than_thirty_positions_are_refused_and_the_rest_register` refuses it first | `test_pyload.py::test_a_malformed_description_is_refused_before_excel_is_told` (a count of 31, made by replacing Pyinex's describer; the script's other function registers) | Left to the fake: the check is the loader's, and Python never hands one over |

**The queue** ([architecture.md](./architecture.md#the-queue))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| A command of Pyinex's own, named for this copy | Left to the fake: it is the loader's | `test_queue.py::test_the_command_is_registered_as_a_command`, `test_lifetime.py::test_unloaded_while_another_copy_is_loaded_the_work_waits_and_runs_no_other_queue` | `test_registration.py::test_excel_lists_the_documented_functions` (the command); another copy is left to the fake, which loads two copies on request |
| Registering comes after the calculation, through `Application.Run` from a timer | Left to the fake, as above | `test_pyload.py::test_a_scripts_functions_register_after_the_calculation`, `test_queue.py::test_the_timer_runs_the_command_through_application_run` | `test_pyload.py::test_registration_queued_during_a_full_recalculation` |
| Work queued while it runs is done by the same run; it never runs twice at once | Left to the fake, as above | `test_queue.py::test_work_queued_while_the_queue_runs_is_done_by_the_same_run`, `test_queue.py::test_the_queue_is_never_run_twice_at_once` | Left to the fake: real Excel cannot be made to start the command inside itself on demand |
| It waits out a cell edit and a modal dialog, however long, reporting each new cause once | Left to the fake, as above | `test_queue.py::test_thousands_of_refusals_then_the_work_is_done` | `test_pyload.py::test_registration_waits_while_excel_is_busy` |
| It reaches the right Excel: no workbook window, Protected View, several Excels | Left to the fake, as above | `test_queue.py::test_with_no_workbook_window_the_work_waits_for_one`, `test_queue.py::test_a_protected_view_window_is_passed_over` | `test_pyload.py::test_with_no_workbook_open`, `test_pyload.py::test_with_only_a_protected_view_window_open`, `test_pyload.py::test_with_a_protected_view_window_in_front_of_a_workbook`, `test_pyload.py::test_two_excels_each_register_their_own` |
| Never inside a calculation, whether from its timer or its command | Left to the fake, as above | `test_range.py::test_the_queue_waits_until_a_stopped_call_returns`, `test_range.py::test_the_queue_s_command_run_inside_a_stopped_call_waits` | Left to the fake: it needs the timer's message inside a call Excel stopped, at a moment real Excel cannot be made to choose |
| A refused registration is reported, its name freed, and tried again by its `PyLoad` cell | `test_funcs.py::test_a_function_that_did_not_register_frees_its_name` | `test_queue.py::test_a_refused_registration_is_reported_and_tried_again_by_pyload` | Fake only: real Excel cannot be made to refuse a registration |
| Nothing lost to an error: a function leaves the queue only once Excel has registered it and its slot holds it, or its failure is reported; its slot is filled before Excel is told, and an error before that leaves it queued | Left to the fake, as above | `test_queue.py::test_thousands_of_refusals_then_the_work_is_done` (kept through Excel's refusals to run the command), `test_queue.py::test_a_refused_registration_is_reported_and_tried_again_by_pyload` (a failure reported), `test_queue.py::test_memory_running_out_before_excel_is_told_leaves_the_function_queued` (memory running out after Excel is asked about the name: the function stays queued and registers once), `test_queue.py::test_memory_running_out_once_excel_has_registered_a_function_loses_nothing` (just after Excel registers it: registered once, in its slot), `test_queue.py::test_pyrepair_run_while_excel_registers_a_function_then_memory_running_out_loses_nothing` (the same, with `PyRepair` run from within the registration), `test_queue.py::test_pyrepair_run_while_a_listed_function_registers_waits_for_the_add_in_to_open` (`PyRepair` run from within a `pyinex.ini` function's registration as the add-in opens: its work waits for the queue's run after, the function registered or refused) | Partly: `test_pyload.py::test_registration_waits_while_excel_is_busy` (kept through Excel's refusals); the rest as for the fake |
| A window or timer that could not be made is made when a `PyLoad` cell calculates again | Left to the fake, as above | `test_queue.py::test_a_queue_window_not_made_is_made_when_pyload_calculates_again` (the window), `test_queue.py::test_a_timer_not_set_is_set_when_pyload_calculates_again` (the timer, which then runs the command itself) | Fake only: Windows cannot be made to refuse the window |
| Opening and closing count as the queue's run; a run started within one, or within a registration, takes none of their work, and an error in it never ends that mark | Left to the fake: it needs Excel to run a command within a call into Excel | `test_queue.py::test_pyrepair_failing_inside_a_registration_leaves_the_queue_its_own`, `test_lifetime.py::test_pyrepair_run_while_the_add_in_opens_again_loses_no_workbook_function`, `::test_pyrepair_run_while_the_add_in_closes_leaves_it_closed` | Fake only: Excel cannot be made to do it on demand; nobody has seen it there |
| A close within a registration loses no function and registers none twice, and is safe if Excel then refuses the registration | Left to the fake, as above | `test_queue.py::test_the_add_in_closed_while_excel_registers_a_function_loses_none`, `::test_the_add_in_closed_while_excel_registers_a_new_name_is_safe` | Fake only: Excel cannot be made to do it on demand; nobody has seen it there |
| The queue's command refused at opening is reported, and later work dropped with a note rather than tried for ever; a refused `PyRepair` or worksheet function is reported | Left to the fake: it is the loader's | `test_queue.py::test_a_queue_command_excel_refuses_at_opening_is_reported_and_its_work_dropped`, `test_repair.py::test_pyrepair_excel_refuses_at_opening_is_reported`, `test_lifetime.py::test_a_worksheet_function_excel_refuses_at_opening_is_reported` | Fake only: Excel cannot be made to do it on demand |
| A run that stopped is reported as such, not as Excel not found | Left to the fake, as above | `test_queue.py::test_a_run_of_the_queue_that_stopped_is_reported_as_such` | Fake only: Excel cannot be made to do it on demand |
| A `PyLoad` cell starts again a timer that failed whatever waits, a repair alone included; a note before Python starts reaches the console | Left to the fake, as above | `test_queue.py::test_pyload_starts_again_a_timer_that_failed_with_only_a_repair_waiting`, `test_console.py::test_a_timer_that_cannot_start_before_python_says_so_on_the_console` | Fake only: Excel cannot be made to do it on demand |
| Rows Python gave that never reached the queue are offered again, a removal waiting for the loader's next call | `test_funcs.py::test_rows_that_did_not_reach_the_loader_are_offered_again`, `::test_changes_that_did_not_reach_the_loader_wait_for_its_next_call` | `test_removal.py::test_memory_running_out_*` | Fake only: Excel cannot be made to do it on demand |

**PyLoad** ([architecture.md](./architecture.md#pyload))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| A relative path is relative to the workbook's folder; in a workbook never saved it is refused; an absolute path works | `test_funcs.py::test_a_relative_path_is_relative_to_the_folder_given`, `test_funcs.py::test_a_relative_path_in_a_workbook_never_saved_is_refused` | `test_pyload.py::test_a_relative_path_is_relative_to_the_workbooks_folder`, `test_pyload.py::test_a_relative_path_in_a_workbook_never_saved_is_refused`, `test_pyload.py::test_a_scripts_functions_register_after_the_calculation` | Partly: `test_pyload.py::test_a_relative_path_is_relative_to_the_workbooks_folder`, `test_pyload.py::test_an_absolute_path` (saved workbooks). A never-saved workbook is left to the fake for `PyLoad`: the folder comes from the same `caller_folder` and the same check in `call.py` as PyCall's, which `test_paths.py::test_an_unsaved_workbook_refuses_a_relative_path` shows real Excel reaching |
| A relative path with no folder says why: no calling cell, or a workbook stored on the web | Left to the fake: it needs `xlfCaller` and `CELL` to answer so | `test_calls.py::test_a_relative_path_with_no_calling_cell_says_so`, `::test_a_workbook_stored_on_the_web_says_a_relative_path_needs_a_folder` | Not checked in real Excel: the VM has no OneDrive workbook |
| A path that is not text is refused | Left to the fake: the loader's check | `test_pyload.py::test_a_path_that_is_not_text_is_refused` | Left to the fake: the same check in any host |
| Not volatile: the cell calculates when entered, when its argument changes, on a full recalculation, and when its workbook opens | Left to the fake: the type string is the loader's | `test_registration.py::test_registers_the_documented_functions` (`PyLoad`'s type string `QQ`, compared whole) | `test_registration.py::test_excel_lists_the_documented_functions` (the same); `test_pyload.py::test_registration_queued_during_a_full_recalculation` (a full recalculation); `test_workbook_open.py::test_a_workbook_opened_in_a_new_excel_loads_its_scripts` (the workbook opening); `test_pyload.py::test_pyload_calculates_again_when_its_argument_changes` (a changed argument) |
| The result counts the script's functions, as text: those accepted as the script is read, not one refused then, and whether or not Excel then registers them | Left to the fake: the loader makes the text | `test_pyload.py::test_the_status_counts_the_scripts_functions`; a function refused as the script is read is not counted: `test_names.py::test_two_names_that_differ_only_in_case` (three marked, `2 functions`), `test_names.py::test_a_name_differing_only_in_case_in_another_script` (`0 functions`); one Excel then refuses is: `test_names.py::test_a_name_of_one_of_excels_functions_is_refused` (`2 functions`) | `test_pyload.py::test_an_absolute_path`, and `direct_functions.xlsx` in `test_examples.py::test_example_workbook_calculates_as_its_sheets_say`; the exclusions are left to the fake: the same count in any host |
| A missing script, a syntax error or an exception at import gives `#VALUE!`, reported, and the fixed file then loads | `test_funcs.py::test_a_missing_script_registers_nothing`, `test_funcs.py::test_a_script_that_raises_after_a_marked_function_registers_nothing`, `test_funcs.py::test_a_script_that_raised_registers_once_it_is_fixed` | `test_pyload.py::test_a_script_that_fails_registers_nothing_until_it_is_fixed` (each failure, then the fixed file) | Partly: `test_pyload.py::test_a_script_that_fails_to_load_is_reported` (each failure reported); `test_pyload.py::test_a_script_that_raises_registers_nothing_until_it_is_fixed` (the fixed file, after an exception at import only). The fixed file after a missing script or a syntax error is left to the fake: the same reload in any host |
| Calculating again registers only what is new; a script loaded twice, and from two workbooks, registers once and works in both | `test_funcs.py::test_loading_again_registers_only_what_is_new` | `test_pyload.py::test_a_script_loaded_twice_registers_once`, `test_pyload.py::test_a_script_loaded_from_two_workbooks_registers_once_for_both` | `test_pyload.py::test_a_script_loaded_twice_and_from_two_workbooks_registers_once` |
| One script through `PyLoad` and PyCall is one module | Left to the fake: it needs both functions | `test_pyload.py::test_one_script_through_pyload_and_pycall_is_one_module` | `test_pyload.py::test_one_script_through_pyload_and_pycall_is_one_module` |
| Blank in the Function Arguments dialog, loading nothing | Left to the fake: it needs the dialog's window | `test_pyload.py::test_in_the_function_wizard_pyload_is_blank_and_loads_nothing` | Left to the fake: the same check of the dialog (`in_function_wizard`, `src/excel/excel.cpp`) a direct function makes, which `test_direct_functions.py::test_the_function_arguments_dialog_shows_the_help_and_runs_no_python` shows real Excel reaching |

**Reloading a script** ([architecture.md](./architecture.md#reloading-a-script))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| A changed body runs at the next call, with no call into Excel | `test_funcs_reload.py::test_a_changed_body_needs_no_change_in_excel` | `test_reload_functions.py::test_a_changed_body_runs_at_the_next_call_with_no_call_into_excel` (the fake counts the calls into Excel) | Partly: `test_reload_functions.py::test_a_changed_body_runs_at_the_next_recalculation` (the new body runs, and Excel's registration table is unchanged); calls into Excel cannot be counted in real Excel, so that is the fake's |
| A change to what Excel is told (the arguments' number, names or types, the name's spelling, the help, the category) is registered again; meanwhile a call gives `#VALUE!` and reads only the registered arguments. Any other change, such as positional to keyword-only with the same arguments, takes effect at the next call | `test_funcs_reload.py::test_a_changed_signature_is_registered_again` (each kind of change but an argument's type, which `test_range.py::test_a_value_becoming_a_reference_registers_again` in the fake and `test_range.py::test_a_value_becoming_a_reference_and_back_registers_again` in real Excel check) | `test_reload_functions.py::test_a_changed_signature_is_registered_again_and_meanwhile_gives_value`, `test_signatures.py::test_positional_to_args_is_registered_again_and_meanwhile_gives_value`; `test_signatures.py::test_positional_to_keyword_only_with_the_same_names_needs_no_registration` | Partly: `test_reload_functions.py::test_a_changed_signature_is_registered_again`, `test_signatures.py::test_positional_to_args_is_registered_again`; a change Excel is not told of, as for a changed body (`test_reload_functions.py::test_a_changed_body_runs_at_the_next_recalculation`, `test_annotations.py::test_a_changed_annotation_converts_at_the_next_recalculation_with_no_registration`) |
| A removed function is removed and its name cleared; a renamed one both; an added one registers | `test_funcs_reload.py::test_a_removed_function_is_unregistered_and_its_name_freed`, `test_funcs_reload.py::test_a_renamed_function_is_removed_and_registered`, `test_funcs_reload.py::test_an_added_function_registers` | `test_reload_functions.py::test_a_removed_function_is_unregistered_and_its_name_cleared`, `test_reload_functions.py::test_a_renamed_function_is_removed_and_registered`, `test_reload_functions.py::test_an_added_function_registers` | `test_reload_functions.py::test_a_removed_function_is_unregistered_and_its_name_cleared`, `test_reload_functions.py::test_a_renamed_function_is_removed_and_registered`, `test_reload_functions.py::test_an_added_function_registers` |
| A script that now fails changes nothing, and recovers through a call or its `PyLoad` cell once fixed | `test_funcs_reload.py::test_a_script_that_now_fails_changes_nothing_until_it_is_fixed`, `test_funcs_reload.py::test_a_script_that_now_fails_recovers_through_pyload` | `test_reload_functions.py::test_a_script_that_now_fails_registers_nothing_and_recovers_once_fixed` | `test_reload_functions.py::test_a_script_that_now_fails_registers_nothing_and_recovers_once_fixed` |
| With `PyReloadOnChange` off, nothing changes | `test_funcs_reload.py::test_with_reloading_off_nothing_changes` | `test_reload_functions.py::test_with_reloading_off_nothing_changes` | `test_reload_functions.py::test_with_reloading_off_nothing_changes` |
| Only scripts `PyLoad` or the list loaded are followed | `test_funcs_reload.py::test_a_script_only_pycall_uses_is_not_followed` | Left to the unit test: decided in Python, the same in any host | As for the fake |
| Several edits while work is queued go in the order of the reloads, the last winning | `test_funcs_reload.py::test_several_edits_before_the_loader_asks`, `test_funcs_reload.py::test_pyload_calculated_again_gives_the_removals_first`, `test_funcs_reload.py::test_a_load_during_a_call_s_reload_follows_that_reload_s_changes` | `test_reload_functions.py::test_several_edits_while_work_is_queued`, `test_reload_functions.py::test_a_pyload_during_a_call_that_reloaded_queues_after_that_reload` | `test_reload_functions.py::test_several_edits_in_one_calculation_while_work_is_queued` |
| A registration again that Excel refuses is reported and tried again by the `PyLoad` cell | `test_funcs_reload.py::test_a_failed_registration_again_frees_the_name_for_the_next_pyload` | `test_reload_functions.py::test_a_registration_again_that_fails_is_reported_and_tried_again_by_pyload` | Fake only: real Excel cannot be made to refuse a registration |
| A changed annotation or a reused object is read again at the next call | `test_funcs_reload.py::test_a_reused_function_is_converted_by_its_annotation_after_the_reload`, `test_funcs_reload.py::test_a_reused_partial_is_placed_by_its_signature_after_the_reload` | Partly: `test_annotations.py::test_a_changed_annotation_converts_at_the_next_call_with_no_registration` (a changed annotation); a reused object is left to the unit tests: which object the reloaded script marks is decided in Python, the same in any host | Partly: `test_annotations.py::test_a_changed_annotation_converts_at_the_next_recalculation_with_no_registration`; a reused object, as for the fake |
| A function whose new signature is refused says why, not that it was removed | `test_funcs_reload.py::test_a_function_whose_new_signature_is_refused_says_why_not_that_it_was_removed` | Left to the unit test | Left to the unit test |
| With reloading off, reopening the add-in keeps a listed script as it loaded | Left to the fake: it opens the add-in again | `test_lifetime.py::test_reopening_with_reloading_off_keeps_a_listed_script_as_it_loaded` | Not checked in real Excel |
| A `PyLoad` within a call runs its script outside any call | Left to the fake: it calls `PyLoad` within a call | `test_reload_functions.py::test_a_pyload_within_a_call_runs_its_script_outside_any_call` | Not checked in real Excel |

**Removing a function** ([architecture.md](./architecture.md#removing-a-function))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| What removal leaves: the registration table, Insert Function, existing cells and new formulas | Left to the fake: it is Excel's | Partly: `test_removal.py::test_what_removal_leaves` (the registration table, and a call by the name giving `#VALUE!`); Insert Function, and a cell's kept value against a formula typed afresh, are Excel's own, so real Excel's | `test_removal.py::test_what_removal_leaves` (each) |
| A step that fails is reported and keeps the slot out of use, with the old name reaching no other function | Left to the fake, as above | `test_removal.py::test_a_removal_that_fails_at_any_step_keeps_the_slot_out_of_use` | Fake only: real Excel cannot be made to fail a step on demand; its answers to a failed step were checked before this was built |
| Memory running out at any point after Excel answers a removal neither frees a slot Excel still sends a name to nor keeps the name from its function | Left to the fake: it needs memory to run out on demand | `test_removal.py::test_memory_running_out_after_a_failed_removal_keeps_the_slot_out_of_use`, `::test_memory_running_out_during_a_removal_never_keeps_the_name_from_its_function` | Fake only: Excel cannot be made to do it on demand |
| A name whose removal failed twice registers again | Left to the fake, as above | `test_removal.py::test_a_name_whose_removal_failed_twice_registers_again` | Fake only: Excel cannot be made to do it on demand |
| Refusals and removals leave no record behind in Python | `test_funcs.py::test_refusals_and_removals_leave_nothing_behind` | Left to the unit test: Python's records | Left to the unit test |
| Closing the add-in removes every function; reopened with another list, an old name never reaches the new function | Left to the fake, as above | `test_direct_functions.py::test_reopened_with_another_list_an_old_name_never_reaches_the_new_function` | `test_direct_functions.py::test_reopened_with_another_list_an_old_name_never_reaches_the_new_function` |

**Names** ([architecture.md](./architecture.md#names))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| Compared without regard to case | `test_funcs_names.py::test_two_names_that_differ_only_in_case_in_one_script`, `test_funcs_names.py::test_two_names_that_differ_only_in_case_in_two_scripts` | `test_names.py::test_two_names_that_differ_only_in_case`, `test_names.py::test_a_name_differing_only_in_case_in_another_script` | `test_names.py::test_two_names_that_differ_only_in_case` |
| The first registration wins; a second script is refused, naming both | `test_funcs.py::test_a_name_another_script_registered_is_refused_naming_both`, `test_funcs.py::test_one_function_under_two_names_registers_once` | `test_pyload.py::test_a_name_another_script_registered_is_refused_naming_both` | `test_names.py::test_one_name_in_two_scripts_the_first_wins` |
| Pyinex's own names are refused | `test_funcs_names.py::test_pyinex_s_own_names_are_refused` | `test_names.py::test_pyinex_s_own_names_are_refused` | `test_names.py::test_pyinex_s_own_names_are_refused` |
| Names Excel cannot take are refused, saying which rule; a Unicode name and others Excel takes register | `test_funcs_names.py::test_names_excel_cannot_take_are_refused`, `test_funcs_names.py::test_names_excel_takes_are_registered` | `test_names.py::test_names_excel_cannot_take_are_refused`, `test_names.py::test_unicode_and_other_names_excel_takes` | `test_names.py::test_names_excel_cannot_take_are_refused`, `test_names.py::test_names_excel_takes_are_registered` |
| Names Excel already knows, its own (`SUM`) or another add-in's, are refused | Left to the fake: it asks Excel (`xlfEvaluate`) | `test_names.py::test_a_name_of_one_of_excels_functions_is_refused`, `test_names.py::test_a_name_another_add_in_uses_is_refused` | `test_names.py::test_a_name_of_one_of_excels_functions_is_refused`, `test_names.py::test_a_name_another_add_in_registered_is_refused` |
| A name a workbook defines is not seen when Pyinex asks Excel about a name; a question that fails refuses the name | Left to the fake, as above | `test_names.py::test_a_name_excel_cannot_answer_about_is_refused` (the question failed; the name asked again by the next `PyLoad`). The fake has no workbook names: left to real Excel | `test_names.py::test_a_name_a_workbook_defines_is_free` (registered, and called from another workbook). A failed question is left to the fake: real Excel cannot be made to fail one |
| A question Excel does not answer is reported as such, not as a name taken | Left to the fake: it needs `xlfEvaluate` to fail | `test_names.py::test_a_name_excel_cannot_answer_about_is_refused` | Fake only: Excel cannot be made to do it on demand |
| Limits count a character beyond U+FFFF as two, as Excel does, and never cut one | `test_funcs.py::test_excel_s_limits_count_characters_beyond_u_ffff_as_two` | Left to the unit test: the text Pyinex sends | Not checked in real Excel |
| A name only typed in formulas is free; a name Pyinex removed is free to it, but not once another add-in took it | Left to the fake, as above | `test_names.py::test_a_name_a_reload_removed_is_free_for_another_script`, `test_names.py::test_a_name_pyinex_no_longer_holds_is_asked_about_again`, `test_removal.py::test_a_removal_that_fails_at_any_step_keeps_the_slot_out_of_use` | `test_names.py::test_a_name_typed_in_formulas_before_it_registers_is_free`, `test_names.py::test_a_name_pyinex_removed_and_another_add_in_took_is_refused` |

**When a workbook opens** ([architecture.md](./architecture.md#when-a-workbook-opens))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| A workbook opened later has its `PyLoad` cells calculated before anything changes, even in a new Excel | Left to the fake: it follows Excel's events | `test_repair.py::test_a_workbook_opened_later_is_calculated_when_excel_reports_it` | `test_workbook_open.py::test_a_workbook_opened_in_a_new_excel_loads_its_scripts` |
| Workbooks open as the add-in loads are handled the same way | Left to the fake, as above | `test_repair.py::test_the_pyload_cells_of_workbooks_open_as_the_addin_opens_are_calculated` | `test_workbook_open.py::test_a_workbook_open_before_the_addin_loads_loads_its_scripts` |
| A workbook made from a template (`NewWorkbook`) is followed as one opened is | Left to the fake, as above | `test_repair.py::test_a_workbook_opened_later_is_calculated_when_excel_reports_it[made from a template]` | `test_workbook_open.py::test_a_workbook_made_from_a_template_loads_its_scripts` |
| Every sheet, hidden and very hidden ones included, and any form of argument: a literal path, a path in a cell; a volatile function in the workbook | Left to the fake, as above | Partly: `test_repair.py::test_the_pyload_cells_of_workbooks_open_as_the_addin_opens_are_calculated` (every sheet; the fake's sheets are never hidden) | `test_workbook_open.py::test_a_workbook_opened_in_a_new_excel_loads_its_scripts` (each case, a hidden and a very hidden sheet among them) |
| The workbook counts as changed | Left to real Excel: whether a workbook is changed is Excel's | Left to real Excel, as for the unit tests | `test_workbook_open.py::test_a_workbook_opened_in_a_new_excel_loads_its_scripts` |
| With nothing to register, only workbooks to follow, and no workbook open, the queue waits without a word | Left to the fake, as above | `test_repair.py::test_with_no_workbook_open_the_queue_waits_without_a_word` | `test_workbook_open.py::test_with_no_workbook_open_the_queue_waits_without_a_word_then_follows` (no console window made, then a workbook opened is loaded). (`test_pyload.py::test_with_no_workbook_open` registers a function with no workbook open, which is another case) |
| A workbook opened with Excel's events off is missed until Ctrl+Alt+F9 or `PyRepair` | Left to real Excel: the limitation is Excel's own | Left to real Excel, as for the unit tests | `test_workbook_open.py::test_a_workbook_opened_with_events_off_loads_by_the_workaround` |
| Closing disconnects from Excel's events; reopening, or a canceled quit, connects again | Left to the fake, as above | `test_repair.py::test_closing_disconnects_from_excels_events_and_reopening_connects_again`, `test_repair.py::test_a_cancelled_quit_with_nothing_registered_follows_workbooks_again` | `test_lifetime.py::test_after_a_cancelled_quit_with_nothing_registered_a_workbook_opened_loads` |

**The `#NAME?` repair** ([architecture.md](./architecture.md#the-name-repair))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| Formulas that call a newly registered name are entered again: on several sheets, with other functions, spilling, inside `IFERROR`, in an untaken `IF` branch; then Excel is asked to calculate | Left to the fake: it searches Excel's formulas | Partly: `test_repair.py::test_only_formulas_that_call_the_name_are_entered_again` (with other functions in the formula, and the calculation asked for once), `test_repair.py::test_a_formula_that_cannot_be_entered_again_is_reported_and_pyrepair_enters_it` (two sheets), `test_repair.py::test_the_pyload_cells_of_workbooks_open_as_the_addin_opens_are_calculated` (a call inside `IFERROR` chosen). Spilling, an untaken branch, and each cell's value after, are real Excel's: the fake does not calculate formulas | `test_repair.py::test_cells_typed_before_registration_are_repaired` (each case, by the values after) |
| Only formulas that call a name: not one holding it in text, a quoted sheet name or a longer name | Left to the fake, as above | `test_repair.py::test_only_formulas_that_call_the_name_are_entered_again` (each case) | Partly: `test_repair.py::test_cells_typed_before_registration_are_repaired` (the name in a text, a longer name, and text that is not a formula); a quoted sheet name is left to the fake: the choice is the add-in's reading of the formula's text, the same in any host |
| Every cell is found before any is entered | Left to the fake, as above | `test_repair.py::test_a_change_handler_that_searches_leaves_the_repair_its_own_cells` | Left to the fake: a workbook's change handler needs VBA, which the tests do not use |
| A change handler that runs `PyRepair` lets the repair end | Each formula entered once; the handler's runs part of the repair | `test_repair.py::test_a_change_handler_that_runs_pyrepair_lets_the_repair_end` | `test_repair.py::test_a_change_handler_that_runs_pyrepair_lets_the_repair_end`, the handler this process's, connected through COM |
| A legacy array formula is entered again whole, once; one longer than 255 characters is reported | Left to the fake, as above | `test_repair.py::test_a_legacy_array_formula_is_entered_again_whole_once`, `test_repair.py::test_a_legacy_array_formula_too_long_to_enter_again_is_reported` | `test_repair.py::test_cells_typed_before_registration_are_repaired` (a legacy array formula); the longer one cannot be entered through COM to make one |
| Only new names: a name registered again is left alone | Left to the fake, as above | `test_repair.py::test_a_name_registered_again_is_left_alone` | Partly: `test_reload_functions.py::test_a_changed_signature_is_registered_again` shows such a cell coming back on an ordinary recalculation |
| A formula that cannot be entered again, or read, is reported, and `PyRepair` enters it later | Left to the fake, as above | `test_repair.py::test_a_formula_that_cannot_be_entered_again_is_reported_and_pyrepair_enters_it` (both halves), `test_repair.py::test_a_formula_that_cannot_be_read_is_reported_and_pyrepair_enters_it` (both halves) | `test_repair.py::test_a_protected_sheet_is_reported_and_pyrepair_repairs_it` (one that cannot be entered again); an unreadable formula is left to the fake: real Excel cannot be made to refuse reading one |
| `PyRepair` inside a worksheet function does nothing and says so | Left to the fake, as above | `test_repair.py::test_pyrepair_refused_inside_a_worksheet_call_leaves_nothing_queued` (the console's message, nothing entered, nothing queued) | Left to the fake: the same check in any host |
| `PyRepair` runs from the Macro dialog (Alt+F8), which lists no direct function and not `PyRepair`; a direct function's name typed there calls it with no arguments | Left to real Excel: the dialog is Excel's | Left to real Excel, as for the unit tests | `test_macro_dialog.py::test_the_macro_dialog_lists_no_direct_function_and_not_pyrepair` (against a macro sheet's command, which it lists), `test_macro_dialog.py::test_pyrepair_typed_in_the_macro_dialog_runs`, `test_macro_dialog.py::test_a_direct_function_typed_in_the_macro_dialog_is_called_with_no_arguments` |

**pyinex.Range** ([architecture.md](./architecture.md#pyinexrange), [conversion.md](./conversion.md#references-pyinexrange))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| A parameter annotated `pyinex.Range` (or `pyinex.Range \| None`) registers as a reference, `U` | `test_funcs_range.py::test_a_reference_parameter_registers_as_type_u` (`pyinex.Range`, quoted, `Optional[pyinex.Range]` and `pyinex.Range \| None`) | Partly: `test_registration.py::test_a_pyinex_range_argument_is_registered_as_a_reference` (`pyinex.Range` only); the optional spelling is left to the unit test, the type string being decided in Python | Partly: `test_range.py::test_a_value_becoming_a_reference_and_back_registers_again` (the type strings Excel lists, `pyinex.Range` only); the optional spelling, as for the fake |
| Every kind of reference arrives with its workbook, sheet, address in both styles and size | `test_funcs_range.py::test_its_workbook_sheet_and_size`, `test_funcs_range.py::test_its_address_as_caller_gives_one`, `test_funcs_range.py::test_a_range_arrives_as_it_is` | `test_range.py::test_each_reference_gives_its_workbook_sheet_address_and_size`, `test_range.py::test_a_reference_with_no_sheet_is_on_the_sheet_calculating` | `test_range.py::test_each_reference_gives_its_workbook_sheet_address_and_size` |
| `values()` converts as an argument's values; `empty()` tells an empty cell from empty text | `test_funcs_range.py::test_values_are_read_from_excel_with_its_sheet`, `test_funcs_range.py::test_empty_tells_an_empty_cell_from_empty_text` | `test_range.py::test_values_and_empty_cells`, `test_range.py::test_values_and_a_reference_in_one_call` | `test_range.py::test_values_and_empty_cells` |
| A whole column or row reads only its used part | `test_funcs_range.py::test_a_whole_column_reads_only_its_used_part`, `test_funcs_range.py::test_a_whole_row_and_the_whole_sheet_read_only_their_used_part`, `test_funcs_range.py::test_the_used_part_ends_at_the_last_row_holding_anything`, `test_funcs_range.py::test_an_empty_whole_column_reads_as_its_first_cell` | `test_range.py::test_a_whole_column_reads_only_its_used_part` | `test_range.py::test_values_and_empty_cells` (whole column, whole row, empty column) |
| Several whole columns or rows holding nothing read as their first cell | `test_funcs_range.py::test_empty_whole_columns_or_rows_read_as_their_first_cell` | Left to the unit test | Left to the unit test |
| A cell holding a value Pyinex does not know reads as `#VALUE!`, said on the console | Left to the fake: it needs Excel to give such a value | `test_range.py::test_a_value_pyinex_does_not_know_is_read_as_value_and_reported` | Fake only: Excel cannot be made to do it on demand |
| Several areas: `areas` gives each; `rows`, `columns`, `values()` and `empty()` raise `ValueError` | `test_funcs_range.py::test_its_areas`, `test_funcs_range.py::test_several_areas_have_no_single_size_or_values` (each of the four) | Partly: `test_range.py::test_several_areas` (`areas`, and `values()` and `rows` refused) | Partly: `test_range.py::test_several_areas` (`areas`, and `values()` refused); the others are left to the unit test: one check in Python, the same in any host |
| A value, an error (`#REF!` for a deleted sheet) or a workbook not open is refused, naming the function and argument | `test_funcs_range.py::test_a_value_that_is_not_a_reference_is_refused`, `test_funcs_range.py::test_an_error_is_refused` | `test_range.py::test_a_value_that_is_not_a_reference_is_refused` | `test_range.py::test_a_value_that_is_not_a_reference_is_refused`, `test_range.py::test_a_deleted_sheet_is_refused` |
| Not an element type: `list[list[pyinex.Range]]` is not converted | `test_funcs_range.py::test_a_range_is_not_an_element_type` | Left to the unit test: decided in Python | As for the fake |
| A function returning a `pyinex.Range` gives `#VALUE!`, as for any object Excel cannot take | Left to the fake: the result's conversion is the core's C++ | `test_range.py::test_a_range_returned_is_value_as_any_object_excel_cannot_take` | `test_range.py::test_a_range_returned_is_value_as_any_object_excel_cannot_take` |
| Recalculated as a function taking values is: when a cell it refers to changes, and only then | Left to real Excel: recalculation is Excel's | Left to real Excel, as for the unit tests | `test_range.py::test_a_function_recalculates_when_a_cell_it_refers_to_changes_and_only_then` |
| After `xlretUncalced` or `xlretAbort`, nothing but `xlFree` reaches Excel in that call, from caught exceptions and cleanup too, and the call gives `#VALUE!`, saying nothing, PyCall's included | `test_funcs_range.py::test_a_call_excel_stopped_fails_saying_nothing` | `test_range.py::test_after_a_stop_nothing_asks_excel_and_the_call_gives_value_silently`, `test_range.py::test_a_stop_in_a_finally_block_and_raised_says_nothing`, `test_range.py::test_a_stop_while_finding_a_whole_column_s_used_part`, `test_range.py::test_a_stop_while_the_result_is_freed_gives_value_silently`, `test_range.py::test_a_stop_while_the_result_is_freed_in_a_pycall_gives_value_silently` | Partly: `test_range.py::test_a_cell_not_yet_calculated_stops_the_call_and_excel_calls_again` (`xlretUncalced`, each later read refused, then the right result); `xlretAbort` and the count of calls into Excel are left to the fake, since real Excel cannot be made to interrupt at a chosen read, nor its calls counted |
| A call Excel stopped says nothing on the console, a PyCall's either | `test_funcs.py::test_a_call_excel_stopped_says_nothing_of_its_failure` | `test_range.py::test_a_pycall_excel_stopped_says_nothing` | Not checked in real Excel |
| A function called from no cell is told so, not that it is a PyCall | Left to the fake: it needs `xlfCaller` to give no cell | `test_range.py::test_a_function_called_with_no_cell_says_so` | Not checked in real Excel |
| `values()` and `empty()` refuse outside a call and on a script's thread | Left to the fake: the refusal is the core's | Partly: `test_range.py::test_reading_outside_a_call_is_refused`, `test_range.py::test_reading_on_a_scripts_thread_is_refused` (`values()`); `empty()` has no test of its own here: it reads through the same function as `values()` | Partly: the same two tests, `values()` only, as for the fake |
| A range kept after its call gives its address, and a later call reads it, unless its sheet is gone or renamed; a sheet put in its place under the same name is read only if Excel gives it the same ID | `test_funcs_range.py::test_a_range_whose_sheet_is_gone_is_not_read` (no sheet of that name, or another ID) | `test_range.py::test_a_range_kept_after_its_call`, `test_range.py::test_a_kept_range_whose_sheet_is_gone_is_not_read` (a sheet deleted, a workbook closed) | `test_range.py::test_a_range_kept_after_its_call`, `test_range.py::test_a_kept_range_whose_sheet_is_gone_is_not_read` (a sheet deleted, a sheet renamed, a workbook closed), `test_range.py::test_a_kept_range_whose_sheet_is_replaced_under_its_name_is_read_only_with_its_id` (a sheet deleted and another added under its name: read only if it has the old ID, as Microsoft 365 gives it) |
| Gaining or losing the annotation registers the function again; meanwhile a call gives `#VALUE!` | `test_funcs_range.py::test_gaining_or_losing_the_annotation_registers_again` | `test_range.py::test_a_value_becoming_a_reference_registers_again`, `test_range.py::test_a_reference_becoming_a_value_registers_again`, `test_range.py::test_a_registration_as_a_reference_that_fails_is_tried_again_by_pyload`, `test_range.py::test_several_edits_between_value_and_reference_while_queued` | `test_range.py::test_a_value_becoming_a_reference_and_back_registers_again` |

**Type annotations** ([conversion.md](./conversion.md#type-annotations-direct-functions))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| `int`, `float`, `str` and `bool`: each row of the table, converting and refusing, with its message | `test_funcs_annotations.py::test_a_scalar_converts`, `test_funcs_annotations.py::test_a_scalar_refuses_naming_the_function_argument_and_value` (each row) | Partly: `test_annotations.py::test_an_annotation_converts`, `test_annotations.py::test_an_annotation_refuses_naming_the_function_argument_and_value` (most rows; not, for one, `float` given an error or a range, `str` given a logical value or a range, or `bool` given an error or a range); the rest are left to the unit tests: the same conversion in Python in any host | Partly: the same two tests, the same rows, as for the fake |
| A list of rows: each element kept, or converted by its type, a refused element named by row and column | `test_funcs_annotations.py::test_a_list_of_rows_keeps_each_element`, `test_funcs_annotations.py::test_a_list_of_typed_rows_converts_each_element` (`float`, `int`, `str`, `bool` and `float \| None` elements), `test_funcs_annotations.py::test_a_refused_element_refuses_the_argument_naming_its_row_and_column` | Partly: `test_annotations.py::test_an_annotation_converts`, `test_annotations.py::test_an_annotation_refuses_naming_the_function_argument_and_value` (untyped, `float` and `int` elements); the other element types are left to the unit tests, as above | Partly: the same two tests, the same element types, as for the fake |
| `numpy.ndarray`: float64 of the range's shape, `nan` for an empty cell, anything but numbers refused | `test_funcs_annotations.py::test_an_array_of_numbers_has_the_ranges_shape`, `test_funcs_annotations.py::test_a_single_value_is_a_one_by_one_array`, `test_funcs_annotations.py::test_an_array_refuses_anything_but_numbers`, `test_funcs_annotations.py::test_an_optional_array_gives_none_for_an_empty_cell` | `test_annotations.py::test_an_annotation_converts`, `test_annotations.py::test_an_annotation_refuses_naming_the_function_argument_and_value` | `test_annotations.py::test_an_annotation_converts`, `test_annotations.py::test_an_annotation_refuses_naming_the_function_argument_and_value` |
| `T \| None` gives `None` for an empty cell, in each spelling | `test_funcs_annotations.py::test_none_for_an_empty_cell_in_each_spelling` | Partly: `test_annotations.py::test_an_annotation_converts` (`Optional[int]` only); the other spellings are left to the unit test: read in Python, the same in any host | Partly: `test_annotations.py::test_an_annotation_converts` (`Optional[int]` only), as for the fake |
| An annotation written as a string, or under `from __future__ import annotations`, is evaluated where the function was defined, again while the result is text, up to five times; text still giving text is an annotation not converted | `test_funcs_annotations.py::test_an_annotation_written_as_a_string_is_evaluated`, `test_funcs_annotations.py::test_every_annotation_a_string_under_future_annotations`, `test_funcs_annotations.py::test_a_quoted_annotation_is_evaluated_until_it_is_not_text`, `test_funcs_annotations.py::test_text_naming_itself_ends_as_an_annotation_not_converted` (text that never ends, not converted), `test_funcs_annotations.py::test_text_is_evaluated_five_times_at_most` (five evaluations convert, a sixth is not made), `test_funcs_annotations.py::test_an_imported_function_s_annotation_is_evaluated_where_it_was_defined` | Partly: `test_annotations.py::test_an_annotation_converts` (one string annotation); the rest are left to the unit tests: evaluated in Python, the same in any host | Partly: `test_annotations.py::test_an_annotation_converts` (one string annotation), as for the fake |
| Any other annotation is left alone, said at registration; `Any` and `object` silently | `test_funcs_annotations.py::test_any_and_object_leave_the_value_silently`, `test_funcs_annotations.py::test_another_annotation_leaves_the_value_and_says_so_at_registration` | Partly: `test_annotations.py::test_an_annotation_pyinex_does_not_convert_is_said_at_registration_and_left_alone` (`Decimal`); `Any` and `object` are left to the unit test: decided in Python, the same in any host | Partly: `test_annotations.py::test_an_annotation_pyinex_does_not_convert_is_said_at_registration_and_left_alone` (`Decimal`), as for the fake |
| An annotation written as a string that cannot be evaluated, whatever it raises (`SystemExit` included), leaves its argument unconverted, said at registration, and the function registers. From Python 3.14, an unquoted one that fails does as [conversion.md](./conversion.md#type-annotations-direct-functions) lists: a missing name leaves its argument unconverted, and `SystemExit` refuses its function while the script's others register | `test_funcs_annotations.py::test_an_annotation_that_cannot_be_evaluated_leaves_the_value_and_says_so`, `test_funcs_annotations.py::test_an_annotation_that_raises_system_exit_leaves_only_its_argument_unconverted`; from 3.14, `test_funcs_annotations.py::test_a_forward_reference_that_cannot_be_evaluated_leaves_the_value_and_says_so` (a missing name, unquoted) and `test_funcs_annotations.py::test_a_lazy_annotation_that_raises_system_exit_refuses_only_its_function` (each skipped before 3.14) | Partly: `test_annotations.py::test_without_numpy_an_annotation_naming_it_is_left_alone` (a string annotation that cannot be evaluated, `NameError`); the rest are left to the unit tests: evaluated in Python, the same in any host | Left to the others: as for the fake, and every real-Excel run has NumPy, so the one case the fake checks does not arise |
| Without NumPy: a script importing it fails to load; a string annotation naming it is left alone | `test_funcs_annotations.py::test_without_numpy_a_script_importing_it_fails_to_load`, `test_funcs_annotations.py::test_without_numpy_a_string_annotation_naming_it_leaves_the_value` | `test_annotations.py::test_without_numpy_an_annotation_naming_it_is_left_alone` | Left to the others: every real-Excel run has NumPy |
| Annotated and plain parameters in one function; a plain one arrives as PyCall passes it | `test_funcs_annotations.py::test_annotated_and_plain_parameters_in_one_function` | `test_annotations.py::test_annotated_and_plain_parameters_in_one_function` | `test_annotations.py::test_annotated_and_plain_parameters_in_one_function` |
| A position left out is never converted | `test_funcs_annotations.py::test_a_position_left_out_is_never_converted` | `test_annotations.py::test_a_position_left_out_is_not_converted` | `test_annotations.py::test_a_position_left_out_is_not_converted` |
| The return annotation, and one on `**kwargs`, are ignored; an annotation on `*args` applies to each value, and one on a keyword-only parameter converts | `test_funcs_annotations.py::test_the_return_annotation_is_ignored`, `test_funcs_annotations.py::test_an_args_value_refused_is_named_by_its_position`, `test_funcs_annotations.py::test_a_keyword_only_value_converts`, `test_funcs_annotations.py::test_an_annotation_on_kwargs_is_ignored` | Left to the unit test: decided in Python, the same in any host | As for the fake |
| From 3.14, an annotation that raises, even an ignored one, leaves the function registered with its arguments unconverted, saying why; a note is said once, not at each load; an inherited constructor's annotations are read where it was defined | `test_funcs_annotations.py::test_an_ignored_annotation_that_raises_refuses_nothing`, `::test_a_parameter_s_annotation_that_raises_leaves_it_unconverted`, `::test_a_note_on_an_annotation_is_said_once_not_at_each_load`, `::test_an_inherited_constructor_s_annotations_are_read_where_it_was_defined` | Left to the unit tests | Left to the unit tests |

**Lifetime and scale** ([architecture.md](./architecture.md#lifetime))

| Rule | Unit | Fake Excel | Real Excel |
|---|---|---|---|
| Closed or unloaded with work queued, then opened again: no stale slot, timer or queued work | Left to the fake: it is the loader's | `test_lifetime.py::test_unloaded_with_work_queued_then_reopened_nothing_is_left_from_before`, `test_lifetime.py::test_closed_with_work_queued_and_reopened_nothing_is_left_from_before` | `test_lifetime.py::test_unloaded_with_work_queued_and_loaded_again` |
| A quit canceled at the save prompt brings the functions back; Protected View is not taken for an unload | Left to the fake, as above | `test_lifetime.py::test_a_quit_cancelled_at_the_save_prompt_brings_the_functions_back`, `test_lifetime.py::test_a_cancelled_quit_brings_back_the_functions_pyinex_ini_lists`, `test_lifetime.py::test_protected_view_after_a_cancelled_quit_is_not_taken_for_an_unload` | Partly: `test_lifetime.py::test_a_quit_cancelled_at_the_save_prompt_brings_the_functions_back` (the functions back); Protected View after a canceled quit is left to the fake: the test would need a Protected View window open while the quit's prompt is canceled |
| A canceled quit in a session that never used Pyinex starts no Python | Left to the fake: it asks whether Excel unloaded the add-in | `test_lifetime.py::test_a_cancelled_quit_in_a_session_that_never_used_pyinex_starts_no_python` | Not checked in real Excel |
| Thousands of calls and hundreds of reloads leave handles and memory flat | Left to the others: it needs the add-in running | Partly: `test_removal.py::test_thousands_of_calls_and_hundreds_of_reloads_leave_handles_flat` (3,000 calls and 300 registrations again add at most 3 handles; the fake does not measure memory) | `test_scale.py::test_ten_thousand_direct_function_calls_leak_nothing` (10,000 calls add at most 20 handles and 32 MB of private memory), `test_scale.py::test_registering_again_hundreds_of_times_leaks_nothing` (180 registrations again add at most 150 handles and 32 MB) |
