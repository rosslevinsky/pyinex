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
count is above 2,500 lines. CI runs it on every push.

The unit and fake-Excel tests below need pytest and openpyxl in each Python that runs
them:

```
python -m pip install pytest openpyxl
```

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
- `test_call.py` and `test_module_cache.py` run PyCall's own code (`src/core/call.cpp`):
  the argument rule, every kind of exception, and loading, reloading and failing to load
  scripts that the tests write into temporary folders. The module's `call` takes the
  workbook's folder as an argument, since there is no Excel to ask.
- `test_examples.py` checks the example workbooks without Excel (see
  [Real Excel](#real-excel-testsexcel)). It needs only openpyxl, so it also runs on macOS
  and Linux, where it is the only unit test collected.
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
code inside the add-in call the functions in `scripts/helpers.py` through PyCall.

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
`LICENSE` and `NOTICE`, with that version's `CHANGELOG.md` entry as its notes.

## Real Excel (`tests/excel`)

These tests run the add-in in Microsoft Excel itself. They make the checks that no other
test can make, because they do not depend on what anyone expected Excel to do.

`tests/excel` is a pytest suite that controls Excel through COM, with pywin32. pytest
collects it only when it is given the add-in to test (`--excel`), so `pytest tests` and
CI never start Excel. Each test works in a new, saved workbook in a temporary folder of
its own, and writes there any script it calls. The session starts one new Excel, leaving
any open Excel alone, loads the add-in, and checks that Excel embedded the intended Python
before any test runs. One script, `scripts/excel_tests.py` (the driver, below), runs the
suite under every Python.

### What each file checks

- `test_examples.py` runs the example workbooks. Every formula on every sheet of
  `examples/*.xlsx` is a check, and the cells just to the right of its result hold the
  value it should show. For a formula entered into a range of cells, the expected block,
  of the same size, is just to its right. Because `addin.xlsx` shows the console and
  prints to it, this file also checks that Pyinex's console window is shown and holds that
  text.

  `scripts/make_examples.py` writes the workbooks from tables in the script, with
  openpyxl, on any system. To change a workbook, change a table and run the script again;
  never edit a workbook in Excel. `tests/unit/test_examples.py` checks, without Excel, that
  every formula has an expected value beside it, uses only the documented functions, and
  names a script and function that exist, and that the committed workbooks match what the
  script writes now.
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
  again) run only in the driver's runs of their kind, each in an Excel of its own.
- `test_threads.py` checks that every call runs on Excel's main thread when multi-threaded
  recalculation is on.
- `test_scale.py` checks that 10,000 calls, each with the right result, add at most 20
  handles and 32 MB of private memory to Excel. In testing, the change ranged from 11
  fewer to 2 more handles, and from 4.5 MB less to 8.6 MB more memory, which is Excel's own
  variation from run to run.
- `test_registration.py` checks that Excel lists the four functions, with their type
  strings, none marked thread-safe.
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

The driver needs Windows, 64-bit Excel, and pytest and pywin32 for the Python that runs it.
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
  sometimes took a minute to quit, with or without an add-in;
- a setting the driver could not restore.

If an Excel is still running after a run and the driver could not end it, the run fails
and the driver stops, because that Excel may change settings as it quits. To end an Excel,
the driver must have its process handle and the start time that the test session
recorded, so it cannot end another process that reused the same process ID.

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
different selection than `tests/excel`.

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
| `test_registration.py::test_registers_the_documented_functions`, `::test_no_function_is_thread_safe` | Partly: `test_registration.py::test_excel_lists_the_documented_functions` (the four functions and their type strings, none thread-safe). Excel's list gives no names of arguments, category or help, which only the fake checks |
| `::test_pycall_argument_text_fits_excels_limit` | Fake only: it checks the text the add-in registers against Excel's 255-character limit, a check of the text itself; `test_registration.py` shows real Excel registered every function |
| `test_threads.py::test_a_thread_keeps_running_between_calls` | `test_output.py::test_a_scripts_thread_prints_between_calls`; and with multi-threaded recalculation, `test_threads.py::test_every_call_runs_on_excels_main_thread` |
