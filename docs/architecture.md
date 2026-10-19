# Architecture

This document describes how Pyinex works, and why it works that way, for anyone who changes
it. It starts with what a user sees: the functions, the argument rule, scripts, the
`pyinex` module and direct functions. It then covers what happens inside: choosing a Python, the add-in's
lifetime, threads and the console. The last sections,
[How a call moves through the code](#how-a-call-moves-through-the-code) and
[Not yet built](#not-yet-built), connect these to the source files and to later work.
Value conversion has its own document, [conversion.md](./conversion.md).

Each decision gives its reason and names the test file that checks it. Where this version
differs from Pyinex 0.83 (2010) for a reason worth knowing, the section says so.

## Excel functions

Pyinex has five worksheet functions of its own, in the function category "Pyinex". A
script can add more, each under its own name: see [Direct functions](#direct-functions).
It also registers two commands. A command is not a worksheet function, so no formula can
use one; Excel runs it by name, from VBA's `Application.Run` or from the Macro dialog
(Alt+F8) with its name typed in. The dialog lists neither command, nor any direct function,
though it lists a command a workbook defines; a direct function's name typed there calls
the function with no arguments, as a formula giving none would, and Excel drops its result
(`tests/excel/test_macro_dialog.py`).

- `PyRepair` loads the scripts of every open workbook's `PyLoad` cells and repairs cells
  left at `#NAME?` (see [The `#NAME?` repair](#the-name-repair)).
- `PyinexQueue` followed by a number registers functions after start-up (see
  [The queue](#the-queue)). It is Pyinex's own, and a user has no reason to run it.

| Function | Purpose |
|---|---|
| `PyCall(file, function, arg1, …, arg30)` | Call a Python function. |
| `PyConsole(show, [clear], [x], [y], [width], [height])` | Show, hide, clear and place the console window. |
| `PyReloadOnChange([on])` | Turn reloading of changed scripts on or off, or report the setting. |
| `PyInfo([item])` | Report which Python and which add-in are loaded. |
| `PyLoad(file)` | Load a script and register the functions it marks `@pyinex.func`. |

None of them is registered as thread-safe (the `$` flag in the type string), so Excel
calls them only on its main thread. See [Threading](#threading).

None of them is volatile. A cell recalculates when its inputs change, or on a full
recalculation (Ctrl+Alt+F9).

**In the Function Arguments dialog.** This dialog, which Excel opens from Insert Function,
calculates a formula while the user types it. There, each function returns a blank and does
nothing else: PyCall runs no Python, PyConsole changes no window, PyReloadOnChange changes
no setting, PyInfo does not start Python, and PyLoad loads nothing. Otherwise every
keystroke would run the function.

Pyinex recognizes the dialog as Microsoft's sample code does: a window in Excel's process
whose class name starts with `bosa_sdm_XL` and whose title is "Function Arguments".
Microsoft's sample also checks that Excel's main window owns the dialog. Pyinex does not,
because since Excel 2013 each workbook has its own window, which may own the dialog
instead. The title is the English one. In Excel in another language, the functions
therefore run in the dialog as they do in a cell: typing is slower, but no result is wrong.

Tests: `tests/host/test_registration.py` (names, argument counts, type strings, no `$`,
the two commands), `tests/host/test_calls.py` (behavior, including each function in the dialog),
`tests/host/test_pyload.py` (PyLoad in the dialog).

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
  from it. An unsaved workbook has no folder, so in one a relative path is an error. So it
  is, saying why, when there is no folder for another reason: a call from no cell, as from
  a macro's `Application.Run`; a question Excel does not answer; or a workbook opened from
  OneDrive or SharePoint, whose `CELL("filename")` is a web address (`https://…/[Book.xlsx]Sheet1`).
  The console names the reason and asks for the script's full path
  (`tests/host/test_calls.py::test_a_relative_path_with_no_calling_cell_says_so`,
  `::test_a_workbook_stored_on_the_web_says_a_relative_path_needs_a_folder`). (In
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
- `x` must be given with `y`, and `width` with `height`. `x` and `y` must lie on a
  monitor, any of them, so they can be negative for a monitor to the left of or above the
  main one; a point in the rectangle around all the monitors that none of them shows, as
  between monitors of different sizes, is refused. `width` and `height` must be more than 0. A value that breaks
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
table of names and values. Asking for any item but `"pyinex"` starts Python if it has not
started. `"pyinex"` needs no Python: the queue asks for it after a close, to see whether
Excel unloaded the add-in ([Lifetime](#lifetime)), which must not start Python in a session
that never used Pyinex (`tests/host/test_lifetime.py::test_a_cancelled_quit_in_a_session_that_never_used_pyinex_starts_no_python`). An
unknown item gives `#VALUE!`.

Tests: `tests/host/test_calls.py`; in real Excel, `tests/excel/test_pyinex_module.py`.

### PyLoad

`PyLoad(file)` loads the Python script `file` at any time and registers the functions it
marks `@pyinex.func` (see [Direct functions](#direct-functions)), so that any cell of any
open workbook can call them by name.

- **`file`** is a path, as for PyCall: a relative path is relative to the folder of the
  workbook that contains the calling cell, and in a workbook never saved a relative path
  is an error.
- **Loading** is PyCall's ([Scripts](#scripts)), with the same cache: a script used
  through both PyLoad and PyCall is one module, and with `PyReloadOnChange` on, a changed
  script loads again when the cell next calculates. The script runs outside any call, as
  one that `pyinex.ini` lists does, so at its top level `pyinex.caller`, `pyinex.sheet`
  and `pyinex.esc_pressed` raise `RuntimeError`, even when the PyLoad cell calculates
  within another function's call; and nothing PyLoad meets changes what that call may ask
  Excel. Within a call Excel has stopped, PyLoad asks Excel nothing either, so a relative
  path finds no workbook folder (`tests/host/test_reload_functions.py::test_a_pyload_within_a_call_runs_its_script_outside_any_call`,
  `::test_a_pyload_within_a_stopped_call_asks_excel_nothing`).
- **Registering comes after the calculation.** Excel refuses to register a function while
  it calculates a cell, so PyLoad hands the script's functions to [the queue](#the-queue),
  which registers them as soon as the calculation is over. A function is registered only
  after the whole script has run without an exception.
- **The result** is a short status, the number of direct functions the script declares,
  registered or waiting to be, as text: `1 function`, `2 functions`. It is the same each
  time the cell calculates. A function refused as the script is read (see below) is not
  counted; one Excel refuses later, as a name Excel already knows, is, and the console
  reports it.
- **A script that fails to load**, because it is missing, has a syntax error or raises at
  its top level (`SystemExit` and `KeyboardInterrupt` included), gives `#VALUE!`, with the
  reason and any traceback on the console, and registers nothing, not even a function it
  marked before it failed. The next calculation of the cell loads the fixed file.
- **The console** names each function as it registers (`Pyinex: registered Discount from
  the script model.py`), and each one refused, with the reason.
- **Calculating again** registers only what is new. A function already registered from the
  same script is left as it is, unless the script has changed (see
  [Reloading a script](#reloading-a-script)). A name already registered from another
  script is refused, with a console message naming both scripts: the first registration of
  a name wins ([Names](#names)). A function whose registration failed is tried again the
  next time its PyLoad cell calculates.
- PyLoad is not volatile. Its cell calculates when it is entered, when its argument
  changes, and on a full recalculation (Ctrl+Alt+F9); and Pyinex has Excel calculate it
  when its workbook opens ([When a workbook opens](#when-a-workbook-opens)).

Tests: `tests/unit/test_funcs.py` (what a second load registers, names already taken, a
failed registration freeing its name, relative paths), `tests/host/test_pyload.py` (each
rule above); in real Excel, `tests/excel/test_pyload.py` (relative and absolute paths,
each kind of failure and the fixed file after it, a script loaded twice and from two
workbooks, one script through PyLoad and PyCall).

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
  name is unique: two paths whose hashes collide get names of their own, the second with a
  number after (`pyinex_model_3f9a1c2e_2`), and each path keeps its name as it reloads
  (`tests/unit/test_funcs.py::test_two_scripts_whose_paths_collide_in_crc32_have_modules_of_their_own`).
  `__name__ == "__main__"` is false, as for any imported module.
- **Other files beside a script.** The first time a script from a folder loads, Pyinex
  puts that folder at the front of `sys.path`, where it stays, as `python script.py` does
  for the script's folder. So `import helper` finds a `helper.py` beside the script,
  including inside functions. These files are ordinary imports: Python caches them in
  `sys.modules` under their plain names, every script that imports one shares it, and
  Pyinex does **not** reload them when they change. Only the scripts named in PyCall,
  loaded by PyLoad or listed in `pyinex.ini` are reloaded. To use a changed helper,
  restart Excel or call `importlib.reload`. As with `python script.py`, a file in the
  script's folder with the name of a standard module, such as `random.py`, hides that
  module. Two workbooks whose folders each hold a `helper.py` share whichever loaded
  first, as one Python process would, so their scripts need helpers of different names.
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
  or direct function is running, in A1 style (`C12`) or, with `r1c1=True`, in R1C1 style
  (`R12C3`). With `full=True` it puts the workbook and sheet in front, as in
  `[Book.xlsx]Sheet1!C12`, without quotes even when the sheet name has spaces. For a
  formula entered into a range of cells (a legacy array formula), it returns that range,
  for example `C12:E14`. Columns run from `A` to `XFD`. `tests/unit/test_colnames.py`
  checks the column letters at every boundary.
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
- **`@pyinex.func`** makes a function a worksheet function under its own name; see
  [Direct functions](#direct-functions).
- **`pyinex.Range`** is the annotation that makes a direct function's argument arrive as
  a reference, and the type of what arrives; see [pyinex.Range](#pyinexrange).

### When they may be called

`caller`, `sheet` and `esc_pressed`, and a `pyinex.Range`'s `values()` and `empty()`, ask
Excel for information. Microsoft allows that only
on Excel's calculation thread, while Excel is calling the add-in. So each of them raises
`RuntimeError`, without asking Excel:

- when called from any other thread, such as one that a script started;
- when called outside a PyCall or a direct function's call, for example by code that runs
  while the add-in closes;
- for the rest of a call in which Excel answered that a cell is not yet calculated, or
  that the calculation was interrupted ([pyinex.Range](#pyinexrange)).

`caller` and `sheet` also raise `RuntimeError` when the running PyCall or direct function
was not called from a cell, as happens through `Application.Run` in VBA or COM (`the
function running was not called from a cell`; `tests/host/test_range.py::test_a_function_called_with_no_cell_says_so`). Excel's
`Application.Evaluate`, however, calls the function as if from the active cell, so there
they return that cell. `ExcelError` works anywhere.

Tests: in real Excel, `tests/excel/test_pyinex_module.py` (every form against the calling
cell, legacy array and spilling formulas, the far corner of the grid, and the refusals); `tests/host/test_pyinex_module.py` (both address styles,
full and short forms, a Unicode sheet name, `esc_pressed` clearing and keeping the key
press), `tests/host/test_lifetime.py` (the refusals); a `pyinex.Range`'s, `tests/host/test_range.py`
and `tests/excel/test_range.py`.

## Direct functions

A function marked `@pyinex.func` in a script becomes a worksheet function under its own
name, so a cell can say `=Discount(A1, B1)` where it would say
`=PyCall("model.py", "Discount", A1, B1)`:

```python
import pyinex

@pyinex.func
def Discount(price, rate):
    """The price after a discount.

    Args:
        price: The price before the discount.
        rate: The discount, as a fraction such as 0.1.
    """
    return price * (1 - rate)
```

**In outline.** Five parts make this work, each described below:

1. **Loading.** `PyLoad(file)` in a cell, or the `scripts` list in `pyinex.ini`, loads a
   script as PyCall would, and Pyinex collects the functions it marks.
2. **The pool.** Excel calls only functions the add-in exports, so `pyinex.xll` exports
   1,000 spare ones, `PyFunc0` to `PyFunc999`. Each Python function is registered with
   Excel under its own name on a free one, its *slot*, and every slot leads to one
   dispatcher that calls the core much as PyCall does (the pool of numbered exports,
   below).
3. **The queue.** Excel lets an add-in register a function only inside a command, never
   inside a worksheet function such as `PyLoad`. So Pyinex queues the work, and a hidden
   window's timer has Excel run Pyinex's queue command, through `Application.Run`, usually
   just after the calculation ends ([The queue](#the-queue)).
4. **Reloading.** When a script reloads, functions added, removed, or changed
   in what Excel is told go through the queue the same way
   ([Reloading a script](#reloading-a-script), [Removing a function](#removing-a-function)).
   A name Excel, another add-in or another script already uses is refused ([Names](#names)).
5. **Workbooks.** When a workbook opens, Pyinex calculates its `PyLoad` cells so its
   functions register first ([When a workbook opens](#when-a-workbook-opens)). It enters
   again any formula left at `#NAME?` because it was typed before its function registered
   ([The `#NAME?` repair](#the-name-repair)). The `PyRepair` command does both on demand.

[How a call moves through the code](#how-a-call-moves-through-the-code) follows one call
through the source files.

**What Excel is told.** Excel shows each of these in Insert Function and the Function
Arguments dialog:

- **The name** is the function's own, or the decorator's `name=`, as in
  `@pyinex.func(name="DISCOUNT")`.
- **The argument names** are the function's parameters.
- **The help text** is the docstring's first paragraph, or the decorator's `help=`.
- **Each argument's help** comes from the docstring's `Args:` (or `Arguments:`) section:
  a line `name: text` or `name (type): text`, with any further lines indented under it. A
  docstring in another style gives the help text only.
- **The category** is the decorator's `category=`, or "Pyinex".

The decorator only marks the function and returns it unchanged.

Excel takes at most 255 characters for each of these texts, counted as Excel counts text,
in UTF-16 code units, so a character beyond U+FFFF (an emoji, a rare CJK character) counts
as two. Pyinex counts the same way. A longer name is refused ([Names](#names)). A longer
help text or category is cut at 255, never within a character, and the list of argument
names after the last whole name that fits, or to nothing when even the first is longer
(`tests/unit/test_funcs_signatures.py`, `tests/unit/test_funcs.py`). Excel shows the last
help text it is given (the last argument's, or the function's for a function with none)
short, so Pyinex gives every function it registers one more argument help text, empty, as
Microsoft advises; the last real one then shows whole, even at 255 characters
(`tests/host/test_direct_functions.py`, and in real Excel Insert Function's help for a
function with no arguments).

**Which scripts: `PyLoad` and `pyinex.ini`.** A cell's `=PyLoad(file)` loads a script at
any time and registers its functions through [the queue](#the-queue) (see
[PyLoad](#pyload)). The `scripts` key in `pyinex.ini`'s `[python]` section
(see [Where Pyinex looks](#where-pyinex-looks-in-order)) lists scripts, separated by
semicolons. A relative path is relative to the add-in's folder, as `executable`'s is.
When the add-in opens, Pyinex starts Python, loads each script in the list, in order, as
PyCall loads a script ([Scripts](#scripts)), and registers its functions there, in
`xlAutoOpen`, where Excel allows registering. Python starts when the add-in opens only
when the list names a script; otherwise it still starts at first use. A script that fails
to load is reported on the console, as for PyCall, and the others still load. A script
loaded then runs outside any call, so at its top level `pyinex.caller`, `pyinex.sheet` and
`pyinex.esc_pressed` raise `RuntimeError` ([When they may be called](#when-they-may-be-called)).

**Registered only after the whole script has run.** Pyinex collects the marked functions
from the script's top level once the script has run without an exception, in the order
the script defines them. A script that raises, even after it has marked a function,
registers nothing, as PyCall's cache keeps no script that failed to load. Only the
script's own functions count, those it marked itself: a marked function it imports from
another file is not registered again.

**Signatures.** A direct function takes up to 30 Excel arguments. Pyinex reads its
signature with `inspect` and gives each parameter Excel positions in a fixed order:

1. its positional parameters (positional-only or ordinary), with or without defaults, in
   signature order;
2. then its keyword-only parameters, in signature order;
3. then `*args`, which takes every remaining position up to 30 and ends at the last one
   the formula gives.

`**kwargs` takes no position and is always empty: a cell passes values by position only
([Argument rule](#argument-rule)). Excel is told the function takes as many arguments as
it has positions, named after their parameters; `*args`' positions are numbered, so for
`*items` they are `items1`, `items2`, and so on, each with the help of `items` (or
`*items`) from the docstring. A keyword-only parameter is given by its position, and
passed as a keyword argument. For example:

```python
@pyinex.func
def f(x, *items, scale=1):
    ...
```

has 30 positions: `x`, `scale`, then `items1` to `items28`. `=f(10, 2)` calls
`f(10, scale=2)`, and `=f(10, , 20, 30)` calls `f(10, 20, 30)` with `scale` left at its
default.

**Positions left out.** A position the formula leaves out takes its parameter's default.
One left out with no default is `None` when a later position is given, as PyCall passes a
missing argument in the middle; after the last position given it is not passed at all,
so Python reports the missing argument (`TypeError`: `#VALUE!`, with the message on the
console), as for PyCall. Within `*args`, a position left out before the last one given is
`None`. A position left out is never converted by its annotation: it is its default, or
`None`. A formula with more arguments than the function has positions gives `#VALUE!`
from Excel itself, which calls nothing.

**Annotations.** A parameter annotated `int`, `float`, `str`, `bool`, a list of rows or
`numpy.ndarray`, or one of these or `None`, has each value Excel passes converted or
checked before the call, by the table in
[conversion.md](./conversion.md#type-annotations-direct-functions). A value that does not
fit gives `#VALUE!` with a console message, and the function is not called. Any other
annotation is left alone, and the console says so when the function is registered.

**Refused.** A signature that needs more than 30 positions (its positional and
keyword-only parameters, and one more for `*args`), or that cannot be read (as for some
built-in functions), is not registered: the console names the function, its script and
the reason (`Pyinex: f in the script model.py is not registered: it has 31 parameters, and
Excel passes at most 30`; `… it has 30 parameters and *items, and Excel passes at most
30`; `… its signature cannot be read (…)`), and the script's other functions still
register. So is a function that fails in any other way as it is read, `SystemExit` and
`KeyboardInterrupt` included (`tests/unit/test_funcs.py::test_one_functions_help_raising_systemexit_refuses_only_that_function`).
A function registered before, whose new signature is refused, is removed from Excel, and its
calls meanwhile give `#VALUE!` with that reason, not as one removed from the script
(`tests/unit/test_funcs_reload.py::test_a_function_whose_new_signature_is_refused_says_why_not_that_it_was_removed`). From 3.14, a
signature is read with annotations in Python's forward-reference form, so an annotation
written without quotes that fails, such as a name imported only under `if TYPE_CHECKING:`,
need not stop the function registering. What each kind of failure does, before 3.14 and
after, is in [conversion.md](./conversion.md#type-annotations-direct-functions). (An
annotation that calls a function still runs it.) A note about an annotation is said as the
function registers, not again each time an unchanged script loads
(`tests/unit/test_funcs_annotations.py::test_a_note_on_an_annotation_is_said_once_not_at_each_load`). The
annotations of a class's inherited constructor are evaluated where that constructor was
defined, as a function's are where it was (`::test_an_inherited_constructor_s_annotations_are_read_where_it_was_defined`).

**What can be marked.** Anything callable that can take the decorator's mark (an
attribute) and whose signature can be read:

- **A function**, the usual case.
- **A class**, registered with its constructor's parameters. A call constructs an
  instance, and the instance is the result, converted as any result is: a class whose
  instances Excel cannot take gives `#VALUE!`, and one deriving from `float` or `str`
  works. Only the class marked is registered, not its subclasses or instances.
- **A `functools.partial`**, as `rate = pyinex.func(functools.partial(discount, years=1))`:
  registered under the name the script gives it (or `name=`), with the parameters the
  partial leaves, read as `inspect` reads them, so one it fixes by keyword becomes
  keyword-only, with that value as its default. A partial of a function whose signature
  cannot be read, such as `functools.partial(max)`, is refused, as above.
- **A built-in function**, such as `pyinex.func(math.sqrt)`, cannot take the mark. The
  decorator returns it unchanged and the console says so (`Pyinex: sqrt in the script
  model.py is not registered: a built-in function cannot be marked; wrap it in a function
  of your own`); the script loads, and its other functions register.

Tests: `tests/unit/test_funcs_signatures.py` (the positions and names, each case of a
position left out, the refusals, what can be marked), `tests/host/test_signatures.py` and,
in real Excel, `tests/excel/test_signatures.py` (0 to 30 arguments, each position checked
by its value, across the fourth and fifth; defaults; `*args` with none, one and many
values; keyword-only parameters; the examples above; a missing argument in the middle;
more arguments than registered; each refusal in a script whose other functions still
register and run; a class, a `functools.partial` and a built-in function; and reloads that
change the positions).

**Calling.** A direct function is PyCall with the script and the function fixed, its
arguments placed by the rules above. The value rules ([conversion.md](./conversion.md)),
the error rule (`#VALUE!`, with the reason and any traceback on the console), the main
thread, `pyinex.caller`, `pyinex.sheet` and `pyinex.esc_pressed` are all PyCall's. Each
call finds the function in its script as PyCall does, so with `PyReloadOnChange` on, a
changed script is loaded again and its changed body runs. Direct functions are not
thread-safe and not volatile, and in the Function Arguments dialog each returns a blank
and runs no Python, as PyCall does ([Excel functions](#excel-functions)). One difference:
a console message about converting an argument or a result starts with the function's
name in Excel (`Pyinex: Discount: argument 2: …`), where PyCall's starts with the script
and the function.

**The pool of numbered exports.** Excel registers only functions that the add-in exports
by name, and a script's function is named long after the add-in was built. So
`pyinex.xll` exports 1,000 spare functions, `PyFunc0` to `PyFunc999`, as Excel-DNA does.
Registering a Python function takes the first free one, its *slot*: Excel is given that
export under the Python function's name, with as many arguments as the function has. The
slot table (`src/loader/functions.cpp`) records which function each slot calls and that
count, and every export passes its slot to one dispatcher, which calls the core
(`pyinex_function`, `src/core/api.h`).

- **A full pool** registers nothing more. The console names each function that did not
  fit (`Pyinex: cannot register F from the script model.py: all 1000 of Pyinex's function
  slots are taken`), and every function already registered keeps working. A function that
  did not fit is tried again the next time its PyLoad cell calculates. Tests:
  `tests/host/test_removal.py` and `tests/excel/test_removal.py` (the pool filled exactly,
  then one function more).
- **A slot is reused at once** once its function has been removed, but only if every step
  of the removal succeeded ([Removing a function](#removing-a-function)). A slot whose
  removal failed is kept out of use until Excel restarts. This was decided in real Excel
  before it was built: without the name cleared, an old name reaches whatever function
  takes its slot next. Tests: `tests/host/test_removal.py` and
  `tests/excel/test_removal.py` (many cycles of registering and removing, each in the one
  slot, and no old name reaching a new function).
- **Why 1,000.** Before this was built, an add-in with 10,000 such exports filled them
  all in real Excel: Excel accepted every one, so its own limit never decides. Each export
  adds about 120 bytes to the add-in; 1,000 add about 120 KB, and registering all of them
  takes well under a tenth of a second. Far fewer are ever registered at once.
- **Generated at build time.** `CMakeLists.txt` sets the pool's size in one place and writes
  one line per export, `PYINEX_SLOT(n)`, into `generated/loader/pool.inc` in the build tree.
  A macro in `src/loader/functions.cpp` makes each line an export.
- **Reading only the arguments Excel passed.** Each export is a variadic C++ function,
  `LPXLOPER12 PyFunc5(LPXLOPER12 first, ...)`. Under the x64 calling convention a variadic
  function receives pointer arguments exactly as a fixed one does (the first four in
  registers, the rest on the stack), and its entry code only stores the four registers in
  space every caller reserves. The dispatcher then reads the arguments one at a time, only
  as many as the slot's count, so it never reads a position Excel did not write. This was
  checked in real Excel, before it was built, at every count from 0 to 30 and across the
  fourth and fifth arguments, where registers end and the stack begins. It needs no assembly:
  Excel-DNA's exports jump to their targets, which in C++ would need a MASM file and a
  target per slot.
- **Checked before Excel is told.** A count outside 0 to 30, or a description of a function
  that does not hold its count's argument help, is refused with a console message, and the
  slot stays free.

Tests: `tests/unit/test_funcs.py` (what the decorator collects, the help texts, the
signatures refused, nothing from a script that raised); `tests/host/test_direct_functions.py`
(every value rule, each also through PyCall and the two results compared; the call
context; the Function Arguments dialog; the `pyinex.ini` list, a script that fails among
others, a listed script's top level run outside any call, and an empty list leaving Python
unstarted), `tests/host/test_registration.py`
(each function's slot, type string, argument names, help and category, none thread-safe).
In real Excel, `tests/excel/test_direct_functions.py` (the same, each value a worksheet can
hold, and what Insert Function and the Function Arguments dialog show) and
`tests/excel/test_registration.py`.

### The queue

Excel lets an add-in register or unregister a function only in a command: in `xlAutoOpen`,
`xlAutoClose`, or a command Excel runs. A worksheet function such as `PyLoad` is refused.
A macro-sheet function (`#` in its type string) would be allowed, but the mark changes how
Excel recalculates its cell. So Pyinex queues each function to register, and registers it
in a command of its own as soon as Excel will run one.

- **The command.** `xlAutoOpen` registers the procedure `PyinexQueue` as a command (macro
  type 2), which registers, or removes, every function waiting, in order. Work queued while
  it runs is done by the same run. If the command is started again while it is running, the
  second run does nothing but say so on the console, so no function is registered twice.
  The add-in opening and closing count as the queue's run from their first call into Excel
  to their last, and an error in a run started meanwhile never ends that mark, so a run
  that Excel starts within one of their calls, or within a registration, never takes their
  work (`tests/host/test_queue.py::test_pyrepair_failing_inside_a_registration_leaves_the_queue_its_own`,
  `tests/host/test_lifetime.py::test_pyrepair_run_while_the_add_in_opens_again_loses_no_workbook_function`,
  `::test_pyrepair_run_while_the_add_in_closes_leaves_it_closed`). A close within a run, from
  within a call into Excel, puts what it removed ahead in the queue; the run takes off the
  queue the function it was registering, not the one now in front, and leaves the rest to
  the queue's next run, once the timer has seen whether Excel unloaded the add-in
  (`tests/host/test_queue.py::test_the_add_in_closed_while_excel_registers_*`). Nobody has seen
  real Excel run a command or close the add-in within `xlfRegister`; the fake Excel does both,
  and Pyinex is safe either way.
- **The command refused.** Excel can refuse to register the command as the add-in opens.
  `Application.Run` could then never run it, so rather than try for ever, the console says
  so as the add-in opens, and each function queued after is dropped with a note to close and
  open the add-in again (`tests/host/test_queue.py::test_a_queue_command_excel_refuses_at_opening_is_reported_and_its_work_dropped`).
  A refused `PyRepair`, or a refused worksheet function such as `PyLoad`, is reported too
  (`tests/host/test_repair.py::test_pyrepair_excel_refuses_at_opening_is_reported`,
  `tests/host/test_lifetime.py::test_a_worksheet_function_excel_refuses_at_opening_is_reported`).
- **The window and its timer.** When Pyinex first queues work, it makes a hidden window
  that only receives messages (`HWND_MESSAGE`), on Excel's main thread, and sets its
  timer. Opening the add-in does too when Excel's `Application` object cannot be reached
  yet, so that the command can start following workbooks as they open
  ([When a workbook opens](#when-a-workbook-opens)). Windows sends a timer's message only when no other message is waiting, so it
  normally arrives once the calculation that queued the work is over; if it comes sooner,
  Excel refuses, as below. Pyinex then asks Excel to run the command through COM, with
  `Application.Run`, which runs it in a command context. The command's name is
  `PyinexQueue` followed by the address where this copy of the add-in is loaded, so another
  copy of Pyinex loaded in the same Excel has a command of its own: running that copy's
  command would leave this copy's work waiting forever. (`Application.Run` also takes a
  registration ID, but real Excel then runs the command even while a cell is being edited.)
- **Never inside a calculation.** The timer's message can also arrive while one of
  Pyinex's worksheet functions is running, when a script handles Windows messages, as a
  message box or a GUI toolkit does. The queue then does nothing but set its timer again,
  and runs once the function has returned: running the command inside the function's call
  could call Excel after Excel told that call to call nothing more
  ([`pyinex.Range`](#pyinexrange)). The command checks the same itself, in case anything
  else runs it there. Tests: `tests/host/test_range.py::test_the_queue_waits_*`.
- **Finding Excel's `Application` object.** Pyinex looks in every workbook window
  (`EXCEL7`) under each of Excel's main windows (`XLMAIN`) on Excel's main thread, since each
  workbook has a main window of its own, and asks the window for its object through
  Windows' accessibility interface (`AccessibleObjectFromWindow`, in `oleacc.dll`), as
  Excel-DNA and xlOil do. A window whose object has no `Application`, as a Protected View
  window's has none, is passed over. With no workbook window, Pyinex takes the Excel in
  Windows' running object table, but only if that Excel's window belongs to this process.
  The table holds one Excel. With several Excels open, one that has no workbook open may
  not be the one in the table; its work then waits until a workbook opens in it, and never
  goes to another Excel. When nothing is queued to register, only workbooks to follow, the
  queue looks only in workbook windows, and waits without a word until one opens: with no
  workbook open there is nothing to follow.
- **Waiting, for as long as it takes.** Excel refuses to run the command while a cell is
  being edited, while a dialog is open, or while it is busy. Pyinex then tries again every
  250 ms. It never gives up and never drops work, because a refusal says only that Excel is
  busy, for however long the user takes. A refusal is reported on the console
  (`Pyinex: Excel cannot register functions yet (…); Pyinex keeps trying`) when its cause
  differs from the last one reported, so a cell edit that lasts minutes is one report, and
  a new cause is not hidden behind it. When Excel ran the command but work is left, as when
  memory ran out in the run, the cause given is that (`its last run of the queue stopped
  before the end`), not how Excel was found (`tests/host/test_queue.py::test_a_run_of_the_queue_that_stopped_is_reported_as_such`).
- **Every result is checked.** Each COM result is initialized before the call and checked
  before it is used. Each function's description is checked before Excel is told. Each
  failure to register, from Excel (`xlfRegister`), from a full pool or from a name Excel
  already knows ([Names](#names)), is reported on the console, and the function's name is
  free again for its PyLoad cell's next calculation. Each failure to remove a function, at
  any step, is reported too ([Removing a function](#removing-a-function)).
- **Removals too.** What a reload decided ([Reloading a script](#reloading-a-script)) goes
  through the queue in the order it was decided: each function to remove, then each to
  register, so a name is cleared before it is registered again.
- **Then workbooks.** Once nothing is left to register, the same run calculates the
  `PyLoad` cells of workbooks that opened ([When a workbook opens](#when-a-workbook-opens)),
  registering what they queue, and then repairs the cells that use names it registered
  ([The `#NAME?` repair](#the-name-repair)).
- **Nothing lost to an error.** A function leaves the queue only once Excel has registered
  it and its slot holds it, or once its failure is reported. Its slot's copy is made, and
  its name added to the repair's list, before Excel is told (the name is taken out again if
  Excel does not register it), so nothing that can fail, such as running out of memory,
  happens after Excel has registered it: a function registered with no slot holding it
  would let its name reach the next function to take the slot, and one registered and left
  queued would be registered again in another. A call into Excel allocates nothing, so it
  cannot fail for want of memory. An error before that leaves it queued for
  the next run. The functions `pyinex.ini` lists register as the add-in opens, outside the
  queue's command, but are marked as the queue's run, so a run started meanwhile (by
  `PyRepair` within a registration) leaves its work to the run that follows; one that an
  error stops before it is applied is offered again when its script next loads. Rows of
  functions that Python gave but that never reached the queue, as when memory ran out
  converting, handing over or queueing them, are handed back: their names are freed, so the
  next load offers them again, and a removal is still made, waiting in Python for the
  loader's next call or, once Python has let the function go, marked on its slot for the
  queue's next run, with nothing to allocate. The names the repair is to enter formulas for
  leave its list only once those formulas are entered, so a run an error stops enters them
  in the next. A close from within a removal's own call into Excel does not queue the
  function being removed to be registered again, and one from within a first registration
  leaves its repair asked for (`tests/host/test_removal.py::test_memory_running_out_*`,
  `::test_the_add_in_closed_while_excel_removes_a_function_does_not_bring_it_back`,
  `tests/host/test_queue.py::test_memory_running_out_after_a_new_name_registers_still_repairs_its_formulas`,
  `::test_the_add_in_closed_while_excel_registers_a_new_name_still_repairs_its_formulas`, `tests/unit/test_funcs.py::test_rows_that_did_not_reach_the_loader_are_offered_again`,
  `::test_changes_that_did_not_reach_the_loader_wait_for_its_next_call`). A result the core
  returned is freed whatever fails before it reaches Excel. If the window or its timer
  cannot be made, the console says so, made first where Python has not started, and each
  later calculation of a PyLoad cell tries again while work is waiting, whatever that
  calculation found, a repair or workbooks to follow as much as functions to register
  (`tests/host/test_queue.py::test_pyload_starts_again_a_timer_that_failed_with_only_a_repair_waiting`,
  `tests/host/test_console.py::test_a_timer_that_cannot_start_before_python_says_so_on_the_console`).

Tests: `tests/host/test_queue.py` (work run by the fake Excel's stand-in for
`Application.Run`, a refused `xlfRegister`, work queued while the queue runs, the queue
started again from inside itself, thousands of refusals in a row and then success, a new
cause of refusal reported, no workbook window, a Protected View window, a window or a timer
that could not be made, memory running out just before Excel is told of a function and just
after), `tests/host/test_pyload.py` (a malformed description refused), `tests/host/test_lifetime.py` (another
copy of Pyinex with the command's name); in real Excel,
`tests/excel/test_pyload.py` (registration queued during a calculation, a long cell edit, a
modal dialog, two Excels, no workbook open, Protected View).

### Reloading a script

Pyinex follows the scripts PyLoad and the `pyinex.ini` list loaded. Each time one of them
loads again, because it changed and `PyReloadOnChange` is on ([Scripts](#scripts)), Pyinex
compares the functions it now marks with those registered from it, by name without regard
to case, and updates Excel to match through [the queue](#the-queue):

- **A changed body**, or anything else Excel is not told, needs nothing. The next call
  finds the new function in the reloaded script, as PyCall does, with no call into Excel,
  and reads its signature and annotations again: an object the script marks again, such
  as a `functools.partial` it imports from a helper, is placed and converted as it now is
  (`tests/unit/test_funcs_reload.py`).
- **A changed signature**, or a change to anything else Excel is told (the name's
  spelling, the argument names, the help texts, the category), removes the function and
  registers it again.
- **A removed function** is removed ([Removing a function](#removing-a-function)).
- **A renamed function** is both: the old name is removed and the new one registered.
- **An added function** is registered.
- **A script that now fails to load** changes nothing in Excel: nothing new registers and
  nothing is removed. Each of its functions gives `#VALUE!`, with the reason and traceback
  on the console, as a PyCall of it would. Once the file is fixed, the next call of one of
  its functions, or the next calculation of its PyLoad cell, loads it and updates Excel to
  match, with no restart.

A script loads again when a cell calls one of its functions, when a PyCall calls a function
in it, or when its PyLoad cell calculates; whichever comes first notices the change. Only
scripts PyLoad or `pyinex.ini` loaded are followed: a script used only through PyCall
registers nothing when it changes. With `PyReloadOnChange` off, nothing changes: the
script stays as it was loaded, and its functions keep running as they were.

**While a function is registered again.** Until the queue has run, Excel still has the
old registration, with its argument count and types, and so does the function's slot: the
dispatcher reads only the arguments Excel was told of. Those arguments were read for the
old signature, so a call meanwhile never reaches the new code. It gives `#VALUE!`, and the
console says `Pyinex: Discount is being registered again, since its script changed; until
then it gives #VALUE!`. A removed function's calls meanwhile give `#VALUE!` too, saying
`… was removed from the script …`. A cell that showed `#VALUE!` meanwhile keeps it until it
calculates again; an ordinary recalculation of the cell, such as a change to one of its
inputs, then reaches the new registration. If the script changes again before the queue
has run, each change is queued in turn, in the order the reloads happened, and the last
one wins. A call that reloads a script hands its changes over when it returns; a PyLoad
cell calculated during that call, as a function calling into Excel can cause, queues them
before its own.

Tests: `tests/unit/test_funcs_reload.py` (each kind of change, the order of removals and
registrations, a load during a call's reload, reloading off, a script that now fails), `tests/host/test_reload_functions.py`
(each kind of change with the cells before and after, and each with reloading off; calls
before the queue has run; a registration again that Excel refuses, tried again by the
PyLoad cell; several edits while work is queued; a PyLoad during a call that reloaded the
script; a script that now fails, recovered through a call and through PyLoad). In real Excel, `tests/excel/test_reload_functions.py`
(the same, but for a registration again that Excel refuses, which cannot be brought about
there; several edits are made within one calculation, while Excel cannot run the queue).

### Removing a function

Pyinex removes a direct function when a reload finds its script no longer has it as
registered, and removes every one when the add-in closes ([Lifetime](#lifetime)). It
removes in a command, the queue's or `xlAutoClose`, in two parts:

- **Unregistering** (`xlfUnregister` with the registration's ID).
- **Clearing the name.** Unregistering alone leaves the name holding the registration's ID,
  which the same numbered export gets again when it is next registered, so the old name
  would call whatever function took its slot (found in real Excel before this was built).
  So Pyinex registers a command of its own, `xlAutoRemove`, under the name, deletes the name
  (`xlfSetName`) and unregisters the command, as Excel-DNA does.

Each step's result is checked: `xlfUnregister` and `xlfSetName` must answer `TRUE`. If any
step fails, the console says `Pyinex: cannot remove Discount from Excel; until Excel
restarts it gives #VALUE!`, and the slot is kept out of use: Excel may still send the name
to it, and it gives `#VALUE!`. The next function takes another slot. What could fail
after Excel is told, as running out of memory, is done before: once Excel has answered, the
slot is kept, or freed with what Excel then makes of the name recorded, and only the
console's note follows. So whatever fails, the slot Excel still sends an old name to is
never taken, and a function removed gets its name back (`tests/host/test_removal.py::test_memory_running_out_*`).
Once a removal is done, no cell can call the function, so Python forgets why it gave
`#VALUE!`; a registration Excel refused is forgotten too, so Python's records stay the same
size over a long session (`tests/unit/test_funcs.py::test_refusals_and_removals_leave_nothing_behind`).

**What removal leaves**, as real Excel shows it:

- Excel's registration table no longer lists the numbered export.
- Insert Function no longer lists the function.
- A cell that calls the function keeps its value until it calculates again, then gives
  `#VALUE!`.
- A formula typed afresh that calls it is accepted, and gives `#VALUE!`: Excel still reads
  the name as a function's, not as an unknown name (`#NAME?`).

Tests: `tests/host/test_removal.py` (what removal leaves; each step failing in turn:
unregistering the function, registering the stand-in command, deleting the name, and
unregistering the command, each keeping the slot out of use with the old name reaching no
other function, and the name then registering again). In real Excel, `tests/excel/test_removal.py` (each item above, before and
after; the registration table is read with a reader first checked to find PyCall's own
registration). A step cannot be made to fail in real Excel on demand; real Excel's answers
to a failed unregistering and a failed deleting of the name were checked, with an add-in
made for it, before this was built.

### Names

A direct function's name follows these rules, each refusal saying why on the console, and
the script's other functions still register:

- **Compared without regard to case**, as Excel compares names.
- **The first registration of a name wins.** A name another script registered, or is
  waiting to register, is refused, naming both scripts (`Pyinex: discount in the script
  b.py is not registered: the script a.py registered that name first`). A script that marks
  two functions under one name registers the first.
- **Pyinex's own names** are refused: `PyCall`, `PyConsole`, `PyReloadOnChange`, `PyInfo`,
  `PyLoad`, and the commands it registers (`PyRepair`, `xlAutoRemove`, and names starting
  `PyinexQueue`). Excel would take them, and Pyinex's functions would be lost.
- **Names Excel cannot take** are refused before Excel is told, saying which rule they
  break: a name is letters (any language's), digits, underscores and periods, starting
  with a letter or underscore, and at most 255 characters, a character beyond U+FFFF
  counting as two, as Excel counts; and it is not a cell reference
  (`A1` to `XFD1048576`, or in R1C1 style `R`, `C`, `R1C1`, `RC5`) or `TRUE` or `FALSE`.
  Excel refuses these itself, but says only that it failed.
- **Names Excel already knows**, as one of its own functions (`SUM`) or another add-in's,
  are refused: `Pyinex: cannot register SUM from the script model.py:
  Excel already knows that name, as one of its functions or another add-in's`. Excel would
  take the registration, but would go on calling its own function, or the other add-in's
  would lose its name. Before registering a name, Pyinex asks Excel, in
  the queue's command, to evaluate the name alone (`xlfEvaluate`). An unknown name gives
  `#NAME?`; anything else means Excel knows it: an add-in's function gives its
  registration's ID, and one of Excel's own `#VALUE!`. Excel reads some names as its own
  that look different: `Ünïcødé` reads as `UNICODE`. A name only typed in formulas before
  it was registered is still unknown, and registers. (Asking Excel to rewrite a formula
  calling the name, `FORMULA.CONVERT`, was tried and does not work: Excel also writes such
  a typed name in its typed spelling.) A name a workbook defines is not seen: Excel
  evaluates a name alone as one of its own or an add-in's, and a workbook's names only
  written with their sheet or `!`. A name a function of Pyinex's holds now is not asked
  about. A name Pyinex removed still evaluates to a registration's ID in Excel, so Pyinex
  notes what it evaluates to just after removing it, or just after failing to (its slot is
  then kept out of use), and takes that answer later as the name still free: a function
  removed and registered again, as after a changed signature, gets its name back, and so
  does one whose removal failed, however often (`tests/host/test_removal.py::test_a_name_whose_removal_failed_twice_registers_again`);
  a name another add-in took meanwhile, which then gives that add-in's ID, is refused. A
  question about the name that fails is never taken as that answer, so it refuses the name
  too, saying so: `Pyinex: cannot register F from the script model.py: Excel could not say
  whether it already knows that name`.

Tests: `tests/unit/test_funcs_names.py` (the rules decided in Python),
`tests/host/test_names.py` (each rule, with the fake Excel's own `xlfEvaluate` and
another add-in's name, including one another add-in took after Pyinex removed it or Excel
refused it), `tests/host/test_removal.py` (a name whose removal failed, registered again,
removed, then taken by another add-in, is refused); in real Excel, `tests/excel/test_names.py` (each rule, `SUM` and
other functions of Excel's, another add-in's function registered with Excel 4's
`REGISTER`, including one taking a name Pyinex removed, a name typed in a formula before it
registered, a Unicode name and the
decorator's `name=`). One name in two scripts: `tests/host/test_pyload.py` and
`tests/unit/test_funcs.py`.

### When a workbook opens

When a workbook opens, Pyinex has Excel calculate each of its cells whose formula calls
`PyLoad`, so the scripts the workbook uses load and their functions register before the
user changes anything or asks for a recalculation. A saved workbook that uses direct
functions therefore works when it is opened again, in a new Excel.

- **Learning of the open.** Pyinex connects to Excel's application events through COM (the
  `AppEvents` interface, which VBA's `Application` events use) and takes `WorkbookOpen`, and
  `NewWorkbook` for a workbook made from a template. It connects as the add-in opens, in
  `xlAutoOpen`, when a workbook window gives it Excel's `Application` object, and otherwise
  when the queue's command first runs, and handles every workbook already open at that
  moment: workbooks open before the add-in was loaded, and the file Excel was started
  with, which opens after Excel loads its add-ins and before its `Application` object can
  be reached. Excel calls the event outside any command, so Pyinex only keeps the workbook and
  sets the queue's timer; the queue's command does the rest ([The queue](#the-queue)).
  While the command runs, Excel tells a program automating it from another process that it
  is busy, as it does while any macro runs, so such a program may be told so just after it
  opens a workbook. It should try again, as the real-Excel tests do through a COM message
  filter (`tests/excel/conftest.py`). Closing the add-in disconnects ([Lifetime](#lifetime)).
- **Why not `ON.WINDOW`.** Excel 4's `ON.WINDOW` (`xlcOnWindow`) runs a command when a
  window is activated, and needs less code. Tried in real Excel before this was built, it
  never ran for a workbook made or opened through COM, by a macro or by another add-in,
  which is how many workbooks open; it runs on every switch of window, so it could not
  tell a workbook closed and opened again; and it is one setting for all of Excel, which
  another add-in's or a macro's replaces.
- **Finding the `PyLoad` cells.** In each worksheet of the workbook, hidden and very hidden
  ones included, Pyinex searches the formulas for `PyLoad(`, in any case, with Excel's own
  search (`Range.Find`, looking in formulas), so it reads no cell but those the search
  finds. Each cell found whose formula calls `PyLoad` (as the repair below decides it) is
  entered again with its own formula (`Range.Formula2`, or `Formula` in an Excel without
  dynamic arrays), and Excel calculates it as it would a formula typed in. So any form of argument works (`=PyLoad("m.py")`,
  `=PyLoad(A1)`, a path built by a formula), and a relative path is relative to the
  workbook's folder.
- **The workbook counts as changed.** Entering a formula again is a change, so Excel asks
  whether to save the workbook when it closes.
- **What Excel calculates first.** Excel may calculate cells while it opens a workbook,
  before Pyinex can act: a workbook saved by another version of Excel, or one with
  volatile functions, recalculates as it opens. A cell calling a direct function not yet
  registered then shows `#NAME?`. [The `#NAME?` repair](#the-name-repair) puts it right
  once the function registers.
- **A workbook opened while Excel's events are off is missed.** A macro may turn Excel's
  events off (`Application.EnableEvents = False`) around its work, as macros often do, and
  open a workbook meanwhile. Excel then reports no `WorkbookOpen`, not even once events
  are back on, so that workbook's `PyLoad` cells do not calculate. This is a limitation, not
  a fault: turning events off asks that no event handler run, and every add-in that follows
  `WorkbookOpen` shares it. The workaround is a full recalculation (Ctrl+Alt+F9), which
  calculates the `PyLoad` cells, or the `PyRepair` command below; opening the workbook
  again with events on also loads it.

Tests: `tests/host/test_repair.py` (the `PyLoad` cells of a workbook open as the add-in
opens, and of one opened later, calculated once and the rest of its cells left alone; on
every sheet; no workbook open; the events disconnected when the add-in closes);
`tests/excel/test_workbook_open.py` (a workbook saved, then opened in a new Excel: the
script ran, shown by a file it writes, before any input changed or any recalculation was
asked for, the workbook counts as changed, and then a changed input recalculates its
cells; with a literal path, a path in a cell, a `PyLoad` cell on a hidden sheet, a
volatile function in the workbook, and the workbook already open when the add-in loads;
and a workbook opened with events off, missed until Ctrl+Alt+F9 or `PyRepair` loads it).

### The `#NAME?` repair

A formula typed while its function is not registered shows `#NAME?`, and goes on showing
it after the function registers, through every kind of recalculation; entering the formula
again resolves it. Excel keeps the call unresolved even where it does not show: inside
`IFERROR(Discount(A1,B1),0)` the cell shows `0`, and in an `IF` branch not yet taken it
shows the other branch, until the branch is taken and shows `#NAME?`. (Each was found in
real Excel before this was built. A formula saved while its function was registered is
different: opened before the function registers, it comes back on an ordinary
recalculation.)

So after the queue's command registers names Excel has not had from this copy of Pyinex
before, it enters again every formula that uses one, in every worksheet of every open
workbook, hidden sheets included, and then asks Excel to calculate what is out of date
(`xlcCalculateNow`, as F9 does), for cells that depend on those formulas when calculation
is manual:

- **Found by the names they contain**, not among the cells Excel reports as errors
  (`SpecialCells`), which would miss the calls hidden in `IFERROR` and untaken `IF`
  branches. For each name it searches the formulas for the name and `(`, in any case, as
  for the `PyLoad` cells above, and reads only the formulas of the cells found.
- **Only formulas that call a name are entered again.** A cell found is entered again only
  if its formula holds the name whole, followed by `(`, outside a text in double quotes
  and a sheet name in single quotes, and not as the end of a longer name or after `!`:
  `=Rate(1)` calls `Rate`, but `=N("Rate(")`, `=MyRate(1)` and `='Rate(s'!A1` do not.
  Text that merely holds a name, formulas that use none of the names, and the cells a
  spilling formula fills (the search finds only the formula's own cell, which is entered
  again) are left alone. Each workbook with such a formula counts as changed.
- **Every cell is found before any is entered.** Entering a formula can run a workbook's
  change handler, and one that searches (`Range.Find`) would change what the next step of
  Pyinex's own search (`FindNext`) continues.
- **A legacy array formula** (entered with Ctrl+Shift+Enter) cannot be changed one cell at
  a time, so Pyinex enters the whole array again once, with its formula and size
  (`Range.CurrentArray`, `FormulaArray`). `FormulaArray` takes at most 255 characters, so a
  longer legacy array formula is left alone and the console names its cells, to be entered
  again by hand with Ctrl+Shift+Enter (`tests/host/test_repair.py::test_a_legacy_array_formula_too_long_*`).
- **Only new names.** A name registered again, after a changed signature or with the
  add-in closed and opened again in the same Excel, was never unknown to Excel, which kept reading it as a function's;
  its cells come back with an ordinary recalculation
  ([While a function is registered again](#reloading-a-script)), and are left alone.
- **A formula that cannot be entered again**, as on a protected sheet, is reported on the
  console (`Pyinex: cannot enter the formula in [Book1.xlsx]Sheet1!$A$1 again (…)`), and
  the rest are still entered.

**`PyRepair`**, a command, does the same on demand. It enters again every `PyLoad` cell
of every open workbook, so their scripts load, then every formula that uses any name a
direct function holds, then asks Excel to calculate what is out of date. It is the way to
load a workbook opened while events were off, and to repair a cell the console reported
it could not enter again, once the cause is gone. Run while the queue's command is
running, or while the add-in opens or closes, it is done once that run reaches it
([The queue](#the-queue)), and run again while a repair is going, it does nothing: a
workbook's change handler that runs `PyRepair`, as each formula the repair enters changes
a cell, would otherwise have each pass ask for another, and Excel would never get control
back (`tests/host/test_repair.py::test_a_change_handler_that_runs_pyrepair_lets_the_repair_end`;
in real Excel, `tests/excel/test_repair.py::test_a_change_handler_that_runs_pyrepair_lets_the_repair_end`,
a handler connected through COM). Run from inside a
worksheet function, as a script could, it does nothing, leaves nothing queued, and says so
on the console (`Pyinex: PyRepair cannot run inside a worksheet function`).

Tests: `tests/host/test_repair.py` (the cells chosen, with formulas and text the host's
search finds, and the cells left alone, among them the name in a text, a quoted sheet
name and a longer name; names registered again left alone; a formula that cannot be
entered again, or read; a change handler that searches; a legacy array formula; the
calculation after; `PyRepair`, and `PyRepair` run inside a worksheet function);
`tests/excel/test_repair.py` (direct-function test group 8: cells typed before
registration on several sheets, in a formula that also uses other functions, in a spilling
formula, in a legacy array formula, inside `IFERROR`, and in an `IF` branch first taken
after registration; cells that must not be touched, checked as untouched, among them
formulas that only hold the name as text or within a longer name, with a function counting
its calls; a protected sheet reported, then repaired by `PyRepair`).

### pyinex.Range

A direct function's parameter annotated `pyinex.Range` receives the reference the formula
gives, rather than the values in it, so the function can ask where its argument is and how
big it is, and read only the part of it that is used:

```python
import pyinex

@pyinex.func
def Where(r: pyinex.Range):
    return f"{r.address(full=True)}: {r.rows} x {r.columns}"
```

**The formula uses Excel's own reference syntax**, with nothing new: an address (`B2`), a
range (`B2:C3`), a whole column (`D:D`) or row (`5:5`), another sheet (`Data!A1:B2`),
another open workbook (`[Prices.xlsx]Sheet1!A1`), a defined name (`Rates`), or several
areas (`(B2:C3,E5)`). Excel is told the parameter takes a reference: its position in the
type string is `U` where any other is `Q` (`tests/host/test_registration.py`).

**What it gives.** None of these ask Excel, so they work at any time, even after the call
has returned:

- `r.workbook` and `r.sheet`: `"Book.xlsx"` and `"Sheet1"`. Excel names the sheet once,
  as the argument arrives.
- `r.address(r1c1=False, full=False)`: as `pyinex.caller` gives an address. `"B2:C3"`, or
  `"R2C2:R3C3"` with `r1c1=True`; with `full=True` the workbook and sheet come first,
  `"[Book.xlsx]Sheet1!B2:C3"`. A whole column is `"D1:D1048576"`. Several areas are joined
  by commas, `"B2:C3,E5"`.
- `r.rows` and `r.columns`: the reference's size, as given: a whole column has 1,048,576
  rows.
- `r.areas`: a tuple of one `pyinex.Range` for each area, in the formula's order; for a
  reference of one area, a tuple of one range with the same reference.

These ask Excel, so they work only during a call (below):

- `r.values()`: the cells' values, converted as an argument's values are
  ([conversion.md](./conversion.md#ranges)): one cell a plain value, more than one a tuple
  of rows. An empty cell is `None`, and empty text (`=""`) is `""`.
- `r.empty()`: whether each cell is empty, in the same shape: `True` for an empty cell,
  `False` for one holding anything, empty text included.

**A whole column or row reads only its used part.** For a reference spanning every row of
the sheet, `values()` and `empty()` stop at the reference's last row holding anything, and
for one spanning every column at its last such column; a reference holding nothing reads
as its first cell alone, whether it is one whole column or several (`D:E`), one whole row
or several (`5:6`) (`tests/unit/test_funcs_range.py::test_empty_whole_columns_or_rows_read_as_their_first_cell`). So for `D:D`, `values()` reads `D1:D3` when `D3` is the last
cell of column D in use, not a million cells. Pyinex finds that row by asking Excel to
count what is in parts of the column (`COUNTA`, about 20 times, halving the part each
time), which Excel does without handing Pyinex the cells.

**Several areas** have no single block of values: `rows`, `columns`, `values()` and
`empty()` raise `ValueError`, saying to use `r.areas`, whose ranges work as any other.

**Values that are not references.** A formula can give a `pyinex.Range` parameter
something other than a reference; Excel then passes a value, and the argument is refused
as any value its annotation does not take ([conversion.md](./conversion.md#references-pyinexrange)):
the call gives `#VALUE!` and the console says `Pyinex: Where: argument 1 (r): pyinex.Range
takes a reference, not 5.0`. That is so for a number or text typed in the formula, an
array constant, a calculation (`B2+1`), a reference to a sheet since deleted, which Excel
turns into `#REF!`, and a reference to a workbook that is not open, for which Excel passes
the cells' values.

**Recalculation.** A function taking a reference is recalculated as one taking values is:
when a cell its formula refers to changes, and not when any other cell does; it is not
volatile (`tests/excel/test_range.py`). One thing differs. Excel need not calculate the
cells a reference refers to before calling the function, since it passes the reference and
not their values: a full calculation, for one, can reach the function first, whether the
reference is typed in the formula, made by `INDIRECT`, or kept by the script from another
call. When `values()` or `empty()` asks for a cell not yet calculated, Excel answers
`xlretUncalced`, and Microsoft requires the add-in then to call nothing in Excel but
`xlFree` before it returns. Excel then calculates that cell and calls the function again,
and the cell shows the second call's result. So Pyinex:

- raises `RuntimeError` from the `values()` or `empty()` that met it (`pyinex.Range cannot
  ask Excel: a cell this call reads is not yet calculated, and Excel calls the function
  again once it is`);
- refuses every later call into Excel in that call, whether from the function's code, from
  code that caught the exception, from a `finally` block, or from Pyinex itself: each
  `values()`, `empty()`, `pyinex.caller`, `pyinex.sheet` and `pyinex.esc_pressed` raises
  that `RuntimeError` again, naming itself, without asking Excel, while what Excel handed over earlier in
  the call is still freed with `xlFree`;
- returns `#VALUE!` to Excel, whatever the function returned or raised, and writes nothing
  to the console: Excel discards that result. Nothing is written for a PyCall either, nor
  for a script that failed to load as the call reloaded it, nor for a reference argument
  Excel stopped reading, nor for a result that cannot be converted
  (`tests/host/test_range.py::test_a_pycall_excel_stopped_says_nothing`). Pyinex decides the `#VALUE!` only once the
  result is converted and every Python object of the call released, so a stop met by a
  finalizer (`__del__`) of what the function returned counts too. A PyCall's result is
  decided the same way, since its function can read a range kept from a direct function's
  call (`tests/host/test_range.py::test_a_stop_while_the_result_is_freed_*`).

`xlretAbort`, Excel's answer when the calculation is interrupted, is handled the same way,
the message saying the calculation was interrupted. The rule is kept in `src/excel/` by the
one function every call into Excel goes through (`call`), for the length of each call of a
function (`ExcelCall`, `src/core/module.h`).

**When they may be called.** `values()` and `empty()` raise `RuntimeError` outside a call
and on a thread a script started, as `pyinex.caller` does ([When they may be
called](#when-they-may-be-called)). A range kept after its call has returned still gives
its address, workbook, sheet and size, and a later call can read it: it holds its sheet's
ID. Excel may give that ID to another sheet once the sheet is deleted or its workbook
closed, so before each `values()` or `empty()` Pyinex asks Excel for the ID of the sheet
by its name again (`xlSheetId`). If Excel has no such sheet open, or gives another ID, the
read raises `RuntimeError` (`pyinex.Range: Excel no longer has the sheet [Book.xlsx]Sheet1
open; it was deleted or renamed, or its workbook closed`) rather than read another sheet.
A sheet put in the deleted one's place under the same name, in the same workbook, is read
only if Excel gives it the same ID; with another ID the read refuses, as for a deleted
sheet. Microsoft 365 gives it the same ID, so it is read, as a formula naming that sheet
by name would read it (`tests/unit/test_funcs_range.py::test_a_range_whose_sheet_*`,
`tests/host/test_range.py::test_a_kept_range_whose_sheet_is_gone_*`; in real Excel,
`tests/excel/test_range.py::test_a_kept_range_whose_sheet_is_gone_*` and
`test_a_kept_range_whose_sheet_is_replaced_under_its_name_*`).

**Reloading.** A parameter that gains or loses the annotation changes the type string Excel
is told, so the function is registered again ([Reloading a script](#reloading-a-script)),
and until then a call gives `#VALUE!`: a reference never reaches code that expects
values, nor values code that expects a reference.

**Only direct functions.** PyCall's arguments are values, as before: a reference reaches
only a direct function's parameter annotated for it. A `pyinex.Range` only reads cells:
Excel forbids a worksheet function to change them. `pyinex.Range | None` is the same as
`pyinex.Range`. A function returning a `pyinex.Range` gives `#VALUE!`, as for any object
Excel cannot take.

Tests: `tests/unit/test_funcs_range.py` (addresses in both styles, sizes, areas, the used
part's search, empty cells, the refusals of values that are not references, the type
codes registered); `tests/host/test_range.py` (each kind of reference as the fake Excel
passes it, with its sheet; several areas; a value, an error and an array constant refused;
a whole column read without reading the column; each stop code, with no call into Excel
after it but `xlFree`, from caught exceptions and from cleanup, everything Excel handed
over freed, and nothing on the console; a range kept after its call; the refusals on a
script's thread; reloads from a value to a reference and back, with calls before the queue
runs, a registration again that fails, and several edits while queued); in real Excel,
`tests/excel/test_range.py` (every kind of reference above, a workbook not open, a deleted
sheet, a range kept after its call, the refusals, cells not yet calculated giving the stop
and then the right result, the recalculation stated above, and the reloads).

## Choosing a Python

One add-in works with any installed x64 CPython 3.10 or later from python.org, because it
uses only Python's stable ABI: the part of Python's C interface that does not change
between versions. The add-in chooses a `python.exe` the first time it needs Python: at
the first PyCall, PyLoad, direct function or PyInfo item other than `"pyinex"`, or as the
add-in opens when `pyinex.ini` lists scripts.

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

### The core's Python files

Part of the core is written in Python, because some of its work takes less code in Python
than through the C API. That Python is in `.py` files in `src/core/python/`:

- `bootstrap.py`, the short program run once when Python starts (step 3 of
  [How a call moves through the code](#how-a-call-moves-through-the-code));
- `excel_error.py`, which defines the `ExcelError` type;
- `call.py`, PyCall's Python half: loading scripts, calling functions, and printing
  exceptions (step 6);
- `funcs.py`, the Python half of [direct functions](#direct-functions): the
  `@pyinex.func` decorator, and reading a script's marked functions, their signatures and
  docstrings. It runs in `call.py`'s namespace, since it loads and calls scripts with
  `call.py`'s functions.

**Built in, not shipped beside the add-in.** The build compiles each file's text into
`pyinex_core.dll`, and also into the unit tests' `_pyinex_testing.pyd`, which offers it as
`sources`, so the add-in is still two files and no installed copy can be
edited, lost or mismatched with its DLL. The core
compiles each text when Python starts. This is a build step of its own, and the reason for
it is that real files can be read, edited and checked as Python, where a C++ string
constant cannot. It needs nothing beyond CMake: for each file, a CMake script
(`cmake/EmbedPython.cmake`) writes a header in the build tree,
`generated/core/python/<name>.py.h`, holding the file's bytes as a C++ array. A change to a
`.py` file rewrites its header and recompiles the code that includes it. The bytes are not
written as a string literal, which Microsoft's compiler limits to 16,380 characters a piece.

**Left out of tracebacks.** Python compiles each text under a file name that starts with
`<pyinex/`, such as `<pyinex/call.py>`. When a script raises an exception, PyCall prints the
traceback without any frame from such a file, since the user did not write them: those
before the user's frames, those after them (`pyinex.ExcelError` refusing its text, in
`excel_error.py`), and those of chained exceptions. Tests:
`tests/unit/test_call.py::test_exception_becomes_value_error` (every frame shown is the
user's script's), `tests/unit/test_call.py::test_pyinex_s_own_frames_are_left_out_after_the_user_s_too`,
`tests/host/test_exceptions.py`, and `tests/unit/test_core_python.py`, which compiles each
file under every Python the tests run under, so a syntax only newer Pythons accept fails on
CI, checks that the build compiled each in whole, and runs `bootstrap.py` from that text.

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
   ; optional: scripts whose direct functions register when the add-in opens
   scripts = model.py; C:\shared\rates.py
   ```

   `scripts` is described in [Direct functions](#direct-functions).

   The file may be saved as UTF-8, with or without a byte-order mark, or as UTF-16. Pyinex
   reads UTF-8 files itself, with Windows' rules: names are compared without regard to
   case, `;` starts a comment, and quotes around a value are removed. Windows' own reader,
   `GetPrivateProfileStringW`, reads a file with no byte-order mark as the 8-bit code page,
   but editors save UTF-8, so a path such as `C:\Users\Zoë\…` would be read as a different
   path. A UTF-16 little-endian file, with its byte-order mark, is still read by Windows'
   reader; a big-endian one, which Windows does not read, is read by Pyinex's. UTF-16 with
   no byte-order mark cannot be told from 8-bit text, so a file holding a NUL byte is not
   read, and the console says to save it as UTF-8 or as UTF-16 with a byte-order mark.
   Tests: `tests/host/test_python_discovery.py::test_ini_names_a_path_beyond_the_ansi_code_page`,
   `::test_ini_saved_as_utf16_big_endian_is_read`, `::test_ini_in_utf16_with_no_byte_order_mark_is_reported`.

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
| It cannot start (see below) | `it failed to start (exit code N): <why>`, or `it did not finish starting within 60 seconds` |
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
as a separate process, it fails without harm, and the refusal says why: Python's own
`Fatal Python error: …` line with the exception it names, when there is one, rather than the
dump of its threads after it, else the last line of its output that is not part of such a
dump (`tests/host/test_python_discovery.py::test_a_python_that_cannot_start_gives_its_own_reason_not_its_thread_dump`).
The run prints the Python's version first, as any Python can, so even Python 2 is refused
as too old rather than with the error its later lines raise there. Once Python has started
in Excel, Pyinex's own set-up can still fail, as when a module on `PYTHONPATH` hides one
Pyinex imports: the console names the step and Python's error, and the status
(`PyInfo("status")`) says `Python started, but Pyinex could not set it up; the console
says why. Restart Excel once the cause is fixed`, since Python cannot start again in the
same process (`::test_a_set_up_python_cannot_finish_says_why_and_to_restart_excel`). The run takes a fraction of a second, once per Excel session. It cannot detect an
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
  so an Excel session that never calls Pyinex does not load Python, unless `pyinex.ini`
  lists scripts whose direct functions register now (see
  [Direct functions](#direct-functions)).
- **First use:** Pyinex chooses and starts Python (above). It then pins the add-in in
  memory (`GetModuleHandleExW` with `GET_MODULE_HANDLE_EX_FLAG_PIN`), so Windows never
  unloads it. This is needed because Python and the `pyinex` module now hold pointers to
  the add-in's code.
- **Closing** (`xlAutoClose`, which Excel also calls when the user only unticks the add-in
  in the Add-ins dialog): Pyinex removes each direct function it registered, then flushes
  Python's console output. Removing one unregisters it and clears its name, so that no old
  name reaches a function that takes its slot later
  ([Removing a function](#removing-a-function)). A formula that still uses the name then
  gives `#VALUE!`. If any step fails, the console says so, and the slot gives `#VALUE!` and
  takes no other function until Excel restarts. Closing also disconnects from Excel's
  events and drops the workbooks waiting for their `PyLoad` cells to calculate
  ([When a workbook opens](#when-a-workbook-opens)).
- **A quit canceled at the save prompt.** When Excel quits, it calls `xlAutoClose`
  *before* it asks whether to save changed workbooks. If the user cancels there, Excel
  carries on, with the add-in loaded and its direct functions removed. So closing also
  queues each function it removed to be registered again, ahead of any work still queued,
  and leaves the queue's timer running, even with nothing to register again. If Excel
  carries on, the queue registers them again (as Excel-DNA does), each under its own name
  in a free slot, and connects to Excel's events again, handling every workbook open then,
  so a workbook opened after a canceled quit loads its scripts as any other does. An old
  name never reaches another function: each was cleared as it was removed. If instead
  Excel has unloaded the add-in, the queue's command is gone with it. When nothing has
  pinned the add-in, Windows then unloads its code too, and the add-in's `DllMain` destroys
  the queue's window and its timer as it goes, calling only Windows' window functions,
  never Python. Pyinex tells the two apart by asking Excel,
  before each try, to calculate `PyInfo("pyinex")` (`Application.Evaluate`). Only a
  positive answer counts: once Excel has unloaded the add-in, that gives `#VALUE!` or
  `#NAME?`. Pyinex then says on the console how many functions it will not register again,
  and destroys the window and its timer. Any other answer says nothing, and Pyinex asks
  again at the next try: a question Excel refuses, as during the save prompt; `#N/A`, which
  Excel gives while a Protected View window is active; or a path, even another copy's,
  since this copy may still be loaded beside it. So with another copy loaded, work left by
  a copy that was really unloaded waits, its timer retrying, until Excel ends; it never runs
  the other copy's command, whose name is that copy's own. Once the queue's command has run,
  Excel evidently has the add-in loaded, and Pyinex stops asking until the next close.
- **Output at closing.** Code that runs while output is flushed, such as the `flush` method of a replaced
  `sys.stdout`, runs outside any PyCall, so a `pyinex` function that would ask Excel for
  information raises `RuntimeError` (see
  [When they may be called](#when-they-may-be-called)). The interpreter and the script cache
  are left as they are, and Python is **never finalized**. After Python has started, the
  add-in stays loaded because it is pinned. It is also pinned once it has made the console
  window, because the window's code and the thread that sends text to it are in the
  add-in. If neither has happened, the add-in unloads normally.
- **Reopening** (`xlAutoOpen` again): Pyinex registers its functions again, and the direct
  functions of the scripts `pyinex.ini` lists, now, in the slots closing freed, and uses
  the running interpreter and its script cache. Changed scripts reload as usual. A Python that
  was refused is tried again. Reopening starts the queue clean: any work still queued from
  before, and the window and its timer, are dropped, and the console says how many
  functions were not registered. It then connects to Excel's events again, as at the first
  opening, and so calculates the `PyLoad` cells of every workbook open then. Their names are free again, so the functions PyLoad
  registered before register again when their PyLoad cells next calculate.
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

Tests: `tests/host/test_direct_functions.py` and `tests/excel/test_direct_functions.py`
(closing, changing `pyinex.ini`'s list and reopening: a cell using a name from before
gives `#VALUE!`, never the function that took its slot); `tests/host/test_lifetime.py` and
`tests/excel/test_lifetime.py` (closing with work queued and reopening: no slot, timer or
queued work left from before; a quit canceled at the save prompt: the functions work again
without reopening, and an old name reaches no other function; in real Excel, with nothing
registered, a workbook opened after the canceled quit loads its script);
`tests/host/test_repair.py` (closing disconnects from Excel's events, and reopening
connects again, as does a canceled quit with nothing registered); and
`tests/host/test_lifetime.py` also (a Protected View window active after a canceled quit,
not taken for an unload; work left waiting, and the other copy's command not run, when
unloaded while another copy of Pyinex is loaded; open, call, close, reopen, call; the
add-in still in memory after the host unloads it, which shows the pin works; refusals
after closing; no Python during unloading; handle and reference counts unchanged across
many calls and reloads). In real Excel, `tests/excel/test_lifetime.py` (unloading the
add-in with Excel's `UNREGISTER`, whose `xlAutoClose` flushes output outside any call, and
loading it again with the same interpreter and module state) and
`tests/excel/test_scale.py` (handles and private memory across 10,000 calls).

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
  that nobody reads would stop Excel once it was full. A character that one read of the
  pipe cuts off waits for the next; only one that can still be finished waits, judged by its
  last bytes alone, so a stray byte that is not UTF-8, as C code may write, shows as a
  replacement character and never holds back the text around it
  (`tests/host/test_console.py::test_a_byte_that_is_not_utf8_never_holds_back_the_line_it_ends`).
- **UTF-8, one line at a time.** At start-up, Pyinex sets `sys.stdout` and `sys.stderr`
  to UTF-8 and to write each line as soon as it ends. So any character can be printed, and
  `print()` from a thread that a script started appears without waiting for a buffer to
  fill. Pyinex's own messages are also written in UTF-8. After each PyCall, Pyinex also
  flushes both streams.
- **A host that already has output keeps it.** When the process already has standard
  output and standard error, as when Excel is started from a terminal, or in the fake Excel
  used by the tests, which captures the output, Pyinex leaves them alone and Python writes
  there. The window still exists for `PyConsole`, but receives no text. A process with only
  one of the two keeps it, and the other goes to the window, so Pyinex's own messages,
  written to standard error, are never lost.
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
  of Excel's grid. A NumPy array returned is not converted; only NumPy's `float64`, a
  subclass of `float`, is accepted as a number ([conversion.md](./conversion.md)). An
  argument annotated `numpy.ndarray` arrives as one, made with the NumPy the script
  imported; Pyinex never imports NumPy itself. CI installs NumPy for `tests/unit` and
  `tests/host`, whose tests of that annotation fail without it rather than skip.

## How a call moves through the code

This section lists what happens, file by file, when a cell calculates
`=PyCall("model.py", "f", A1)`. The other four functions take steps 1 and 2 (PyLoad, and PyInfo for any item
but `"pyinex"`, also take step 3) and then do their own small amount of work in
`src/loader/addin.cpp`; PyLoad's is in `src/loader/functions.cpp`, which hands the
script's functions to [the queue](#the-queue).

1. **Opening.** Excel loads `pyinex.xll` and calls `xlAutoOpen` (`src/loader/addin.cpp`),
   which registers the five functions from the table at the top of that file through
   `register_functions` (`src/excel/excel.cpp`), the two commands, and then the direct
   functions of any scripts `pyinex.ini` lists (`src/loader/functions.cpp`; see
   [Direct functions](#direct-functions)). A direct function's call is described after
   these steps.
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
   - finds the core's exported functions (`src/core/api.h`);
   - pins the add-in in memory;
   - calls the core's `pyinex_start` (`src/core/core.cpp`).

   `pyinex_start` starts Python and runs a short Python program
   (`src/core/python/bootstrap.py`) that corrects `sys._base_executable` and `sys.path`
   and sets the output streams to UTF-8. It then creates the `ExcelError` type
   (`src/core/convert.cpp`, from `src/core/python/excel_error.py`), PyCall's and direct
   functions' Python code (`src/core/call.cpp`, from `src/core/python/call.py` and
   `funcs.py`) and the `pyinex` module (`src/core/module.cpp`), and releases the GIL. Each `.py` file is compiled into the
   core ([The core's Python files](#the-cores-python-files)). If Python was refused, the refusal is kept, and every later call returns
   `#VALUE!` at once.
4. **Into the core.** `pyinex_call` (`src/core/core.cpp`) takes the GIL, records that Excel
   is calling (so that the `pyinex` module may ask Excel for information), and calls
   `core::call` (`src/core/call.cpp`).
5. **Arguments.** `core::call` drops the missing arguments at the end and converts the rest
   with `to_python` (`src/core/convert.cpp`). It asks Excel for the workbook's folder only
   when the path is relative (`caller_folder`, `src/excel/excel.cpp`: `xlfCaller`, then
   `CELL("filename")`).
6. **The script and the function.** The rest of PyCall is written in Python, in
   `src/core/python/call.py`, compiled into the core and run at start-up by
   `src/core/call.cpp`, because loading files and catching every exception take less code
   in Python than through the C API. Its `call`
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

**A direct function's call**, such as `=Discount(A1, B1)`, takes the same steps, with
these differences:

- **Step 2.** Excel calls the numbered export registered under the name, such as
  `PyFunc7`, generated from `PYINEX_SLOT` (`src/loader/functions.cpp`), which hands its
  slot to `dispatch`. In the Function Arguments dialog, `dispatch` returns a blank. It
  reads only as many arguments as the slot's count, the count Excel was told.
- **Step 4.** `dispatch` calls the core's `pyinex_function` (`src/core/core.cpp`) with the
  slot's key, the name and the arguments, which takes the GIL and records the call as
  `pyinex_call` does.
- **Step 5.** `call_function` (`src/core/call.cpp`) drops the arguments left out at the
  end, converts the rest with `to_python`, makes a reference a `pyinex.Range`, and marks
  each one left out before the last one given, so that its parameter's default applies.
- **Step 6.** `call_function` in `src/core/python/funcs.py` finds the script and the
  function by the key, loads the script again if it changed (`_load`, PyCall's cache),
  and refuses a function that a reload changed or removed. It places the arguments by the
  function's signature, converts them by their annotations, and calls the function. Any
  failure prints its reason, and its traceback without Pyinex's frames.
- **After step 7.** Back in `dispatch`, `take_changes` hands anything a reload changed to
  [the queue](#the-queue), which registers or removes functions after the calculation.

**Closing.** `xlAutoClose` removes every direct function and clears its name
(`close_functions`, `src/loader/functions.cpp`; [Lifetime](#lifetime)), then calls the
core's `pyinex_close`, which flushes Python's output. Python itself keeps running.

The tests divide the code at the same points: `tests/unit` calls the code of steps 5 to 7
directly, through an extension module built from the same source files, and `tests/host`
runs steps 1 to 8 in a fake Excel (see [testing.md](./testing.md)).

## Not yet built

These are possible later work:

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
