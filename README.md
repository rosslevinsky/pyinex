# Pyinex

Pyinex lets an Excel formula call a Python function:

```
=PyCall("model.py", "discount", A1, B1)
```

or, once a script marks the function `@pyinex.func`, by the function's own name:

```
=Discount(A1, B1)
```

Pyinex is an Excel add-in (an XLL) that runs Python inside Excel's own process. This has
two consequences:

- **Calls are fast and need no setup.** The usual way to link Excel and Python runs Python
  in a separate COM server, and every call crosses from one process to the other. Pyinex
  needs only the add-in and an installed Python.
- **The code is in ordinary files.** A script is a `.py` file beside the workbook, not VBA
  stored inside it. You can read it, test it, and keep it under version control like any
  other code.

Pyinex is deliberately small: five worksheet functions of its own, one Python module, and
under 3,000 lines of C++. It was first written in 2009–2010 and rewritten in 2026 for
current Excel and Python.

## Requirements

- Windows, with 64-bit Excel. Pyinex is tested with Microsoft 365. It uses only the part of
  Excel's C interface that 64-bit Excel has had since Excel 2010, so other 64-bit versions
  should also work. Windows on Arm works, because Excel there loads x64 add-ins.
- Microsoft's Visual C++ runtime, x64. Microsoft 365 includes it. With another version of
  Excel, if Excel cannot load `pyinex.xll`, install Microsoft's Visual C++ Redistributable
  (x64).
- 64-bit (x64) CPython 3.10 or later, from
  [python.org](https://www.python.org/downloads/windows/). Pyinex is tested with 3.10 to
  3.15. One add-in works with all of them.

## Installing

1. Install an x64 Python from python.org, and any packages your scripts need
   (`py -m pip install ...`). Pyinex finds it: see
   [Choosing a Python](#choosing-a-python) to choose another.
2. Get `pyinex.xll` and `pyinex_core.dll`: download the zip from the repository's
   Releases page, or build them (see [docs/testing.md](./docs/testing.md#building)). Keep
   the two files in the same folder. The first file loads the second, so they must stay
   together.
3. If you downloaded the zip, unblock it before you extract it: in its Properties, tick
   **Unblock**, or run `Unblock-File` on it in PowerShell. Excel refuses a blocked add-in.
   To check that a download is the one this repository's release workflow built, run
   `gh attestation verify <file> -R rosslevinsky/pyinex` with the
   [GitHub CLI](https://cli.github.com/) (releases from 2026.10.1 on).
4. In Excel, choose **File → Options → Add-ins**, choose **Excel Add-ins** under Manage,
   click **Go**, then click **Browse** and select `pyinex.xll`.
5. Type `=PyInfo("status")` into a cell. It shows `running` once Python has started. If it
   shows a reason instead, the reason says what failed. Usually Pyinex found no Python it
   could use (see [Choosing a Python](#choosing-a-python)); `cannot load …pyinex_core.dll`
   means that file is missing from the add-in's folder or is still blocked (steps 2 and 3).
   `=PyInfo("executable")` shows which `python.exe` Pyinex chose. If you have several
   Pythons, install your scripts' packages with that one: `<that path> -m pip install ...`.

To build the add-in yourself, see [docs/testing.md](./docs/testing.md#building).

## Using it

Save a workbook. In the same folder, save a script named `model.py`:

```python
def discount(price, rate):
    return price * (1 - rate)
```

Then type this in a cell:

```
=PyCall("model.py", "discount", A1, B1)
```

A relative path, such as `model.py` or `scripts\model.py`, is relative to the workbook's
folder, so you can move the workbook and its scripts together. A full path also works. A
workbook that has no folder needs the full path: one never saved, or one opened from
OneDrive or SharePoint, whose location is a web address.

**Editing a script.** After you save a change to `model.py`, the next time a cell that uses
it calculates, Pyinex loads the new version. Excel does not know that the file changed, and
recalculates a cell only when the cell's inputs change, so press Ctrl+Alt+F9 to recalculate
every formula. A reload resets the script's module-level variables.

**Other files a script imports.** A script can `import` another file in its own folder, as
it could when run with `python model.py`. Unlike separate runs of `python`, every script in
one Excel shares Python's imports: once one script imports `helpers`, every other script
that imports `helpers` gets that same module, even with a different `helpers.py` beside
it. Give unrelated workbooks' helper files names of their own. Pyinex reloads only the
scripts named in `PyCall` or loaded by `PyLoad`. To use a changed imported file, restart
Excel.

**Seeing output.** `print()` output, and the traceback of each failing call, go to
Pyinex's console window. The window is hidden until `=PyConsole(TRUE)` shows it. It uses
Windows' light or dark mode.

**Ranges and blocks.** A range passed to a function arrives as a tuple of rows. A function
that returns a list of rows fills a block of cells. In Microsoft 365, a formula that
returns a block spills into the cells around it. In an Excel without spilling, select the
cells the block should fill, type the formula, and press Ctrl+Shift+Enter; only the cells
selected show the result.

```python
def transpose(cells):
    return [list(column) for column in zip(*cells)]
```

**NumPy and pandas** arrays and frames are not converted. Convert a result to lists before
returning it (`array.tolist()`, `frame.values.tolist()`), and make an array from the rows
you receive (`numpy.array(cells)`), or annotate a direct function's parameter
`numpy.ndarray` (below). Of single NumPy numbers, only `numpy.float64` is a kind of Python
`float` and is returned as one; any other, such as `numpy.int64` or `numpy.float32`, gives
`#VALUE!`, so return `int(n)`, `float(n)` or `n.item()`.

## Functions under their own names

A script can make its functions worksheet functions of their own. Mark each with
`@pyinex.func`:

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

Then load the script from any cell, once:

```
=PyLoad("model.py")
```

and use the function by its name in any cell of any open workbook:

```
=Discount(A1, B1)
```

A direct function is PyCall with the script and the function fixed: the same values, the
same `#VALUE!` and console message for every failure, the same reloading of a changed
script. What it adds:

- **Insert Function knows it.** Its argument names come from the signature, its help from
  the docstring's first paragraph, and each argument's help from an `Args:` section. The
  decorator can set the name Excel uses, the category and the help:
  `@pyinex.func(name="DISCOUNT", category="Pricing", help="...")`. The default category is
  "Pyinex".
- **Any signature that fits Excel's 30 arguments.** Positional parameters come first, then
  keyword-only ones, then `*args`, which takes the rest. So for
  `def f(x, *items, scale=1)`, `=f(10, 2)` calls `f(10, scale=2)`. An argument left out
  takes its default.
- **Annotations convert arguments.** A parameter annotated `int`, `float`, `str`, `bool`,
  a list of rows (`list[list[float]]`) or `numpy.ndarray` gets its value converted, or the
  call gives `#VALUE!` and the console says which argument did not fit. `int` refuses
  `2.5`; `bool` refuses the text `"FALSE"`. An argument without an annotation arrives as it
  would through PyCall.
  [docs/conversion.md](./docs/conversion.md#type-annotations-direct-functions) has the
  table.
- **`pyinex.Range` passes a reference.** A parameter annotated `pyinex.Range` receives the
  reference itself, not its values, written in the formula as any reference is (`B2:C3`,
  `D:D`, `Data!A1`, a defined name):

  ```python
  @pyinex.func
  def Where(cells: pyinex.Range):
      return f"{cells.address(full=True)}: {cells.rows} x {cells.columns}"
  ```

  It gives its `workbook`, `sheet`, `address()` (as `pyinex.caller` gives one), `rows`,
  `columns` and `areas`. During the call, `values()` reads its cells, and `empty()` tells
  an empty cell from empty text. A whole column reads only its used part. It only reads:
  Excel lets a worksheet function change no cell.

**Loading.** `PyLoad(file)` loads the script and returns how many of its marked functions
it accepted, such as `2 functions` ([PyLoad](#pyloadfile) has the details). A relative path
is relative to the workbook's folder, as for PyCall. Excel lets an add-in add a function
only between calculations, so the functions register a
moment after the cell calculates, and the console names each one. A script that fails to
load registers nothing; fix it, and the next calculation of the `PyLoad` cell loads it.

**When the workbook opens again**, even in a new Excel, Pyinex calculates its `PyLoad`
cells, so the scripts load before you change anything. Excel then counts the workbook as
changed, and asks to save it when it closes. A cell typed while its function was not yet
registered shows `#NAME?`; once the function registers, Pyinex enters such formulas again,
so they calculate. The `PyRepair` command does both on demand (below).

**Loading when Excel starts.** To have scripts' functions in every workbook from the start,
list the scripts in `pyinex.ini` beside the add-in
([Choosing a Python](#choosing-a-python)), separated by semicolons. A relative path is
relative to the add-in's folder:

```ini
[python]
scripts = model.py; C:\shared\rates.py
```

Python then starts when the add-in opens, rather than at the first call.

**Editing the script.** With `PyReloadOnChange` on, a changed function body takes effect
at the next calculation, and so does any other change Excel is not told of, such as a
default value, or a parameter made keyword-only with the same arguments in Excel. A change
to what Excel is told (the number of arguments, their names or types, the name's spelling,
the help or the category), a removed function, a renamed one or a new one is registered or
removed a moment later. Until then, a call of a function being registered again gives
`#VALUE!`, and the cell keeps it after the registration: recalculate (Ctrl+Alt+F9) or
change one of its inputs once the console says the function is registered.

**Names** belong to all of Excel, not to one workbook, and are compared without regard to
case. The first script to register a name keeps it, and Pyinex refuses, on the console, a
name another script, Excel itself (`SUM`) or another add-in already uses. A name stays
registered after the workbook that loaded its script closes, until Excel exits or the
add-in is unticked in the Add-ins dialog, so another workbook whose script marks the same
name gets the first script's function. Give the functions of different scripts different
names.

[docs/architecture.md](./docs/architecture.md#direct-functions) has every rule.

## Functions

All five functions are in Excel's function category "Pyinex".

### PyCall(file, function, arg1, …, arg30)

Loads the script `file`, calls `function` in it with up to 30 arguments, and returns the
result.

- `function` names anything callable at the top level of the script: a function, a class,
  or an object such as `max` assigned to a name.
- Pyinex drops the arguments left out at the end, so Python's default values and `*args`
  apply. An argument left out in the middle, as in `=PyCall("s.py", "f", A1, , C1)`,
  arrives as `None`.
- Arguments are passed by position only. A cell holding `rate=0.1` is text, not a keyword
  argument.
- Every failure gives `#VALUE!` and writes the reason to the console. Failures include a
  missing script or function, an exception, and a result that Excel cannot hold.

### PyConsole(show, [clear], [x], [y], [width], [height])

Shows the console when `show` is TRUE, hides it when `show` is FALSE, and returns `show`.

- When `clear` is TRUE, the console's text is cleared each time the formula calculates. Set
  `clear` back to FALSE afterwards. Otherwise the console is cleared on every
  recalculation.
- `x`, `y`, `width` and `height` set the window's position and size in pixels each time the
  window changes from hidden to shown. Give `x` together with `y`, and `width` together with
  `height`. If you leave them out, the window stays where it was.
- Closing the console window hides it. `PyConsole(TRUE)` shows it again.

### PyReloadOnChange([on])

When `on` is TRUE (the default), Pyinex reloads a script when its file changes. When `on`
is FALSE, Pyinex loads each script once and keeps it, which saves checking the file on
every call. With no argument, the function changes nothing. It returns the current setting.

### PyLoad(file)

Loads the script `file` and registers the functions it marks `@pyinex.func` (see
[Functions under their own names](#functions-under-their-own-names)). It returns the number
of those functions it accepted as text, such as `2 functions` (a function refused as the
script is read is not counted, and Excel may still refuse one it counts), or `#VALUE!` if
the script fails to load, with the reason on the console.

### PyInfo([item])

Reports which Python and which add-in are loaded:

| `item` | Returns |
|---|---|
| `"status"` | `running`, or the reason Python could not start |
| `"python"` | the path of the Python DLL in use |
| `"pyinex"` | the path of the add-in |
| `"version"` | Python's `sys.version` |
| `"executable"` | the `python.exe` Pyinex chose |
| `"prefix"` | `sys.prefix` |
| `"venv"` | TRUE when that Python is a virtual environment |
| `"addin_version"` | the add-in's version |

With no `item`, it returns all of them as a table of two columns.

### The PyRepair command

`PyRepair` is a command, not a worksheet function: run it from the Macro dialog (Alt+F8,
type `PyRepair`, then Run; the dialog does not list it) or from VBA's
`Application.Run "PyRepair"`. It calculates the
`PyLoad` cells of every open workbook, so their scripts load, and enters again every
formula that calls a direct function, so cells left at `#NAME?` calculate. A workbook
opened while a macro had turned Excel's events off is not loaded by itself; `PyRepair`, or
Ctrl+Alt+F9, loads it.

### In the Function Arguments dialog

Excel's Function Arguments dialog (opened from Insert Function) calculates a formula while
you type it. In that dialog, Pyinex's functions and direct functions return a blank and do
nothing else, so typing does not run your scripts. Pyinex recognizes the dialog by its
English title. In Excel in another language, the functions run in the dialog as they do in
a cell.

## How values cross

| In Excel | In Python |
|---|---|
| a number | `float`, always, even `3` (use `int(x)` where you need an integer) |
| text | `str` |
| TRUE / FALSE | `bool` |
| an empty cell | `None` |
| an error such as `#N/A` | `pyinex.ExcelError` |
| a range of two or more cells | a tuple of rows, each a tuple of values: `((1.0, 2.0), (3.0, 4.0))` |

A range of two or more cells is always a tuple of rows, whatever its shape, so
`cells[r][c]` works for one row, one column, or a block. A single cell arrives as a plain
value.

| Python returns | Excel shows |
|---|---|
| `bool`, `int`, `float`, `str` | the value; NaN and infinity show `#NUM!` |
| `None` | nothing (empty text) |
| `pyinex.ExcelError("#N/A")` | that error |
| a flat list or tuple | one row |
| a list of lists or tuples | a block; short rows are filled out with blanks |
| anything else (`dict`, a NumPy array, ...) | `#VALUE!`, with the reason on the console |

[docs/conversion.md](./docs/conversion.md) lists every rule.

## The pyinex module

Scripts can `import pyinex`. The module is built into the add-in, so it exists only inside
Excel.

- `pyinex.caller(r1c1=False, full=False)` returns the calling cell's address: `C12`, or
  `R12C3` with `r1c1=True`, or `[Book.xlsx]Sheet1!C12` with `full=True`.
- `pyinex.sheet()` returns the calling cell's workbook and sheet: `[Book.xlsx]Sheet1`.
- `pyinex.esc_pressed(clear=False)` returns True once the user has pressed Esc during the
  current calculation. Pyinex cannot stop Python code that is running, so a long function
  should check this and return early.
- `pyinex.ExcelError` is the type that Excel errors arrive as. Returning one puts that
  error in the cell.
- `@pyinex.func` and `pyinex.Range`: see
  [Functions under their own names](#functions-under-their-own-names).

`caller`, `sheet`, `esc_pressed`, and a `pyinex.Range`'s `values()` and `empty()`, work
only while Excel is calculating a `PyCall` or a direct function, and only on Excel's own
thread. Anywhere else they raise `RuntimeError`: in a script's top level as `PyLoad` or
`pyinex.ini` loads it, for one.

**Running a script outside Excel.** To test a script with plain `python`, give it a
stand-in for the module when the real one is missing:

```python
try:
    import pyinex
except ImportError:
    class pyinex:
        @staticmethod
        def caller(r1c1=False, full=False):
            return "A1"
```

## Choosing a Python

Pyinex uses the first of these that names a Python:

1. The environment variable `PYINEX_PYTHON`, holding the path of a `python.exe` or of the
   folder that contains it.
2. A file named `pyinex.ini` in the add-in's folder:

   ```ini
   [python]
   executable = C:\Users\me\venvs\finance\Scripts\python.exe
   ```

   Save it as UTF-8 (Notepad's default), or as UTF-16 with a byte-order mark (Notepad's
   "UTF-16 LE" and "UTF-16 BE"). The path may contain any character.
   To find a Python's path, run `py -3.14 -c "import sys; print(sys.executable)"` with
   its version.

3. The first `python.exe` on `PATH`, skipping any under `Microsoft\WindowsApps`, where
   the Microsoft Store's placeholder and app aliases are.
4. The newest x64 Python that python.org's installer recorded in the registry. The
   installer does not put Python on `PATH` unless you ask it to, so this is usually how
   Pyinex finds a newly installed Python.

A virtual environment's `python.exe` works as that environment: its own packages import,
and the base Python's packages do not, unless the environment was made with
`--system-site-packages`.

Pyinex refuses a Python that is too old, is not an x64 build, is a free-threaded build, or
fails to start. It writes the reason to the console and to `PyInfo("status")`, and every
`PyCall` then gives `#VALUE!`. After you fix the cause, reopen the add-in (untick it and
tick it again in the Add-ins dialog) or restart Excel. One refusal needs Excel restarted:
`Python started, but Pyinex could not set it up`, as when a module on `PYTHONPATH` hides
one Pyinex imports. The console says which step failed and why; Python, once started,
cannot start again in the same Excel.

## Limits

- **Python runs on Excel's calculation thread.** While a function runs, Excel waits. Keep
  calls short, and check `pyinex.esc_pressed()` in long ones.
- **A call that opens its own window waits for that window to close.** For example,
  `matplotlib.pyplot.show()` and a Tk `mainloop()` do not return until their window
  closes, and Excel waits for them. Save a chart to a file instead.
- **Python is never shut down.** It starts at the first call, or when the add-in opens if
  `pyinex.ini` lists scripts, and runs until Excel exits.
  To start a fresh Python, or to replace the add-in's files after Python has started,
  restart Excel. `atexit` handlers do not run, so close any file a script writes.
- **One Python per Excel process.** Pyinex refuses to start if another add-in has already
  loaded a Python into Excel.
- **Open only workbooks you trust.** A formula can call any Python script on your
  computer, so a workbook with Pyinex formulas runs code, as a workbook with macros does.

## Examples

The [`examples/`](./examples) folder holds five workbooks, each with its script beside it.
On every sheet, each formula's result is next to the value it should show.

| Workbook | Shows |
|---|---|
| `basics.xlsx` | calling functions, arguments, single values and errors |
| `ranges.xlsx` | ranges in, blocks out, and matrix arithmetic |
| `pyinex_module.xlsx` | `pyinex.caller`, `pyinex.sheet` and `pyinex.esc_pressed` |
| `addin.xlsx` | `PyInfo`, `PyReloadOnChange`, `PyConsole` and `print()` |
| `direct_functions.xlsx` | `PyLoad`, functions called by their own names, annotations and `pyinex.Range` |

Formulas that return a block are entered as legacy array formulas (shown in braces), so
that they work in every Excel version.

## Documentation

- [docs/architecture.md](./docs/architecture.md): how Pyinex works and why, for anyone
  changing it.
- [docs/conversion.md](./docs/conversion.md): every rule for converting values.
- [docs/testing.md](./docs/testing.md): building, and the three kinds of tests.
- [AGENTS.md](./AGENTS.md): rules for working on the code.
- [CHANGELOG.md](./CHANGELOG.md): what changed in each version.

## License

BSD 3-Clause; see [LICENSE](./LICENSE) and [NOTICE](./NOTICE).
