# Pyinex

Pyinex lets an Excel formula call a Python function:

```
=PyCall("model.py", "discount", A1, B1)
```

Pyinex is an Excel add-in (an XLL) that runs Python inside Excel's own process. This has
two consequences:

- **Calls are fast and need no setup.** The usual way to link Excel and Python runs Python
  in a separate COM server, and every call crosses from one process to the other. Pyinex
  needs only the add-in and an installed Python.
- **The code is in ordinary files.** A script is a `.py` file beside the workbook, not VBA
  stored inside it. You can read it, test it, and keep it under version control like any
  other code.

Pyinex is deliberately small: four worksheet functions, one Python module, and under 2,500
lines of C++. It was first written in 2009–2010 and rewritten in 2026 for current Excel and
Python.

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
4. In Excel, choose **File → Options → Add-ins**, choose **Excel Add-ins** under Manage,
   click **Go**, then click **Browse** and select `pyinex.xll`.
5. Type `=PyInfo("status")` into a cell. It shows `running` once Python has started. If it
   shows a reason instead, the reason says what failed. Usually Pyinex found no Python it
   could use (see [Choosing a Python](#choosing-a-python)); `cannot load …pyinex_core.dll`
   means that file is missing from the add-in's folder or is still blocked (steps 2 and 3).

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
folder, so you can move the workbook and its scripts together. A full path also works.

**Editing a script.** After you save a change to `model.py`, the next time a cell that uses
it calculates, Pyinex loads the new version. Excel does not know that the file changed, and
recalculates a cell only when the cell's inputs change, so press Ctrl+Alt+F9 to recalculate
every formula. A reload resets the script's module-level variables.

**Other files a script imports.** A script can `import` another file in its own folder, as
it could when run with `python model.py`. Pyinex reloads only the scripts named in
`PyCall`. To use a changed imported file, restart Excel.

**Seeing output.** `print()` output, and the traceback of each failing call, go to
Pyinex's console window. The window is hidden until `=PyConsole(TRUE)` shows it. It uses
Windows' light or dark mode.

**Ranges and blocks.** A range passed to a function arrives as a tuple of rows. A function
that returns a list of rows fills a block of cells. In Microsoft 365, a formula that
returns a block spills into the cells around it.

```python
def transpose(cells):
    return [list(column) for column in zip(*cells)]
```

**NumPy and pandas** values are not converted. Convert a result to lists before returning
it (`array.tolist()`, `frame.values.tolist()`), and make an array from the rows you
receive (`numpy.array(cells)`).

## Functions

All four functions are in Excel's function category "Pyinex".

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

### In the Function Arguments dialog

Excel's Function Arguments dialog (opened from Insert Function) calculates a formula while
you type it. In that dialog, all four functions return a blank and do nothing else, so
typing does not run your scripts. Pyinex recognizes the dialog by its English title. In
Excel in another language, the functions run in the dialog as they do in a cell.

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

`caller`, `sheet` and `esc_pressed` work only while Excel is calculating a Pyinex formula,
and only on Excel's own thread. Anywhere else they raise `RuntimeError`.

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

   Save it as UTF-8 (Notepad's default) or UTF-16. The path may contain any character.
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
tick it again in the Add-ins dialog) or restart Excel.

## Limits

- **Python runs on Excel's calculation thread.** While a function runs, Excel waits. Keep
  calls short, and check `pyinex.esc_pressed()` in long ones.
- **A call that opens its own window waits for that window to close.** For example,
  `matplotlib.pyplot.show()` and a Tk `mainloop()` do not return until their window
  closes, and Excel waits for them. Save a chart to a file instead.
- **Python is never shut down.** It starts at the first call and runs until Excel exits.
  To start a fresh Python, or to replace the add-in's files after Python has started,
  restart Excel. `atexit` handlers do not run, so close any file a script writes.
- **One Python per Excel process.** Pyinex refuses to start if another add-in has already
  loaded a Python into Excel.
- **Open only workbooks you trust.** A formula can call any Python script on your
  computer, so a workbook with Pyinex formulas runs code, as a workbook with macros does.

## Examples

The [`examples/`](./examples) folder holds four workbooks, each with its script beside it.
On every sheet, each formula's result is next to the value it should show.

| Workbook | Shows |
|---|---|
| `basics.xlsx` | calling functions, arguments, single values and errors |
| `ranges.xlsx` | ranges in, blocks out, and matrix arithmetic |
| `pyinex_module.xlsx` | `pyinex.caller`, `pyinex.sheet` and `pyinex.esc_pressed` |
| `addin.xlsx` | `PyInfo`, `PyReloadOnChange`, `PyConsole` and `print()` |

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
