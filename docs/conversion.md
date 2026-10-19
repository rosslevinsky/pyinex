# Value conversion

This document gives the rules for converting values between Excel and Python. It is the
only full statement of them: the code follows it, `tests/unit/test_conversion.py` checks
every rule, and the README summarizes it. Each rule names the tests that check it.

## Excel to Python

### Single values

| In the cell | Arrives in Python as | Notes |
|---|---|---|
| A number | `float` | Always `float`, whole numbers too: `3` arrives as `3.0`. Excel stores every number as a double, and a type that changed with the value (`int` for `3`, `float` for `3.5`) would make code behave differently on different data. Convert with `int(x)` where an integer is needed. |
| Text | `str` | Excel text is UTF-16; it decodes to `str` with surrogate pairs combined, so characters outside the Basic Multilingual Plane (emoji, for example) arrive as one character. An unpaired surrogate arrives unchanged as that code point and goes back unchanged. |
| Empty text (`=""`) | `""` | Distinct from an empty cell. |
| `TRUE` / `FALSE` | `True` / `False` | |
| An Excel error | `pyinex.ExcelError` | See [Excel errors](#excel-errors). |
| An empty cell | `None` | |

Tests: `test_conversion.py::test_number_*`, `::test_text_*`, `::test_bool`, `::test_empty_*`;
in real Excel, `tests/excel/test_values_in.py`. (A worksheet cannot hold an unpaired
surrogate, so only the unit tests reach that rule.)

### Ranges

A single cell arrives as a plain value (above). **Any larger range arrives as a tuple of
rows, each row a tuple of values**, whatever the range's shape:

| Range | Arrives as |
|---|---|
| `A1` | `1.0` |
| `A1:C1` (one row) | `((1.0, 2.0, 3.0),)` |
| `A1:A3` (one column) | `((1.0,), (2.0,), (3.0,))` |
| `A1:B2` (block) | `((1.0, 2.0), (3.0, 4.0))` |

So `len(x)` is the number of rows and `len(x[0])` the number of columns, for every range.
A function written for a block works unchanged when the user selects one row or one
column. Empty cells inside a range are `None`. Array constants typed into a formula, such
as `{1,2;3,4}`, follow the same rule.

An argument Excel cannot turn into values — a reference to several separate areas, for
example `(A1:B2,D4)` — arrives as `pyinex.ExcelError("#VALUE!")`.
A direct function's parameter annotated `pyinex.Range` receives the reference itself
instead, several areas included ([References](#references-pyinexrange)).

Tests: `test_conversion.py::test_range_*`; in real Excel,
`tests/excel/test_values_in.py::test_a_range_*`, `::test_an_array_constant_*` and
`::test_a_reference_to_several_areas_*`.

### Missing arguments

For PyCall, missing arguments at the end are dropped, and an argument left out in the
middle arrives as `None`, the same as an empty cell
([Argument rule](./architecture.md#argument-rule)). A direct function gives any argument
left out its parameter's default, in the middle too; one with no default is `None` before
the last argument given, and not passed after it
([Signatures](./architecture.md#direct-functions)).

### Excel errors

An Excel error arrives as an instance of `pyinex.ExcelError`, not as text, so a function
can tell the error `#N/A` from a cell containing the text "#N/A":

- `str(e)` and `e.text` give the error's text: one of `#NULL!`, `#DIV/0!`, `#VALUE!`,
  `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#GETTING_DATA`.
- `pyinex.ExcelError("#N/A")` makes one; any other text raises `ValueError`.
- Two are equal when their text is equal. An `ExcelError` never equals a `str`. It is
  hashable.
- `repr(e)` is `ExcelError('#N/A')`.
- **Returning one puts that error in the cell** (see below), so a function can pass an
  error through or produce one deliberately.

Any error code or value type Pyinex does not recognize arrives as
`pyinex.ExcelError("#VALUE!")`, and the console says what was received. Pyinex never passes
Python a missing (NULL) value. Real Excel never sends one: it passes an add-in the errors
newer than the Excel SDK, such as `#SPILL!` and `#CALC!`, as `#VALUE!` itself, so they
arrive as `pyinex.ExcelError("#VALUE!")` with no message.

Tests: `test_conversion.py::test_error_*`, `::test_unrecognized_*`; in real Excel,
`tests/excel/test_values_in.py::test_an_error_*`.

### Type annotations (direct functions)

A [direct function](./architecture.md#direct-functions)'s parameter may carry a type
annotation that asks for its values converted or checked before the call. PyCall ignores
annotations, and so does a direct function for a parameter without one: its values arrive
exactly as above. Each value is first what the rules above make it (a number a `float`, a
range a tuple of rows, and so on); the annotation then works on that.

**What each value becomes.** "Refused" means the call gives `#VALUE!`, the function is not
called, and the console names the function, the argument and the value, with the rule it
broke: `Pyinex: Discount: argument 2 (years): int takes a whole number, not 2.5`.

| In the cell | `int` | `float` | `str` | `bool` |
|---|---|---|---|---|
| A whole number, `3` | `3` | `3.0` | refused | refused |
| A fraction, `2.5` | refused | `2.5` | refused | refused |
| Text, such as `"3"` or `"FALSE"` | refused | refused | the text | refused |
| Empty text (`=""`) | refused | refused | `""` | refused |
| `TRUE` / `FALSE` | refused | refused | refused | `True` / `False` |
| An empty cell | `0` | `0.0` | `""` | `False` |
| An Excel error | refused | refused | refused | refused |
| A range of more than one cell | refused | refused | refused | refused |

- **Text is never read as a number or a logical value.** `"3"` is refused for `int`, and
  `"FALSE"` for `bool`, where `bool("FALSE")` would be `True`.
- **An empty cell** is what Excel itself makes of one in arithmetic and text: `0`, `""`,
  `FALSE`. To tell an empty cell apart, annotate the parameter `int | None` (below).
- **An Excel error** is refused, since it is not a value of the type. A parameter without
  an annotation receives it as a `pyinex.ExcelError`.
- **A number** must be whole for `int`, as `3.0` is; Excel has no other integers.

| In the cell | A list of rows | `numpy.ndarray` |
|---|---|---|
| A range | a `list` of rows, each a `list`: `[[1.0, 2.0], [3.0, 4.0]]` | a two-dimensional array of `float64`, of the range's shape |
| A single value, from one cell or a constant | one row of one: `[[5.0]]` | a 1 × 1 array |
| An empty cell, alone or in a range | `None`, or by the element type's rule | `nan` |
| Text, `TRUE` / `FALSE` or an Excel error, alone or in a range | as it arrives, or by the element type's rule | refused |
| Mixed values in a range, such as numbers and text | each element by itself, as above | refused, naming the first cell that is not a number or empty |

- **A list of rows** is always rows, whatever the range's shape, as a range arrives
  unannotated, but lists instead of tuples. `list` and `list[list]` keep each element as it
  arrives. `list[list[T]]`, for `T` one of `int`, `float`, `str` and `bool` (or `T | None`),
  converts each element by `T`'s rule above, and a refused element refuses the argument,
  naming its row and column: `Pyinex: Total: argument 1 (rows): row 2, column 1: float
  takes a number, not the text 'x'`.
- **`numpy.ndarray`** takes numbers and empty cells only, and the console names the first
  cell that is neither: `Pyinex: Solve: argument 1 (a): row 1, column 2: numpy.ndarray
  takes numbers and empty cells, not TRUE`. An array of text or of mixed values is easier
  made in the function from a list of rows.

**`None` for an empty cell.** `T | None`, `Optional[T]` or `Union[T, None]`, for any `T`
above, gives `None` for an empty cell where `T` would apply to it, and otherwise converts
as `T`: `int | None` gives `None` for an empty cell, `list[list[float | None]]` gives
`None` for each empty cell in the range, and `list[list[float]] | None` gives `None` for an
argument that is a single empty cell.

**A position left out** is never converted: it takes its parameter's default, or `None`
([architecture.md](./architecture.md#direct-functions)).

**Spellings accepted.** The annotation is read as Python reads it, so these are the same:

- `int`, `float`, `str`, `bool`, and `numpy.ndarray` however NumPy was imported
  (`np.ndarray`). It is recognized by its module and name, so Pyinex never imports NumPy.
- A list of rows: `list`, `list[list]`, `list[list[float]]` and the like, and the same
  with `typing.List`. `list[float]`, a list of single values, is not a list of rows: it is
  an annotation Pyinex does not convert.
- With `None`: `int | None`, `None | int`, `Optional[int]`, `Union[int, None]`.
- **Written as a string**, as `"list[list[float]]"`, or with `from __future__ import
  annotations`, which makes every annotation a string: the string is evaluated among the
  names of the module that defined the function (for a `functools.partial` or a decorated
  function, the function underneath; a script's own function, the script's) when the
  function is read, and the result is read as above. A result
  that is a string again, as a quoted `"int"` is under that import, is evaluated in turn,
  up to five times; text that still gives text, such as a name whose value is its own
  name, is an annotation Pyinex does not convert. From
  Python 3.14 an annotation is not evaluated until asked for, and is read the same way.

**Annotations Pyinex does not convert.** Any other annotation, such as `dict`, `Decimal`,
`int | str` or `list[float]`, leaves the parameter's values as they arrive unannotated; so
do `typing.Any` and `object`, silently. For the others, the console says so when the
function is registered: `Pyinex: Discount in the script model.py: Pyinex does not convert
the annotation of rate, <class 'decimal.Decimal'>, so rate arrives unconverted`. So does an
annotation written as a string that cannot be evaluated, whatever its evaluation raises,
`SystemExit` included: `… the annotation of rate, 'Decimal', cannot be evaluated
(NameError: …), so rate arrives unconverted`. The function still registers. A name imported
only under `if TYPE_CHECKING:` is such an annotation when quoted.

**Annotations written without quotes that fail.** Before Python 3.14, Python evaluates an
annotation as the script defines the function, so one that fails, a name imported only
under `if TYPE_CHECKING:` among them, makes the script fail to load. From 3.14 Python
evaluates annotations only when asked, and Pyinex asks for them as forward references.
What happens then depends on what Python can do with the one that fails:

- **Python leaves it for Pyinex**, as a name it cannot find (`x: Missing`) or a call that
  raises (`x: bad()`): it is an annotation that cannot be evaluated, as above. Only that
  argument arrives unconverted: `… the annotation of x, 'bad()', cannot be evaluated
  (RuntimeError: boom), so x arrives unconverted`.
- **Python cannot leave it**, as `x: 1/0`: no annotation of that function can be read,
  since Python evaluates them together, even the return annotation or one on `**kwargs`.
  The function still registers, every argument arrives unconverted, and the console says
  why (`… its annotations cannot be evaluated (ZeroDivisionError: division by zero), so its
  arguments arrive unconverted`).
- **It raises `SystemExit` or `KeyboardInterrupt`**: the function is refused, with the
  reason on the console, and the script's other functions register.

Tests: `tests/unit/test_funcs_annotations.py::test_a_forward_reference_that_cannot_be_evaluated_leaves_the_value_and_says_so`,
`::test_a_parameter_s_annotation_that_raises_leaves_it_unconverted`,
`::test_an_ignored_annotation_that_raises_refuses_nothing`,
`::test_a_lazy_annotation_that_raises_system_exit_refuses_only_its_function`.

Each note about an annotation is said as the function registers, not again each time its
unchanged script loads.

**`numpy.ndarray` without NumPy installed.** A script that imports NumPy fails to load
without it, as any script whose import fails: the console shows the traceback, and nothing
registers. One that names NumPy only in a string annotation, or under `if TYPE_CHECKING:`,
has an annotation that cannot be evaluated, so its parameter arrives unconverted, as above.
Its other annotations convert as usual: Pyinex itself never needs NumPy.

The return annotation, and an annotation on `**kwargs`, are ignored. An annotation on
`*args` applies to each of its values.

Tests: `tests/unit/test_funcs_annotations.py` (every row and rule here, under plain
Python, and without NumPy by making `import numpy` fail in the test);
`tests/host/test_annotations.py` (each annotation converting and refusing, with its
message; annotated and unannotated parameters in one function; without NumPy); in real
Excel, `tests/excel/test_annotations.py` (the same, from cells). Every real-Excel run has
NumPy, so the rows without it are tested in `tests/unit` and `tests/host` only.

### References (`pyinex.Range`)

A direct function's parameter annotated `pyinex.Range` (or `pyinex.Range | None`, the
same) is registered to take a reference, and receives what Excel passes for it
([architecture.md](./architecture.md#pyinexrange) says what the object gives and when):

| In the formula | Arrives as |
|---|---|
| A reference: `B2`, `B2:C3`, `D:D`, `5:5`, `Data!A1`, `[Prices.xlsx]Sheet1!A1`, a defined name, `(B2:C3,E5)` | a `pyinex.Range` |
| A value: a number, text, `TRUE`, an array constant, a calculation such as `B2+1` | refused |
| An Excel error, such as `#REF!` for a reference to a sheet since deleted | refused |
| A reference to a workbook that is not open | refused: Excel passes its cells' values |
| Left out | its default, or `None` ([Signatures](./architecture.md#direct-functions)) |

"Refused" means what it means for the other annotations: the call gives `#VALUE!`, the
function is not called, and the console names the function, the argument and the value
(`Pyinex: Where: argument 1 (r): pyinex.Range takes a reference, not 5.0`).

A range's `values()` converts the cells by the rules for an unannotated argument
([Ranges](#ranges)): one cell is a plain value, more than one a tuple of rows, an empty
cell `None` and empty text `""`. `empty()` gives `True` for each empty cell, in the same
shape. A whole column or row gives only its used part. A cell holding a value Pyinex does
not know, such as an error code Excel added later, arrives as `ExcelError("#VALUE!")`, and
the console says what it was (`Pyinex: pyinex.Range: Excel error code 99 is not one Pyinex
knows; it arrives as #VALUE!`; `tests/host/test_range.py::test_a_value_pyinex_does_not_know_is_read_as_value_and_reported`).
If Python fails converting the cells, as when memory runs out, `values()` raises
`RuntimeError` saying so. `pyinex.Range` is not an element
type: `list[list[pyinex.Range]]` is an annotation Pyinex does not convert.

Tests: `tests/unit/test_funcs_range.py`, `tests/host/test_range.py`; in real Excel,
`tests/excel/test_range.py`.

## Python to Excel

### Single values

| Python returns | The cell shows | Notes |
|---|---|---|
| `bool` | `TRUE` / `FALSE` | Checked before `int`, since `bool` is a kind of `int`. |
| `int` | a number | Converted to the nearest double, so integers beyond 2^53 lose precision. An integer too large for a double (beyond about 1.8 × 10^308) is an error. |
| `float` | a number | NaN and infinity cannot be stored in a cell, so they become `#NUM!`, Excel's own error for an invalid number. That is a value, like returning `pyinex.ExcelError("#NUM!")`, not a failure: inside a block, only that cell shows `#NUM!`. |
| `str` | text | At most 32,767 UTF-16 code units, Excel's limit for a cell. A longer string is an error, never silently cut short. |
| `None` | nothing | The cell shows nothing. It holds empty text, because Excel displays a truly empty result as `0`. So `=A1=""` is TRUE and `ISBLANK(A1)` is FALSE. |
| `pyinex.ExcelError` | that error | |

Subclasses count as their base type: a `float` subclass (such as numpy's `float64`) is a
number, an `IntEnum` member is an integer, and so on. Tests:
`test_conversion.py::test_return_*`; in real Excel, `tests/excel/test_values_out.py`.

### Sequences

A `list` or `tuple` fills a block of cells:

- **All elements are single values** → one row: `[1, 2, 3]` fills three cells across.
- **At least one element is a list or tuple** → each element is a row. A list or tuple
  element supplies that row's values; a single value is a one-cell row.
  `[[1], [2], [3]]` fills three cells down; `[[1, 2], [3, 4]]` fills a 2 × 2 block.
- **Short rows are padded** with blank cells (as for `None`) to the longest row's length.
  The block is as wide as its longest row, and at least one cell wide.
- **An empty list or tuple** shows as one blank cell. An empty row inside a block is a row
  of blanks: `[[1, 2], []]` fills 2 × 2 with the second row blank, `[[]]` is one blank
  cell, and `[[], []]` is two blank cells, one above the other.
- **Nesting deeper than rows of values** — a list inside a row — is an error.
- A block larger than Excel's grid (1,048,576 rows × 16,384 columns) is an error.

So whatever arrives from a range (rows of rows) goes back out the same shape.

Tests: `test_conversion.py::test_return_row`, `::test_return_column`,
`::test_return_block`, `::test_return_ragged`, `::test_return_empty*`,
`::test_return_too_deep`; in real Excel, `tests/excel/test_values_out.py` and
`tests/excel/test_blocks.py` (array formulas, spilling, and blocks the grid's full height
and width).

### Anything else

Any other type, such as `dict`, `set`, `complex`, a custom object or a NumPy array, is an
error. Convert it to lists first. Pyinex does not convert NumPy or pandas results (an
argument can arrive as a NumPy array:
[Type annotations](#type-annotations-direct-functions)). NumPy 2.5's `np.linalg.eig` gives
complex numbers even for a real symmetric matrix, so take `np.real_if_close` of its
results before `.tolist()`. Tests: in real Excel, `tests/excel/test_numpy.py` (eigenvalues and matrix multiplication, and a raw array refused).

### What "an error" means

Every failure to convert a result for Excel gives `#VALUE!`: an unsupported type, a
string that is too long, a number that is too large, a block that is too big, or nesting
that is too deep. The console names the problem and where in the result it was, for
example "row 3, column 2: dict cannot be returned to Excel".

**A failure anywhere fails the whole result.** If any element cannot be converted, the
formula returns a single `#VALUE!`, not a block with one bad cell. A user who saw a block
that was mostly right could trust values in it that are wrong. So `[[1, 2], [3, {}]]`
returns one `#VALUE!`.

Values that *are* errors are different: a returned `pyinex.ExcelError`, or NaN and
infinity, which become `#NUM!`. Each takes only its own cell, so
`[[1, 2], [3, float("nan")]]` fills 2 × 2 with `#NUM!` in the last cell.

When the Python function raises an exception, the formula also returns
`#VALUE!`, and the traceback goes to the console. The result is an error value, not
message text, so that formulas that depend on the cell fail too: `SUM` over a range that
contains an error is an error, but `SUM` ignores text.

Tests: `test_conversion.py::test_return_unsupported*`, `::test_return_failure_is_whole`,
`::test_return_error_values_per_cell`, `tests/unit/test_call.py` (exceptions); in real
Excel, `tests/excel/test_failures.py`, each failure's console line included.
