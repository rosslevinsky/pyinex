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

Tests: `test_conversion.py::test_range_*`; in real Excel,
`tests/excel/test_values_in.py::test_a_range_*`, `::test_an_array_constant_*` and
`::test_a_reference_to_several_areas_*`.

### Missing arguments

See [Argument rule](./architecture.md#argument-rule) in `docs/architecture.md`. Missing
arguments at the end are dropped. An argument left out in the middle arrives as `None`,
the same as an empty cell.

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
error. Convert it to lists first. Pyinex does not convert NumPy or pandas values. NumPy
2.5's `np.linalg.eig` gives complex numbers even for a real symmetric matrix, so take
`np.real_if_close` of its results before `.tolist()`. Tests: in real Excel,
`tests/excel/test_numpy.py` (eigenvalues and matrix multiplication, and a raw array refused).

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
