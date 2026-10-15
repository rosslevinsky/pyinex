"""Functions the real-Excel value tests call (test_values_in.py, test_values_out.py,
test_failures.py, test_blocks.py), copied beside each test's workbook first.

The 2010 functions are carried over from 0.83's Examples/PyinexTest.py and
Examples/UnicodeTest.py, without their print() calls and Python 2 branches.
"""

import enum
import math

import numpy
import pyinex

# What arrives.


def describe(x):
    """What arrived: its type's name and its repr, as text."""
    return f"{type(x).__name__}: {x!r}"


def describe_each(*args):
    """describe() of each argument, in one row."""
    return [describe(a) for a in args]


def count(*args):
    return len(args)


def length(text):
    return len(text)


def echo(x):
    return x


def error_facts(e):
    """What the documents promise of an ExcelError that arrived, one fact per cell."""
    return [
        isinstance(e, pyinex.ExcelError),
        str(e),
        e.text,
        repr(e),
        e == pyinex.ExcelError(e.text),
        e == e.text,
        hash(e) == hash(pyinex.ExcelError(e.text)),
        e == pyinex.ExcelError("#NULL!" if e.text != "#NULL!" else "#N/A"),
    ]


# What goes back.


class Small(enum.IntEnum):
    SEVEN = 7


class Text(str):
    pass


class Row(list):
    pass


RETURNS = {
    "none": None,
    "true": True,
    "false": False,
    "text": "text",
    "float": 2.5,
    "int": 666,
    "huge_int": 1234567891011121314,
    "too_large_int": 10**400,
    "nan": math.nan,
    "inf": math.inf,
    "-inf": -math.inf,
    "nan_in_block": [[1, 2], [3, math.nan]],
    "longest_text": "x" * 32767,
    "long_text": "x" * 32768,
    "longest_emoji": "😀" * 16383 + "x",  # 32,767 UTF-16 code units.
    "emoji_too_long": "😀" * 16384,  # 32,768.
    "row": [1, 2, 3],
    "column": [[1], [2], [3]],
    "ragged": [[1, 2, 3], [4], 5],
    "empty": [],
    "empty_rows": [[], []],
    "empty_in_list": [[]],
    "empty_row_in_block": [[1, 2], []],
    "too_deep": [[1, [2]]],
    "dict": {"server": "mpilgrim"},
    "set": {1},
    "complex": 1 + 2j,
    "object": object(),
    "bad_in_block": [[1, 2], [3, {}]],
    "intenum": Small.SEVEN,
    "float64": numpy.float64(1.25),
    "unicode": "é ✓ 😀",
    "str_subclass": Text("text"),
    "list_subclass": Row([1, True, "x"]),
    "bools_in_block": [[True, False], [1, 0]],
}


def give(name):
    return RETURNS[name]


def is_given(name, x):
    """Whether x, read back from a cell, is exactly what give(name) returned."""
    return x == RETURNS[name]


def error(text):
    return pyinex.ExcelError(text)


def block(rows, columns):
    """rows × columns numbers, counting from 1 across each row."""
    rows, columns = int(rows), int(columns)
    if columns == 1:
        return [[r + 1] for r in range(rows)]
    return [[r * columns + c + 1 for c in range(columns)] for r in range(rows)]


# Failures.


def raises():
    raise ValueError("raised on purpose")


def exits():
    raise SystemExit(3)


def interrupts():
    raise KeyboardInterrupt


def two(a, b):
    return a + b


def bad_error():
    return pyinex.ExcelError("#BAD!")


NOT_A_FUNCTION = 3


# 2010: Examples/PyinexTest.py.


def TypesOfReturnValues(idx):
    retArray = [
        [(2, 2), ("be", 1.0), "me", 123.4, (1, 2, 3)],
        ((2, 2), ("be", 1.0), "me", 123.4, (1, 2, 3)),
        ((2, 2), ("be", 1.0), "me", 123.4, (1, 2, 3), (1, ("too nested",))),
        {"server": "mpilgrim", "database": "master"},
        [{"server": "pilgrim", "database": "master"}],
        None,
        True,
        False,
        "me",
        3.14,
        666,
        1234567891011121314,
    ]
    return retArray[int(idx)]


def stringize(x):
    if type(x) == list:
        return list(map(stringize, x))
    elif type(x) == tuple:
        return tuple(map(stringize, x))
    elif x is None:
        return ""
    else:
        return str(x)


def SortList(unsortedList):
    return sorted(unsortedList, key=stringize)


def CellWordcount(sentence):
    if type(sentence) != str:
        return 0
    return len(sentence.split())


def Columnize(tup):
    return [(x,) for x in tup]


def HasVarargs(firstParam, *allTheRest):
    r1 = "First params: " + str(firstParam)
    r2 = str(len(allTheRest)) + " varargs"
    r3 = "Varargs: " + ",".join([str(x) for x in allTheRest]) + ")"
    return Columnize((r1, r2, r3))


# 2010: Examples/UnicodeTest.py.


def ReturnConstString():
    return "Hello...ρεαλλξ Unicode"


def ReturnString(s):
    return s
