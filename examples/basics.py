"""Calling Python from a cell, and how single values cross between Excel and Python.

basics.xlsx calls these functions with PyCall. docs/conversion.md has the full rules.
"""

import pyinex  # Built into the add-in; used here only for ExcelError.


def add(a, b):
    return a + b


def kind(value):
    """The name of the Python type a value arrives as."""
    return type(value).__name__


def greet(name, greeting="Hello"):
    return f"{greeting}, {name}!"


def count(*values):
    """How many arguments arrived. A missing one in the middle arrives as None."""
    return len(values)


def is_even(n):
    return int(n) % 2 == 0


def shout(text):
    return text.upper()


def length(text):
    return len(text)


def nothing():
    """Returning None leaves the cell showing nothing."""
    return None


def not_a_number():
    """A cell cannot hold NaN, so it shows #NUM!."""
    return float("nan")


def error_text(value):
    """An Excel error arrives as pyinex.ExcelError; its text names it."""
    return value.text if isinstance(value, pyinex.ExcelError) else "not an error"


def not_available():
    """Returning an ExcelError puts that error in the cell."""
    return pyinex.ExcelError("#N/A")


def fail():
    """An exception gives #VALUE!, and its traceback goes to the console."""
    raise ValueError("this function always fails")


# Any callable at the top level of a script works, not only a def.
biggest = max
whole = int
