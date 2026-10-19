"""Functions a cell calls by their own names.

direct_functions.xlsx loads this script with =PyLoad("direct_functions.py"). Each function
marked @pyinex.func then becomes a worksheet function under its own name, so a cell can
say =Discount(100, 0.1) where it would say =PyCall("direct_functions.py", "Discount",
100, 0.1). When the workbook opens again, Pyinex loads the script again. The docstrings
are the help Insert Function shows. docs/architecture.md, "Direct functions", has the
rules.
"""

import pyinex


@pyinex.func
def Discount(price, rate):
    """The price after a discount.

    Args:
        price: The price before the discount.
        rate: The discount, as a fraction such as 0.1.
    """
    return price * (1 - rate)


@pyinex.func(name="CompoundGrowth", category="Pyinex examples")
def compound(start: float, rate: float, years: int = 1):
    """A starting amount grown at a yearly rate for whole years.

    Args:
        start: The amount at the start.
        rate: The yearly rate, as a fraction such as 0.05.
        years: How many whole years; 1 if left out.
    """
    return start * (1 + rate) ** years


@pyinex.func
def Total(*values):
    """The sum of any number of values; one left out is skipped.

    Args:
        values: A number to add.
    """
    return sum(v for v in values if v is not None)


@pyinex.func
def Where(cells: pyinex.Range):
    """The workbook, sheet and address of a reference.

    Args:
        cells: A reference, such as H4:H7.
    """
    return cells.address(full=True)


@pyinex.func
def UsedRows(cells: pyinex.Range):
    """How many rows of a reference are read: for a whole column, only its used part.

    Args:
        cells: A reference, such as H:H.
    """
    values = cells.values()
    return len(values) if isinstance(values, tuple) else 1


@pyinex.func
def FilledCells(cells: pyinex.Range):
    """How many cells of a reference hold something.

    Args:
        cells: A reference, such as H:H.
    """
    empty = cells.empty()
    rows = empty if isinstance(empty, tuple) else ((empty,),)
    return sum(not e for row in rows for e in row)
