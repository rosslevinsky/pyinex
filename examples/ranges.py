"""Ranges in and blocks of cells out.

ranges.xlsx calls these functions with PyCall. Any range larger than one cell arrives as
a tuple of rows, each row a tuple of values, so cells[r][c] works whatever its shape. A
list or tuple returned fills cells: a flat one fills a row, a list of rows fills a block.
docs/conversion.md has the full rules.
"""


def size(cells):
    """The range's rows and columns, as text."""
    if not isinstance(cells, tuple):
        return "one value"
    return f"{len(cells)} x {len(cells[0])}"


def total(cells):
    """The sum of the numbers in a range, skipping empty cells."""
    return sum(v for row in cells for v in row if v is not None)


def empty_cells(cells):
    """Empty cells in a range arrive as None."""
    return sum(1 for row in cells for v in row if v is None)


def transpose(cells):
    return [list(column) for column in zip(*cells)]


def multiply(a, b):
    """The matrix product of two ranges."""
    return [[sum(x * y for x, y in zip(row, column)) for column in zip(*b)] for row in a]


def as_row(cells):
    """A flat list fills one row."""
    return [v for row in cells for v in row]


def as_column(cells):
    """A list of one-value rows fills one column."""
    return [[v] for row in cells for v in row]


def ragged():
    """Short rows are padded with blank cells."""
    return [[1, 2, 3], [4]]
