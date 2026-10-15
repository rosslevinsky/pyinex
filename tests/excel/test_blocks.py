"""Blocks of values in real Excel (docs/conversion.md, "Sequences"): filling a legacy array
formula, spilling as Microsoft 365 does, and blocks as tall and as wide as the grid.
"""

import pytest
from conftest import E, pycall, typed


@pytest.fixture
def values(book):
    book.copy("values.py")
    return book


def block(rows, columns):
    return pycall("values.py", "block", str(rows), str(columns))


def test_a_block_smaller_than_its_array_formula_leaves_n_a(values):
    na = E("#N/A")
    assert values.array("A1:C3", block(2, 2)) == [[1.0, 2.0, na], [3.0, 4.0, na], [na, na, na]]


def test_a_block_larger_than_its_array_formula_is_cut_to_it(values):
    assert typed(values.array("A1:B2", block(4, 4))) == typed([[1.0, 2.0], [5.0, 6.0]])


def test_a_block_spills(values):
    values.formula("A1", block(3, 2))
    assert typed(values.values("A1:B3")) == typed([[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]])
    assert values.values("A4:C4") == [[None, None, None]]


def test_a_spill_blocked_by_a_value(values):
    values.put("B2", "in the way")
    assert values.formula("A1", block(3, 2)) == E("#SPILL!")


def test_a_block_as_tall_as_the_grid(values):
    values.formula("A1", block(1_048_576, 1))
    assert values.value("A1") == 1.0
    assert values.value("A1048576") == 1_048_576.0


def test_a_block_as_wide_as_the_grid(values):
    values.formula("A1", block(1, 16_384))
    assert values.value("A1") == 1.0
    assert values.value("XFD1") == 16_384.0


@pytest.mark.parametrize("rows, columns", [(1_048_577, 1), (1, 16_385)])
def test_a_block_larger_than_the_grid(values, cleared, rows, columns):
    assert values.formula("A1", block(rows, columns)) == E("#VALUE!")
    said = f"Pyinex: values.py, block: a block of {rows} rows by {columns} columns is larger than Excel's grid of 1,048,576 by 16,384"
    window = cleared(said)
    assert window and said in window.text
