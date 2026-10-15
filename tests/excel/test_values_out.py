"""Values from Python to Excel, in real Excel (docs/conversion.md, "Python to Excel"): what
each return value shows as, single values and blocks, and the 2010 return-value checks.

Each formula is entered as Microsoft 365 enters it, so a block spills from its cell.
"""

import pytest
from conftest import E, pycall, typed


@pytest.fixture
def values(book):
    book.copy("values.py")
    return book


def give(book, name, cell="A1"):
    return book.formula(cell, pycall("values.py", "give", f'"{name}"'))


# Single values.


@pytest.mark.parametrize(
    "name, shows",
    [
        ("true", True),
        ("false", False),
        ("text", "text"),
        ("float", 2.5),
        ("int", 666.0),
        ("huge_int", float(1234567891011121314)),
        ("intenum", 7.0),
        ("float64", 1.25),
        ("unicode", "é ✓ 😀"),
        ("str_subclass", "text"),
        ("emoji_too_long", E("#VALUE!")),
        ("nan", E("#NUM!")),
        ("inf", E("#NUM!")),
        ("-inf", E("#NUM!")),
        ("too_large_int", E("#VALUE!")),
        ("long_text", E("#VALUE!")),
        ("dict", E("#VALUE!")),
        ("set", E("#VALUE!")),
        ("complex", E("#VALUE!")),
        ("object", E("#VALUE!")),
    ],
)
def test_a_return_value_shows_as_documented(values, name, shows):
    got = give(values, name)
    assert (type(got), got) == (type(shows), shows)


def test_none_shows_nothing_but_is_not_blank(values):
    assert give(values, "none") == ""
    assert values.formula("B1", '=A1=""') is True
    assert values.formula("B2", "=ISBLANK(A1)") is False


@pytest.mark.parametrize("name", ["longest_text", "longest_emoji"])
def test_text_at_excels_limit_is_returned_whole(values, name):
    # Excel's limit is 32,767 UTF-16 code units, and an emoji takes two.
    give(values, name)
    assert len(values.value("A1").encode("utf-16-le")) // 2 == 32767
    assert values.formula("B2", pycall("values.py", "is_given", f'"{name}"', "A1")) is True


@pytest.mark.parametrize("error", ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"])
def test_a_returned_excel_error_shows_that_error(values, error):
    assert values.formula("A1", pycall("values.py", "error", f'"{error}"')) == E(error)


# Blocks.


@pytest.mark.parametrize(
    "name, cells, shows",
    [
        ("row", "A1:C1", [[1.0, 2.0, 3.0]]),
        ("column", "A1:A3", [[1.0], [2.0], [3.0]]),
        ("ragged", "A1:C3", [[1.0, 2.0, 3.0], [4.0, "", ""], [5.0, "", ""]]),
        ("empty", "A1:A1", [[""]]),
        ("empty_in_list", "A1:A1", [[""]]),
        ("empty_rows", "A1:A2", [[""], [""]]),
        ("empty_row_in_block", "A1:B2", [[1.0, 2.0], ["", ""]]),
        ("nan_in_block", "A1:B2", [[1.0, 2.0], [3.0, E("#NUM!")]]),
        ("list_subclass", "A1:C1", [[1.0, True, "x"]]),
        ("bools_in_block", "A1:B2", [[True, False], [1.0, 0.0]]),
    ],
)
def test_a_list_fills_a_block(values, name, cells, shows):
    give(values, name)
    assert typed(values.values(cells)) == typed(shows)


def test_a_block_is_exactly_its_size(values):
    give(values, "ragged")
    assert values.values("D1:D3") == [[None], [None], [None]]
    assert values.values("A4:C4") == [[None, None, None]]


@pytest.mark.parametrize("name", ["too_deep", "bad_in_block"])
def test_a_failure_anywhere_fails_the_whole_result(values, name):
    assert give(values, name) == E("#VALUE!")
    assert values.values("A1:B2") == [[E("#VALUE!"), None], [None, None]]


# 2010: TypesOfReturnValues, SortList, UnicodeTest.

NESTED = [[2.0, 2.0, ""], ["be", 1.0, ""], ["me", "", ""], [123.4, "", ""], [1.0, 2.0, 3.0]]


@pytest.mark.parametrize(
    "index, cells, shows",
    [
        (0, "A1:C5", NESTED),
        (1, "A1:C5", NESTED),
        (2, "A1:A1", [[E("#VALUE!")]]),
        (3, "A1:A1", [[E("#VALUE!")]]),
        (4, "A1:A1", [[E("#VALUE!")]]),
        (5, "A1:A1", [[""]]),
        (6, "A1:A1", [[True]]),
        (7, "A1:A1", [[False]]),
        (8, "A1:A1", [["me"]]),
        (9, "A1:A1", [[3.14]]),
        (10, "A1:A1", [[666.0]]),
        (11, "A1:A1", [[1.2345678910111214e18]]),
    ],
)
def test_types_of_return_values_2010(values, index, cells, shows):
    values.formula("A1", pycall("values.py", "TypesOfReturnValues", str(index)))
    assert typed(values.values(cells)) == typed(shows)


def test_sort_list_2010(values):
    for cell, value in zip(["A1", "A2", "A3", "A5"], ["3", "b", "TRUE", "1.5"]):
        values.put(cell, value)
    values.formula("C1", pycall("values.py", "SortList", "A1:A5"))
    # Sorted by each value's text, an empty cell's being empty: '', '1.5', '3.0', 'True', 'b'.
    assert typed(values.values("C1:C5")) == typed([[""], [1.5], [3.0], [True], ["b"]])


def test_unicode_2010(values):
    assert values.formula("A1", pycall("values.py", "ReturnConstString")) == "Hello...ρεαλλξ Unicode"
    values.put("B1", "Ünïcödé ✓ 😀 中文")
    assert values.formula("C1", pycall("values.py", "ReturnString", "B1")) == "Ünïcödé ✓ 😀 中文"
