"""Values from Excel to Python, in real Excel (docs/conversion.md, "Excel to Python", and
docs/architecture.md, "Argument rule"): what each argument arrives as, its type and value.

Each test calls tests/excel/scripts/values.py, whose describe() returns the type's name
and the repr of what arrived.
"""

import pytest
from conftest import E, pycall, typed


@pytest.fixture
def values(book):
    book.copy("values.py")
    return book


def describe(book, arg, cell="C1"):
    return book.formula(cell, pycall("values.py", "describe", arg))


# Single values.


@pytest.mark.parametrize(
    "constant, arrives",
    [
        ("3", "float: 3.0"),
        ("3.5", "float: 3.5"),
        ("-2.25", "float: -2.25"),
        ("1E+300", "float: 1e+300"),
        ("TRUE", "bool: True"),
        ("FALSE", "bool: False"),
        ("text", "str: 'text'"),
        ("é ✓ 😀", "str: 'é ✓ 😀'"),
    ],
)
def test_a_cell_arrives_as_its_python_value(values, constant, arrives):
    values.put("A1", constant)
    assert describe(values, "A1") == arrives


@pytest.mark.parametrize(
    "formula, arrives",
    [
        ('""', "str: ''"),
        ("1/4", "float: 0.25"),
        ("DATE(2026,1,2)", "float: 46024.0"),
        ("1=1", "bool: True"),
        ('"a"&"b"', "str: 'ab'"),
    ],
)
def test_a_formula_argument_arrives_as_its_value(values, formula, arrives):
    assert describe(values, formula) == arrives


def test_an_empty_cell_arrives_as_none(values):
    assert describe(values, "A1") == "NoneType: None"


def test_empty_text_is_not_an_empty_cell(values):
    values.put("A1", '=""')
    assert describe(values, "A1") == "str: ''"


def test_text_beyond_the_bmp_arrives_one_character_each(values):
    values.put("A1", "😀😀")
    assert values.formula("C1", pycall("values.py", "length", "A1")) == 2.0


def test_text_beyond_the_bmp_at_excels_limit_arrives_whole(values):
    # 16,383 emoji are 32,766 UTF-16 code units; Python counts 16,383 characters. (Excel's
    # LEN counts characters too, so the units are counted here.)
    values.put("A1", '=REPT("😀",16383)')
    assert len(values.value("A1").encode("utf-16-le")) // 2 == 32766
    assert values.formula("C1", pycall("values.py", "length", "A1")) == 16383.0


def test_text_at_excels_limit_arrives_whole(values):
    values.put("A1", '=REPT("é",32767)')
    assert values.formula("C1", pycall("values.py", "length", "A1")) == 32767.0


@pytest.mark.parametrize("error", ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"])
def test_an_error_arrives_as_an_excel_error(values, error):
    values.put("A1", f"={error}")
    assert describe(values, "A1") == f"ExcelError: ExcelError('{error}')"


def test_an_excel_error_behaves_as_documented(values):
    values.put("A1", "=NA()")
    values.formula("C1", pycall("values.py", "error_facts", "A1"))
    assert typed(values.values("C1:J1")) == typed([[True, "#N/A", "#N/A", "ExcelError('#N/A')", True, False, True, False]])


@pytest.mark.parametrize(
    "setup, formula, error",
    [
        ([("A2", "in the way")], "=SEQUENCE(3)", "#SPILL!"),
        ([], "=FILTER({1,2},{FALSE,FALSE})", "#CALC!"),
    ],
)
def test_an_error_newer_than_the_sdk_arrives_as_value(values, cleared, mark, setup, formula, error):
    # Excel itself passes an add-in errors newer than the SDK it was built with as #VALUE!,
    # so Pyinex never sees their codes (docs/conversion.md, "Excel errors").
    for cell, value in setup:
        values.put(cell, value)
    values.put("A1", formula)
    assert values.value("A1") == E(error)
    assert describe(values, "A1") == "ExcelError: ExcelError('#VALUE!')"
    text = mark()  # Everything written since the console was cleared.
    assert "Pyinex:" not in text, text


# Ranges and arrays: rows of rows, whatever the shape.


@pytest.mark.parametrize(
    "cells, arrives",
    [
        ("A1:A1", "float: 1.0"),
        ("A1:C1", "tuple: ((1.0, 2.0, 3.0),)"),
        ("A1:A3", "tuple: ((1.0,), (4.0,), (7.0,))"),
        ("A1:B2", "tuple: ((1.0, 2.0), (4.0, 5.0))"),
        ("A1:C3", "tuple: ((1.0, 2.0, 3.0), (4.0, 5.0, 6.0), (7.0, None, 'nine'))"),
    ],
)
def test_a_range_arrives_as_rows_of_rows(values, cells, arrives):
    for cell, value in zip(["A1", "B1", "C1", "A2", "B2", "C2", "A3", "C3"], ["1", "2", "3", "4", "5", "6", "7", "nine"]):
        values.put(cell, value)
    assert describe(values, cells, cell="E1") == arrives


@pytest.mark.parametrize(
    "constant, arrives",
    [
        ("{1}", "float: 1.0"),
        ("{1,2,3}", "tuple: ((1.0, 2.0, 3.0),)"),
        ("{1;2;3}", "tuple: ((1.0,), (2.0,), (3.0,))"),
        ('{1,"a";TRUE,#N/A}', "tuple: ((1.0, 'a'), (True, ExcelError('#N/A')))"),
    ],
)
def test_an_array_constant_arrives_as_rows_of_rows(values, constant, arrives):
    assert describe(values, constant) == arrives


def test_a_reference_to_several_areas_arrives_as_value(values):
    values.put("A1", "1")
    values.put("D4", "2")
    assert describe(values, "(A1:B2,D4)", cell="F1") == "ExcelError: ExcelError('#VALUE!')"


def test_a_range_goes_back_out_the_same_shape(values):
    for cell, value in zip(["A1", "B1", "A2", "B2", "A3"], ["1", "two", "TRUE", '=""', "=NA()"]):
        values.put(cell, value)
    values.formula("D1", pycall("values.py", "echo", "A1:B3"))
    assert typed(values.values("D1:E3")) == typed([[1.0, "two"], [True, ""], [E("#N/A"), ""]])


# Arguments.


def test_trailing_missing_arguments_are_dropped(values):
    assert values.formula("A1", pycall("values.py", "count", "1", "2")) == 2.0
    assert values.formula("A2", pycall("values.py", "count")) == 0.0


def test_a_missing_argument_in_the_middle_is_none(values):
    values.formula("A1", pycall("values.py", "describe_each", "1", "", "3"))
    assert values.values("A1:C1") == [["float: 1.0", "NoneType: None", "float: 3.0"]]


def test_thirty_arguments_arrive(values):
    assert values.formula("A1", pycall("values.py", "count", *[str(i) for i in range(30)])) == 30.0


def test_varargs_2010(values):
    values.formula("A1", pycall("values.py", "HasVarargs", '"first"', "1", "2", '"three"'))
    assert values.values("A1:A3") == [["First params: first"], ["3 varargs"], ["Varargs: 1.0,2.0,three)"]]


@pytest.mark.parametrize("value, words", [("the quick brown fox", 4.0), ("", 0.0), ("'42", 1.0), ("42", 0.0), ("TRUE", 0.0)])
def test_cell_wordcount_2010(values, value, words):
    values.put("A1", value)
    assert values.formula("C1", pycall("values.py", "CellWordcount", "A1")) == words
