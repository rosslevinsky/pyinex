"""The pyinex module's functions that ask Excel (docs/architecture.md, "The pyinex module"),
called from scripts through the fake Excel."""

import pytest
from conftest import attempt, evaluate


def caller(host, **options):
    return evaluate(host, f"__import__('pyinex').caller(**{options!r})")


@pytest.mark.parametrize(
    "options, expected",
    [
        ({}, "C12"),
        ({"r1c1": True}, "R12C3"),
        ({"full": True}, "[Book.xlsx]Sheet 1!C12"),
        ({"r1c1": True, "full": True}, "[Book.xlsx]Sheet 1!R12C3"),
    ],
)
def test_caller_in_both_address_styles(host, options, expected):
    host.set_caller("[Book.xlsx]Sheet 1", row=12, column=3)
    assert caller(host, **options) == expected


def test_caller_given_as_a_reference_with_no_sheet(host):
    # Excel may describe the calling cell either way; the sheet then comes from xlSheetNm.
    host.set_caller("[Book.xlsx]Sheet1", row=4, column=27, sref=True)
    assert caller(host) == "AA4"
    assert caller(host, full=True) == "[Book.xlsx]Sheet1!AA4"


def test_caller_of_a_formula_in_a_range_is_the_range(host):
    host.set_caller("[Book.xlsx]Sheet1", row=12, column=3, rows=3, columns=3)
    assert caller(host) == "C12:E14"
    assert caller(host, r1c1=True, full=True) == "[Book.xlsx]Sheet1!R12C3:R14C5"


def test_caller_at_the_grids_far_corner(host):
    host.set_caller("[Book.xlsx]Sheet1", row=1048576, column=16384)
    assert caller(host) == "XFD1048576"


def test_caller_outside_a_cell_raises(host):
    host.set_caller(None)
    assert attempt(host, "__import__('pyinex').caller()").startswith("RuntimeError: ")
    assert attempt(host, "__import__('pyinex').sheet()").startswith("RuntimeError: ")


def test_sheet_keeps_every_character(host):
    name = "[Bücher ✓.xlsx]Лист 😀"
    host.set_caller(name, row=2, column=2)
    assert evaluate(host, "__import__('pyinex').sheet()") == name
    assert caller(host, full=True) == name + "!B2"


def test_esc_pressed_keeps_the_key_press_by_default(host):
    assert evaluate(host, "__import__('pyinex').esc_pressed()") == "False"
    host.set_escape(True)
    assert evaluate(host, "__import__('pyinex').esc_pressed()") == "True"
    assert evaluate(host, "__import__('pyinex').esc_pressed(clear=False)") == "True"
    assert evaluate(host, "__import__('pyinex').esc_pressed()") == "True"


def test_esc_pressed_can_clear_the_key_press(host):
    host.set_escape(True)
    assert evaluate(host, "__import__('pyinex').esc_pressed(clear=True)") == "True"
    assert evaluate(host, "__import__('pyinex').esc_pressed()") == "False"


def test_excel_error_is_the_type_values_arrive_as(host):
    assert evaluate(host, "__import__('pyinex').ExcelError('#N/A')") == "#N/A"
    assert evaluate(host, "__import__('pyinex').ExcelError.__module__") == "pyinex"
