"""Every rule in docs/conversion.md, in both directions, through the add-in's own code."""

import enum
import math

import pytest

# Excel's error codes, from XLCALL.H, written out here rather than taken from the code
# under test.
ERRORS = {
    "#NULL!": 0,
    "#DIV/0!": 7,
    "#VALUE!": 15,
    "#REF!": 23,
    "#NAME?": 29,
    "#NUM!": 36,
    "#N/A": 42,
    "#GETTING_DATA": 43,
}
VALUE = ("err", 15)
BLANK = ("str", "")  # What a returned None, or padding, puts in a cell.
XLTYPE_REF, XLTYPE_FLOW, XLTYPE_SREF, XLTYPE_BIGDATA = 0x0008, 0x0020, 0x0400, 0x0802

MAX_TEXT = 32767


def multi(rows):
    """The description of an Excel array holding these rows of descriptions."""
    return ("multi", len(rows), len(rows[0]), [item for row in rows for item in row])


def nums(*values):
    return [("num", float(v)) for v in values]


@pytest.fixture
def to_python(testing):
    """Excel to Python; fails the test if the conversion reported a problem."""

    def convert(desc):
        value, problem = testing.to_python(desc)
        assert problem is None, problem
        return value

    return convert


@pytest.fixture
def to_excel(testing):
    """Python to Excel; fails the test if the conversion reported a problem."""

    def convert(value):
        desc, problem = testing.to_excel(value)
        assert problem is None, problem
        return desc

    return convert


@pytest.fixture
def refused(testing):
    """Python to Excel for a value that must fail: returns the problem reported."""

    def convert(value):
        desc, problem = testing.to_excel(value)
        assert desc == VALUE
        assert problem
        return problem

    return convert


# Excel to Python: single values


def test_number_is_float(to_python):
    assert to_python(("num", 3.5)) == 3.5


def test_number_whole_is_still_float(to_python):
    value = to_python(("num", 3.0))
    assert type(value) is float and value == 3.0


def test_number_from_integer_type_is_float(to_python):
    value = to_python(("int", 7))
    assert type(value) is float and value == 7.0


def test_text(to_python):
    assert to_python(("str", "hello")) == "hello"


def test_text_unicode(to_python):
    assert to_python(("str", "naïve 日本 \U0001F600")) == "naïve 日本 \U0001F600"


def test_text_surrogate_pair_is_one_character(to_python):
    assert len(to_python(("str", "\U0001F600"))) == 1


def test_text_unpaired_surrogate_is_kept(to_python, to_excel):
    assert to_python(("str", "a\ud800b")) == "a\ud800b"
    assert to_excel("a\udc00b") == ("str", "a\udc00b")


def test_text_empty(to_python):
    assert to_python(("str", "")) == ""


def test_text_at_the_limit(to_python):
    assert to_python(("str", "x" * MAX_TEXT)) == "x" * MAX_TEXT


def test_text_with_nul(to_python, to_excel):
    assert to_python(("str", "a\0b")) == "a\0b"
    assert to_excel("a\0b") == ("str", "a\0b")


def test_bool(to_python):
    assert to_python(("bool", True)) is True
    assert to_python(("bool", False)) is False


def test_empty_cell_is_none(to_python):
    assert to_python(("nil",)) is None


def test_empty_missing_argument_is_none(to_python):
    assert to_python(("missing",)) is None


# Excel to Python: ranges


def test_range_one_cell_array_is_a_plain_value(to_python):
    assert to_python(multi([nums(1)])) == 1.0


def test_range_row(to_python):
    assert to_python(multi([nums(1, 2, 3)])) == ((1.0, 2.0, 3.0),)


def test_range_column(to_python):
    assert to_python(multi([nums(1), nums(2), nums(3)])) == ((1.0,), (2.0,), (3.0,))


def test_range_block(to_python):
    assert to_python(multi([nums(1, 2), nums(3, 4)])) == ((1.0, 2.0), (3.0, 4.0))


def test_range_mixed_values(to_python, testing):
    value = to_python(multi([[("str", "a"), ("bool", True)], [("nil",), ("err", 42)]]))
    assert value == (("a", True), (None, testing.ExcelError("#N/A")))


def test_range_is_tuples(to_python):
    value = to_python(multi([nums(1, 2), nums(3, 4)]))
    assert type(value) is tuple and all(type(row) is tuple for row in value)


def test_range_multi_area_reference_is_value_error(testing):
    for xltype in (XLTYPE_REF, XLTYPE_SREF):
        value, problem = testing.to_python(("raw", xltype))
        assert value == testing.ExcelError("#VALUE!")
        assert problem


# Excel to Python: errors


@pytest.mark.parametrize("text", ERRORS)
def test_error_arrives_as_excel_error(to_python, testing, text):
    value = to_python(("err", ERRORS[text]))
    assert type(value) is testing.ExcelError
    assert value == testing.ExcelError(text)
    assert str(value) == text and value.text == text


def test_error_is_not_text(to_python):
    assert to_python(("err", 42)) != "#N/A"
    assert "#N/A" != to_python(("err", 42))


def test_error_type_rules(testing):
    ExcelError = testing.ExcelError
    assert repr(ExcelError("#N/A")) == "ExcelError('#N/A')"
    assert ExcelError("#N/A") == ExcelError("#N/A")
    assert ExcelError("#N/A") != ExcelError("#REF!")
    assert len({ExcelError("#N/A"), ExcelError("#N/A"), ExcelError("#REF!")}) == 2
    assert ExcelError.__module__ == "pyinex"
    for bad in ("#n/a", "N/A", "", 42, None):
        with pytest.raises(ValueError):
            ExcelError(bad)

    class Agreeable(str):
        def __eq__(self, other):
            return True

        __hash__ = str.__hash__

    with pytest.raises(ValueError):
        ExcelError(Agreeable("not an Excel error"))
    assert type(ExcelError(Agreeable("#N/A")).text) is str

    class Unprintable(str):
        def __repr__(self):
            raise RuntimeError("no repr")

    with pytest.raises(ValueError):
        ExcelError(Unprintable("invalid"))


def test_unrecognized_error_code(testing):
    value, problem = testing.to_python(("err", 99))
    assert value == testing.ExcelError("#VALUE!")
    assert "99" in problem


@pytest.mark.parametrize("xltype", [XLTYPE_FLOW, XLTYPE_BIGDATA, 0x0200, 0x2000])
def test_unrecognized_type(testing, xltype):
    value, problem = testing.to_python(("raw", xltype))
    assert value == testing.ExcelError("#VALUE!")
    assert problem


def test_unrecognized_inside_a_range(testing):
    value, problem = testing.to_python(multi([[("num", 1.0), multi([nums(2)])]]))
    assert value == ((1.0, testing.ExcelError("#VALUE!")),)
    assert problem


def test_unrecognized_empty_array(testing):
    value, problem = testing.to_python(("multi", 0, 0, []))
    assert value == testing.ExcelError("#VALUE!")
    assert problem


# Python to Excel: single values


def test_return_bool(to_excel):
    assert to_excel(True) == ("bool", True)
    assert to_excel(False) == ("bool", False)


def test_return_int(to_excel):
    assert to_excel(3) == ("num", 3.0)
    assert to_excel(-(2**53)) == ("num", float(-(2**53)))


def test_return_int_beyond_double_precision_rounds(to_excel):
    assert to_excel(2**53 + 1) == ("num", float(2**53))


def test_return_int_too_large(refused):
    assert "int" in refused(10**400)


def test_return_float(to_excel):
    assert to_excel(2.5) == ("num", 2.5)
    assert to_excel(-0.0) == ("num", -0.0)


@pytest.mark.parametrize("value", [math.nan, math.inf, -math.inf])
def test_return_float_not_a_number_is_num_error(to_excel, value):
    assert to_excel(value) == ("err", ERRORS["#NUM!"])


def test_return_str(to_excel):
    assert to_excel("naïve 日本 \U0001F600") == ("str", "naïve 日本 \U0001F600")


def test_return_str_at_the_limit(to_excel):
    assert to_excel("x" * MAX_TEXT) == ("str", "x" * MAX_TEXT)


def test_return_str_limit_counts_utf16_units(to_excel, refused):
    # Each emoji is two UTF-16 code units: 16,383 of them and one letter is exactly 32,767.
    at_limit = "x" + "\U0001F600" * (MAX_TEXT // 2)
    assert to_excel(at_limit) == ("str", at_limit)
    assert "32,767" in refused("x" + at_limit)


def test_return_str_too_long(refused):
    assert "32,767" in refused("x" * (MAX_TEXT + 1))


def test_return_none_is_blank(to_excel):
    assert to_excel(None) == BLANK


@pytest.mark.parametrize("text", ERRORS)
def test_return_excel_error(to_excel, testing, text):
    assert to_excel(testing.ExcelError(text)) == ("err", ERRORS[text])


def test_return_subclasses_count_as_their_base(to_excel, testing):
    class Color(enum.IntEnum):
        RED = 3

    class Money(float):
        pass

    class Name(str):
        pass

    class Row(list):
        pass

    class Pair(tuple):
        pass

    class Missing(testing.ExcelError):
        pass

    assert to_excel(Color.RED) == ("num", 3.0)
    assert to_excel(Money(1.25)) == ("num", 1.25)
    assert to_excel(Name("x")) == ("str", "x")
    assert to_excel(Row([1, 2])) == multi([nums(1, 2)])
    assert to_excel(Pair((Row([1]), Row([2])))) == multi([nums(1), nums(2)])
    assert to_excel(Missing("#N/A")) == ("err", ERRORS["#N/A"])


# Python to Excel: sequences


def test_return_row(to_excel):
    assert to_excel([1, 2, 3]) == multi([nums(1, 2, 3)])
    assert to_excel((1, 2, 3)) == multi([nums(1, 2, 3)])


def test_return_single_element_row(to_excel):
    assert to_excel([7]) == multi([nums(7)])


def test_return_column(to_excel):
    assert to_excel([[1], [2], [3]]) == multi([nums(1), nums(2), nums(3)])


def test_return_block(to_excel):
    assert to_excel([[1, 2], [3, 4]]) == multi([nums(1, 2), nums(3, 4)])


def test_return_block_mixed_values(to_excel, testing):
    assert to_excel([["a", True], [None, testing.ExcelError("#N/A")]]) == multi(
        [[("str", "a"), ("bool", True)], [BLANK, ("err", 42)]]
    )


def test_return_ragged(to_excel):
    assert to_excel([[1, 2, 3], [4]]) == multi([nums(1, 2, 3), [("num", 4.0), BLANK, BLANK]])


def test_return_ragged_single_values_are_one_cell_rows(to_excel):
    assert to_excel([1, [2, 3]]) == multi([[("num", 1.0), BLANK], nums(2, 3)])


def test_return_empty_is_one_blank_cell(to_excel):
    assert to_excel([]) == BLANK
    assert to_excel(()) == BLANK


def test_return_empty_rows(to_excel):
    assert to_excel([[]]) == multi([[BLANK]])
    assert to_excel([[], []]) == multi([[BLANK], [BLANK]])
    assert to_excel([[1, 2], []]) == multi([nums(1, 2), [BLANK, BLANK]])


def test_return_too_deep(refused):
    assert "row 2, column 1" in refused([[1, 2], [[3], 4]])


def test_return_round_trips_a_range(testing, to_excel):
    block = multi([[("num", 1.0), ("str", "a")], [("bool", False), ("err", 7)]])
    value, problem = testing.to_python(block)
    assert problem is None
    assert to_excel(value) == block


def test_return_largest_row_and_column(to_excel):
    assert to_excel([0] * 16384)[:3] == ("multi", 1, 16384)
    assert to_excel([[0]] * 1048576)[:3] == ("multi", 1048576, 1)


def test_return_larger_than_the_grid(refused):
    assert "grid" in refused([0] * 16385)
    assert "grid" in refused([[0]] * 1048577)


# Python to Excel: anything else, and what an error means


@pytest.mark.parametrize("value", [{"a": 1}, {1}, object(), b"bytes", 1j, range(3)])
def test_return_unsupported(refused, value):
    assert type(value).__name__ in refused(value)


def test_return_unsupported_names_where(refused):
    assert refused([[1, 2], [3, {}]]) == "row 2, column 2: dict cannot be returned to Excel"
    assert refused([1, {}]).startswith("row 1, column 2: dict")


def test_return_failure_is_whole(refused):
    refused([[1, 2], [3, {}]])
    refused([[1, 2], [3, "x" * (MAX_TEXT + 1)]])
    refused([1, 10**400])


def test_return_error_values_per_cell(to_excel, testing):
    assert to_excel([[1, 2], [3, math.nan]]) == multi([nums(1, 2), [("num", 3.0), ("err", 36)]])
    assert to_excel([1, testing.ExcelError("#REF!")]) == multi([[("num", 1.0), ("err", 23)]])


def test_return_list_changed_by_conversion_is_safe(to_excel, testing):
    # Converting an ExcelError subclass can run its code; the result reflects the list as
    # it was when conversion began.
    rows = [[1, 2], [3, 4]]

    class Meddling(testing.ExcelError):
        def __getattribute__(self, name):
            for row in rows:
                row.clear()
            rows.clear()
            return super().__getattribute__(name)

    rows[0][0] = Meddling("#N/A")
    desc, problem = testing.to_excel(rows)
    assert rows == []  # The conversion did run the meddling code.
    assert problem is None
    assert desc == multi([[("err", 42), ("num", 2.0)], nums(3, 4)])


# The test module itself: descriptions it must refuse rather than misread.


def test_harness_refuses_sizes_beyond_the_grid(testing):
    # 262,145 x 16,384 wraps to 16,384 in 32-bit arithmetic.
    with pytest.raises(ValueError):
        testing.to_python(("multi", 262145, 16384, [("nil",)] * 16384))


def test_harness_ignores_items_added_while_building(testing):
    items = []

    class Growing:
        def __bool__(self):
            items.append(("nil",))
            return True

    items.append(("bool", Growing()))
    value, problem = testing.to_python(("multi", 1, 1, items))
    assert value is True and problem is None
    assert len(items) == 2


def test_range_conversion_survives_code_keeping_objects(testing, monkeypatch):
    # Making an ExcelError runs Python code; here it keeps every object Python tracks,
    # which would include any tuple still being filled.
    import gc

    kept = []
    original = testing.ExcelError.__init__

    def keeping(self, text):
        original(self, text)
        kept.append(gc.get_objects())

    monkeypatch.setattr(testing.ExcelError, "__init__", keeping)
    value, problem = testing.to_python(multi([[("err", 42), ("num", 1.0)], nums(2, 3)]))
    assert problem is None
    assert value == ((testing.ExcelError("#N/A"), 1.0), (2.0, 3.0))
