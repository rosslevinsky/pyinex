"""pyinex.Range in pure Python (src/core/python/funcs.py; docs/architecture.md,
"pyinex.Range"; docs/conversion.md, "References"): addresses in both styles, sizes, areas,
the used part of a whole column or row found by counting, empty cells, the type codes a
function taking a reference registers with, values that are not references refused, and a
call Excel stopped.

A Range asks Excel through the pyinex module's private _sheet_id, _range_values and
_range_count, which the core provides; here a stand-in sheet answers them, and records each
question but _sheet_id's.
"""

import pytest

# The fixtures core and load are in conftest.py.

ROWS, COLUMNS = 1048576, 16384


class Sheet:
    """Answers a Range's questions from cells, {(row, column): value}, counted from 0, and
    records them."""

    def __init__(self, cells):
        self.cells = cells
        self.asked = []
        self.ids = {}  # Each sheet's ID now, if not 7, or None for a sheet Excel has not open.

    def sheet_id(self, name):
        return self.ids.get(name, 7)

    def values(self, sheet, r1, r2, c1, c2):
        self.asked.append(("values", sheet, r1, r2, c1, c2))
        assert r2 - r1 + 1 < ROWS and c2 - c1 + 1 < COLUMNS, "read a whole column or row"
        if (r2 - r1 + 1) * (c2 - c1 + 1) > 1000:
            return "many"  # Only which cells were asked for is checked.
        rows = tuple(tuple(self.cells.get((r, c)) for c in range(c1, c2 + 1)) for r in range(r1, r2 + 1))
        return rows[0][0] if len(rows) == 1 and len(rows[0]) == 1 else rows

    def count(self, sheet, r1, r2, c1, c2):
        self.asked.append(("count", sheet, r1, r2, c1, c2))
        return sum(1 for (r, c) in self.cells if r1 <= r <= r2 and c1 <= c <= c2)


@pytest.fixture
def sheet(core, monkeypatch):
    """A stand-in sheet the pyinex module's private functions read; sheet.cells to fill."""
    s = Sheet({})
    pyinex = __import__("sys").modules["pyinex"]
    monkeypatch.setattr(pyinex, "_range_values", s.values, raising=False)
    monkeypatch.setattr(pyinex, "_range_count", s.count, raising=False)
    monkeypatch.setattr(pyinex, "_sheet_id", s.sheet_id, raising=False)
    return s


def make(core, *areas, sheet="[Book.xlsx]Sheet1", id=7):
    return core["Range"](sheet, id, areas)


B2_C3 = (1, 2, 1, 2)
COLUMN_D = (0, ROWS - 1, 3, 3)


def test_its_workbook_sheet_and_size(core):
    r = make(core, B2_C3, sheet="[My Book.xlsx]Data Sheet")
    assert (r.workbook, r.sheet, r.rows, r.columns) == ("My Book.xlsx", "Data Sheet", 2, 2)


@pytest.mark.parametrize(
    "areas, a1, r1c1",
    [
        ([(1, 1, 1, 1)], "B2", "R2C2"),
        ([B2_C3], "B2:C3", "R2C2:R3C3"),
        ([COLUMN_D], "D1:D1048576", "R1C4:R1048576C4"),
        ([(4, 4, 0, COLUMNS - 1)], "A5:XFD5", "R5C1:R5C16384"),
        ([(ROWS - 1, ROWS - 1, COLUMNS - 1, COLUMNS - 1)], "XFD1048576", "R1048576C16384"),
        ([B2_C3, (4, 4, 4, 4)], "B2:C3,E5", "R2C2:R3C3,R5C5"),
    ],
    ids=["cell", "range", "column", "row", "last cell", "areas"],
)
def test_its_address_as_caller_gives_one(core, areas, a1, r1c1):
    r = make(core, *areas)
    assert (r.address(), r.address(r1c1=True)) == (a1, r1c1)
    assert r.address(full=True) == "[Book.xlsx]Sheet1!" + a1
    assert r.address(True, True) == "[Book.xlsx]Sheet1!" + r1c1
    assert repr(r) == f"pyinex.Range('[Book.xlsx]Sheet1!{a1}')"


def test_its_areas(core):
    r = make(core, B2_C3, (4, 4, 4, 4))
    assert [a.address(full=True) for a in r.areas] == ["[Book.xlsx]Sheet1!B2:C3", "[Book.xlsx]Sheet1!E5"]
    assert [(a.rows, a.columns) for a in r.areas] == [(2, 2), (1, 1)]
    one = make(core, B2_C3)
    assert [a.address() for a in one.areas] == ["B2:C3"]


@pytest.mark.parametrize("ask", ["rows", "columns", "values", "empty"])
def test_several_areas_have_no_single_size_or_values(core, sheet, ask):
    r = make(core, B2_C3, (4, 4, 4, 4))
    with pytest.raises(ValueError, match=r"B2:C3,E5 has 2 areas: use its areas"):
        value = getattr(r, ask)
        if callable(value):
            value()
    assert sheet.asked == []


def test_values_are_read_from_excel_with_its_sheet(core, sheet):
    sheet.cells.update({(1, 1): 1.0, (1, 2): "", (2, 2): "x"})
    r = make(core, B2_C3)
    assert r.values() == ((1.0, ""), (None, "x"))
    assert sheet.asked == [("values", 7, 1, 2, 1, 2)]
    assert make(core, (1, 1, 1, 1)).values() == 1.0


@pytest.mark.parametrize("now", [None, 8], ids=["gone", "another sheet's"])
def test_a_range_whose_sheet_is_gone_is_not_read(core, sheet, now):
    # Excel may give a deleted sheet's ID, or a closed workbook's sheet's, to another sheet.
    sheet.ids["[Book.xlsx]Sheet1"] = now
    r = make(core, B2_C3)
    for read in (r.values, r.empty, r.areas[0].values):
        with pytest.raises(RuntimeError) as raised:
            read()
        assert str(raised.value) == ("pyinex.Range: Excel no longer has the sheet [Book.xlsx]Sheet1 open; it was "
                                     "deleted or renamed, or its workbook closed")
    assert sheet.asked == []


def test_empty_tells_an_empty_cell_from_empty_text(core, sheet):
    sheet.cells.update({(1, 1): 1.0, (1, 2): "", (2, 2): "x"})
    assert make(core, B2_C3).empty() == ((False, False), (True, False))
    assert make(core, (2, 2, 1, 1)).empty() is True
    assert make(core, (1, 1, 2, 2)).empty() is False


def test_a_whole_column_reads_only_its_used_part(core, sheet):
    sheet.cells.update({(0, 3): 1.0, (2, 3): 3.0, (5, 4): "beside"})
    r = make(core, COLUMN_D)
    assert r.values() == ((1.0,), (None,), (3.0,))
    reads = [a for a in sheet.asked if a[0] == "values"]
    assert reads == [("values", 7, 0, 2, 3, 3)]
    counts = [a for a in sheet.asked if a[0] == "count"]
    assert len(counts) <= 22 and all(a[3] == ROWS - 1 for a in counts)  # Each counts to the column's end.


def test_a_whole_row_and_the_whole_sheet_read_only_their_used_part(core, sheet):
    sheet.cells.update({(4, 0): "a", (4, 2): "c", (9, 5): "f"})
    assert make(core, (4, 4, 0, COLUMNS - 1)).values() == (("a", None, "c"),)
    sheet.asked.clear()
    assert len(make(core, (0, ROWS - 1, 0, COLUMNS - 1)).values()) == 10
    assert [a for a in sheet.asked if a[0] == "values"] == [("values", 7, 0, 9, 0, 5)]


@pytest.mark.parametrize("last", [0, 1, 2, 1000, ROWS - 2])
def test_the_used_part_ends_at_the_last_row_holding_anything(core, sheet, last):
    sheet.cells.update({(last, 3): "last"})
    if last != 0:
        sheet.cells[(0, 3)] = "first"
    make(core, COLUMN_D).values()
    assert [a for a in sheet.asked if a[0] == "values"] == [("values", 7, 0, last, 3, 3)]


def test_an_empty_whole_column_reads_as_its_first_cell(core, sheet):
    assert make(core, COLUMN_D).values() is None
    assert make(core, COLUMN_D).empty() is True


@pytest.mark.parametrize("area", [(0, ROWS - 1, 3, 4), (4, 5, 0, COLUMNS - 1)], ids=["D:E", "5:6"])
def test_empty_whole_columns_or_rows_read_as_their_first_cell(core, sheet, area):
    assert make(core, area).values() is None
    assert make(core, area).empty() is True
    assert [a for a in sheet.asked if a[0] == "values"] == [("values", 7, area[0], area[0], area[2], area[2])] * 2


def test_a_range_that_is_not_whole_is_read_as_given(core, sheet):
    sheet.cells.update({(0, 3): 1.0})
    assert make(core, (0, 3, 3, 3)).values() == ((1.0,), (None,), (None,), (None,))
    assert [a[0] for a in sheet.asked] == ["values"]


# Registering and calling a function that takes a reference.

SCRIPT = """
from typing import Optional
import pyinex

@pyinex.func
def where(a, r: pyinex.Range, n: int = 0):
    return r

@pyinex.func
def quoted(r: "pyinex.Range"):
    return r

@pyinex.func
def optional(r: Optional[pyinex.Range]):
    return r

@pyinex.func
def or_none(r: pyinex.Range | None):
    return r

@pyinex.func
def spread(first: pyinex.Range, *more: pyinex.Range):
    return [first, more]

@pyinex.func
def rows_of_ranges(r: list[list[pyinex.Range]]):
    return r

@pyinex.func
def values_only(a, b):
    return a
"""


@pytest.fixture
def rows(load, capfd):
    loaded = load(SCRIPT, name="ranges.py")
    return {row[1]: row for row in loaded}, capfd.readouterr().err


def test_a_reference_parameter_registers_as_type_u(rows):
    rows, _ = rows
    assert rows["where"][2] == "QUQ"
    assert rows["quoted"][2] == "U"
    assert rows["optional"][2] == "U"
    assert rows["or_none"][2] == "U"
    assert rows["spread"][2] == "U" * 30
    assert rows["values_only"][2] == 2  # Values only: the count, as before.


def test_a_range_is_not_an_element_type(rows):
    rows, err = rows
    assert rows["rows_of_ranges"][2] == 1
    assert "Pyinex does not convert the annotation of r, list[list[" in err


def test_a_range_arrives_as_it_is(core, rows):
    rows, _ = rows
    r = make(core, B2_C3)
    assert core["call_function"](rows["where"][0], ("a", r, 2.0), True) is r
    assert core["call_function"](rows["spread"][0], (r, r), True) == [r, (r,)]


@pytest.mark.parametrize(
    "value, shown",
    [(5.0, "5.0"), ("B2", "the text 'B2'"), (((1.0, 2.0),), "a range of 1 × 2 cells"), (None, "an empty cell")],
)
def test_a_value_that_is_not_a_reference_is_refused(core, rows, capfd, value, shown):
    rows, _ = rows
    assert core["call_function"](rows["quoted"][0], (value,), True) is core["FAILED"]
    assert f"Pyinex: quoted: argument 1 (r): pyinex.Range takes a reference, not {shown}" in capfd.readouterr().err


def test_an_error_is_refused(core, rows, capfd):
    rows, _ = rows
    error = type("ExcelError", (), {"__str__": lambda self: "#REF!"})()  # As pyinex.ExcelError shows.
    assert core["call_function"](rows["where"][0], ("a", error), True) is core["FAILED"]
    assert "Pyinex: where: argument 2 (r): pyinex.Range takes a reference, not the error #REF!" in capfd.readouterr().err


def test_gaining_or_losing_the_annotation_registers_again(core, script, capfd):
    import os

    path = script("import pyinex\n\n@pyinex.func\ndef f(a):\n    return a\n", name="switch.py")
    (first,) = core["load_functions"](str(path))
    assert first[2] == 1
    stat = path.stat()
    path.write_text("import pyinex\n\n@pyinex.func\ndef f(a: pyinex.Range):\n    return a\n", encoding="utf-8")
    os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns + 2_000_000_000))
    removal, second = core["load_functions"](str(path))
    assert removal == [first[0]] and second[2] == "U"
    # Meanwhile the old registration's values never reach the new code.
    assert core["call_function"](first[0], (1.0,), True) is core["FAILED"]
    assert "f is being registered again" in capfd.readouterr().err
    path.write_text("import pyinex\n\n@pyinex.func\ndef f(a):\n    return a\n", encoding="utf-8")
    os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns + 4_000_000_000))
    removal, third = core["load_functions"](str(path))
    assert removal == [second[0]] and third[2] == 1


# A call Excel stopped (docs/architecture.md, "Recalculation").

STOPPING = """
import pyinex

@pyinex.func
def raises(r: pyinex.Range):
    r.values()

@pyinex.func
def returns(r: pyinex.Range):
    try:
        r.values()
    except RuntimeError:
        pass
    return 'kept going'
"""


@pytest.mark.parametrize("function", ["raises", "returns"])
def test_a_call_excel_stopped_fails_saying_nothing(core, load, sheet, monkeypatch, capfd, function):
    rows = {row[1]: row[0] for row in load(STOPPING, name="stopping.py")}
    pyinex = __import__("sys").modules["pyinex"]

    def stopped(*args):
        monkeypatch.setattr(pyinex, "_stopped", lambda: True)
        raise RuntimeError("stopped")

    monkeypatch.setattr(pyinex, "_range_values", stopped)
    capfd.readouterr()
    assert core["call_function"](rows[function], (make(core, B2_C3),), True) is core["FAILED"]
    assert capfd.readouterr().err == ""
