"""The example workbooks (examples/*.xlsx) in real Excel.

Every formula on every sheet is a check: the cells just to the right of its result hold
the value it should show (scripts/make_examples.py); for a formula entered into a range of
cells, the expected block of the same size sits just to its right. tests/unit/test_examples.py
checks the workbooks' shape without Excel.

direct_functions.xlsx loads its script with PyLoad, so its direct functions register a
moment after the workbook opens (docs/architecture.md, "When a workbook opens"); its cells
are read once none shows #NAME?.
"""

import time

import pytest
from conftest import ROOT, E, rows_of, same, shown

WORKBOOKS = sorted(p.name for p in (ROOT / "examples").glob("*.xlsx"))
PRINTED = "Hello from Python, Excel."  # What examples/addin.py prints.


def shows_name_error(workbook):
    """Whether any cell of the workbook shows #NAME?."""
    return any(shown(v) == E("#NAME?") for sheet in workbook.Worksheets for row in rows_of(sheet.UsedRange.Value2) for v in row)


def check(app, path):
    """Opens and recalculates a workbook; returns (number of checks, failures as text)."""
    checked, failures = 0, []
    workbook = app.Workbooks.Open(str(path), ReadOnly=True)
    try:
        app.CalculateFull()
        deadline = time.monotonic() + 20
        while shows_name_error(workbook) and time.monotonic() < deadline:
            time.sleep(0.1)
        for sheet in workbook.Worksheets:
            used = sheet.UsedRange
            for r in range(used.Row, used.Row + used.Rows.Count):
                for c in range(used.Column, used.Column + used.Columns.Count):
                    cell = sheet.Cells(r, c)
                    if not cell.HasFormula:
                        continue
                    result = cell
                    if cell.HasArray:
                        result = cell.CurrentArray
                        if (result.Row, result.Column) != (r, c):
                            continue  # Checked from the array's first cell.
                    rows, columns = result.Rows.Count, result.Columns.Count
                    expected = sheet.Range(sheet.Cells(r, c + columns), sheet.Cells(r + rows - 1, c + 2 * columns - 1))
                    got, want = rows_of(result.Value2), rows_of(expected.Value2)
                    checked += 1
                    if not all(same(g, w) for gr, wr in zip(got, want) for g, w in zip(gr, wr)):
                        failures.append(
                            f"{sheet.Name}!{cell.Address.replace('$', '')} {cell.Formula}: got "
                            f"{[[shown(v) for v in row] for row in got]}, expected "
                            f"{[[shown(v) for v in row] for row in want]}"
                        )
    finally:
        workbook.Close(SaveChanges=False)
    return checked, failures


@pytest.mark.parametrize("name", WORKBOOKS)
def test_example_workbook_calculates_as_its_sheets_say(app, examples, name):
    checked, failures = check(app, examples / name)
    assert checked > 0, f"{name} has no formulas to check"
    assert not failures, "\n".join(failures)


def test_addin_example_shows_its_output_in_the_console(app, examples, console):
    check(app, examples / "addin.xlsx")
    window = console(holding=PRINTED)
    assert window is not None, "Excel has no Pyinex console window"
    assert window.title == "Pyinex console"
    assert window.shown
    assert PRINTED in window.text
