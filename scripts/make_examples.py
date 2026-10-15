"""Writes the example workbooks, examples/*.xlsx, from the tables below.

Each workbook calls the script of the same name beside it (examples/*.py, written by
hand). On every sheet, each formula is a check: the cells just to the right of its result
hold the value it should show, which tests/excel/test_examples.py compares in real Excel and
tests/unit/test_examples.py checks the shape of. Run it on any machine with openpyxl after
changing a table, and commit the workbooks it writes:

    python scripts/make_examples.py
"""

import argparse
import datetime
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

ROOT = Path(__file__).resolve().parent.parent
VERSION = re.search(r"project\(pyinex VERSION ([\d.]+)", (ROOT / "CMakeLists.txt").read_text()).group(1)

# Columns: what the check shows, its formula as text, its result, then the expected value.
RESULT_COLUMN = 3


class Text(str):
    """An expected value that is text, even if it reads like an error, such as "#N/A"."""


@dataclass
class Place:
    """Where a check's result is, for expected values that depend on it."""

    row: int
    column: int
    rows: int
    columns: int

    @property
    def a1(self):
        return f"{get_column_letter(self.column)}{self.row}"

    @property
    def r1c1(self):
        return f"R{self.row}C{self.column}"

    @property
    def ref(self):
        """The result's range, as pyinex.caller() gives it."""
        last = f"{get_column_letter(self.column + self.columns - 1)}{self.row + self.rows - 1}"
        return self.a1 if (self.rows, self.columns) == (1, 1) else f"{self.a1}:{last}"


@dataclass
class Check:
    what: str
    formula: str  # Without the "=".
    expected: Any  # A value; a list of rows for a block, entered as an array formula; or a function of the Place.
    rows: int = 1  # The block's size, when expected is a function.
    columns: int = 1


@dataclass
class Sheet:
    name: str
    title: str
    note: str
    inputs: list = field(default_factory=list)  # (label cell, label, top-left cell, rows of values)
    checks: list = field(default_factory=list)


BASICS = [
    Sheet(
        "Basics",
        "Calling a function, and single values",
        "PyCall(file, function, arguments...) calls a function in basics.py, beside this workbook.",
        inputs=[("A4", "Inputs:", "B4", [["A number", "Text", "TRUE", "An empty cell", "An error"], [3, "text", True, None, "#N/A"]])],
        checks=[
            Check("A function's result", 'PyCall("basics.py","add",2,3)', 5),
            Check("A number arrives as float", 'PyCall("basics.py","kind",B5)', "float"),
            Check("Text arrives as str", 'PyCall("basics.py","kind",C5)', "str"),
            Check("TRUE arrives as bool", 'PyCall("basics.py","kind",D5)', "bool"),
            Check("An empty cell arrives as None", 'PyCall("basics.py","kind",E5)', "NoneType"),
            Check("An error arrives as pyinex.ExcelError", 'PyCall("basics.py","kind",F5)', "ExcelError"),
            Check("... whose text names it", 'PyCall("basics.py","error_text",F5)', Text("#N/A")),
            Check("A default argument", 'PyCall("basics.py","greet","Ada")', "Hello, Ada!"),
            Check("Every argument given", 'PyCall("basics.py","greet","Ada","Welcome")', "Welcome, Ada!"),
            Check("*args", 'PyCall("basics.py","count",1,2,3)', 3),
            Check("One left out in the middle arrives as None", 'PyCall("basics.py","count",1,,3)', 3),
            Check("Trailing ones left out are dropped", 'PyCall("basics.py","count",1,2)', 2),
            Check("A builtin (biggest = max)", 'PyCall("basics.py","biggest",3,9,4)', 9),
            Check("A class (whole = int)", 'PyCall("basics.py","whole",3.7)', 3),
            Check("Returning a bool", 'PyCall("basics.py","is_even",4)', True),
            Check("Unicode text, both ways", 'PyCall("basics.py","shout","héllo, κόσμε")', "héllo, κόσμε".upper()),
            Check("An emoji is one character", 'PyCall("basics.py","length","😀")', 1),
            Check("None shows as nothing (empty text)", 'PyCall("basics.py","nothing")=""', True),
            Check("NaN shows as #NUM!", 'PyCall("basics.py","not_a_number")', "#NUM!"),
            Check("Returning an ExcelError", 'PyCall("basics.py","not_available")', "#N/A"),
            Check("An exception gives #VALUE!", 'PyCall("basics.py","fail")', "#VALUE!"),
            Check("A function that does not exist", 'PyCall("basics.py","no_such_function")', "#VALUE!"),
            Check("A script that does not exist", 'PyCall("no_such_script.py","add",1,2)', "#VALUE!"),
        ],
    )
]

RANGES = [
    Sheet(
        "Ranges",
        "Ranges in, blocks out",
        "A range arrives as rows of values; a list of rows fills a block. "
        "Formulas returning a block are entered into a range of cells.",
        inputs=[
            ("B4", "A:", "B5", [[1, 2], [3, 4]]),
            ("E4", "B:", "E5", [[5, 6], [7, 8]]),
            ("H4", "With empty cells:", "H5", [[1, None], [None, 4]]),
        ],
        checks=[
            Check("One cell is a plain value", 'PyCall("ranges.py","size",B5)', "one value"),
            Check("A row", 'PyCall("ranges.py","size",B5:C5)', "1 x 2"),
            Check("A column", 'PyCall("ranges.py","size",B5:B6)', "2 x 1"),
            Check("A block", 'PyCall("ranges.py","size",B5:C6)', "2 x 2"),
            Check("An array constant", 'PyCall("ranges.py","size",{1,2,3;4,5,6})', "2 x 3"),
            Check("Summing a block", 'PyCall("ranges.py","total",B5:C6)', 10),
            Check("Empty cells arrive as None", 'PyCall("ranges.py","empty_cells",H5:I6)', 2),
            Check("... and can be skipped", 'PyCall("ranges.py","total",H5:I6)', 5),
            Check("Transposing A", 'PyCall("ranges.py","transpose",B5:C6)', [[1, 3], [2, 4]]),
            Check("A times B", 'PyCall("ranges.py","multiply",B5:C6,E5:F6)', [[19, 22], [43, 50]]),
            Check("A flat list fills a row", 'PyCall("ranges.py","as_row",B5:C6)', [[1, 2, 3, 4]]),
            Check("One-value rows fill a column", 'PyCall("ranges.py","as_column",B5:C6)', [[1], [2], [3], [4]]),
            Check("Short rows are padded with blanks", 'PyCall("ranges.py","ragged")', [[1, 2, 3], [4, None, None]]),
        ],
    )
]


def module_checks(sheet):
    full = f"[pyinex_module.xlsx]{sheet}"
    return [
        Check("The calling cell", 'PyCall("pyinex_module.py","where")', lambda p: p.a1),
        Check("... in R1C1 style", 'PyCall("pyinex_module.py","where_r1c1")', lambda p: p.r1c1),
        Check("... with its workbook and sheet", 'PyCall("pyinex_module.py","where_full")', lambda p: f"{full}!{p.a1}"),
        Check("The calling sheet", 'PyCall("pyinex_module.py","which_sheet")', full),
        Check(
            "In a range of cells, that range",
            'PyCall("pyinex_module.py","where_block")',
            lambda p: [[p.ref] * 2] * 2,
            rows=2,
            columns=2,
        ),
    ]


MODULE = [
    Sheet(
        "Δοκιμή",
        "The pyinex module",
        "Scripts import pyinex to ask which cell is calling. This sheet's name is Greek, to show Unicode names.",
        checks=[
            *module_checks("Δοκιμή"),
            Check("Checking for Esc in a long calculation", 'PyCall("pyinex_module.py","slow_count",0.5)', "finished"),
        ],
    ),
    Sheet(
        "Second sheet",
        "The pyinex module, on a sheet whose name has a space",
        "Full addresses carry no quotes, even around a name with a space.",
        checks=module_checks("Second sheet"),
    ),
]

ADDIN = [
    Sheet(
        "Add-in",
        "The add-in's other functions",
        "PyInfo reports which Python is loaded, PyReloadOnChange controls reloading, and PyConsole "
        "shows the console that print() writes to.",
        checks=[
            Check("Python is running", 'PyInfo("status")', "running"),
            Check("Item names ignore case", 'PyInfo("STATUS")', "running"),
            Check("The add-in's version", 'PyInfo("addin_version")', VERSION),
            Check("The Python DLL in use", 'ISNUMBER(SEARCH("python3",PyInfo("python")))', True),
            Check("PyInfo() is a table of every item", 'ROWS(PyInfo())', 8),
            Check("... with names and values", 'COLUMNS(PyInfo())', 2),
            Check("An unknown item", 'PyInfo("nonsense")', "#VALUE!"),
            Check("Reloading changed scripts is on", "PyReloadOnChange()", True),
            Check("Show the console at 40,40, 800 x 400 pixels", "PyConsole(TRUE,FALSE,40,40,800,400)", True),
            Check("print() writes to the console", 'PyCall("addin.py","hello","Excel")', "printed"),
        ],
    )
]

WORKBOOKS = {"basics": BASICS, "ranges": RANGES, "pyinex_module": MODULE, "addin": ADDIN}


def write_sheet(ws, sheet):
    bold = Font(bold=True)
    ws["A1"] = sheet.title
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = sheet.note
    row = 4
    for label_cell, label, top_left, values in sheet.inputs:
        ws[label_cell] = label
        ws[label_cell].font = bold
        top = ws[top_left]
        for r, line in enumerate(values):
            for c, value in enumerate(line):
                ws.cell(top.row + r, top.column + c, value)
        row = max(row, top.row + len(values) + 1)
    for column, heading in enumerate(["What it shows", "Formula", "Result", "Expected"], start=1):
        ws.cell(row, column, heading).font = bold
    row += 1
    for check in sheet.checks:
        expected = check.expected
        if callable(expected):
            expected = expected(Place(row, RESULT_COLUMN, check.rows, check.columns))
        block = expected if isinstance(expected, list) else None
        ws.cell(row, 1, check.what)
        text = ws.cell(row, 2, f"{{={check.formula}}}" if block else f"={check.formula}")
        text.data_type = "s"  # Shown as text, not calculated.
        result = ws.cell(row, RESULT_COLUMN)
        if block:
            place = Place(row, RESULT_COLUMN, len(block), len(block[0]))
            result.value = ArrayFormula(place.ref, f"={check.formula}")
            for r, line in enumerate(block):
                for c, value in enumerate(line):
                    ws.cell(row + r, RESULT_COLUMN + place.columns + c, value)
            row += place.rows
        else:
            result.value = f"={check.formula}"
            cell = ws.cell(row, RESULT_COLUMN + 1, expected)
            if isinstance(expected, Text):
                cell.data_type = "s"  # openpyxl would otherwise store "#N/A" as an error.
            row += 1
    for column, width in zip("ABCDEFGHI", [44, 52, 22, 22, 22, 22, 22, 22, 22]):
        ws.column_dimensions[column].width = width


def build(out_dir):
    """Writes every workbook into out_dir; returns their paths."""
    paths = []
    for name, sheets in WORKBOOKS.items():
        wb = Workbook()
        wb.remove(wb.active)
        for sheet in sheets:
            write_sheet(wb.create_sheet(sheet.name), sheet)
        # Fixed, so the workbooks change only when a table does.
        wb.properties.creator = "scripts/make_examples.py"
        wb.properties.created = wb.properties.modified = datetime.datetime(2026, 1, 1)
        path = Path(out_dir) / f"{name}.xlsx"
        wb.save(path)
        paths.append(path)
    return paths


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--out", default=ROOT / "examples", help="folder to write into (default: examples/)")
    for path in build(parser.parse_args().out):
        print(path)


if __name__ == "__main__":
    main()
