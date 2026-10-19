"""Type annotations of direct functions in real Excel (docs/conversion.md, "Type annotations
(direct functions)"): the direct-functions test group 3, as tests/host/test_annotations.py
has it, from cells. Each annotation converting and refusing, with its console message naming
the function, the argument and the value; annotated and unannotated parameters in one
function, the unannotated one also through PyCall; an annotation Pyinex does not convert; and
a changed annotation, which needs no registration. Every real-Excel run has NumPy, so the
rows without it are tested in tests/unit and tests/host only.

A name registered stays registered for the rest of the session's Excel, so every function
here has a name starting "Xa". One script holds the functions, loaded once by PyLoad in a
workbook of the module's own, whose D1:E2 holds 1, an empty cell, 3 and 4. Z99 is empty.
"""

import time

import pytest
from conftest import E, Book, pycall, registered_types, shown
from test_pyload import when_registered
from test_reload_functions import loaded

SCRIPT = """from decimal import Decimal
from typing import Optional
import numpy as np
import pyinex

def shown(x):
    return f'{type(x).__name__}: {x!r}'

@pyinex.func
def XaInt(x: int): return shown(x)
@pyinex.func
def XaFloat(x: float): return shown(x)
@pyinex.func
def XaStr(x: str): return shown(x)
@pyinex.func
def XaBool(x: bool): return shown(x)
@pyinex.func
def XaRows(x: list): return shown(x)
@pyinex.func
def XaFloatRows(x: list[list[float]]): return shown(x)
@pyinex.func
def XaIntRows(x: list[list[int]]): return shown(x)
@pyinex.func
def XaOptional(x: Optional[int]): return shown(x)
@pyinex.func
def XaString(x: "list[list[float]]"): return shown(x)
@pyinex.func
def XaArray(x: np.ndarray): return f'{x.dtype} {x.shape} {x.tolist()!r}'
@pyinex.func
def XaDecimal(x: Decimal): return shown(x)
@pyinex.func
def XaMixed(plain, n: int, *rest: float, flag: bool = False):
    return repr([shown(plain), shown(n), shown(rest), shown(flag)])
"""


@pytest.fixture(scope="module")
def sheet(excel, embedded, tmp_path_factory):
    folder = tmp_path_factory.mktemp("annotations")
    (folder / "xa_annotated.py").write_text(SCRIPT, encoding="utf-8")
    workbook = excel.Workbooks.Add()
    while workbook.Worksheets.Count > 1:
        workbook.Worksheets(workbook.Worksheets.Count).Delete()
    # Not book.xlsx: Excel opens no two workbooks of one name, and each test's own is that.
    workbook.SaveAs(str(folder / "annotations.xlsx"), FileFormat=51)  # 51: .xlsx
    book = Book(excel, workbook)
    try:
        for cell, value in (("D1", 1), ("D2", 3), ("E2", 4)):
            book.put(cell, value)
        book.put("A1", '=PyLoad("xa_annotated.py")')
        assert when_registered(excel, "XaInt(3)") == "int: 3"
        book.path = folder / "xa_annotated.py"
        yield book
    finally:
        workbook.Close(SaveChanges=False)


@pytest.mark.parametrize(
    "function, argument, shows",
    [
        ("XaInt", "3", "int: 3"),
        ("XaInt", "-0", "int: 0"),
        ("XaInt", "Z99", "int: 0"),
        ("XaFloat", "2.5", "float: 2.5"),
        ("XaFloat", "Z99", "float: 0.0"),
        ("XaStr", '"FALSE"', "str: 'FALSE'"),
        ("XaStr", '""', "str: ''"),
        ("XaStr", "Z99", "str: ''"),
        ("XaBool", "TRUE", "bool: True"),
        ("XaBool", "FALSE", "bool: False"),
        ("XaBool", "Z99", "bool: False"),
        ("XaRows", '{1,"a";TRUE,2}', "list: [[1.0, 'a'], [True, 2.0]]"),
        ("XaRows", "D1:E1", "list: [[1.0, None]]"),
        ("XaRows", "5", "list: [[5.0]]"),
        ("XaRows", "NA()", "list: [[ExcelError('#N/A')]]"),
        ("XaFloatRows", "D1:E1", "list: [[1.0, 0.0]]"),
        ("XaFloatRows", "D1:D2", "list: [[1.0], [3.0]]"),
        ("XaIntRows", "{1,2}", "list: [[1, 2]]"),
        ("XaOptional", "Z99", "NoneType: None"),
        ("XaOptional", "3", "int: 3"),
        ("XaString", "D2:E2", "list: [[3.0, 4.0]]"),
        ("XaArray", "D1:E2", "float64 (2, 2) [[1.0, nan], [3.0, 4.0]]"),
        ("XaArray", "5", "float64 (1, 1) [[5.0]]"),
        ("XaDecimal", "2.5", "float: 2.5"),  # An annotation Pyinex does not convert: left alone.
    ],
)
def test_an_annotation_converts(sheet, function, argument, shows):
    assert sheet.formula("B1", f"={function}({argument})") == shows


@pytest.mark.parametrize(
    "function, argument, said",
    [
        ("XaInt", "2.5", "int takes a whole number, not 2.5"),
        ("XaInt", '"3"', "int takes a whole number, not the text '3'"),
        ("XaInt", "TRUE", "int takes a whole number, not TRUE"),
        ("XaInt", "NA()", "int takes a whole number, not the error #N/A"),
        ("XaInt", "{1,2}", "int takes a whole number, not a range of 1 × 2 cells"),
        ("XaFloat", '"x"', "float takes a number, not the text 'x'"),
        ("XaFloat", "TRUE", "float takes a number, not TRUE"),
        ("XaStr", "3", "str takes text, not 3.0"),
        ("XaStr", "NA()", "str takes text, not the error #N/A"),
        ("XaBool", '"FALSE"', "bool takes TRUE or FALSE, not the text 'FALSE'"),
        ("XaBool", "1", "bool takes TRUE or FALSE, not 1.0"),
        ("XaFloatRows", '{1,"x"}', "row 1, column 2: float takes a number, not the text 'x'"),
        ("XaIntRows", "{1;2.5}", "row 2, column 1: int takes a whole number, not 2.5"),
        ("XaOptional", "2.5", "int takes a whole number, not 2.5"),
        ("XaArray", '{1,"x"}', "row 1, column 2: numpy.ndarray takes numbers and empty cells, not the text 'x'"),
        ("XaArray", "{1;TRUE}", "row 2, column 1: numpy.ndarray takes numbers and empty cells, not TRUE"),
        ("XaArray", "NA()", "row 1, column 1: numpy.ndarray takes numbers and empty cells, not the error #N/A"),
    ],
)
def test_an_annotation_refuses_naming_the_function_argument_and_value(sheet, cleared, function, argument, said):
    assert sheet.formula("B1", f"={function}({argument})") == E("#VALUE!")
    message = f"Pyinex: {function}: argument 1 (x): {said}"
    window = cleared(holding=message)
    assert window is not None and message in window.text, window and window.text


def test_annotated_and_plain_parameters_in_one_function(sheet, cleared):
    # plain, n, flag, then rest1 to rest27.
    assert sheet.formula("B1", '=XaMixed("as is",3,,1,2)') == repr(["str: 'as is'", "int: 3", "tuple: (1.0, 2.0)", "bool: False"])
    for plain in ('"text"', "2.5", "Z99", "NA()", '{1,"a"}', "D1:E2"):
        direct = eval(sheet.formula("B1", f"=XaMixed({plain},3)"))[0]
        by_pycall = eval(sheet.formula("B2", pycall(str(sheet.path), "XaMixed", plain, "3")))[0]  # PyCall converts nothing.
        assert direct == by_pycall, plain
    assert sheet.formula("B1", '=XaMixed("p",3,TRUE,"x")') == E("#VALUE!")
    message = "Pyinex: XaMixed: argument 4 (rest1): float takes a number, not the text 'x'"
    assert message in cleared(holding=message).text
    assert sheet.formula("B1", '=XaMixed("p",3,1)') == E("#VALUE!")
    message = "Pyinex: XaMixed: argument 3 (flag): bool takes TRUE or FALSE, not 1.0"
    assert message in cleared(holding=message).text


def test_a_position_left_out_is_not_converted(sheet):
    assert eval(sheet.formula("B1", '=XaMixed("p",,,1)'))[1:] == ["NoneType: None", "tuple: (1.0,)", "bool: False"]


def test_an_annotation_pyinex_does_not_convert_is_said_at_registration_and_left_alone(app, book, cleared):
    loaded(app, book, "xa_decimal.py",
           "from decimal import Decimal\nimport pyinex\n\n@pyinex.func\ndef XaDecimalOwn(x: Decimal):\n    return repr(x)\n",
           "XaDecimalOwn(1)")
    said = (f"Pyinex: XaDecimalOwn in the script {book.folder / 'xa_decimal.py'}: Pyinex does not convert the "
            "annotation of x, <class 'decimal.Decimal'>, so x arrives unconverted")
    window = cleared(holding=said)
    assert window is not None and said in window.text, window and window.text
    assert book.formula("B1", "=XaDecimalOwn(2.5)") == "2.5"


def test_a_changed_annotation_converts_at_the_next_recalculation_with_no_registration(app, book, pytestconfig):
    from pathlib import Path

    xll = Path(pytestconfig.getoption("--excel")).resolve()
    loaded(app, book, "xa_changing.py", "import pyinex\n\n@pyinex.func\ndef XaChanging(x: int):\n    return repr(x)\n",
           "XaChanging(1)")
    book.put("A1", 3)
    assert book.formula("B1", "=XaChanging(A1)") == "3"
    before = registered_types(app, xll)
    book.script("xa_changing.py", "import pyinex\n\n@pyinex.func\ndef XaChanging(x: float):\n    return repr(x)\n")
    book.put("A1", 4)
    assert book.value("B1") == "4.0"
    time.sleep(1)
    assert registered_types(app, xll) == before  # Excel was told nothing.
    assert shown(app.Evaluate("XaChanging(5)")) == "5.0"
