"""The pyinex module's caller functions and PyInfo in real Excel (docs/architecture.md, "The
pyinex module" and "PyInfo"), with the 2010 caller checks (PythonExtensionTest).
"""

import re
import time
from pathlib import Path

import pytest
from conftest import ITEMS, ROOT, E, pycall, typed


@pytest.fixture
def module(book):
    book.copy("module.py")
    return book


def caller(book, cell, *options):
    return book.formula(cell, pycall("module.py", "caller", *options))


@pytest.mark.parametrize(
    "options, address",
    [
        ((), "C12"),
        (("TRUE",), "R12C3"),
        (("FALSE", "TRUE"), "[book.xlsx]Sheet1!C12"),
        (("TRUE", "TRUE"), "[book.xlsx]Sheet1!R12C3"),
    ],
)
def test_caller_in_both_address_styles(module, options, address):
    assert caller(module, "C12", *options) == address


def test_sheet(module):
    assert module.formula("B2", pycall("module.py", "sheet")) == "[book.xlsx]Sheet1"


def test_the_calling_sheet_not_the_active_one(module, workbook, tmp_path):
    # The formula's sheet and workbook, not whichever Excel shows: another sheet of the
    # workbook is made active, then another workbook.
    other = module.workbook.Worksheets.Add(After=module.sheet)
    other.Name = "Other"
    other.Activate()
    assert caller(module, "C12", "FALSE", "TRUE") == "[book.xlsx]Sheet1!C12"
    workbook(tmp_path / "elsewhere" / "active.xlsx").sheet.Activate()
    assert module.workbook.Application.ActiveWorkbook.Name == "active.xlsx"
    assert caller(module, "D4", "FALSE", "TRUE") == "[book.xlsx]Sheet1!D4"
    assert module.formula("E5", pycall("module.py", "sheet")) == "[book.xlsx]Sheet1"


def test_a_sheet_name_with_spaces_is_not_quoted(module):
    module.sheet.Name = "My Sheet"
    assert caller(module, "B2", "FALSE", "TRUE") == "[book.xlsx]My Sheet!B2"


def test_caller_at_the_grids_far_corner(module):
    assert caller(module, "XFD1048576") == "XFD1048576"
    assert caller(module, "XFD1048575", "TRUE") == "R1048575C16384"


def test_caller_of_a_legacy_array_formula_is_its_range(module):
    assert module.array("C12:E14", pycall("module.py", "caller")) == [["C12:E14"] * 3] * 3
    module.array("G1:H2", pycall("module.py", "caller", "TRUE", "TRUE"))
    assert module.value("H2") == "[book.xlsx]Sheet1!R1C7:R2C8"


def test_caller_of_a_spilling_formula_is_its_cell(module):
    module.formula("B3", '=PyCall("module.py", "caller") & {"", ""}')
    assert module.values("B3:C3") == [["B3", "B3"]]


@pytest.mark.parametrize("name", ["caller", "sheet"])
def test_caller_and_sheet_refuse_outside_a_cell(module, app, name):
    # Application.Run calls PyCall from no cell.
    assert app.Run("PyCall", str(module.folder / "module.py"), "refused", name) == "RuntimeError"


def test_evaluate_calls_as_if_from_the_active_cell(module, app):
    # Excel's Application.Evaluate gives its calls the active cell as their caller.
    module.sheet.Range("E5").Select()
    path = module.folder / "module.py"
    assert app.Evaluate(f'PyCall("{path}", "caller")') == "E5"
    assert app.Evaluate(f'PyCall("{path}", "sheet")') == "[book.xlsx]Sheet1"


@pytest.mark.parametrize("name", ["caller", "sheet", "esc_pressed"])
def test_caller_sheet_and_esc_pressed_refuse_on_another_thread(module, name):
    assert module.formula("A1", pycall("module.py", "refused_on_a_thread", f'"{name}"')) == "RuntimeError"


def test_caller_sheet_and_esc_pressed_refuse_between_calls(module):
    go, done = module.folder / "go", module.folder / "done"
    assert module.formula("A1", pycall("module.py", "refused_between_calls", f'"{go}"', f'"{done}"')) == "started"
    go.write_text("go", encoding="utf-8")  # No call is running now.
    deadline = time.monotonic() + 10
    while not done.exists() and time.monotonic() < deadline:
        time.sleep(0.05)
    assert done.read_text("utf-8").splitlines() == ["caller: RuntimeError", "sheet: RuntimeError", "esc_pressed: RuntimeError"]


@pytest.mark.parametrize(
    "function, address",
    [
        ("CallerA1Test", "D7"),
        ("CallerA1FullTest", "[book.xlsx]Sheet1!D7"),
        ("CallerR1C1Test", "R7C4"),
        ("CallerR1C1FullTest", "[book.xlsx]Sheet1!R7C4"),
        ("CallerSheetTest", "[book.xlsx]Sheet1"),
    ],
)
def test_python_extension_test_2010(module, function, address):
    assert module.formula("D7", pycall("module.py", function)) == address


# PyInfo.


def version():
    return re.search(r"project\(pyinex VERSION (\S+)", (ROOT / "CMakeLists.txt").read_text("utf-8")).group(1)


def test_every_pyinfo_item(book, pytestconfig, expected):
    info = {item: book.formula("A1", f'=PyInfo("{item}")') for item in ITEMS}
    assert info["status"] == "running"
    assert Path(info["pyinex"]) == Path(pytestconfig.getoption("--excel")).resolve()
    assert info["addin_version"] == version()
    assert re.fullmatch(r"python3\d+\.dll", Path(info["python"]).name, re.IGNORECASE) and Path(info["python"]).is_file()
    assert Path(info["executable"]).is_file() and Path(info["prefix"]).is_dir()
    assert isinstance(info["venv"], bool)
    if expected is not None:
        assert info["executable"].lower() == expected["executable"].lower()
        assert info["version"] == expected["version"]
        assert info["venv"] is expected["venv"]
        # A virtual environment's prefix is its own folder; a Python's, its own.
        prefix = Path(expected["executable"]).parent
        assert Path(info["prefix"]) == (prefix.parent if expected["venv"] else prefix)
        # The DLL of that Python: beside its python.exe, or a virtual environment's base
        # Python's, named in its pyvenv.cfg.
        home = prefix
        if expected["venv"]:
            cfg = (prefix.parent / "pyvenv.cfg").read_text("utf-8")
            home = Path(re.search(r"^home\s*=\s*(.+?)\s*$", cfg, re.MULTILINE).group(1))
        minor = expected["version"].split(".")[1]
        assert Path(info["python"]) == home / f"python3{minor}.dll"


def test_pyinfo_names_ignore_case(book):
    assert book.formula("A1", '=PyInfo("Status")') == "running"
    assert book.formula("A2", '=PyInfo("VERSION")') == book.formula("A3", '=PyInfo("version")')


def test_pyinfo_with_no_item_is_the_whole_table(book):
    book.formula("A1", "=PyInfo()")
    table = book.values("A1:B8")
    assert [row[0] for row in table] == ITEMS
    for row in table:
        assert typed([[row[1]]]) == typed([[book.formula("D1", f'=PyInfo("{row[0]}")')]]), row[0]
    assert book.values("A9:B9") == [[None, None]]


@pytest.mark.parametrize("item", ['"colour"', "1"])
def test_an_unknown_pyinfo_item(book, item):
    assert book.formula("A1", f"=PyInfo({item})") == E("#VALUE!")
