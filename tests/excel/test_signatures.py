"""Signatures of direct functions in real Excel (docs/architecture.md, "Signatures" to "What
can be marked"): the direct-functions test group 2, as tests/host/test_signatures.py has it,
from cells. 0 to 30 arguments, each position checked by its value, across the fourth and
fifth; defaults; *args with none, one and many; keyword-only parameters; the plan's examples;
a missing argument in the middle; more arguments than registered, which Excel takes and gives
#VALUE! without calling the function; each
refusal in a script whose other functions still register and run; a class, a
functools.partial and a built-in function. Then reloads that change a function's positions: a
call before the queue has run, and several edits within one calculation, while Excel cannot
run the queue.

A name registered stays registered for the rest of the session's Excel, so every function
here has a name starting "Xs". One script holds the functions of the first part, loaded once
by PyLoad in a workbook of the module's own.
"""

import pytest
from conftest import E, Book, shown
from test_pyload import when_registered
from test_reload_functions import loaded, when

SCRIPT = """import functools, math
import pyinex

{numbered}

@pyinex.func
def XsDefaults(a, b=2.0, c=3.0):
    return repr([a, b, c])

@pyinex.func
def XsArgs(*items):
    return repr(list(items))

@pyinex.func
def XsKw(a, *, k, scale=1.0):
    return repr([a, k, scale])

@pyinex.func
def XsPlan(x, *items, scale=1):
    return repr([x, list(items), scale])

@pyinex.func
def XsMiddle(a, b, c):
    return repr([a, b, c])

@pyinex.func
class XsMoney(float):
    def __new__(cls, amount, rate=1.0):
        return super().__new__(cls, amount * rate)

@pyinex.func
class XsPoint:
    def __init__(self, x, y):
        self.x, self.y = x, y

def _discount(price, rate, years=2.0):
    return price * (1 - rate) ** years

XsOnce = pyinex.func(functools.partial(_discount, years=1.0))

calls = 0

@pyinex.func
def XsCounted(a, b):
    global calls
    calls += 1
    return repr([a, b])

@pyinex.func
def XsCalls():
    return calls
"""


def numbered():
    """XsCount0 to XsCount30, each returning the repr of its arguments in order."""
    out = []
    for n in range(31):
        params = ", ".join(f"p{i}" for i in range(1, n + 1))
        out.append(f"@pyinex.func\ndef XsCount{n}({params}):\n    return repr([{params}])\n")
    return "\n".join(out)


@pytest.fixture(scope="module")
def sheet(excel, embedded, tmp_path_factory):
    """A workbook whose A1 loads the script by PyLoad, its functions registered."""
    folder = tmp_path_factory.mktemp("signatures")
    (folder / "xs_signatures.py").write_text(SCRIPT.format(numbered=numbered()), encoding="utf-8")
    workbook = excel.Workbooks.Add()
    while workbook.Worksheets.Count > 1:
        workbook.Worksheets(workbook.Worksheets.Count).Delete()
    # Not book.xlsx: Excel opens no two workbooks of one name, and each test's own is that.
    workbook.SaveAs(str(folder / "signatures.xlsx"), FileFormat=51)  # 51: .xlsx
    book = Book(excel, workbook)
    try:
        book.put("A1", '=PyLoad("xs_signatures.py")')
        assert when_registered(excel, "XsOnce(1,0)") == 1.0
        yield book
    finally:
        workbook.Close(SaveChanges=False)


@pytest.mark.parametrize("n", range(31))
def test_each_count_from_0_to_30_passes_each_position_its_value(sheet, n):
    values = [float(10 + i) for i in range(n)]
    formula = f"=XsCount{n}(" + ",".join(str(int(v)) for v in values) + ")"
    assert sheet.formula("B1", formula) == repr(values)


def test_the_fourth_and_fifth_positions_each_have_their_own_value(sheet):
    assert sheet.formula("B1", "=XsCount5(1,2,3,4,5)") == "[1.0, 2.0, 3.0, 4.0, 5.0]"
    assert sheet.formula("B1", "=XsCount5(,,,4,5)") == "[None, None, None, 4.0, 5.0]"
    assert sheet.formula("B1", "=XsCount5(1,2,3,,5)") == "[1.0, 2.0, 3.0, None, 5.0]"


def test_defaults(sheet):
    assert sheet.formula("B1", "=XsDefaults(1)") == "[1.0, 2.0, 3.0]"
    assert sheet.formula("B1", "=XsDefaults(1,,9)") == "[1.0, 2.0, 9.0]"
    assert sheet.formula("B1", "=XsDefaults(1,8,9)") == "[1.0, 8.0, 9.0]"


def test_args_with_none_one_and_many(sheet):
    assert sheet.formula("B1", "=XsArgs()") == "[]"
    assert sheet.formula("B1", "=XsArgs(1)") == "[1.0]"
    assert sheet.formula("B1", "=XsArgs(" + ",".join(str(i) for i in range(30)) + ")") == repr([float(i) for i in range(30)])
    assert sheet.formula("B1", "=XsArgs(1,,3)") == "[1.0, None, 3.0]"


def test_keyword_only_parameters(sheet, cleared):
    assert sheet.formula("B1", "=XsKw(1,2)") == "[1.0, 2.0, 1.0]"
    assert sheet.formula("B1", "=XsKw(1,2,3)") == "[1.0, 2.0, 3.0]"
    assert sheet.formula("B1", "=XsKw(1,,3)") == "[1.0, None, 3.0]"
    assert sheet.formula("B1", "=XsKw(1)") == E("#VALUE!")
    said = "TypeError: XsKw() missing 1 required keyword-only argument: 'k'"
    assert said in cleared(holding=said).text


def test_args_and_keyword_only_together_the_plans_examples(sheet):
    assert sheet.formula("B1", "=XsPlan(10,2)") == "[10.0, [], 2.0]"  # f(10, scale=2)
    assert sheet.formula("B1", "=XsPlan(10,,20,30)") == "[10.0, [20.0, 30.0], 1]"  # f(10, 20, 30)


def test_a_missing_argument_in_the_middle_is_none(sheet):
    assert sheet.formula("B1", "=XsMiddle(1,,3)") == "[1.0, None, 3.0]"


def test_more_arguments_than_registered_give_value_without_a_call(sheet):
    calls = shown(sheet.excel.Evaluate("XsCalls()"))
    assert sheet.formula("B1", "=XsCounted(1,2)") == "[1.0, 2.0]"
    assert shown(sheet.excel.Evaluate("XsCalls()")) == calls + 1
    assert sheet.formula("B1", "=XsCounted(1,2,3)") == E("#VALUE!")  # Excel takes the formula,
    assert shown(sheet.excel.Evaluate("XsCalls()")) == calls + 1  # but calls nothing.


def test_a_class(sheet, cleared):
    assert sheet.formula("B1", "=XsMoney(3,2)") == 6.0
    assert sheet.formula("B1", "=XsMoney(3)") == 3.0
    assert sheet.formula("B1", "=XsPoint(1,2)") == E("#VALUE!")
    assert "Pyinex: XsPoint: XsPoint cannot be returned to Excel" in cleared(holding="XsPoint cannot be returned").text


def test_a_partial(sheet):
    assert sheet.formula("B1", "=XsOnce(100,0.5)") == 50.0
    assert sheet.formula("B1", "=XsOnce(100,0.5,2)") == 25.0


REFUSED = """import functools, math
import pyinex

@pyinex.func
def XsBefore(a):
    return 'before'

@pyinex.func
def XsOver({over}):
    return 'never'

@pyinex.func
def XsOverArgs({thirty}, *rest):
    return 'never'

XsUnreadable = pyinex.func(functools.partial(max))
XsSqrt = pyinex.func(math.sqrt)

@pyinex.func
def XsAfter(a):
    return 'after'
"""


def test_each_refusal_leaves_the_rest_registered_and_working(app, book, cleared):
    text = REFUSED.format(over=", ".join(f"p{i}" for i in range(31)), thirty=", ".join(f"p{i}" for i in range(30)))
    loaded(app, book, "xs_refused.py", text, "XsAfter(1)")
    path = book.folder / "xs_refused.py"
    window = cleared(holding="sqrt in the script")
    for name, said in [
        ("XsOver", "it has 31 parameters, and Excel passes at most 30"),
        ("XsOverArgs", "it has 30 parameters and *rest, and Excel passes at most 30"),
        ("XsUnreadable", "its signature cannot be read"),
        ("sqrt", "a built-in function cannot be marked; wrap it in a function of your own"),
    ]:
        assert f"Pyinex: {name} in the script {path} is not registered: {said}" in window.text, window.text
    for name in ("XsOver(1)", "XsOverArgs(1)", "XsUnreadable(1)", "XsSqrt(1)"):
        assert shown(app.Evaluate(name)) == E("#NAME?"), name
    assert book.formula("B1", "=XsBefore(1)") == "before"
    assert book.formula("B2", "=XsAfter(1)") == "after"


# Reloads that change a function's positions (docs/architecture.md, "Reloading a script").

ONE = "import pyinex\n\n@pyinex.func\ndef {0}({1}):\n    return repr([{2}])\n"


def test_positional_to_args_is_registered_again(app, book, cleared):
    loaded(app, book, "xs_shape.py", ONE.format("XsShape", "a, b", "a, b"), "XsShape(0,0)")
    book.put("A1", 1)
    assert book.formula("B1", "=XsShape(A1,2)") == "[1.0, 2.0]"
    book.script("xs_shape.py", ONE.format("XsShape", "a, *rest", "a, rest"))
    book.put("A1", 2)  # A call before the queue has run: the new code never sees its arguments.
    assert book.value("B1") == E("#VALUE!")
    assert "Pyinex: XsShape is being registered again" in cleared(holding="being registered again").text
    assert when(lambda: shown(app.Evaluate("XsShape(1,2,3)")) == "[1.0, (2.0, 3.0)]")
    book.put("A1", 3)  # An ordinary recalculation reaches the new registration.
    assert book.value("B1") == "[3.0, (2.0,)]"


EDITOR = """from pathlib import Path

def write(path, parameters, body, trigger):
    # Writes xs_edits.py with these parameters when trigger is 7, then returns trigger.
    if trigger == 7:
        Path(path).write_text(
            "import pyinex\\n\\n@pyinex.func\\ndef XsEdits(" + parameters + "):\\n    return repr([" + body + "])\\n\\n"
            + "@pyinex.func\\ndef XsEditsOther(a):\\n    return a\\n", encoding="utf-8")
    return trigger
"""


def test_several_edits_of_the_positions_in_one_calculation(app, book):
    loaded(app, book, "xs_edits.py",
           ONE.format("XsEdits", "a", "a") + "\n@pyinex.func\ndef XsEditsOther(a):\n    return a\n", "XsEdits(0)")
    book.script("xs_editor.py", EDITOR)
    script = book.folder / "xs_edits.py"
    book.put("A1", 0)
    book.put("C1", f'=PyCall("xs_editor.py", "write", "{script}", "a, *rest", "a, rest", A1)')
    book.put("C2", "=XsEditsOther(C1)")
    book.put("C3", f'=PyCall("xs_editor.py", "write", "{script}", "a, b, *rest, k=3.0", "a, b, rest, k", C2)')
    book.put("C4", "=XsEditsOther(C3)")
    assert shown(app.Evaluate("XsEdits(1)")) == "[1.0]"
    book.put("A1", 7)  # One calculation: C1 to C4 in order.
    assert book.values("C1:C4") == [[7.0], [7.0], [7.0], [7.0]]
    assert when(lambda: shown(app.Evaluate("XsEdits(1,2,9,4,5)")) == "[1.0, 2.0, (4.0, 5.0), 9.0]")
    assert book.formula("B1", "=XsEdits(1,2)") == "[1.0, 2.0, (), 3.0]"
