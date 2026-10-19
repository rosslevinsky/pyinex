"""Type annotations of direct functions in the fake Excel (docs/conversion.md, "Type
annotations (direct functions)"): the direct-functions test group 3. Each annotation
converting and refusing, with its console message naming the function, the argument and the
value; annotated and unannotated parameters in one function, the unannotated one also
through PyCall; an annotation Pyinex does not convert; and numpy.ndarray without NumPy, made
missing in the host's own Python.

One script holds the functions, loaded once by PyLoad in a host of the module's own. Each
returns its argument's type and repr, so a conversion is checked by both.
"""

import pytest
from conftest import evaluate, host_env
from hostdriver import MISSING, NA, VALUE, Err, Host
from test_direct_functions import wait_for
from test_reload_functions import edit, loaded

SCRIPT = """from decimal import Decimal
from typing import Optional
import numpy as np
import pyinex

def shown(x):
    return f'{type(x).__name__}: {x!r}'

@pyinex.func
def AnInt(x: int): return shown(x)
@pyinex.func
def AnFloat(x: float): return shown(x)
@pyinex.func
def AnStr(x: str): return shown(x)
@pyinex.func
def AnBool(x: bool): return shown(x)
@pyinex.func
def AnRows(x: list): return shown(x)
@pyinex.func
def AnFloatRows(x: list[list[float]]): return shown(x)
@pyinex.func
def AnIntRows(x: list[list[int]]): return shown(x)
@pyinex.func
def AnOptional(x: Optional[int]): return shown(x)
@pyinex.func
def AnString(x: "list[list[float]]"): return shown(x)
@pyinex.func
def AnArray(x: np.ndarray): return f'{x.dtype} {x.shape} {x.tolist()!r}'
@pyinex.func
def AnDecimal(x: Decimal): return shown(x)
@pyinex.func
def AnMixed(plain, n: int, *rest: float, flag: bool = False):
    return repr([shown(plain), shown(n), shown(rest), shown(flag)])
"""


@pytest.fixture(scope="module")
def ann(xllhost, pyinex_xll, tmp_path_factory):
    path = tmp_path_factory.mktemp("annotations") / "annotated.py"
    path.write_text(SCRIPT, encoding="utf-8")
    with Host(xllhost, pyinex_xll, env=host_env()) as h:
        h.open()
        h.call("PyLoad", str(path))
        h.tick()
        h.path = path
        yield h


@pytest.mark.parametrize(
    "function, value, shown",
    [
        ("AnInt", 3.0, "int: 3"),
        ("AnInt", -0.0, "int: 0"),
        ("AnInt", None, "int: 0"),
        ("AnFloat", 2.5, "float: 2.5"),
        ("AnFloat", None, "float: 0.0"),
        ("AnStr", "FALSE", "str: 'FALSE'"),
        ("AnStr", "", "str: ''"),
        ("AnStr", None, "str: ''"),
        ("AnBool", True, "bool: True"),
        ("AnBool", False, "bool: False"),
        ("AnBool", None, "bool: False"),
        ("AnRows", ((1.0, None), ("a", True)), "list: [[1.0, None], ['a', True]]"),
        ("AnRows", 5.0, "list: [[5.0]]"),
        ("AnRows", Err(NA), "list: [[ExcelError('#N/A')]]"),
        ("AnFloatRows", ((1.0, None),), "list: [[1.0, 0.0]]"),
        ("AnFloatRows", ((1.0,), (2.0,)), "list: [[1.0], [2.0]]"),
        ("AnIntRows", ((1.0, 2.0),), "list: [[1, 2]]"),
        ("AnOptional", None, "NoneType: None"),
        ("AnOptional", 3.0, "int: 3"),
        ("AnString", ((1.0,),), "list: [[1.0]]"),
        ("AnArray", ((1.0, None), (3.0, 4.0)), "float64 (2, 2) [[1.0, nan], [3.0, 4.0]]"),
        ("AnArray", 5.0, "float64 (1, 1) [[5.0]]"),
    ],
)
def test_an_annotation_converts(ann, function, value, shown):
    assert ann.call(function, value) == shown


@pytest.mark.parametrize(
    "function, value, said",
    [
        ("AnInt", 2.5, "int takes a whole number, not 2.5"),
        ("AnInt", "3", "int takes a whole number, not the text '3'"),
        ("AnInt", True, "int takes a whole number, not TRUE"),
        ("AnInt", Err(NA), "int takes a whole number, not the error #N/A"),
        ("AnInt", ((1.0, 2.0),), "int takes a whole number, not a range of 1 × 2 cells"),
        ("AnFloat", "x", "float takes a number, not the text 'x'"),
        ("AnFloat", True, "float takes a number, not TRUE"),
        ("AnStr", 3.0, "str takes text, not 3.0"),
        ("AnStr", Err(NA), "str takes text, not the error #N/A"),
        ("AnBool", "FALSE", "bool takes TRUE or FALSE, not the text 'FALSE'"),
        ("AnBool", 1.0, "bool takes TRUE or FALSE, not 1.0"),
        ("AnFloatRows", ((1.0, "x"),), "row 1, column 2: float takes a number, not the text 'x'"),
        ("AnIntRows", ((1.0,), (2.5,)), "row 2, column 1: int takes a whole number, not 2.5"),
        ("AnOptional", 2.5, "int takes a whole number, not 2.5"),
        ("AnArray", ((1.0, "x"),), "row 1, column 2: numpy.ndarray takes numbers and empty cells, not the text 'x'"),
        ("AnArray", ((1.0,), (True,)), "row 2, column 1: numpy.ndarray takes numbers and empty cells, not TRUE"),
        ("AnArray", Err(NA), "row 1, column 1: numpy.ndarray takes numbers and empty cells, not the error #N/A"),
    ],
)
def test_an_annotation_refuses_naming_the_function_argument_and_value(ann, function, value, said):
    assert ann.call(function, value) == Err(VALUE)
    message = f"Pyinex: {function}: argument 1 (x): {said}"
    assert message in wait_for(ann, message)


def test_an_annotation_pyinex_does_not_convert_is_said_at_registration_and_left_alone(ann):
    said = (f"Pyinex: AnDecimal in the script {ann.path}: Pyinex does not convert the annotation of x, "
            "<class 'decimal.Decimal'>, so x arrives unconverted")
    assert said in ann.stderr
    assert ann.call("AnDecimal", 2.5) == "float: 2.5"


def test_annotated_and_plain_parameters_in_one_function(ann):
    # plain, n, flag, then rest1 to rest27.
    result = ann.call("AnMixed", "as is", 3.0, MISSING, 1.0, 2.0)
    assert result == repr(["str: 'as is'", "int: 3", "tuple: (1.0, 2.0)", "bool: False"])
    for plain in ("text", 2.5, None, Err(NA), ((1.0, "a"),)):
        direct = eval(ann.call("AnMixed", plain, 3.0))[0]
        pycall = eval(ann.call("PyCall", str(ann.path), "AnMixed", plain, 3.0))[0]  # PyCall converts nothing.
        assert direct == pycall, plain
    assert ann.call("AnMixed", "p", 3.0, True, "x") == Err(VALUE)
    message = "Pyinex: AnMixed: argument 4 (rest1): float takes a number, not the text 'x'"
    assert message in wait_for(ann, message)
    assert ann.call("AnMixed", "p", 3.0, 1.0) == Err(VALUE)
    message = "Pyinex: AnMixed: argument 3 (flag): bool takes TRUE or FALSE, not 1.0"
    assert message in wait_for(ann, message)


def test_a_position_left_out_is_not_converted(ann):
    assert eval(ann.call("AnMixed", "p", MISSING, MISSING, 1.0))[1:] == ["NoneType: None", "tuple: (1.0,)", "bool: False"]


def test_a_changed_annotation_converts_at_the_next_call_with_no_registration(host, tmp_path):
    script = tmp_path / "changing.py"
    loaded(host, script, "import pyinex\n\n@pyinex.func\ndef AnChanging(x: int):\n    return repr(x)\n")
    assert host.call("AnChanging", 3.0) == "3"
    edit(script, "import pyinex\n\n@pyinex.func\ndef AnChanging(x: float):\n    return repr(x)\n")
    host.excel_calls()
    assert host.call("AnChanging", 3.0) == "3.0"
    assert host.excel_calls() == 0
    assert host.tick()[1:] == (0, 0)  # Excel is told nothing new: the queue was not asked to run.


def test_without_numpy_an_annotation_naming_it_is_left_alone(host, tmp_path):
    evaluate(host, "0", "import sys; sys.modules['numpy'] = None")  # import numpy now fails.
    needs = tmp_path / "needs_numpy.py"
    loaded(host, needs, "import numpy\nimport pyinex\n\n@pyinex.func\ndef AnNeeds(a: numpy.ndarray):\n    return 1\n")
    assert f"Pyinex: the script {needs} failed to load:" in wait_for(host, "ModuleNotFoundError")
    names = tmp_path / "names_numpy.py"
    loaded(host, names, "import pyinex\n\n@pyinex.func\ndef AnNames(a: 'numpy.ndarray', n: int):\n    return repr([a, n])\n")
    said = (f"Pyinex: AnNames in the script {names}: the annotation of a, 'numpy.ndarray', cannot be evaluated "
            "(NameError: name 'numpy' is not defined), so a arrives unconverted")
    assert said in wait_for(host, said)
    assert host.call("AnNames", ((1.0, 2.0),), 2.0) == "[((1.0, 2.0),), 2]"
    assert "AnNeeds" not in [r.name for r in host.registrations()]
