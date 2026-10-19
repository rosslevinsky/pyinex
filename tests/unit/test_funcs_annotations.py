"""Type annotations of direct functions in pure Python (src/core/python/funcs.py;
docs/conversion.md, "Type annotations (direct functions)"): every row of the table, each
refusal with its message, the spellings accepted, annotations Pyinex does not convert, and
numpy.ndarray with NumPy and without it.

Values are given as the core gives them, already converted by the rules for values: numbers
as floats, a range as a tuple of rows, an empty cell as None. NumPy must be installed: these
tests fail without it, rather than skip (docs/testing.md).
"""

import math
import sys

import numpy
import pytest

# The fixtures core and load are in conftest.py.

SCRIPT = """
import typing
from decimal import Decimal
from typing import Any, List, Optional, Union
import numpy as np
import pyinex

def _shown(x):
    if isinstance(x, list):
        return [_shown(r) for r in x]
    return [type(x).__name__, x]

@pyinex.func
def a_int(x: int): return _shown(x)
@pyinex.func
def a_float(x: float): return _shown(x)
@pyinex.func
def a_str(x: str): return _shown(x)
@pyinex.func
def a_bool(x: bool): return _shown(x)
@pyinex.func
def a_list(x: list): return _shown(x)
@pyinex.func
def a_list_list(x: list[list]): return _shown(x)
@pyinex.func
def a_rows_float(x: list[list[float]]): return _shown(x)
@pyinex.func
def a_rows_int(x: List[List[int]]): return _shown(x)
@pyinex.func
def a_rows_str(x: list[list[str]]): return _shown(x)
@pyinex.func
def a_rows_bool(x: list[list[bool]]): return _shown(x)
@pyinex.func
def a_rows_optional(x: list[list[float | None]]): return _shown(x)
@pyinex.func
def a_optional_rows(x: list[list[float]] | None): return _shown(x)
@pyinex.func
def a_array(x: np.ndarray): return x
@pyinex.func
def a_optional_array(x: Optional[np.ndarray]): return _shown(x)
@pyinex.func
def a_or_none(x: int | None): return _shown(x)
@pyinex.func
def a_none_or(x: None | int): return _shown(x)
@pyinex.func
def a_optional(x: Optional[int]): return _shown(x)
@pyinex.func
def a_union(x: Union[int, None]): return _shown(x)
@pyinex.func
def a_string(x: "list[list[float]]"): return _shown(x)
@pyinex.func
def a_string_int(x: "int"): return _shown(x)
@pyinex.func
def a_any(x: Any): return _shown(x)
@pyinex.func
def a_object(x: object): return _shown(x)
@pyinex.func
def a_dict(x: dict): return _shown(x)
@pyinex.func
def a_decimal(x: Decimal): return _shown(x)
@pyinex.func
def a_int_or_str(x: int | str): return _shown(x)
@pyinex.func
def a_flat(x: list[float]): return _shown(x)
@pyinex.func
def a_unknown(x: "NotDefinedAnywhere"): return _shown(x)
@pyinex.func
def a_mixed(plain, n: int, *rest: float, flag: bool = False) -> int:
    return [_shown(plain), _shown(n), _shown(list(rest)), _shown(flag)]
"""


class Err:
    """Stands in for pyinex.ExcelError, which the core makes in C++."""

    def __init__(self, text):
        self.text = text

    def __str__(self):
        return self.text

    def __repr__(self):
        return f"ExcelError({self.text!r})"


NA = Err("#N/A")


@pytest.fixture
def rows(load, capfd):
    """The script loaded: each function's key by its name, and what the console said."""
    loaded = load(SCRIPT, name="annotated.py")
    return {row[1]: row[0] for row in loaded}, capfd.readouterr().err


@pytest.fixture
def call(core, rows, capfd):
    """call(function, *args) -> (result, console text)."""

    def run(function, *args):
        capfd.readouterr()
        result = core["call_function"](rows[0][function], args, True)
        return (None if result is core["FAILED"] else result), capfd.readouterr().err

    return run


def converts(call, function, value, shown):
    result, err = call(function, value)
    assert result == shown and err == "", (function, value, result, err)


def refuses(call, function, value, said, argument="1 (x)"):
    result, err = call(function, value)
    assert result is None, (function, value, result)
    assert f"Pyinex: {function}: argument {argument}: {said}" in err, err


# The scalar table, row by row.

BLOCK = ((1.0, 2.0), (3.0, 4.0))


@pytest.mark.parametrize(
    "function, value, shown",
    [
        ("a_int", 3.0, ["int", 3]),
        ("a_int", -0.0, ["int", 0]),
        ("a_int", 1e20, ["int", 10**20]),
        ("a_int", None, ["int", 0]),
        ("a_float", 3.0, ["float", 3.0]),
        ("a_float", 2.5, ["float", 2.5]),
        ("a_float", None, ["float", 0.0]),
        ("a_str", "FALSE", ["str", "FALSE"]),
        ("a_str", "", ["str", ""]),
        ("a_str", None, ["str", ""]),
        ("a_bool", True, ["bool", True]),
        ("a_bool", False, ["bool", False]),
        ("a_bool", None, ["bool", False]),
    ],
)
def test_a_scalar_converts(call, function, value, shown):
    converts(call, function, value, shown)


@pytest.mark.parametrize(
    "function, value, said",
    [
        ("a_int", 2.5, "int takes a whole number, not 2.5"),
        ("a_int", "3", "int takes a whole number, not the text '3'"),
        ("a_int", "", "int takes a whole number, not the text ''"),
        ("a_int", True, "int takes a whole number, not TRUE"),
        ("a_int", NA, "int takes a whole number, not the error #N/A"),
        ("a_int", BLOCK, "int takes a whole number, not a range of 2 × 2 cells"),
        ("a_float", "3", "float takes a number, not the text '3'"),
        ("a_float", "", "float takes a number, not the text ''"),
        ("a_float", False, "float takes a number, not FALSE"),
        ("a_float", NA, "float takes a number, not the error #N/A"),
        ("a_float", ((1.0, 2.0),), "float takes a number, not a range of 1 × 2 cells"),
        ("a_str", 3.0, "str takes text, not 3.0"),
        ("a_str", 2.5, "str takes text, not 2.5"),
        ("a_str", True, "str takes text, not TRUE"),
        ("a_str", NA, "str takes text, not the error #N/A"),
        ("a_str", BLOCK, "str takes text, not a range of 2 × 2 cells"),
        ("a_bool", "FALSE", "bool takes TRUE or FALSE, not the text 'FALSE'"),
        ("a_bool", "", "bool takes TRUE or FALSE, not the text ''"),
        ("a_bool", 1.0, "bool takes TRUE or FALSE, not 1.0"),
        ("a_bool", 0.0, "bool takes TRUE or FALSE, not 0.0"),
        ("a_bool", NA, "bool takes TRUE or FALSE, not the error #N/A"),
        ("a_bool", BLOCK, "bool takes TRUE or FALSE, not a range of 2 × 2 cells"),
    ],
)
def test_a_scalar_refuses_naming_the_function_argument_and_value(call, function, value, said):
    refuses(call, function, value, said)


# Lists of rows.


@pytest.mark.parametrize("function", ["a_list", "a_list_list"])
def test_a_list_of_rows_keeps_each_element(call, function):
    converts(call, function, ((1.0, "a"), (None, True)), [[["float", 1.0], ["str", "a"]], [["NoneType", None], ["bool", True]]])
    converts(call, function, ((1.0,), (2.0,)), [[["float", 1.0]], [["float", 2.0]]])
    converts(call, function, 5.0, [[["float", 5.0]]])  # A single value: one row of one.
    converts(call, function, None, [[["NoneType", None]]])
    result, _ = call(function, NA)
    assert result[0][0][0] == "Err"  # As it arrives: the core's ExcelError, here the stand-in.


@pytest.mark.parametrize(
    "function, value, shown",
    [
        ("a_rows_float", ((1.0, 2.0), (3.0, None)), [[["float", 1.0], ["float", 2.0]], [["float", 3.0], ["float", 0.0]]]),
        ("a_rows_float", 5.0, [[["float", 5.0]]]),
        ("a_rows_float", None, [[["float", 0.0]]]),
        ("a_rows_int", ((1.0, 2.0),), [[["int", 1], ["int", 2]]]),
        ("a_rows_str", (("a", None),), [[["str", "a"], ["str", ""]]]),
        ("a_rows_bool", ((True, None),), [[["bool", True], ["bool", False]]]),
        ("a_rows_optional", ((1.0, None),), [[["float", 1.0], ["NoneType", None]]]),
        ("a_optional_rows", None, ["NoneType", None]),
        ("a_optional_rows", ((1.0, None),), [[["float", 1.0], ["float", 0.0]]]),
        ("a_string", ((1.0,),), [[["float", 1.0]]]),
    ],
)
def test_a_list_of_typed_rows_converts_each_element(call, function, value, shown):
    converts(call, function, value, shown)


@pytest.mark.parametrize(
    "function, value, said",
    [
        ("a_rows_float", ((1.0, 2.0), ("x", 4.0)), "row 2, column 1: float takes a number, not the text 'x'"),
        ("a_rows_float", "x", "row 1, column 1: float takes a number, not the text 'x'"),
        ("a_rows_int", ((1.0, 2.5),), "row 1, column 2: int takes a whole number, not 2.5"),
        ("a_rows_str", ((1.0,),), "row 1, column 1: str takes text, not 1.0"),
        ("a_rows_bool", (("TRUE",),), "row 1, column 1: bool takes TRUE or FALSE, not the text 'TRUE'"),
        ("a_rows_float", ((1.0, NA),), "row 1, column 2: float takes a number, not the error #N/A"),
    ],
)
def test_a_refused_element_refuses_the_argument_naming_its_row_and_column(call, function, value, said):
    refuses(call, function, value, said)


# numpy.ndarray.


def test_an_array_of_numbers_has_the_ranges_shape(call):
    result, err = call("a_array", ((1.0, 2.0, 3.0), (4.0, None, 6.0)))
    assert err == "" and isinstance(result, numpy.ndarray)
    assert result.dtype == numpy.float64 and result.shape == (2, 3)
    assert result[0].tolist() == [1.0, 2.0, 3.0] and math.isnan(result[1, 1])


def test_a_single_value_is_a_one_by_one_array(call):
    result, _ = call("a_array", 5.0)
    assert result.shape == (1, 1) and result[0, 0] == 5.0
    result, _ = call("a_array", None)
    assert result.shape == (1, 1) and math.isnan(result[0, 0])


@pytest.mark.parametrize(
    "value, said",
    [
        (((1.0, "x"),), "row 1, column 2: numpy.ndarray takes numbers and empty cells, not the text 'x'"),
        (((1.0, 2.0), (True, 4.0)), "row 2, column 1: numpy.ndarray takes numbers and empty cells, not TRUE"),
        (((NA,),), "row 1, column 1: numpy.ndarray takes numbers and empty cells, not the error #N/A"),
        ("x", "row 1, column 1: numpy.ndarray takes numbers and empty cells, not the text 'x'"),
    ],
    ids=["text", "mixed", "error", "single text"],
)
def test_an_array_refuses_anything_but_numbers(call, value, said):
    refuses(call, "a_array", value, said)


def test_an_optional_array_gives_none_for_an_empty_cell(call):
    converts(call, "a_optional_array", None, ["NoneType", None])


# None for an empty cell, and the spellings.


@pytest.mark.parametrize("function", ["a_or_none", "a_none_or", "a_optional", "a_union"])
def test_none_for_an_empty_cell_in_each_spelling(call, function):
    converts(call, function, None, ["NoneType", None])
    converts(call, function, 3.0, ["int", 3])
    refuses(call, function, 2.5, "int takes a whole number, not 2.5")


def test_an_annotation_written_as_a_string_is_evaluated(call):
    converts(call, "a_string_int", 3.0, ["int", 3])
    refuses(call, "a_string", "x", "row 1, column 1: float takes a number, not the text 'x'")


def test_every_annotation_a_string_under_future_annotations(core, load, capfd):
    (row,) = load(
        "from __future__ import annotations\nimport pyinex\n\n"
        "@pyinex.func\ndef f(n: int, rows: list[list[float]]):\n    return [n, rows]\n",
        name="future.py",
    )
    assert core["call_function"](row[0], (3.0, ((1.0, None),)), True) == [3, [[1.0, 0.0]]]
    assert capfd.readouterr().err == ""


@pytest.mark.parametrize("future", [True, False], ids=["future", "plain"])
def test_a_quoted_annotation_is_evaluated_until_it_is_not_text(core, load, capfd, future):
    # Under the future import, "int" is text within text: evaluated once, it is still text.
    head = "from __future__ import annotations\n" if future else ""
    quote = '"' if future else "'\"" 
    end = '"' if future else "\"'"
    (row,) = load(
        f"{head}import pyinex\n\n@pyinex.func\n"
        f"def f(n: {quote}int{end}, rows: {quote}list[list[float]]{end}):\n    return [n, rows]\n",
        name="quoted.py",
    )
    assert capfd.readouterr().err == ""
    assert core["call_function"](row[0], (3.0, ((1.0, None),)), True) == [3, [[1.0, 0.0]]]
    assert core["call_function"](row[0], (2.5, 1.0), True) is core["FAILED"]
    assert "argument 1 (n): int takes a whole number, not 2.5" in capfd.readouterr().err


def test_text_naming_itself_ends_as_an_annotation_not_converted(core, load, capfd):
    (row,) = load('import pyinex\n\nLoop = "Loop"\n\n@pyinex.func\ndef f(x: "Loop"):\n    return x\n', name="loop.py")
    assert "Pyinex does not convert the annotation of x" in capfd.readouterr().err
    assert core["call_function"](row[0], (2.5,), True) == 2.5


def test_text_is_evaluated_five_times_at_most(core, load, capfd):
    # "T1" gives "T2" ... "T4" gives "int", then int: five evaluations. "T0" needs six.
    (row,) = load(
        'import pyinex\n\nT0, T1, T2, T3, T4 = "T1", "T2", "T3", "T4", "int"\n\n@pyinex.func\n'
        'def f(five: "T1", six: "T0"):\n    return [five, six]\n',
        name="chain.py",
    )
    err = capfd.readouterr().err
    assert "the annotation of five" not in err
    assert "Pyinex does not convert the annotation of six, 'T0'" in err
    assert core["call_function"](row[0], (3.0, 2.5), True) == [3, 2.5]


# Annotations Pyinex does not convert.


@pytest.mark.parametrize("function", ["a_any", "a_object"])
def test_any_and_object_leave_the_value_silently(rows, call, function):
    assert f"{function} in the script" not in rows[1]
    converts(call, function, "x", ["str", "x"])
    converts(call, function, 2.5, ["float", 2.5])


@pytest.mark.parametrize(
    "function, shown",
    [
        ("a_dict", "<class 'dict'>"),
        ("a_decimal", "<class 'decimal.Decimal'>"),
        ("a_int_or_str", "int | str"),
        ("a_flat", "list[float]"),
    ],
)
def test_another_annotation_leaves_the_value_and_says_so_at_registration(rows, call, tmp_path, function, shown):
    path = tmp_path / "annotated.py"
    said = f"Pyinex: {function} in the script {path}: Pyinex does not convert the annotation of x, {shown}, so x arrives unconverted"
    assert said in rows[1], rows[1]
    converts(call, function, 2.5, ["float", 2.5])
    converts(call, function, ((1.0,),), ["tuple", ((1.0,),)])


def test_an_annotation_that_cannot_be_evaluated_leaves_the_value_and_says_so(rows, call, tmp_path):
    path = tmp_path / "annotated.py"
    said = (f"Pyinex: a_unknown in the script {path}: the annotation of x, 'NotDefinedAnywhere', cannot be "
            "evaluated (NameError: name 'NotDefinedAnywhere' is not defined), so x arrives unconverted")
    assert said in rows[1], rows[1]
    converts(call, "a_unknown", "x", ["str", "x"])


def test_annotated_and_plain_parameters_in_one_function(call, core):
    # plain, n, flag, then rest1 to rest27.
    result, err = call("a_mixed", "as is", 3.0, core["MISSING"], 1.0, 2.0)
    assert err == "" and result == [["str", "as is"], ["int", 3], [["float", 1.0], ["float", 2.0]], ["bool", False]]
    result, err = call("a_mixed", NA, 3.0)
    assert result[0][0] == "Err" and result[1] == ["int", 3]


def test_an_args_value_refused_is_named_by_its_position(core, rows, capfd):
    capfd.readouterr()
    # a_mixed: plain, n, flag, then rest1 to rest27.
    args = ("as is", 3.0, core["MISSING"], 1.0, "x")
    assert core["call_function"](rows[0]["a_mixed"], args, True) is core["FAILED"]
    assert "Pyinex: a_mixed: argument 5 (rest2): float takes a number, not the text 'x'" in capfd.readouterr().err


def test_a_keyword_only_value_converts(core, rows):
    assert core["call_function"](rows[0]["a_mixed"], ("p", 1.0, True), True)[3] == ["bool", True]


def test_a_position_left_out_is_never_converted(core, rows):
    m = core["MISSING"]
    # n has no default, and a later position is given: None, not int's 0.
    result = core["call_function"](rows[0]["a_mixed"], ("p", m, m, 1.0), True)
    assert result[1] == ["NoneType", None] and result[3] == ["bool", False]


def test_the_return_annotation_is_ignored(core, rows):
    assert core["call_function"](rows[0]["a_mixed"], ("p", 1.0), True)[1] == ["int", 1]


def test_an_annotation_on_kwargs_is_ignored(core, load, capfd):
    # dict is an annotation Pyinex does not convert: looked at, it would be said so.
    (row,) = load("import pyinex\n\n@pyinex.func\ndef f(x: int, **more: dict):\n    return [x, more]\n", name="kwargs.py")
    assert capfd.readouterr().err == ""
    assert row[2] == 1  # One position: **more takes none.
    assert core["call_function"](row[0], (3.0,), True) == [3, {}]


# Without NumPy.


def test_without_numpy_a_script_importing_it_fails_to_load(core, script, monkeypatch, capfd):
    monkeypatch.setitem(sys.modules, "numpy", None)
    path = script("import numpy\nimport pyinex\n\n@pyinex.func\ndef f(a: numpy.ndarray):\n    return 1\n", name="needs_numpy.py")
    assert core["load_functions"](str(path)) is core["FAILED"]
    err = capfd.readouterr().err
    assert "failed to load:" in err and "ModuleNotFoundError" in err


def test_without_numpy_a_string_annotation_naming_it_leaves_the_value(core, script, monkeypatch, capfd):
    monkeypatch.setitem(sys.modules, "numpy", None)
    path = script(
        "import pyinex\n\n@pyinex.func\ndef f(a: 'numpy.ndarray', n: int):\n    return [a, n]\n", name="names_numpy.py"
    )
    (row,) = core["load_functions"](str(path))
    assert "the annotation of a, 'numpy.ndarray', cannot be evaluated (NameError" in capfd.readouterr().err
    # Its other annotations convert: Pyinex itself never imports NumPy.
    assert core["call_function"](row[0], (((1.0,),), 2.0), True) == [((1.0,),), 2]


@pytest.mark.skipif(sys.version_info < (3, 14), reason="before 3.14 the annotation is evaluated as the script runs")
def test_a_forward_reference_that_cannot_be_evaluated_leaves_the_value_and_says_so(core, load, capfd):
    (row,) = load("import pyinex\n\n@pyinex.func\ndef f(x: Later, n: int):\n    return [x, n]\n", name="later.py")
    assert "the annotation of x, 'Later', cannot be evaluated (NameError" in capfd.readouterr().err
    assert core["call_function"](row[0], ("x", 2.0), True) == ["x", 2]


def test_an_imported_function_s_annotation_is_evaluated_where_it_was_defined(core, load, capfd, monkeypatch):
    # The helper's Alias is int; the script's own Alias, str, must not be used for it, either
    # for the function marked directly or through a partial.
    import types

    helper = types.ModuleType("annotation_helper")
    exec('Alias = int\n\ndef target(x: "Alias"):\n    return [type(x).__name__, x]\n', vars(helper))
    monkeypatch.setitem(sys.modules, "annotation_helper", helper)
    rows = load(
        "import functools\nimport pyinex\nfrom annotation_helper import target\n\nAlias = str\n\n"
        'direct = pyinex.func(name="direct")(target)\n'
        'partly = pyinex.func(name="partly")(functools.partial(target))\n',
        name="imported.py",
    )
    keys = {row[1]: row[0] for row in rows}
    assert set(keys) == {"direct", "partly"}
    assert capfd.readouterr().err == ""
    for name in keys:
        assert core["call_function"](keys[name], (2.0,), True) == ["int", 2]
        assert core["call_function"](keys[name], ("text",), True) is core["FAILED"]
        assert "int takes a whole number, not the text 'text'" in capfd.readouterr().err


def test_an_annotation_that_raises_system_exit_leaves_only_its_argument_unconverted(core, load, capfd):
    rows = load(
        "import pyinex\n\ndef stop():\n    raise SystemExit(3)\n\n"
        "@pyinex.func\ndef before(x: int):\n    return x\n\n"
        '@pyinex.func\ndef bad(x: "stop()"):\n    return x\n\n'
        "@pyinex.func\ndef after(x: int):\n    return x\n",
        name="stopping.py",
    )
    keys = {row[1]: row[0] for row in rows}
    assert set(keys) == {"before", "bad", "after"}
    assert "the annotation of x, 'stop()', cannot be evaluated (SystemExit" in capfd.readouterr().err
    assert core["call_function"](keys["bad"], (2.5,), True) == 2.5
    assert core["call_function"](keys["after"], (2.0,), True) == 2


@pytest.mark.skipif(sys.version_info < (3, 14), reason="annotations are evaluated lazily from 3.14")
def test_a_lazy_annotation_that_raises_system_exit_refuses_only_its_function(core, load, capfd):
    rows = load(
        "import pyinex\n\ndef stop():\n    raise SystemExit(3)\n\n"
        "@pyinex.func\ndef before(x: int):\n    return x\n\n"
        "@pyinex.func\ndef bad(x: stop()):\n    return x\n\n"
        "@pyinex.func\ndef after(x: int):\n    return x\n",
        name="lazy.py",
    )
    assert {row[1] for row in rows} == {"before", "after"}
    assert "bad in the script" in capfd.readouterr().err


def test_a_note_on_an_annotation_is_said_once_not_at_each_load(core, script, capfd):
    path = script("import pyinex\n\n@pyinex.func\ndef F(x: dict):\n    return x\n", name="noted.py")
    rows = core["load_functions"](str(path))
    core["registered"](rows[0][0], True)
    assert "does not convert the annotation of x" in capfd.readouterr().err
    assert core["load_functions"](str(path)) == [[-1]]
    assert capfd.readouterr().err == ""


def test_an_inherited_constructor_s_annotations_are_read_where_it_was_defined(core, script):
    # Unit is int where Base is defined, and float in the script that marks its subclass.
    script("from __future__ import annotations\nUnit = int\nclass Base(float):\n"
           "    def __new__(cls, x: Unit):\n        return super().__new__(cls, x if isinstance(x, int) else 999)\n",
           name="annotated_base.py")
    sys.path.insert(0, str(script("", name="unused.py").parent))
    try:
        rows = core["load_functions"](str(script(
            "import pyinex\nfrom annotated_base import Base\nUnit = float\n\n@pyinex.func\nclass Sample(Base):\n    pass\n",
            name="annotated_child.py")))
        core["registered"](rows[0][0], True)
        assert core["call_function"](rows[0][0], (3.0,), False) == 3.0
    finally:
        sys.path.pop(0)
        sys.modules.pop("annotated_base", None)


@pytest.mark.skipif(sys.version_info < (3, 14), reason="annotations are evaluated only when asked for from 3.14")
@pytest.mark.parametrize("ignored", ["-> 1/0", "**kwargs: 1/0"])
def test_an_ignored_annotation_that_raises_refuses_nothing(core, script, capfd, ignored):
    if ignored.startswith("->"):
        source = f"import pyinex\n\n@pyinex.func\ndef Sample(x: int) {ignored}:\n    return x\n"
    else:
        source = f"import pyinex\n\n@pyinex.func\ndef Sample(x: int, {ignored}):\n    return x\n"
    rows = core["load_functions"](str(script(source, name="ignored.py")))
    assert [row[1] for row in rows] == ["Sample"]
    err = capfd.readouterr().err
    assert "not registered" not in err
    # Python evaluates them together, so none is: x arrives unconverted, and the console says so.
    assert "its annotations cannot be evaluated (ZeroDivisionError: division by zero)" in err


@pytest.mark.skipif(sys.version_info < (3, 14), reason="annotations are evaluated only when asked for from 3.14")
def test_a_parameter_s_annotation_that_raises_leaves_it_unconverted(core, script, capfd):
    source = ("import pyinex\n\ndef bad():\n    raise RuntimeError('boom')\n\n"
              "@pyinex.func\ndef F(x: bad()):\n    return x\n")
    rows = core["load_functions"](str(script(source, name="raises.py")))
    assert [row[1] for row in rows] == ["F"]
    assert "cannot be evaluated (RuntimeError: boom), so x arrives unconverted" in capfd.readouterr().err


@pytest.mark.skipif(sys.version_info < (3, 14), reason="annotations are evaluated only when asked for from 3.14")
def test_a_wrapper_whose_inner_function_s_annotations_raise_keeps_the_inner_signature(core, script):
    source = ("import functools, pyinex\n\n"
              "def target(*, x=1) -> 1/0:\n    return x\n\n"
              "@pyinex.func(name='Wrapped')\n@functools.wraps(target)\n"
              "def wrapper(*args, **kwargs):\n    return target(*args, **kwargs)\n")
    rows = core["load_functions"](str(script(source, name="wrapped.py")))
    assert [(row[1], row[3]) for row in rows] == [("Wrapped", "x")]  # Not *args.
    core["registered"](rows[0][0], True)
    assert core["call_function"](rows[0][0], (2.0,), False) == 2.0
