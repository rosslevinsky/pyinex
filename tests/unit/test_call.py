"""PyCall's argument rule and exception boundary (docs/architecture.md, "PyCall" and
"Argument rule"), through the add-in's own code."""

import os
import re
from pathlib import Path

import pytest

CORE_PYTHON = Path(__file__).resolve().parents[2] / "src" / "core" / "python"

MISSING = ("missing",)
NIL = ("nil",)
VALUE_ERROR = ("err", 15)

FUNCTIONS = """
import functools, sys

def add(a, b):
    return a + b

def count(*args):
    return len(args)

def shown(*args):
    return repr(args)

def default(a, b=10):
    return a + b

class Twice(float):
    def __new__(cls, x):
        return float.__new__(cls, 2 * x)

class Adder:
    def __call__(self, a, b):
        return a + b

adder = Adder()
power_of_two = functools.partial(pow, 2)
biggest = max
whole = int
not_callable = 42

def fails():
    raise ValueError("deliberately")

def exits():
    sys.exit(3)

def interrupted():
    raise KeyboardInterrupt

def unreturnable():
    return [1, {}]

def ok():
    return "still working"
"""


@pytest.fixture
def functions(script):
    return script(FUNCTIONS, name="functions.py")


def test_plain_function(pycall, functions):
    assert pycall(functions, "add", 2, 3) == ("num", 5.0)


def test_star_args(pycall, functions):
    assert pycall(functions, "count", 1, 2, 3) == ("num", 3.0)


def test_default(pycall, functions):
    assert pycall(functions, "default", 1) == ("num", 11.0)
    assert pycall(functions, "default", 1, 2) == ("num", 3.0)


def test_trailing_missing_arguments_are_dropped(pycall, functions):
    assert pycall(functions, "count", 1, *[MISSING] * 29) == ("num", 1.0)
    assert pycall(functions, "count", *[MISSING] * 30) == ("num", 0.0)


def test_middle_missing_argument_is_none(pycall, functions):
    assert pycall(functions, "shown", 1, MISSING, 3, MISSING) == ("str", "(1.0, None, 3.0)")


def test_trailing_empty_cell_is_passed_as_none(pycall, functions):
    assert pycall(functions, "shown", 1, NIL) == ("str", "(1.0, None)")
    assert pycall(functions, "shown", 1, None) == ("str", "(1.0, None)")


def test_thirty_arguments(pycall, functions):
    assert pycall(functions, "count", *range(30)) == ("num", 30.0)


def test_range_argument(pycall, functions):
    assert pycall(functions, "shown", ((1, 2), (3, 4))) == ("str", "(((1.0, 2.0), (3.0, 4.0)),)")


def test_class(pycall, functions):
    assert pycall(functions, "Twice", 4) == ("num", 8.0)


def test_object_with_call(pycall, functions):
    assert pycall(functions, "adder", 1, 2) == ("num", 3.0)


def test_partial(pycall, functions):
    assert pycall(functions, "power_of_two", 3) == ("num", 8.0)


def test_builtins_with_no_inspectable_signature(pycall, functions):
    assert pycall(functions, "biggest", 1, 7, 3) == ("num", 7.0)
    assert pycall(functions, "whole", "42") == ("num", 42.0)


def test_wrong_argument_count(pycall, functions, capfd):
    assert pycall(functions, "add", 1) == VALUE_ERROR
    assert "TypeError" in capfd.readouterr().err
    assert pycall(functions, "add", 1, 2, 3) == VALUE_ERROR


@pytest.mark.parametrize("name, shown", [("fails", "ValueError: deliberately"), ("exits", "SystemExit: 3"),
                                         ("interrupted", "KeyboardInterrupt")])
def test_exception_becomes_value_error(pycall, functions, capfd, name, shown):
    assert pycall(functions, name) == VALUE_ERROR
    err = capfd.readouterr().err
    assert f"Pyinex: {name} in the script" in err
    assert "Traceback (most recent call last)" in err and shown in err
    assert "<pyinex>" not in err, "Pyinex's own frames are left out"
    # Every frame shown is the script's: none from the core's Python files, under any name
    # they are compiled with (docs/architecture.md, "The core's Python files").
    frames = re.findall(r'File "([^"]*)"', err.split("Traceback (most recent call last)")[1])
    assert frames and {os.path.basename(f) for f in frames} == {"functions.py"}, frames
    assert not any(p.name in err for p in CORE_PYTHON.glob("*.py")), err
    assert pycall(functions, "ok") == ("str", "still working")


def test_unknown_function(pycall, functions, capfd):
    assert pycall(functions, "nonexistent") == VALUE_ERROR
    assert "has no function nonexistent" in capfd.readouterr().err


def test_not_callable(pycall, functions, capfd):
    assert pycall(functions, "not_callable") == VALUE_ERROR
    assert "not_callable in the script" in capfd.readouterr().err


def test_dotted_name_is_not_split(pycall, functions, capfd):
    assert pycall(functions, "functools.partial") == VALUE_ERROR
    assert "has no function functools.partial" in capfd.readouterr().err


def test_unreturnable_result_names_where(pycall, functions, capfd):
    assert pycall(functions, "unreturnable") == VALUE_ERROR
    assert "functions.py, unreturnable: row 1, column 2: dict cannot be returned to Excel" in capfd.readouterr().err


def test_file_and_function_must_be_text(testing, functions, capfd):
    assert testing.call(("num", 1.0), ("str", "ok")) == VALUE_ERROR
    assert testing.call(("str", str(functions)), MISSING) == VALUE_ERROR
    assert "as text" in capfd.readouterr().err


def test_unrecognized_argument_arrives_as_value_error(pycall, functions, capfd):
    assert pycall(functions, "shown", ("raw", 0x0020)) == ("str", "(ExcelError('#VALUE!'),)")
    assert "argument 1" in capfd.readouterr().err


def test_argument_conversion_survives_code_keeping_objects(pycall, functions, testing, monkeypatch):
    # Making an ExcelError for an argument runs Python code; here it keeps every object
    # Python tracks, which would include an argument tuple still being filled.
    import gc

    kept = []
    original = testing.ExcelError.__init__

    def keeping(self, text):
        original(self, text)
        kept.append(gc.get_objects())

    monkeypatch.setattr(testing.ExcelError, "__init__", keeping)
    assert pycall(functions, "shown", ("err", 42), 1, ("err", 7)) == (
        "str", "(ExcelError('#N/A'), 1.0, ExcelError('#DIV/0!'))"
    )


def test_traceback_keeps_user_code_from_eval(pycall, script, capfd):
    path = script("f = eval('lambda: 1 / 0')\n", name="evaluated.py")
    assert pycall(path, "f") == VALUE_ERROR
    err = capfd.readouterr().err
    assert 'File "<string>"' in err and "ZeroDivisionError" in err


def test_pyinex_s_own_frames_are_left_out_after_the_user_s_too(pycall, script, capfd):
    # pyinex.ExcelError's constructor refuses the text in a frame of excel_error.py, after the
    # user's; a chained exception's frames are filtered too.
    path = script(
        "import _pyinex_testing\n\n"
        "def invalid():\n    return _pyinex_testing.ExcelError('invalid')\n\n"
        "def chained():\n    try:\n        invalid()\n    except ValueError as e:\n"
        "        raise RuntimeError('again') from e\n",
        name="own_frames.py",
    )
    for function, shown in [("invalid", "ValueError: not an Excel error"), ("chained", "RuntimeError: again")]:
        assert pycall(path, function) == VALUE_ERROR
        err = capfd.readouterr().err
        assert shown in err and "ValueError: not an Excel error" in err
        frames = re.findall(r'File "([^"]*)"', err)
        assert frames and {os.path.basename(f) for f in frames} == {"own_frames.py"}, err
        assert "<pyinex/" not in err, err


def test_a_traceback_keeps_a_cause_that_is_also_a_context(pycall, script, capfd):
    # outer's cause is middle; root is both outer's context and middle's cause.
    path = script(
        "def shared():\n    root = KeyError('root')\n    middle = ValueError('middle')\n"
        "    middle.__cause__ = root\n    outer = RuntimeError('outer')\n"
        "    outer.__cause__ = middle\n    outer.__context__ = root\n    raise outer\n",
        name="shared.py",
    )
    assert pycall(path, "shared") == VALUE_ERROR
    err = capfd.readouterr().err
    assert "KeyError: 'root'" in err and "ValueError: middle" in err and "RuntimeError: outer" in err, err


def test_a_long_chain_of_causes_is_still_shown(pycall, script, capfd):
    path = script(
        "def chain():\n    error = ValueError(0)\n    for n in range(1, 1500):\n"
        "        error, cause = ValueError(n), error\n        error.__cause__ = cause\n    raise error\n",
        name="long_chain.py",
    )
    assert pycall(path, "chain") == VALUE_ERROR
    err = capfd.readouterr().err
    assert "ValueError: 1499" in err and "ValueError: 0" in err
