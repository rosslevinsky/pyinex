"""Signatures of direct functions in pure Python (src/core/python/funcs.py;
docs/architecture.md, "Direct functions", "Signatures" to "What can be marked"): the Excel
positions each parameter takes and the names Excel is told, each case of a position left
out, the signatures refused, and what the decorator can mark.

The core runs funcs.py in call.py's namespace, and so do these tests, under plain Python. A
call's arguments are given as the core gives them: arguments left out at the end dropped, and
one left out before the last given as funcs.py's MISSING.
"""

import pytest

# The fixtures core and load are in conftest.py.

POSITIONS = """
import functools, math
import pyinex

@pyinex.func
def f(x, *items, scale=1):
    '''The plan's example.

    Args:
        x: The first.
        *items: Each of the rest.
        scale: How much.
    '''
    return [x, list(items), scale]

@pyinex.func
def kinds(a, b=2, /, c=3, *, d, e=5, **options):
    return [a, b, c, d, e, len(options)]

@pyinex.func
def only_args(*values):
    return list(values)

@pyinex.func
def no_arguments():
    return 'none'
"""


@pytest.fixture
def keys(load):
    """The script POSITIONS loaded: each function's row by its name."""
    return {row[1]: row for row in load(POSITIONS)}


def call(core, row, *args):
    return core["call_function"](row[0], args, True)


def test_positions_are_positional_then_keyword_only_then_args(keys):
    row = keys["f"]
    assert row[2] == 30
    assert row[3] == ",".join(["x", "scale"] + [f"items{i}" for i in range(1, 29)])
    assert keys["kinds"][2:4] == [5, "a,b,c,d,e"]  # **options takes no position.
    # Excel takes at most 255 characters of names: the list is cut after the last whole one.
    names = ",".join(f"values{i}" for i in range(1, 31))
    assert keys["only_args"][2:4] == [30, names[: names.rindex(",", 0, 256)]]
    assert keys["no_arguments"][2:4] == [0, ""]


def test_args_positions_each_have_its_help(keys):
    row = keys["f"]
    assert row[6:8] == ["The first.", "How much."]
    assert row[8:] == ["Each of the rest."] * 28


def test_the_plans_examples(core, keys):
    m = core["MISSING"]
    assert call(core, keys["f"], 10.0, 2.0) == [10.0, [], 2.0]  # f(10, scale=2)
    assert call(core, keys["f"], 10.0, m, 20.0, 30.0) == [10.0, [20.0, 30.0], 1]  # f(10, 20, 30)


@pytest.mark.parametrize(
    "args, called",
    [
        ((), []),
        ((1.0,), [1.0]),
        (tuple(float(i) for i in range(30)), [float(i) for i in range(30)]),
        ((1.0, "MISSING", 3.0), [1.0, None, 3.0]),  # Left out before the last given: None.
    ],
    ids=["none", "one", "thirty", "missing in the middle"],
)
def test_args_take_none_one_or_many(core, keys, args, called):
    args = tuple(core["MISSING"] if a == "MISSING" else a for a in args)
    assert call(core, keys["only_args"], *args) == called


def test_a_position_left_out_takes_its_default(core, keys):
    m = core["MISSING"]
    # a, b (positional-only, default 2), c (default 3), d (keyword-only), e (default 5).
    assert call(core, keys["kinds"], 1.0, m, m, 4.0) == [1.0, 2, 3, 4.0, 5, 0]
    assert call(core, keys["kinds"], 1.0, 9.0, 8.0, 4.0, 7.0) == [1.0, 9.0, 8.0, 4.0, 7.0, 0]
    assert call(core, keys["kinds"], 1.0, m, m, 4.0, m) == [1.0, 2, 3, 4.0, 5, 0]


def test_one_left_out_with_no_default_before_the_last_given_is_none(core, keys):
    m = core["MISSING"]
    assert call(core, keys["kinds"], m, m, m, m, 7.0) == [None, 2, 3, None, 7.0, 0]


def test_one_left_out_with_no_default_after_the_last_given_is_pythons_error(core, keys, capfd):
    assert call(core, keys["kinds"], 1.0) is core["FAILED"]
    err = capfd.readouterr().err
    assert "TypeError" in err and "'d'" in err
    assert call(core, keys["kinds"]) is core["FAILED"]
    assert "TypeError" in capfd.readouterr().err


def test_more_arguments_than_positions_are_never_passed(core, keys):
    # Excel refuses such a formula; the core never has more than the registered count.
    assert keys["no_arguments"][2] == 0 and call(core, keys["no_arguments"]) == "none"


@pytest.mark.parametrize(
    "signature, why",
    [
        (", ".join(f"a{i}" for i in range(31)), "it has 31 parameters, and Excel passes at most 30"),
        (", ".join(f"a{i}" for i in range(30)) + ", *items", "it has 30 parameters and *items, and Excel passes at most 30"),
        (", ".join(f"a{i}" for i in range(29)) + ", *items, k", "it has 30 parameters and *items, and Excel passes at most 30"),
        ("*, " + ", ".join(f"k{i}" for i in range(31)), "it has 31 parameters, and Excel passes at most 30"),
    ],
    ids=["31 positional", "30 and args", "29, args and a keyword-only", "31 keyword-only"],
)
def test_more_than_thirty_positions_are_refused_and_the_rest_register(load, capfd, signature, why):
    rows = load(
        f"import pyinex\n\n@pyinex.func\ndef before(a):\n    pass\n\n"
        f"@pyinex.func\ndef refused({signature}):\n    pass\n\n@pyinex.func\ndef after(b):\n    pass\n"
    )
    assert [row[1] for row in rows] == ["before", "after"]
    assert "Pyinex: refused in the script" in (err := capfd.readouterr().err) and f"is not registered: {why}" in err


@pytest.mark.parametrize(
    "signature, count",
    [
        (", ".join(f"a{i}" for i in range(29)) + ", *items", 30),
        (", ".join(f"a{i}" for i in range(15)) + ", *, " + ", ".join(f"k{i}" for i in range(15)), 30),
        (", ".join(f"a{i}" for i in range(30)) + ", **options", 30),
    ],
    ids=["29 and args", "15 and 15 keyword-only", "30 and kwargs"],
)
def test_thirty_positions_are_accepted(load, signature, count):
    (row,) = load(f"import pyinex\n\n@pyinex.func\ndef f({signature}):\n    pass\n")
    assert row[2] == count and len(row) == 6 + count


# What can be marked.

MARKED = """
import functools, math
import pyinex

def discount(price, rate, years=1):
    return price * (1 - rate) ** years

@pyinex.func
class Money(float):
    def __new__(cls, amount, rate=1):
        return super().__new__(cls, amount * rate)

class Euros(Money):
    pass

@pyinex.func
class Point:
    def __init__(self, x, y):
        self.x, self.y = x, y

origin = Point(0, 0)

once = pyinex.func(functools.partial(discount, years=1))
half = pyinex.func(functools.partial(discount, 100, 0.5))
root = pyinex.func(math.sqrt)

@pyinex.func
def after(b):
    return b
"""


@pytest.fixture
def marked(load, capfd):
    rows = load(MARKED, name="marked_model.py")
    return {row[1]: row for row in rows}, capfd.readouterr().err


def test_a_class_is_registered_with_its_constructors_parameters(core, marked):
    rows, _ = marked
    assert rows["Money"][2:4] == [2, "amount,rate"] and rows["Point"][2:4] == [2, "x,y"]
    money = call(core, rows["Money"], 3.0, 2.0)
    assert type(money).__name__ == "Money" and money == 6.0  # A float: Excel takes it.
    assert type(call(core, rows["Point"], 1.0, 2.0)).__name__ == "Point"  # Excel refuses it.


def test_only_the_class_marked_is_registered_not_its_subclasses_or_instances(marked):
    rows, _ = marked
    assert list(rows) == ["Money", "Point", "once", "half", "after"]


def test_a_partial_has_the_parameters_it_leaves(core, marked):
    rows, _ = marked
    # years, fixed by keyword, becomes keyword-only, with that value as its default.
    assert rows["once"][2:4] == [3, "price,rate,years"]
    assert call(core, rows["once"], 100.0, 0.5) == 50.0
    assert call(core, rows["once"], 100.0, 0.5, 2.0) == 25.0
    assert rows["half"][2:4] == [1, "years"]
    assert call(core, rows["half"]) == 50.0


def test_a_built_in_function_cannot_be_marked_and_the_script_loads(core, marked, tmp_path):
    rows, err = marked
    assert "root" not in rows and "sqrt" not in rows and "after" in rows
    path = tmp_path / "marked_model.py"
    assert (
        f"Pyinex: sqrt in the script {path} is not registered: a built-in function cannot be marked; "
        "wrap it in a function of your own" in err
    )


def test_the_decorator_returns_a_built_in_function_unchanged(core):
    import math

    assert core["func"](math.sqrt) is math.sqrt
