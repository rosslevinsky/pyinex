"""Signatures of direct functions in the fake Excel (docs/architecture.md, "Signatures" to
"What can be marked"): the direct-functions test group 2. 0 to 30 arguments, each position
checked by its value, across the fourth and fifth; defaults; *args with none, one and many;
keyword-only parameters; the plan's examples; a missing argument in the middle; more
arguments than registered; each refusal in a script whose other functions still register and
run; a class, a functools.partial and a built-in function. Then reloads that change a
function's positions: calls before the queue has run, a registration again that fails, and
several edits while work is queued.

One script holds every function of the first part, loaded once by PyLoad in a host of the
module's own. Functions return the repr of what they received, so a value is checked by its
position.
"""

import pytest
from conftest import host_env
from hostdriver import MISSING, VALUE, Err, Host, HostError
from test_direct_functions import wait_for
from test_reload_functions import edit, loaded, types

SCRIPT = """import functools, math
import pyinex

{numbered}

@pyinex.func
def SgDefaults(a, b=2.0, c=3.0):
    return repr([a, b, c])

@pyinex.func
def SgArgs(*items):
    return repr(list(items))

@pyinex.func
def SgKw(a, *, k, scale=1.0):
    return repr([a, k, scale])

@pyinex.func
def SgPlan(x, *items, scale=1):
    '''The plan's example.

    Args:
        x: The first.
        *items: Each of the rest.
        scale: How much.
    '''
    return repr([x, list(items), scale])

@pyinex.func
def SgMiddle(a, b, c):
    return repr([a, b, c])

@pyinex.func
def SgBefore(a):
    return 'before'

@pyinex.func
def SgOver({over}):
    return 'never'

@pyinex.func
def SgOverArgs({thirty}, *rest):
    return 'never'

SgUnreadable = pyinex.func(functools.partial(max))
SgSqrt = pyinex.func(math.sqrt)

@pyinex.func
class SgMoney(float):
    def __new__(cls, amount, rate=1.0):
        return super().__new__(cls, amount * rate)

@pyinex.func
class SgPoint:
    def __init__(self, x, y):
        self.x, self.y = x, y

def _discount(price, rate, years=2.0):
    return price * (1 - rate) ** years

SgOnce = pyinex.func(functools.partial(_discount, years=1.0))

@pyinex.func
def SgAfter(a):
    return 'after'
"""


def numbered():
    """Count0 to Count30, each returning the repr of its arguments in order. (Not Sg1 and so
    on: SG1 is a cell, a name Excel cannot take.)"""
    out = []
    for n in range(31):
        params = ", ".join(f"p{i}" for i in range(1, n + 1))
        out.append(f"@pyinex.func\ndef Count{n}({params}):\n    return repr([{params}])\n")
    return "\n".join(out)


@pytest.fixture(scope="module")
def sig(xllhost, pyinex_xll, tmp_path_factory):
    """A host with the script loaded by PyLoad and its functions registered."""
    path = tmp_path_factory.mktemp("signatures") / "signatures.py"
    path.write_text(
        SCRIPT.format(
            numbered=numbered(),
            over=", ".join(f"p{i}" for i in range(31)),
            thirty=", ".join(f"p{i}" for i in range(30)),
        ),
        encoding="utf-8",
    )
    with Host(xllhost, pyinex_xll, env=host_env()) as h:
        h.open()
        h.loaded = h.call("PyLoad", str(path))
        h.tick()
        h.path = path
        yield h


def registration(host, name):
    (r,) = [r for r in host.registrations() if r.name == name]
    return r


@pytest.mark.parametrize("n", range(31))
def test_each_count_from_0_to_30_passes_each_position_its_value(sig, n):
    values = [float(10 + i) for i in range(n)]  # 10.0 at the first position, 11.0 at the next...
    assert registration(sig, f"Count{n}").type == "Q" * (n + 1)
    assert sig.call(f"Count{n}", *values) == repr(values)


def test_the_fourth_and_fifth_positions_each_have_their_own_value(sig):
    # Registers end after the fourth; the fifth is the first on the stack.
    assert sig.call("Count5", 1.0, 2.0, 3.0, 4.0, 5.0) == "[1.0, 2.0, 3.0, 4.0, 5.0]"
    assert sig.call("Count5", MISSING, MISSING, MISSING, 4.0, 5.0) == "[None, None, None, 4.0, 5.0]"
    assert sig.call("Count5", 1.0, 2.0, 3.0, MISSING, 5.0) == "[1.0, 2.0, 3.0, None, 5.0]"


def test_defaults(sig):
    assert sig.call("SgDefaults", 1.0) == "[1.0, 2.0, 3.0]"
    assert sig.call("SgDefaults", 1.0, MISSING, 9.0) == "[1.0, 2.0, 9.0]"  # Left out in the middle: its default.
    assert sig.call("SgDefaults", 1.0, 8.0, 9.0) == "[1.0, 8.0, 9.0]"


def test_args_with_none_one_and_many(sig):
    assert registration(sig, "SgArgs").type == "Q" * 31
    assert sig.call("SgArgs") == "[]"
    assert sig.call("SgArgs", 1.0) == "[1.0]"
    many = [float(i) for i in range(30)]
    assert sig.call("SgArgs", *many) == repr(many)
    assert sig.call("SgArgs", 1.0, MISSING, 3.0) == "[1.0, None, 3.0]"


def test_keyword_only_parameters(sig):
    assert registration(sig, "SgKw").arguments == "a,k,scale"
    assert sig.call("SgKw", 1.0, 2.0) == "[1.0, 2.0, 1.0]"
    assert sig.call("SgKw", 1.0, 2.0, 3.0) == "[1.0, 2.0, 3.0]"
    assert sig.call("SgKw", 1.0, MISSING, 3.0) == "[1.0, None, 3.0]"  # No default, a later one given.
    assert sig.call("SgKw", 1.0) == Err(VALUE)  # No default, none later: Python's error.
    assert "TypeError: SgKw() missing 1 required keyword-only argument: 'k'" in wait_for(sig, "keyword-only argument: 'k'")


def test_args_and_keyword_only_together_the_plans_examples(sig):
    r = registration(sig, "SgPlan")
    assert r.type == "Q" * 31
    assert r.arguments == ",".join(["x", "scale"] + [f"items{i}" for i in range(1, 29)])
    assert r.argument_help[:3] == ("The first.", "How much.", "Each of the rest.")
    assert sig.call("SgPlan", 10.0, 2.0) == "[10.0, [], 2.0]"  # f(10, scale=2)
    assert sig.call("SgPlan", 10.0, MISSING, 20.0, 30.0) == "[10.0, [20.0, 30.0], 1]"  # f(10, 20, 30)


def test_a_missing_argument_in_the_middle_is_none(sig):
    assert sig.call("SgMiddle", 1.0, MISSING, 3.0) == "[1.0, None, 3.0]"


def test_more_arguments_than_registered_never_reach_the_function(sig):
    # Real Excel gives #VALUE! without calling the function (tests/excel/test_signatures.py);
    # the fake Excel refuses the call.
    with pytest.raises(HostError, match="too many arguments for Count2"):
        sig.call("Count2", 1.0, 2.0, 3.0)
    assert sig.call("Count2", 1.0, 2.0) == "[1.0, 2.0]"


@pytest.mark.parametrize(
    "name, said",
    [
        ("SgOver", "it has 31 parameters, and Excel passes at most 30"),
        ("SgOverArgs", "it has 30 parameters and *rest, and Excel passes at most 30"),
        ("SgUnreadable", "its signature cannot be read"),
        ("sqrt", "a built-in function cannot be marked; wrap it in a function of your own"),
    ],
)
def test_a_refused_signature_leaves_the_rest_registered_and_working(sig, name, said):
    assert f"Pyinex: {name} in the script {sig.path} is not registered: {said}" in sig.stderr
    names = [r.name for r in sig.registrations()]
    assert name not in names and "SgSqrt" not in names
    assert sig.call("SgBefore", 1.0) == "before" and sig.call("SgAfter", 1.0) == "after"


def test_a_class(sig):
    assert registration(sig, "SgMoney").arguments == "amount,rate"
    assert sig.call("SgMoney", 3.0, 2.0) == 6.0  # A float: Excel takes it.
    assert sig.call("SgMoney", 3.0) == 3.0
    assert sig.call("SgPoint", 1.0, 2.0) == Err(VALUE)  # Its instances Excel cannot take.
    assert "Pyinex: SgPoint: SgPoint cannot be returned to Excel" in wait_for(sig, "SgPoint cannot be returned")


def test_a_partial(sig):
    assert registration(sig, "SgOnce").arguments == "price,rate,years"
    assert sig.call("SgOnce", 100.0, 0.5) == 50.0  # years fixed at 1 by the partial.
    assert sig.call("SgOnce", 100.0, 0.5, 2.0) == 25.0


# Reloads that change a function's positions (docs/architecture.md, "Reloading a script").

ONE = "import pyinex\n\n@pyinex.func\ndef {0}({1}):\n    return repr([{2}])\n"


@pytest.fixture
def script(tmp_path):
    return tmp_path / "reshaped.py"


def test_positional_to_args_is_registered_again_and_meanwhile_gives_value(host, script):
    loaded(host, script, ONE.format("Shape", "a, b", "a, b"))
    assert types(host) == {"Shape": "QQQ"}
    host.set_busy(True)  # So the queue waits while calls are made.
    edit(script, ONE.format("Shape", "a, *rest", "a, rest"))
    assert host.call("Shape", 1.0, 2.0) == Err(VALUE)  # Two arguments read, for the old registration.
    assert "Pyinex: Shape is being registered again" in wait_for(host, "being registered again")
    assert types(host) == {"Shape": "QQQ"}
    host.set_busy(False)
    host.tick()
    assert types(host) == {"Shape": "Q" * 31}
    assert host.call("Shape", 1.0, 2.0, 3.0) == "[1.0, (2.0, 3.0)]"


def test_positional_to_keyword_only_with_the_same_names_needs_no_registration(host, script):
    # The names and their positions are the same, so Excel is told nothing new; the new
    # function places the same positions by its own signature.
    loaded(host, script, ONE.format("Kinds", "a, b", "a, b"))
    edit(script, ONE.format("Kinds", "a, *, b=5.0", "a, b"))
    assert host.call("Kinds", 1.0) == "[1.0, 5.0]"
    assert host.call("Kinds", 1.0, 2.0) == "[1.0, 2.0]"
    assert host.tick()[1:] == (0, 0)
    assert types(host) == {"Kinds": "QQQ"}


def test_a_registration_again_with_new_positions_that_fails_is_tried_again_by_pyload(host, script):
    loaded(host, script, ONE.format("Retry", "a", "a"))
    edit(script, ONE.format("Retry", "a, *, k", "a, k"))
    host.fail_registrations(1)
    assert host.call("Retry", 1.0) == Err(VALUE)
    host.tick()
    assert f"Pyinex: Excel refused to register Retry from the script {script}" in wait_for(host, "refused")
    assert "Retry" not in types(host)
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick()
    assert types(host) == {"Retry": "QQQ"}
    assert host.call("Retry", 1.0, 2.0) == "[1.0, 2.0]"


def test_several_edits_of_the_positions_while_work_is_queued(host, script):
    loaded(host, script, ONE.format("Edited", "a", "a"))
    host.set_busy(True)  # Excel refuses to run the queue, so each change waits.
    edit(script, ONE.format("Edited", "a, *rest", "a, rest"))
    assert host.call("Edited", 1.0) == Err(VALUE)
    edit(script, ONE.format("Edited", "a, *, k=2.0", "a, k"))
    assert host.call("PyLoad", str(script)) == "1 function"
    edit(script, ONE.format("Edited", "a, b, *rest, k=3.0", "a, b, rest, k"))
    assert host.call("PyLoad", str(script)) == "1 function"
    host.tick(3)
    assert types(host) == {"Edited": "QQ"}
    host.set_busy(False)
    host.tick()
    assert types(host) == {"Edited": "Q" * 31}
    assert host.call("Edited", 1.0, 2.0, 9.0, 4.0, 5.0) == "[1.0, 2.0, (4.0, 5.0), 9.0]"
    assert len([r for r in host.registrations() if r.procedure.startswith("PyFunc")]) == 1
