"""PyLoad in the fake Excel (docs/architecture.md, "PyLoad"): a script loaded at any time,
whose functions register through the queue once the calculation is over (the queue itself:
test_queue.py), and test group 5 of the plan: relative and absolute paths, each way a
script fails to load, a fixed file, a script loaded twice and from two workbooks, one
script through both PyLoad and PyCall, and a malformed description the loader refuses.

The host refuses registration from a worksheet function, as Excel does, so every test here
also shows PyLoad registers nothing while it calculates.
"""

import pytest
from hostdriver import VALUE, Err
from test_direct_functions import wait_for

GOOD = "import pyinex\n\n@pyinex.func\ndef {0}(x):\n    return '{0}: ' + str(x)\n"


# A script whose top level asks Excel, as only a call may; its function says what it got.
TOP_LEVEL = """import pyinex

seen = []
for ask in (pyinex.caller, pyinex.sheet, pyinex.esc_pressed):
    try:
        ask()
        seen.append("answered")
    except RuntimeError:
        seen.append("RuntimeError")

@pyinex.func
def {0}():
    return ", ".join(seen)
"""


def direct_names(host):
    return [r.name for r in host.registrations() if r.procedure.startswith("PyFunc")]


def queue_command(host):
    """The name of the queue's command, PyinexQueue followed by a number of this copy's own."""
    (queue,) = [r.name for r in host.registrations() if r.procedure == "PyinexQueue"]
    return queue


def load(host, path):
    """PyLoad's result for path, after the queue's timer has fired once."""
    status = host.call("PyLoad", str(path))
    host.tick()
    return status


def test_a_scripts_functions_register_after_the_calculation(host, tmp_path):
    script = tmp_path / "two.py"
    script.write_text(GOOD.format("first") + "\n" + GOOD.format("second"), encoding="utf-8")
    assert host.call("PyLoad", str(script)) == "2 functions"
    host.tick()
    assert direct_names(host) == ["first", "second"]
    assert host.call("first", 1.0) == "first: 1.0"
    assert host.call("second", "x") == "second: x"
    err = wait_for(host, "registered second")
    assert f"Pyinex: registered first from the script {script}" in err
    assert f"Pyinex: registered second from the script {script}" in err


def test_the_status_counts_the_scripts_functions(host, tmp_path):
    one, none = tmp_path / "one.py", tmp_path / "none.py"
    one.write_text(GOOD.format("only"), encoding="utf-8")
    none.write_text("x = 1\n", encoding="utf-8")
    assert load(host, one) == "1 function"
    assert load(host, none) == "0 functions"


def test_a_scripts_top_level_runs_outside_any_call(host, tmp_path):
    script = tmp_path / "top.py"
    script.write_text(TOP_LEVEL.format("top_level"), encoding="utf-8")
    assert load(host, script) == "1 function"
    assert host.call("top_level") == "RuntimeError, RuntimeError, RuntimeError"


def test_a_relative_path_is_relative_to_the_workbooks_folder(host, tmp_path):
    (tmp_path / "near.py").write_text(GOOD.format("near"), encoding="utf-8")
    host.set_folder(tmp_path)
    assert load(host, "near.py") == "1 function"
    assert host.call("near", 2.0) == "near: 2.0"


def test_a_relative_path_in_a_workbook_never_saved_is_refused(host):
    host.set_folder(None)
    assert load(host, "near.py") == Err(VALUE)
    assert "has not been saved" in wait_for(host, "has not been saved")
    assert direct_names(host) == []


def test_a_path_that_is_not_text_is_refused(host):
    assert host.call("PyLoad", 1.0) == Err(VALUE)
    assert "Pyinex: PyLoad needs the script's path as text" in wait_for(host, "PyLoad needs")


@pytest.mark.parametrize(
    "broken, said",
    [
        (None, "Pyinex: cannot find the script"),
        ("def f(:\n", "SyntaxError"),
        ("import pyinex\n\n@pyinex.func\ndef lost(x):\n    return x\n\nraise ValueError('at import')\n", "ValueError: at import"),
        ("import pyinex\n\n@pyinex.func\ndef lost(x):\n    return x\n\nraise SystemExit(4)\n", "SystemExit: 4"),
        ("import pyinex\n\n@pyinex.func\ndef lost(x):\n    return x\n\nraise KeyboardInterrupt\n", "KeyboardInterrupt"),
    ],
    ids=["missing", "syntax error", "ValueError", "SystemExit", "KeyboardInterrupt"],
)
def test_a_script_that_fails_registers_nothing_until_it_is_fixed(host, tmp_path, broken, said):
    script = tmp_path / "broken.py"
    if broken is not None:
        script.write_text(broken, encoding="utf-8")
    assert load(host, script) == Err(VALUE)
    err = wait_for(host, said)
    assert said in err and "broken.py" in err
    assert direct_names(host) == []
    # The next calculation of the cell loads the fixed file, with no restart.
    script.write_text(GOOD.format("lost"), encoding="utf-8")
    assert load(host, script) == "1 function"
    assert host.call("lost", 3.0) == "lost: 3.0"


def test_a_script_loaded_twice_registers_once(host, tmp_path):
    script = tmp_path / "twice.py"
    script.write_text(GOOD.format("twice"), encoding="utf-8")
    assert load(host, script) == "1 function"
    assert load(host, script) == "1 function"
    assert direct_names(host) == ["twice"]


def test_a_script_loaded_from_two_workbooks_registers_once_for_both(host, tmp_path):
    (tmp_path / "shared.py").write_text(GOOD.format("shared"), encoding="utf-8")
    host.set_folder(tmp_path)
    host.set_caller("[Book1]Sheet1", 1, 1)
    assert load(host, "shared.py") == "1 function"
    host.set_caller("[Book2]Sheet1", 1, 1)
    assert load(host, "shared.py") == "1 function"
    assert direct_names(host) == ["shared"]
    assert host.call("shared", 4.0) == "shared: 4.0"


def test_one_script_through_pyload_and_pycall_is_one_module(host, tmp_path):
    script = tmp_path / "counted.py"
    script.write_text("import pyinex\n\nn = 0\n\n@pyinex.func\ndef count():\n    global n\n    n += 1\n    return n\n",
                      encoding="utf-8")
    assert load(host, script) == "1 function"
    assert host.call("count") == 1.0
    assert host.call("PyCall", str(script), "count") == 2.0
    assert host.call("count") == 3.0


def test_a_name_another_script_registered_is_refused_naming_both(host, tmp_path):
    first, second = tmp_path / "first.py", tmp_path / "second.py"
    first.write_text(GOOD.format("Taken"), encoding="utf-8")
    second.write_text(GOOD.format("taken") + "\n" + GOOD.format("free"), encoding="utf-8")
    assert load(host, first) == "1 function"
    assert load(host, second) == "1 function"
    assert direct_names(host) == ["Taken", "free"]
    err = wait_for(host, "registered that name first")
    assert f"Pyinex: taken in the script {second} is not registered: the script {first}" in err
    assert host.call("taken", 1.0) == "Taken: 1.0"  # Still the first script's.


def test_in_the_function_wizard_pyload_is_blank_and_loads_nothing(host, tmp_path):
    marker = tmp_path / "loaded"
    script = tmp_path / "wizard.py"
    script.write_text(f"open({str(marker)!r}, 'w').close()\n" + GOOD.format("wizardly"), encoding="utf-8")
    host.set_wizard(True)
    assert load(host, script) == ""
    host.set_wizard(False)
    assert not marker.exists() and direct_names(host) == []
    assert load(host, script) == "1 function" and marker.exists()


# Python never gives the loader a malformed description, so this script makes one: it has
# Pyinex's describer, found through pyinex.func's globals, give the first function a count
# of arguments past Excel's 30, once.
MALFORMED = """import pyinex

describe = pyinex.func.__globals__['_description']

def without_help(*args):
    pyinex.func.__globals__['_description'] = describe
    row, notes = describe(*args)
    row[1] = 31
    return row, notes

pyinex.func.__globals__['_description'] = without_help

@pyinex.func
def Malformed(a, b):
    return a

@pyinex.func
def WellFormed(a, b):
    return b
"""


def test_a_malformed_description_is_refused_before_excel_is_told(host, tmp_path):
    script = tmp_path / "malformed.py"
    script.write_text(MALFORMED, encoding="utf-8")
    assert host.call("PyLoad", str(script)) == "2 functions"
    host.tick()
    assert direct_names(host) == ["WellFormed"]
    said = wait_for(host, "malformed")
    assert f"Pyinex: cannot register a function of the script {script}: its description is malformed" in said
    assert host.call("WellFormed", 1.0, 2.0) == 2.0
