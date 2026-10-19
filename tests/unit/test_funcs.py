"""The Python half of direct functions (src/core/python/funcs.py; docs/architecture.md,
"Direct functions" and "PyLoad"): what @pyinex.func collects from a script, its help from the
docstring, a signature refused, nothing from a script that raised, and which names
are registered already.

The core runs funcs.py in call.py's namespace, and so do these tests, under plain Python:
they need no Excel and no extension module. A stand-in pyinex module offers the decorator.
"""

import sys

import pytest

# The fixtures core and load are in conftest.py.


def test_collects_the_scripts_marked_functions_in_order(load):
    rows = load(
        """
        import pyinex

        @pyinex.func
        def second(b, a):
            return a

        def plain(x):
            return x

        @pyinex.func()
        def first():
            return 1
        """
    )
    assert [row[1:4] for row in rows] == [["second", 2, "b,a"], ["first", 0, ""]]
    assert [row[4] for row in rows] == ["Pyinex", "Pyinex"]
    assert len({row[0] for row in rows}) == 2


def test_the_decorator_returns_the_function_unchanged(core):
    def f(x):
        return x

    assert core["func"](f) is f
    assert core["func"](name="g")(f) is f


def test_name_category_and_help_from_the_decorator(load):
    (row,) = load(
        """
        import pyinex

        @pyinex.func(name="Discounted", category="Finance", help="Given help.")
        def discount(price, rate):
            '''Docstring help.'''
            return price * (1 - rate)
        """
    )
    assert row[1:6] == ["Discounted", 2, "price,rate", "Finance", "Given help."]


def test_help_and_argument_help_from_the_docstring(load):
    (row,) = load(
        '''
        import pyinex

        @pyinex.func
        def discount(price, rate, years=1):
            """Price after a discount,
            compounded.

            More detail, not shown.

            Args:
                price: The price before
                    the discount.
                rate (float): The discount, as a fraction.
                unknown: Not a parameter.

            Returns:
                The price.
            """
        '''
    )
    assert row[5] == "Price after a discount, compounded."
    assert row[6:] == ["The price before the discount.", "The discount, as a fraction.", ""]


def test_a_docstring_without_args_gives_function_help_only(load):
    (row,) = load(
        """
        import pyinex

        @pyinex.func
        def f(a, b):
            '''One line.

            a and b are numbers.'''
        """
    )
    assert row[5:] == ["One line.", "", ""]


def test_no_docstring_gives_empty_help(load):
    (row,) = load("import pyinex\n\n@pyinex.func\ndef f(a):\n    pass\n")
    assert row[5:] == ["", ""]


def test_texts_longer_than_excel_takes_are_cut(load):
    # A name that long is refused instead (test_funcs_names.py): cut, it would be another name.
    long = "x" * 300
    (row,) = load(
        f"import pyinex\n\n@pyinex.func(category='{long}', help='{long}')\n"
        f"def f(a):\n    '''Help.\n\n    Args:\n        a: {long}\n    '''\n"
    )
    assert [len(t) for t in (row[4], row[5], row[6])] == [255] * 3


@pytest.mark.parametrize("length, names", [(255, "a" * 255), (256, "")], ids=["fits", "too long"])
def test_argument_names_are_given_only_whole(load, length, names):
    # A first name too long for Excel's 255 characters gives no names, not part of one.
    (row,) = load(f"import pyinex\n\n@pyinex.func\ndef f({'a' * length}, b):\n    pass\n")
    assert row[2:4] == [2, names]


# Every refusal, and the *args, keyword-only parameters and **kwargs accepted:
# test_funcs_signatures.py.
@pytest.mark.parametrize(
    "signature, why",
    [
        (", ".join(f"a{i}" for i in range(31)), "it has 31 parameters, and Excel passes at most 30"),
    ],
)
def test_signatures_refused_are_refused_and_the_rest_register(load, capfd, signature, why):
    rows = load(
        f"import pyinex\n\n@pyinex.func\ndef before(a):\n    pass\n\n"
        f"@pyinex.func\ndef refused({signature}):\n    pass\n\n@pyinex.func\ndef after(b):\n    pass\n"
    )
    assert [row[1] for row in rows] == ["before", "after"]
    err = capfd.readouterr().err
    assert "Pyinex: refused in the script" in err
    assert f"is not registered: {why}" in err


def test_thirty_positional_parameters_are_accepted(load):
    (row,) = load("import pyinex\n\n@pyinex.func\ndef f(" + ", ".join(f"a{i}" for i in range(30)) + "):\n    pass\n")
    assert row[2] == 30 and len(row) == 6 + 30


def test_an_unreadable_signature_is_refused(load, capfd):
    rows = load("import functools, pyinex\n\nunreadable = pyinex.func(functools.partial(max))\n")
    assert rows == []
    assert "its signature cannot be read" in capfd.readouterr().err


def test_a_signature_that_raises_as_it_is_read_is_refused_and_the_rest_register(load, capfd):
    rows = load(
        "import pyinex\n\n@pyinex.func\ndef before(a):\n    pass\n\n"
        "class Odd:\n    def __call__(self, x):\n        return x\n\n"
        "    @property\n    def __signature__(self):\n        raise RuntimeError('no signature here')\n\n"
        "odd = pyinex.func(Odd())\n\n@pyinex.func\ndef after(b):\n    pass\n"
    )
    assert [row[1] for row in rows] == ["before", "after"]
    err = capfd.readouterr().err
    assert "odd in the script" in err and "its signature cannot be read" in err and "no signature here" in err


def test_any_other_failure_describing_a_function_refuses_only_that_function(load, capfd):
    rows = load(
        "import pyinex\n\nclass Name:\n    def __str__(self):\n        raise RuntimeError('no name here')\n\n"
        "@pyinex.func(name=Name())\ndef odd(a):\n    pass\n\n@pyinex.func\ndef after(b):\n    pass\n"
    )
    assert [row[1] for row in rows] == ["after"]
    assert "odd in the script" in (err := capfd.readouterr().err) and "RuntimeError: no name here" in err


@pytest.mark.skipif(sys.version_info < (3, 14), reason="before 3.14 the annotation is evaluated as the script runs")
def test_an_annotation_only_a_type_checker_imports_does_not_stop_registering(load):
    rows = load(
        "from typing import TYPE_CHECKING\nimport pyinex\n\nif TYPE_CHECKING:\n    from decimal import Decimal\n\n"
        "@pyinex.func\ndef price(x: Decimal) -> Decimal:\n    return x\n\n@pyinex.func\ndef after(b):\n    pass\n"
    )
    assert [row[1:4] for row in rows] == [["price", 1, "x"], ["after", 1, "b"]]


@pytest.mark.parametrize("statement, shown", [("raise ValueError('at import')", "ValueError: at import"),
                                              ("raise SystemExit(3)", "SystemExit: 3"),
                                              ("raise KeyboardInterrupt", "KeyboardInterrupt")])
def test_a_script_that_raises_after_a_marked_function_registers_nothing(core, load, capfd, statement, shown):
    rows = load(f"import pyinex\n\n@pyinex.func\ndef f(a):\n    return a\n\n{statement}\n")
    assert rows is core["FAILED"]
    err = capfd.readouterr().err
    assert "failed to load:" in err and shown in err


def test_a_missing_script_registers_nothing(core, tmp_path, capfd):
    assert core["load_functions"](str(tmp_path / "missing.py")) is core["FAILED"]
    assert "cannot find the script" in capfd.readouterr().err


def test_functions_the_script_imports_are_not_its_own(load, script):
    script("import pyinex\n\n@pyinex.func\ndef helped(a):\n    return a\n", name="funcs_helper_module.py")
    rows = load("import pyinex\nfrom funcs_helper_module import helped\n\n@pyinex.func\ndef own(a):\n    return a\n")
    assert [row[1] for row in rows] == ["own"]


def test_one_function_under_two_names_registers_once(load):
    rows = load("import pyinex\n\n@pyinex.func\ndef f(a):\n    return a\n\ng = f\n")
    assert [row[1] for row in rows] == ["f"]


def test_a_collected_function_is_called_as_pycall_calls_it(core, load, capfd):
    rows = load("import pyinex\n\n@pyinex.func(name='Add')\ndef add(a, b=10):\n    return a + b\n\n"
                "@pyinex.func\ndef fails():\n    raise ValueError('deliberately')\n")
    add, fails = (row[0] for row in rows)
    assert core["call_function"](add, (1.0,), True) == 11.0
    assert core["call_function"](add, (1.0, 2.0), True) == 3.0
    assert core["call_function"](fails, (), True) is core["FAILED"]
    err = capfd.readouterr().err
    assert "Pyinex: fails in the script" in err and "ValueError: deliberately" in err


# PyLoad's bookkeeping: which names are registered, or waiting to be, and from which script
# (docs/architecture.md, "PyLoad").

ONE = "import pyinex\n\n@pyinex.func\ndef {0}(a):\n    return a\n"


def test_loading_again_registers_only_what_is_new(core, script):
    path = script(ONE.format("one"), name="again.py")
    (first,) = core["load_functions"](str(path))
    assert first[1] == "one" and first[0] >= 0
    # Already registered from this script: a row of the key -1 alone, counted in the status.
    assert core["load_functions"](str(path)) == [[-1]]
    path.write_text(ONE.format("one") + "\n@pyinex.func\ndef two():\n    return 2\n", encoding="utf-8")
    rows = core["load_functions"](str(path))
    assert rows[0] == [-1] and rows[1][1] == "two" and rows[1][0] >= 0


def test_a_name_another_script_registered_is_refused_naming_both(core, script, capfd):
    first = script(ONE.format("Shared"), name="first_owner.py")
    second = script(ONE.format("shared") + "\n@pyinex.func\ndef other():\n    pass\n", name="second_owner.py")
    assert [row[1] for row in core["load_functions"](str(first))] == ["Shared"]
    capfd.readouterr()
    # Names compare without regard to case, as Excel compares them.
    assert [row[1] for row in core["load_functions"](str(second))] == ["other"]
    err = capfd.readouterr().err
    assert f"shared in the script {second} is not registered" in err and str(first) in err


def test_a_function_that_did_not_register_frees_its_name(core, script):
    path = script(ONE.format("lost"), name="lost.py")
    (row,) = core["load_functions"](str(path))
    core["registered"](row[0], False)
    (again,) = core["load_functions"](str(path))
    assert again[1] == "lost" and again[0] >= 0
    core["registered"](again[0], True)
    assert core["load_functions"](str(path)) == [[-1]]


def test_a_relative_path_is_relative_to_the_folder_given(core, script):
    path = script(ONE.format("near"), name="near.py")
    assert [row[1] for row in core["load_functions"]("near.py", str(path.parent))] == ["near"]


def test_a_relative_path_in_a_workbook_never_saved_is_refused(core, capfd):
    assert core["load_functions"]("near.py", "") is core["FAILED"]
    assert "this workbook has not been saved" in capfd.readouterr().err


def test_a_script_that_raised_registers_once_it_is_fixed(core, script):
    path = script(ONE.format("later") + "\nraise ValueError('not yet')\n", name="later.py")
    assert core["load_functions"](str(path)) is core["FAILED"]
    path.write_text(ONE.format("later"), encoding="utf-8")
    assert [row[1] for row in core["load_functions"](str(path))] == ["later"]


def test_one_functions_help_raising_systemexit_refuses_only_that_function(core, script, capfd):
    # Help that raises SystemExit as it is read refuses its own function, which then gives no
    # row, and lets the script's other functions register.
    path = script(
        """
        import pyinex

        class Help:
            def __str__(self):
                raise SystemExit("bad help")

        @pyinex.func
        def Good(x):
            return x

        @pyinex.func(help=Help())
        def Broken(x):
            return x
        """,
        name="exits.py",
    )
    rows = core["load_functions"](str(path))
    assert [row[1] for row in rows] == ["Good"]
    assert "Broken in the script" in capfd.readouterr().err
    core["registered"](rows[0][0], True)
    assert core["load_functions"](str(path)) == [[-1]]


@pytest.mark.parametrize("what", ["name", "argument names", "help"])
def test_excel_s_limits_count_characters_beyond_u_ffff_as_two(core, script, what, capfd):
    # Excel counts text in UTF-16 code units: 200 such characters are 400, past its 255.
    wide = "\U00020000" * 200
    if what == "name":
        source = f"import pyinex\n\n@pyinex.func(name={wide!r})\ndef f(x):\n    return x\n"
    elif what == "argument names":
        source = f"import pyinex\n\n@pyinex.func\ndef f({wide[:100]}, {wide[100:150]}b):\n    return 0\n"
    else:
        source = f"import pyinex\n\n@pyinex.func(help={wide!r})\ndef f(x):\n    return x\n"
    rows = core["load_functions"](str(script(source, name="wide.py")))
    if what == "name":
        assert rows == [] and "at most 255 characters" in capfd.readouterr().err
        return
    (row,) = rows
    units = lambda t: len(t.encode("utf-16-le")) // 2
    if what == "argument names":
        assert row[3] == wide[:100]  # The first name alone fits: 200 units.
    else:
        assert units(row[5]) <= 255 and row[5] == wide[:127]  # Never half a character.


def test_refusals_and_removals_leave_nothing_behind(core, script):
    # Each refusal of a registration, and each removal Excel has made, is forgotten: Python's
    # records stay the same size however often it happens.
    path = script(ONE.format("Discount"), name="discount.py")
    for _ in range(1000):
        (row,) = core["load_functions"](str(path))
        core["registered"](row[0], False)
    assert len(core["_live"]) == 0 and len(core["_gone"]) == 0


def test_rows_that_did_not_reach_the_loader_are_offered_again(core, script):
    # As when memory runs out converting load_functions' rows: the new function is offered
    # again, and the removal waits for changes.
    path = script(ONE.format("Kept") + "\n" + ONE.format("Lost"), name="offered.py")
    rows = core["load_functions"](str(path))
    for row in rows:
        core["registered"](row[0], True)
    lost = rows[1][0]
    path.write_text(ONE.format("Kept") + "\n" + ONE.format("Added"), encoding="utf-8")
    import os
    os.utime(path, ns=(path.stat().st_atime_ns, path.stat().st_mtime_ns + 2_000_000_000))
    rows = core["load_functions"](str(path))
    assert rows[0] == [lost] and [r[1] for r in rows[1:] if len(r) > 1] == ["Added"]
    core["_unoffer"]()
    assert core["changes"]() == [[str(path), lost]]
    again = core["load_functions"](str(path))
    assert [r[1] for r in again if len(r) > 1] == ["Added"]


def test_changes_that_did_not_reach_the_loader_wait_for_its_next_call(core):
    core["_pending"].extend([["a.py", 1], ["b.py", 2]])
    first = core["changes"]()
    core["_pending"].append(["c.py", 3])
    core["_restore_changes"](first)
    assert core["changes"]() == [["a.py", 1], ["b.py", 2], ["c.py", 3]]


def test_two_scripts_whose_paths_collide_in_crc32_have_modules_of_their_own(core, script, monkeypatch):
    import types

    monkeypatch.setitem(core, "zlib", types.SimpleNamespace(crc32=lambda b: 0))
    first = script("import sys\nVALUE = 1\ndef f():\n    return sys.modules[__name__].VALUE\n")
    a = core["_load"](str(first), True)
    second = first.parent / "other" / "model.py"
    second.parent.mkdir()
    second.write_text("import sys\nVALUE = 2\ndef f():\n    return sys.modules[__name__].VALUE\n", encoding="utf-8")
    b = core["_load"](str(second), True)
    assert a.__name__ != b.__name__
    assert (a.f(), b.f()) == (1, 2)


def test_a_call_excel_stopped_says_nothing_of_its_failure(core, script, capfd, monkeypatch):
    path = script("def f():\n    raise RuntimeError('stopped')\n", name="stopped.py")
    monkeypatch.setattr(sys.modules["pyinex"], "_stopped", lambda: True)
    assert core["call"](str(path), "f", (), True, None) is core["FAILED"]
    assert capfd.readouterr().err == ""
