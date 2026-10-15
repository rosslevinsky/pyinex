"""Loading scripts by path, reloading them when they change, and failing to load
(docs/architecture.md, "Scripts"), through the add-in's own code."""

import os

VALUE_ERROR = ("err", 15)


def test_load_by_path(pycall, script):
    path = script("def f():\n    return 'loaded'\n")
    assert pycall(path, "f") == ("str", "loaded")


def test_same_named_scripts_in_different_folders(pycall, script, tmp_path):
    a = script("def f():\n    return 'a'\n", folder=tmp_path / "a")
    b = script("def f():\n    return 'b'\n", folder=tmp_path / "b")
    assert pycall(a, "f") == ("str", "a")
    assert pycall(b, "f") == ("str", "b")
    assert pycall(a, "f") == ("str", "a")


def test_module_name(pycall, script):
    path = script("def f():\n    return __name__\n")
    kind, name = pycall(path, "f")
    assert name.startswith("pyinex_model_") and name != "__main__"


def test_dotted_file_name_is_not_a_package(pycall, script):
    path = script("def f():\n    return __name__ + '|' + str(__package__)\n", name="my.model.py")
    kind, shown = pycall(path, "f")
    name, package = shown.split("|")
    assert "." not in name and package in ("", "None")


def test_helper_beside_the_script(pycall, script, tmp_path):
    folder = tmp_path / "with_helper"
    script("VALUE = 'from helper'\n", name="pyinex_test_helper.py", folder=folder)
    path = script(
        """
        import pyinex_test_helper

        def at_top():
            return pyinex_test_helper.VALUE

        def inside():
            import pyinex_test_helper as h
            return h.VALUE
        """,
        folder=folder,
    )
    assert pycall(path, "at_top") == ("str", "from helper")
    assert pycall(path, "inside") == ("str", "from helper")


def test_unchanged_script_is_kept(pycall, script):
    path = script("calls = 0\n\ndef f():\n    global calls\n    calls += 1\n    return calls\n")
    assert pycall(path, "f") == ("num", 1.0)
    assert pycall(path, "f") == ("num", 2.0)


def test_changed_script_reloads_with_fresh_state(pycall, script):
    path = script("calls = 0\n\ndef f():\n    global calls\n    calls += 1\n    return calls\n")
    assert pycall(path, "f") == ("num", 1.0)
    assert pycall(path, "f") == ("num", 2.0)
    script("calls = 100\n\ndef f():\n    global calls\n    calls += 1\n    return calls\n")
    assert pycall(path, "f") == ("num", 101.0)


def test_same_size_edit_within_the_same_second(pycall, script):
    path = script("def f():\n    return 1\n")
    assert pycall(path, "f") == ("num", 1.0)
    before = os.stat(path).st_mtime_ns
    script("def f():\n    return 2\n")
    # The same size, and a modification time one microsecond later: the same second, so
    # a cache keyed on whole seconds would miss it.
    os.utime(path, ns=(before, before + 1000))
    assert os.stat(path).st_size == len("def f():\n    return 1\n")
    assert pycall(path, "f") == ("num", 2.0)


def test_reload_off_keeps_the_loaded_version(pycall, script):
    path = script("def f():\n    return 'first'\n")
    assert pycall(path, "f", reload_on_change=False) == ("str", "first")
    script("def f():\n    return 'second, longer'\n")
    assert pycall(path, "f", reload_on_change=False) == ("str", "first")
    assert pycall(path, "f") == ("str", "second, longer")


def test_no_bytecode_cache_written(pycall, script, tmp_path):
    path = script("def f():\n    return 1\n", folder=tmp_path / "clean")
    assert pycall(path, "f") == ("num", 1.0)
    assert sorted(p.name for p in path.parent.iterdir()) == ["model.py"]


def test_missing_script(pycall, tmp_path, capfd):
    assert pycall(tmp_path / "nowhere.py", "f") == VALUE_ERROR
    assert "cannot find the script" in capfd.readouterr().err


def test_relative_path_uses_the_workbook_folder(pycall, script, tmp_path):
    script("def f():\n    return 'relative'\n", name="rel.py", folder=tmp_path / "book")
    assert pycall("rel.py", "f", folder=str(tmp_path / "book") + "\\") == ("str", "relative")


def test_relative_path_in_an_unsaved_workbook(pycall, capfd):
    assert pycall("rel.py", "f") == VALUE_ERROR
    assert "has not been saved" in capfd.readouterr().err


def test_script_that_fails_to_import(pycall, script, capfd):
    path = script("def f(:\n")
    assert pycall(path, "f") == VALUE_ERROR
    err = capfd.readouterr().err
    assert "failed to load" in err and "SyntaxError" in err
    script("def f():\n    return 'fixed'\n")
    assert pycall(path, "f") == ("str", "fixed")


def test_failed_import_leaves_no_module(pycall, script, testing, capfd):
    path = script("import sys\nsys.modules['pyinex_seen'] = __name__\nraise RuntimeError('top level')\n")
    assert pycall(path, "f") == VALUE_ERROR
    assert "RuntimeError: top level" in capfd.readouterr().err
    import sys

    assert sys.modules.pop("pyinex_seen") not in sys.modules


def test_failed_reload_drops_the_old_version(pycall, script, capfd):
    path = script("def f():\n    return 'old'\n")
    assert pycall(path, "f") == ("str", "old")
    script("def f():\n    return 'new'\nraise SystemExit(1)\n")
    assert pycall(path, "f") == VALUE_ERROR
    assert pycall(path, "f") == VALUE_ERROR  # Tried again, not the old version.
    assert "SystemExit" in capfd.readouterr().err


def test_exception_while_importing_then_a_good_call(pycall, script, capfd):
    bad = script("raise KeyboardInterrupt\n", name="bad.py")
    good = script("def f():\n    return 'good'\n", name="good.py")
    assert pycall(bad, "f") == VALUE_ERROR
    assert "KeyboardInterrupt" in capfd.readouterr().err
    assert pycall(good, "f") == ("str", "good")


def test_upper_case_suffix(pycall, script):
    path = script("def f():\n    return 'shouted'\n", name="LOUD.PY")
    assert pycall(path, "f") == ("str", "shouted")
    lower = path.with_name("loud.py")  # The same file, named differently.
    assert pycall(lower, "f") == ("str", "shouted")
