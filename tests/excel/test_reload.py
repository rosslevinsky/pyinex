"""Reloading a changed script in real Excel (docs/architecture.md, "Scripts" and
"PyReloadOnChange"): what reloads, what does not, and module variables across both.

Each change is a real save of the file, as an editor would make it; the formula is then
entered again, which recalculates it.
"""

import os
import time

import pytest
from conftest import E, pycall, shown


def returns(value):
    return f"def f():\n    return {value!r}\n"


@pytest.fixture
def setting(app):
    """A function setting PyReloadOnChange, which is one setting for the whole Excel: it is
    turned back on after the test."""

    def set_to(on):
        return app.Evaluate(f"PyReloadOnChange({'TRUE' if on else 'FALSE'})")

    try:
        yield set_to
    finally:
        set_to(True)


def test_a_changed_script_is_reloaded(book):
    book.script("changing.py", returns("first"))
    assert book.formula("A1", pycall("changing.py", "f")) == "first"
    book.script("changing.py", returns("second, and longer"))
    assert book.formula("A1", pycall("changing.py", "f")) == "second, and longer"


def test_an_edit_that_keeps_the_size_within_the_same_second_is_reloaded(book):
    # Python's compiled-file cache records a source's time only to the second, which is
    # why Pyinex compiles each script from its source.
    path = book.folder / "same_size.py"
    for _ in range(10):
        book.script("same_size.py", returns(1))
        assert book.formula("A1", pycall("same_size.py", "f")) == 1.0
        first = os.stat(path)
        time.sleep(0.05)  # Past the clock's tick, so the times differ.
        book.script("same_size.py", returns(2))
        second = os.stat(path)
        if first.st_mtime_ns // 10**9 == second.st_mtime_ns // 10**9:
            break
        time.sleep(0.5)  # The second turned over between the saves; try again.
    else:
        pytest.fail("could not save twice within one second")
    assert first.st_size == second.st_size and first.st_mtime_ns != second.st_mtime_ns
    assert book.formula("A1", pycall("same_size.py", "f")) == 2.0


def test_reload_on_change_off_keeps_the_loaded_script(book, setting, app):
    book.script("kept.py", returns("first"))
    assert book.formula("A1", pycall("kept.py", "f")) == "first"
    assert app.Evaluate("PyReloadOnChange()") is True
    assert setting(False) is False
    assert app.Evaluate("PyReloadOnChange()") is False
    book.script("kept.py", returns("second, and longer"))
    assert book.formula("A1", pycall("kept.py", "f")) == "first"
    assert setting(True) is True
    # Turning it on recalculates nothing: the cell still shows what it showed.
    assert book.value("A1") == "first"
    assert app.Evaluate("PyReloadOnChange()") is True
    assert book.formula("A1", pycall("kept.py", "f")) == "second, and longer"


def test_reload_on_change_refuses_text(app, setting):
    assert shown(app.Evaluate('PyReloadOnChange("yes")')) == E("#VALUE!")
    assert app.Evaluate("PyReloadOnChange()") is True


def test_a_change_of_size_alone_is_reloaded(book):
    # The time is put back as it was, so only the size says the file changed.
    path = book.folder / "sized.py"
    book.script("sized.py", returns("short"))
    assert book.formula("A1", pycall("sized.py", "f")) == "short"
    before = os.stat(path)
    book.script("sized.py", returns("rather longer"))
    os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns))
    assert os.stat(path).st_mtime_ns == before.st_mtime_ns
    assert book.formula("A1", pycall("sized.py", "f")) == "rather longer"


def test_no_compiled_file_is_written_for_a_script(book):
    book.script("uncached.py", returns("compiled from source"))
    assert book.formula("A1", pycall("uncached.py", "f")) == "compiled from source"
    assert not (book.folder / "__pycache__").exists()


def test_a_function_removed_by_a_reload_is_gone(book, cleared):
    book.script("shrinks.py", returns("kept") + "\ndef g():\n    return 'removed later'\n")
    assert book.formula("A1", pycall("shrinks.py", "g")) == "removed later"
    book.script("shrinks.py", returns("kept"))
    assert book.formula("A1", pycall("shrinks.py", "g")) == E("#VALUE!")
    window = cleared("has no function g")
    assert window and "has no function g" in window.text
    assert book.formula("A2", pycall("shrinks.py", "f")) == "kept"


def test_a_module_variable_persists_between_calls_2010(book):
    book.copy("module.py")
    assert [book.formula(f"A{i}", pycall("module.py", "PlayWithGlobalVariable")) for i in (1, 2, 3)] == [5.0, 6.0, 7.0]


def test_a_module_variable_starts_over_on_a_reload(book):
    book.script("counter.py", "n = 0\n\ndef count():\n    global n\n    n += 1\n    return n\n")
    assert [book.formula(f"A{i}", pycall("counter.py", "count")) for i in (1, 2)] == [1.0, 2.0]
    book.script("counter.py", "n = 0\n\ndef count():\n    global n\n    n += 1\n    return n  # changed\n")
    assert book.formula("A3", pycall("counter.py", "count")) == 1.0


def test_a_change_to_a_syntax_error_drops_the_old_version(book, cleared, setting):
    book.script("breaks.py", returns("working"))
    assert book.formula("A1", pycall("breaks.py", "f")) == "working"
    book.script("breaks.py", "def f(:\n    return 'broken'\n")
    assert book.formula("A1", pycall("breaks.py", "f")) == E("#VALUE!")
    said = f"Pyinex: the script {book.folder / 'breaks.py'} failed to load:"
    window = cleared("SyntaxError")
    assert window and said.lower() in window.text.lower() and "SyntaxError" in window.text
    # The old version is gone, not kept for later: a second call fails too, and so does
    # one with reloading off, which would use a kept module without looking at the file.
    assert book.formula("A1", pycall("breaks.py", "f")) == E("#VALUE!")
    assert setting(False) is False
    assert book.formula("A1", pycall("breaks.py", "f")) == E("#VALUE!")
    assert setting(True) is True
    book.script("breaks.py", returns("fixed"))
    assert book.formula("A1", pycall("breaks.py", "f")) == "fixed"


@pytest.mark.parametrize(
    "statement, shown_as",
    [("raise ValueError('on reload')", "ValueError: on reload"), ("raise SystemExit(5)", "SystemExit: 5"), ("raise KeyboardInterrupt", "KeyboardInterrupt")],
)
def test_a_reload_that_raises_drops_the_old_version(book, cleared, setting, statement, shown_as):
    book.script("raises.py", returns("working"))
    assert book.formula("A1", pycall("raises.py", "f")) == "working"
    book.script("raises.py", returns("working") + statement + "\n")
    assert book.formula("A1", pycall("raises.py", "f")) == E("#VALUE!")
    window = cleared(shown_as)
    assert window and "failed to load:" in window.text and "Traceback (most recent call last)" in window.text and shown_as in window.text
    assert setting(False) is False  # Reloading off would use a kept module: there is none.
    assert book.formula("A1", pycall("raises.py", "f")) == E("#VALUE!")


def test_a_deleted_script(book, cleared):
    book.script("goes.py", returns("here"))
    assert book.formula("A1", pycall("goes.py", "f")) == "here"
    (book.folder / "goes.py").unlink()
    assert book.formula("A1", pycall("goes.py", "f")) == E("#VALUE!")
    said = f"Pyinex: cannot find the script {book.folder / 'goes.py'}"
    window = cleared("cannot find the script")
    assert window and said.lower() in window.text.lower()


def test_an_edited_helper_is_not_reloaded(book):
    helper = book.folder / "helper_reload.py"
    helper.write_text("VALUE = 'the first helper'\n", encoding="utf-8")
    book.script("uses.py", "import helper_reload\n\ndef f():\n    return helper_reload.VALUE\n")
    assert book.formula("A1", pycall("uses.py", "f")) == "the first helper"
    helper.write_text("VALUE = 'the second helper, longer'\n", encoding="utf-8")
    assert book.formula("A1", pycall("uses.py", "f")) == "the first helper"
    # Even reloading the script that imports it keeps the helper Python already imported.
    book.script("uses.py", "import helper_reload\n\ndef f():\n    return helper_reload.VALUE  # changed\n")
    assert book.formula("A1", pycall("uses.py", "f")) == "the first helper"
