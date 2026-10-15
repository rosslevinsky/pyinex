"""Each registered function called through the fake Excel (docs/architecture.md, "Excel
functions"): results, error cells, the function wizard and freeing of returned values.

Every call also checks Excel's memory rules: the driver fails the test on any fault, so a
returned value that is not freed exactly as Excel frees it fails here too."""

import os
import re
import sys
import time

import pytest
from conftest import HELPERS, ROOT, host_env
from hostdriver import VALUE, Err, Host


def echo(host, *args):
    return host.call("PyCall", str(HELPERS), "echo", *args)


def wait_for(host, text, timeout=10):
    """The host's standard error once it contains text; output arrives on a thread."""
    deadline = time.monotonic() + timeout
    while text not in host.stderr and time.monotonic() < deadline:
        time.sleep(0.05)
    return host.stderr


# PyCall


@pytest.mark.parametrize("value", [1.5, -0.0, 1e308, "text", "", "é ✓ 😀", "a\ud800b", True, False, Err(VALUE), Err(42)])
def test_pycall_round_trips_values(host, value):
    assert echo(host, value) == value


def test_pycall_round_trips_text_at_excels_limit(host):
    assert echo(host, "x" * 32767) == "x" * 32767


def test_pycall_round_trips_ranges(host):
    block = ((1.0, "two"), (True, ""), (Err(7), 6.0))
    assert echo(host, block) == block
    assert echo(host, ((1.0, 2.0, 3.0),)) == ((1.0, 2.0, 3.0),)
    assert echo(host, ((1.0,), (2.0,))) == ((1.0,), (2.0,))


def test_pycall_none_is_a_blank_cell(host):
    assert echo(host) == ""
    assert echo(host, None) == ""


def test_pycall_result_is_freed_by_excel(host):
    result, dll_free, faults = host.call_detail("PyCall", str(HELPERS), "echo", ((1.0, "a"), ("b", 2.0)))
    assert dll_free and not faults


def test_pycall_exception_is_value_error_with_traceback(host, tmp_path):
    path = tmp_path / "raises.py"
    path.write_text("def f():\n    raise ValueError('from the script')\n", encoding="utf-8")
    assert host.call("PyCall", str(path), "f") == Err(VALUE)
    err = wait_for(host, "ValueError: from the script")
    assert "Traceback (most recent call last)" in err


def test_pycall_returned_excel_error(host, tmp_path):
    path = tmp_path / "errors.py"
    path.write_text("import math\ndef f():\n    return [1, math.nan]\n", encoding="utf-8")
    assert host.call("PyCall", str(path), "f") == ((1.0, Err(36)),)


def test_pycall_relative_path_uses_the_workbook_folder(host, tmp_path):
    (tmp_path / "rel.py").write_text("def f():\n    return 'beside the workbook'\n", encoding="utf-8")
    host.set_folder(tmp_path)
    assert host.call("PyCall", "rel.py", "f") == "beside the workbook"


def test_pycall_relative_path_when_excel_gives_the_caller_without_its_sheet(host, tmp_path):
    # xlfCaller may give the calling cell as an xltypeSRef, which CELL takes as it is.
    (tmp_path / "rel.py").write_text("def f():\n    return 'found by an SRef'\n", encoding="utf-8")
    host.set_folder(tmp_path)
    host.set_caller("[Book.xlsx]Sheet1", row=2, column=3, sref=True)
    assert host.call("PyCall", "rel.py", "f") == "found by an SRef"


def test_pycall_relative_path_with_brackets_in_the_folder(host, tmp_path):
    folder = tmp_path / "data[old]"
    folder.mkdir()
    (folder / "rel.py").write_text("def f():\n    return 'bracketed folder'\n", encoding="utf-8")
    host.set_folder(folder)
    assert host.call("PyCall", "rel.py", "f") == "bracketed folder"


def test_pycall_relative_path_when_excel_gives_the_workbook_without_its_sheet(host, tmp_path):
    # CELL("filename") gives C:\folder\Book.xlsx, with no [Book.xlsx]Sheet1, for a workbook
    # whose only sheet has the workbook's name; the folder may still hold brackets.
    folder = tmp_path / "data[old]"
    folder.mkdir()
    (folder / "rel.py").write_text("def f():\n    return 'found without a sheet'\n", encoding="utf-8")
    host.set_folder(folder, alone=True)
    host.set_caller("[Book.xlsx]Book", row=1, column=1)
    assert host.call("PyCall", "rel.py", "f") == "found without a sheet"


def test_pycall_relative_path_in_an_unsaved_workbook(host):
    host.set_folder(None)
    assert host.call("PyCall", "rel.py", "f") == Err(VALUE)
    assert "has not been saved" in wait_for(host, "has not been saved")


def test_pycall_file_not_text(host):
    assert host.call("PyCall", 1.0, "f") == Err(VALUE)


def test_pycall_reloads_a_changed_script(host, tmp_path):
    path = tmp_path / "changing.py"
    path.write_text("def f():\n    return 1\n", encoding="utf-8")
    assert host.call("PyCall", str(path), "f") == 1.0
    before = os.stat(path).st_mtime_ns
    path.write_text("def f():\n    return 2\n", encoding="utf-8")
    os.utime(path, ns=(before, before + 1000))
    assert host.call("PyCall", str(path), "f") == 2.0


# PyReloadOnChange


def test_reload_on_change_setting(host, tmp_path):
    assert host.call("PyReloadOnChange") is True
    assert host.call("PyReloadOnChange", False) is False
    assert host.call("PyReloadOnChange") is False
    path = tmp_path / "kept.py"
    path.write_text("def f():\n    return 'first'\n", encoding="utf-8")
    assert host.call("PyCall", str(path), "f") == "first"
    path.write_text("def f():\n    return 'second, longer'\n", encoding="utf-8")
    assert host.call("PyCall", str(path), "f") == "first"
    assert host.call("PyReloadOnChange", True) is True
    assert host.call("PyCall", str(path), "f") == "second, longer"


def test_reload_on_change_refuses_text(host):
    assert host.call("PyReloadOnChange", "yes") == Err(VALUE)
    assert host.call("PyReloadOnChange") is True


# PyConsole (the window itself: test_console.py)


def test_console_returns_show(host):
    assert host.call("PyConsole", True) is True
    assert host.call("PyConsole", False) is False
    assert host.call("PyConsole", "show") == Err(VALUE)


# PyInfo


def test_info_items(host):
    assert host.call("PyInfo", "status") == "running"
    assert host.call("PyInfo", "pyinex") == str(host.xll)
    assert os.path.normcase(host.call("PyInfo", "executable")) == os.path.normcase(sys.executable)
    python = host.call("PyInfo", "python")
    assert os.path.basename(python).lower() == f"python3{sys.version_info.minor}.dll" and os.path.isfile(python)
    assert host.call("PyInfo", "version") == sys.version
    assert host.call("PyInfo", "prefix") == sys.prefix
    assert host.call("PyInfo", "venv") is (sys.prefix != sys.base_prefix)
    cmake = (ROOT / "CMakeLists.txt").read_text("utf-8")
    assert host.call("PyInfo", "addin_version") == re.search(r"project\(pyinex VERSION (\S+)", cmake).group(1)


def test_info_names_ignore_case(host):
    assert host.call("PyInfo", "Python") == host.call("PyInfo", "python")
    assert host.call("PyInfo", "STATUS") == "running"


def test_info_table(host):
    table = host.call("PyInfo")
    names = ["status", "python", "pyinex", "version", "executable", "prefix", "venv", "addin_version"]
    assert [row[0] for row in table] == names
    assert all(row[1] == host.call("PyInfo", name) for row, name in zip(table, names))


def test_info_unknown_item(host):
    assert host.call("PyInfo", "colour") == Err(VALUE)
    assert host.call("PyInfo", 1.0) == Err(VALUE)


# The function wizard


def test_every_function_is_blank_in_the_function_wizard(xllhost, pyinex_xll, tmp_path):
    marker = tmp_path / "ran.txt"
    script = tmp_path / "marks.py"
    script.write_text(f"def f():\n    open({str(marker)!r}, 'w').close()\n    return 'ran'\n", encoding="utf-8")
    # A Python that cannot start: if anything in the wizard tried, its refusal would show.
    with Host(xllhost, pyinex_xll, env=host_env(PYINEX_PYTHON=str(tmp_path / "none" / "python.exe"))) as h:
        h.open()
        h.set_wizard(True)
        assert h.call("PyCall", str(script), "f") == ""
        assert h.call("PyInfo", "status") == ""
        assert h.call("PyInfo") == ""
        assert h.call("PyConsole", True) == ""
        assert h.call("PyReloadOnChange", False) == ""
        h.set_wizard(False)
        assert h.call("PyReloadOnChange") is True, "the wizard changed the setting"
        time.sleep(0.5)
        assert "Pyinex:" not in h.stderr, "Python was started in the wizard"
        assert "file not found" in h.call("PyInfo", "status")
    assert not marker.exists()


def test_functions_run_once_the_wizard_closes(host):
    host.set_wizard(True)
    assert host.call("PyInfo", "status") == ""
    host.set_wizard(False)
    assert host.call("PyInfo", "status") == "running"
