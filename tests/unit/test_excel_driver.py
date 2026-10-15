"""The real-Excel driver's record and restore of the user's settings (scripts/excel_tests.py).

The driver puts back the registry keys a real-Excel run may change, and a run whose keys it
cannot put back fails. These tests use a scratch key under HKEY_CURRENT_USER\\Software\\
PyinexTests, removed afterwards, and need no Excel.
"""

import importlib.util
import os
import sys
from pathlib import Path

import pytest

if sys.platform != "win32":
    pytest.skip("the driver runs only on Windows", allow_module_level=True)

import winreg  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
SCRATCH = rf"Software\PyinexTests\excel-driver-{os.getpid()}"


@pytest.fixture(scope="module")
def driver():
    spec = importlib.util.spec_from_file_location("excel_tests", ROOT / "scripts" / "excel_tests.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def delete_tree(path):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_ALL_ACCESS) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                delete_tree(rf"{path}\{child}")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
    except FileNotFoundError:
        pass


@pytest.fixture
def scratch():
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, SCRATCH, 0, winreg.KEY_ALL_ACCESS) as key:
        winreg.SetValueEx(key, "Mode", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "OPEN", 0, winreg.REG_SZ, "/R addin.xll")
    yield SCRATCH
    delete_tree(SCRATCH)
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, r"Software\PyinexTests")
    except OSError:
        pass  # Another test's keys are still there.


def values(path):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
        return {winreg.EnumValue(key, i)[0]: winreg.EnumValue(key, i)[1] for i in range(winreg.QueryInfoKey(key)[1])}


def test_restore_puts_back_changed_added_and_removed_values(driver, scratch):
    recorded = driver.record([scratch])
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, scratch, 0, winreg.KEY_ALL_ACCESS) as key:
        winreg.SetValueEx(key, "Mode", 0, winreg.REG_DWORD, 0)
        winreg.SetValueEx(key, "OPEN1", 0, winreg.REG_SZ, "/R other.xll")
        winreg.DeleteValue(key, "OPEN")
    assert driver.restore(recorded) == []
    assert values(scratch) == {"Mode": 1, "OPEN": "/R addin.xll"}


def test_restore_removes_a_key_that_did_not_exist(driver, scratch):
    path = rf"{scratch}\New"
    recorded = driver.record([path])
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_ALL_ACCESS) as key:
        winreg.SetValueEx(key, "MultiThreaded", 0, winreg.REG_DWORD, 1)
    assert driver.restore(recorded) == []
    assert driver.read_key(path) is None


def test_restore_reports_a_key_it_cannot_put_back(driver, scratch, monkeypatch):
    recorded = driver.record([scratch])
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, scratch, 0, winreg.KEY_ALL_ACCESS) as key:
        winreg.SetValueEx(key, "Mode", 0, winreg.REG_DWORD, 0)

    def refuse(*args):
        raise PermissionError(5, "Access is denied")

    monkeypatch.setattr(driver.winreg, "SetValueEx", refuse)
    problems = driver.restore(recorded)
    assert len(problems) == 1 and scratch in problems[0]


def test_an_unchanged_key_is_left_alone(driver, scratch, monkeypatch):
    recorded = driver.record([scratch])
    monkeypatch.setattr(driver.winreg, "SetValueEx", lambda *a: pytest.fail("wrote an unchanged key"))
    assert driver.restore(recorded) == []
