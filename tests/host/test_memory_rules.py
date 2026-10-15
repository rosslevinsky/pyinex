"""The fake Excel enforces Excel's memory rules: pyinex.xll passes, and each deliberate
break in badxll.xll is reported."""

import pytest
from conftest import HELPERS
from hostdriver import HostError, HostFault, Int


def test_pyinex_passes(host):
    # Every call checks for faults and raises HostFault on any, so these pass only if
    # registration and the calls broke no rule.
    host.call("PyInfo")
    host.call("PyCall", str(HELPERS), "echo", ((1.0, "a"), ("b", None)))
    assert host.close() == 1.0


def test_result_without_dll_free_is_reported(bad_host):
    _, dll_free, faults = bad_host.call_detail("BadNoDllFree", expect_faults=True)
    assert not dll_free
    assert any("without xlbitDLLFree" in f for f in faults)


def test_double_free_is_reported(bad_host):
    _, _, faults = bad_host.call_detail("BadDoubleFree", expect_faults=True)
    assert any("already freed" in f for f in faults)


def test_leak_is_reported(bad_host):
    _, _, faults = bad_host.call_detail("BadLeak", expect_faults=True)
    assert any("never freed with xlFree" in f for f in faults)


def test_unfreed_text_in_a_result_is_reported(bad_host):
    result, dll_free, faults = bad_host.call_detail("BadNestedLeak", expect_faults=True)
    assert (result, dll_free) == ("hi", True)
    assert any("did not free the result's text" in f for f in faults)
    assert any("xlAutoFree12 released 1 heap blocks; the result had 2" in f for f in faults)


def test_a_free_that_balances_the_count_but_leaks_is_reported(bad_host):
    # Frees an unrelated block in place of the result's text: the count matches, the text leaks.
    _, _, faults = bad_host.call_detail("BadSwapFree", expect_faults=True)
    assert faults == ["xlAutoFree12 did not free the result's text"]


def test_a_function_reading_an_unregistered_argument_crashes(bad_host):
    with pytest.raises(HostError, match="stopped"):
        bad_host.call("BadArity", 1.0)


def test_changing_an_argument_is_reported(bad_host):
    _, _, faults = bad_host.call_detail("BadChangesArgument", "text", expect_faults=True)
    assert faults == ["the function changed or freed argument 1"]


def test_32_bit_integers_pass_through(host):
    assert host.call("PyCall", str(HELPERS), "echo", Int(65536)) == 65536.0


def test_faults_raise_by_default(bad_host):
    with pytest.raises(HostFault):
        bad_host.call("BadLeak")
