"""Exceptions of every kind, while a script imports, reloads and runs, in one fake Excel
process: each gives #VALUE! and a diagnostic, and the next call works (docs/architecture.md,
"Scripts")."""

import os
import time

from conftest import HELPERS
from hostdriver import VALUE, Err

RAISES = {
    "exception": ("raise ValueError('deliberate')", "ValueError: deliberate"),
    "SystemExit": ("raise SystemExit(3)", "SystemExit: 3"),
    "KeyboardInterrupt": ("raise KeyboardInterrupt", "KeyboardInterrupt"),
}


def wait_for(host, text, timeout=10):
    deadline = time.monotonic() + timeout
    while text not in host.stderr and time.monotonic() < deadline:
        time.sleep(0.05)
    return host.stderr


def still_works(host):
    assert host.call("PyCall", str(HELPERS), "echo", "alive") == "alive"


def test_every_kind_in_every_place_in_one_process(host, tmp_path):
    for kind, (statement, shown) in RAISES.items():
        # While importing.
        path = tmp_path / f"imports_{kind}.py"
        path.write_text(f"{statement}\n", encoding="utf-8")
        assert host.call("PyCall", str(path), "f") == Err(VALUE), kind
        assert "failed to load" in wait_for(host, shown)
        still_works(host)

        # While reloading: a good version first, then a changed one that raises.
        path = tmp_path / f"reloads_{kind}.py"
        path.write_text("def f():\n    return 'good'\n", encoding="utf-8")
        assert host.call("PyCall", str(path), "f") == "good"
        before = os.stat(path).st_mtime_ns
        path.write_text(f"def f():\n    return 'good'\n{statement}\n", encoding="utf-8")
        os.utime(path, ns=(before, before + 10_000_000))
        assert host.call("PyCall", str(path), "f") == Err(VALUE), kind
        still_works(host)

        # During a call.
        path = tmp_path / f"calls_{kind}.py"
        path.write_text(f"def f():\n    {statement}\n", encoding="utf-8")
        assert host.call("PyCall", str(path), "f") == Err(VALUE), kind
        still_works(host)

    err = host.stderr
    for kind, (statement, shown) in RAISES.items():
        assert err.count(shown) >= 3, (kind, err)
    assert err.count("Traceback (most recent call last)") >= 9
