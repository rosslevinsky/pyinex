"""Where the add-in's output goes, and PyConsole (docs/architecture.md, "Console" and
"PyConsole").

The console is a window of Pyinex's own. Windows the fake Excel makes are never visible in
a session with no desktop, such as SSH or CI, so showing and hiding are checked through
placement, which applies only as the window goes from hidden to shown.
"""

import ctypes
import time

import pytest
from conftest import HELPERS, host_env
from hostdriver import VALUE, Err, Host

TITLE = "Pyinex console"


def wait_until(condition, timeout=10):
    deadline = time.monotonic() + timeout
    while not condition() and time.monotonic() < deadline:
        time.sleep(0.05)
    return condition()


@pytest.fixture
def chatty(tmp_path):
    path = tmp_path / "chatty.py"
    path.write_text(
        "import threading, time\n\n"
        "def talk():\n    print('hello from print')\n    return 1\n\n"
        "def fail():\n    raise ValueError('from the script')\n\n"
        "def unicode():\n    print('h\u00e9llo \u03ba\u03cc\u03c3\u03bc\u03b5 \U0001f600')\n    print('\u00e9' * 5000)\n    print('before\\x00after')\n    return 1\n\n"
        "def later():\n"
        "    threading.Thread(target=lambda: (time.sleep(0.2), print('from a thread'))).start()\n"
        "    return 1\n",
        encoding="utf-8",
    )
    return path


@pytest.fixture
def desktop():
    """The desktop's left, top, width and height, spanning every monitor."""
    metric = ctypes.windll.user32.GetSystemMetrics
    return metric(76), metric(77), metric(78), metric(79)  # SM_XVIRTUALSCREEN and so on.


@pytest.fixture
def no_output_host(xllhost, pyinex_xll):
    """A fake Excel started as Excel is, with no standard output or error."""
    with Host(xllhost, pyinex_xll, env=host_env(), output="none") as h:
        h.open()
        yield h


def console_text(host):
    console = host.console()
    return console.text if console else ""


# Where output goes


def test_print_and_tracebacks_reach_the_hosts_output(host, chatty):
    assert host.call("PyCall", str(chatty), "talk") == 1.0
    assert wait_until(lambda: "hello from print" in host.stdout)
    assert host.call("PyCall", str(chatty), "fail") == Err(VALUE)
    assert wait_until(lambda: "ValueError: from the script" in host.stderr)
    assert "Traceback" in host.stderr


def test_a_host_with_output_keeps_it_after_pyconsole_makes_a_window(host, chatty):
    assert host.call("PyConsole", True) is True
    assert host.call("PyCall", str(chatty), "talk") == 1.0
    assert wait_until(lambda: "hello from print" in host.stdout)
    assert "hello" not in host.console().text


def test_with_no_output_python_writes_to_the_console(no_output_host, chatty):
    h = no_output_host
    assert h.call("PyCall", str(chatty), "talk") == 1.0
    assert h.call("PyCall", str(chatty), "fail") == Err(VALUE)
    assert wait_until(lambda: "Pyinex: fail in the script" in console_text(h))
    console = h.console()
    assert console.title == TITLE
    assert "hello from print" in console.text
    assert "ValueError: from the script" in console.text
    # C code's output, through the C runtime the add-in and Python share.
    h.call("PyCall", str(HELPERS), "evaluate", "0", "import ctypes\nctypes.cdll.ucrtbase.puts(b'from puts')\nctypes.cdll.ucrtbase.fflush(None)")
    assert wait_until(lambda: "from puts" in console_text(h))


def test_the_console_shows_any_character(no_output_host, chatty):
    # 5,000 two-byte characters span several reads of the pipe, so some arrive split.
    assert no_output_host.call("PyCall", str(chatty), "unicode") == 1.0
    assert wait_until(lambda: "\u00e9" * 5000 in console_text(no_output_host))
    assert "h\u00e9llo \u03ba\u03cc\u03c3\u03bc\u03b5 \U0001f600\n" in console_text(no_output_host)
    # A NUL would end the text box's text there, so it shows as a replacement character.
    assert wait_until(lambda: "before\ufffdafter" in console_text(no_output_host))


def test_print_from_a_scripts_thread_reaches_the_console_between_calls(no_output_host, chatty):
    assert no_output_host.call("PyCall", str(chatty), "later") == 1.0
    assert wait_until(lambda: "from a thread" in console_text(no_output_host))


def test_with_no_output_pyconsoles_complaints_reach_the_console(no_output_host):
    assert no_output_host.call("PyConsole", True, False, 10) == Err(VALUE)
    assert wait_until(lambda: "Pyinex: PyConsole: give x and y together" in console_text(no_output_host))


# PyConsole


def test_pyconsole_makes_the_window(host):
    assert host.console() is None
    assert host.call("PyConsole", False) is False
    assert host.console().title == TITLE


def colors(mode):
    """The console's text and background colors in Windows' light or dark app mode."""
    if mode == "dark":
        return 0xDCDCDC, 0x1E1E1E
    system = ctypes.windll.user32.GetSysColor  # Windows' own, for a read-only text box.
    return system(8), system(15)  # COLOR_WINDOWTEXT, COLOR_3DFACE


@pytest.mark.parametrize("mode", ["light", "dark"])
def test_the_console_follows_windows_light_or_dark_mode(host, mode):
    host.set_theme(mode)
    host.call("PyConsole", False)
    c = host.console()
    assert (c.text_color, c.background) == colors(mode)


def test_the_console_follows_a_change_of_mode_while_open(host):
    host.set_theme("light")
    host.call("PyConsole", True)
    host.set_theme("dark")
    c = host.console()
    assert (c.text_color, c.background) == colors("dark")
    host.set_theme("light")
    c = host.console()
    assert (c.text_color, c.background) == colors("light")


def test_closing_the_console_hides_it(host, desktop):
    left, top, _, _ = desktop
    assert host.call("PyConsole", True, False, left + 10, top + 20, 400, 300) is True
    host.close_console()
    assert host.console() is not None  # Hidden, not gone, and Excel carries on.
    # Hidden, so showing it again places it.
    assert host.call("PyConsole", True, False, left + 50, top + 60, 500, 400) is True
    c = host.console()
    assert (c.left, c.top, c.width, c.height) == (left + 50, top + 60, 500, 400)


def test_pyconsole_does_not_start_python(host):
    assert host.call("PyConsole", True) is True
    host.close()
    # The add-in stays loaded to remember the console; the core would be loaded too had
    # Python started.
    assert host.unload() == (True, False)


def test_the_console_stays_pyinexs_across_unloading(host, desktop):
    left, top, _, _ = desktop
    assert host.call("PyConsole", False) is False
    host.close()
    host.unload()
    host.open()
    # Still managed: placement applies as it is shown.
    assert host.call("PyConsole", True, False, left + 30, top + 40, 400, 300) is True
    c = host.console()
    assert (c.left, c.top, c.width, c.height) == (left + 30, top + 40, 400, 300)


def test_pyconsole_places_the_window_as_it_is_shown(host, desktop):
    left, top, _, _ = desktop
    assert host.call("PyConsole", True, False, left + 10, top + 20, 400, 300) is True
    first = host.console()
    assert (first.left, first.top, first.width, first.height) == (left + 10, top + 20, 400, 300)
    # While shown, the window is left where the user put it.
    assert host.call("PyConsole", True, False, left + 50, top + 60, 500, 400) is True
    assert host.console() == first
    # Leaving the position out leaves the window where it was.
    host.call("PyConsole", False)
    host.call("PyConsole", True)
    assert host.console() == first
    host.call("PyConsole", False)
    host.call("PyConsole", True, False, left + 50, top + 60, 500, 400)
    moved = host.console()
    assert (moved.left, moved.top, moved.width, moved.height) == (left + 50, top + 60, 500, 400)


def test_pyconsole_rounds_halves_away_from_zero(host, desktop):
    left, top, _, _ = desktop
    host.call("PyConsole", True, False, left + 10.5, top + 20.4, 400.5, 300.5)
    c = host.console()
    assert (c.left, c.top, c.width, c.height) == (left + 11, top + 20, 401, 301)
    host.call("PyConsole", False)
    # Just left of the desktop: -0.4 rounds onto it, -0.5 off it.
    assert host.call("PyConsole", True, False, left - 0.4, top, 400, 300) is True
    host.call("PyConsole", False)
    assert host.call("PyConsole", True, False, left - 0.5, top, 400, 300) == Err(VALUE)
    # A width of 0.4 rounds to 0, and 0.5 to 1.
    assert host.call("PyConsole", True, False, None, None, 0.4, 300) == Err(VALUE)
    assert host.call("PyConsole", True, False, None, None, 0.5, 300) is True


@pytest.mark.parametrize(
    "args",
    [
        ("show",),
        (1.0,),
        (True, "yes"),
        (True, False, 10),  # x without y
        (True, False, None, 10),  # y without x
        (True, False, None, None, 400),  # width without height
        (True, False, None, None, None, 300),  # height without width
        (True, False, None, None, 0, 300),
        (True, False, None, None, 400, -1),
        (True, False, "10", 20),
        (True, False, 1e300, 20),
    ],
)
def test_pyconsole_refuses_bad_arguments(host, args):
    assert host.call("PyConsole", *args) == Err(VALUE)
    assert wait_until(lambda: "Pyinex: PyConsole" in host.stderr)


def test_pyconsole_refuses_a_place_off_the_desktop(host, desktop):
    left, top, width, height = desktop
    for x, y in [(left + width, top), (left, top + height), (left - 1, top), (left, top - 1)]:
        assert host.call("PyConsole", True, False, x, y, 400, 300) == Err(VALUE), (x, y)
    assert host.call("PyConsole", True, False, left + width - 1, top + height - 1, 400, 300) is True


def test_pyconsole_clears_the_text(no_output_host):
    h = no_output_host
    h.call("PyConsole", False)
    h.write_output("some earlier output")
    assert wait_until(lambda: "some earlier output" in console_text(h))
    assert h.call("PyConsole", False, False) is False
    assert "some earlier output" in console_text(h)
    assert h.call("PyConsole", False, True) is False
    assert console_text(h) == ""
