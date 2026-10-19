"""Pyinex's console window in real Excel (docs/architecture.md, "PyConsole" and "Console"):
showing, hiding, clearing, placing and closing it, each refusal, and its colors in light
and dark mode.

The console is one window for the whole Excel, so each test leaves it hidden and cleared.
"""

import ctypes
import math
import time

import pytest
from conftest import E, foreground, pycall, shown

TITLE = "Pyinex console"
BLOCKS = "█" * 40  # Solid glyphs: pixels of exactly the text's color.


@pytest.fixture
def pyconsole(app, console, mark):
    """A function calling PyConsole with formula-text arguments, returning its result; the
    console is hidden and cleared after the test."""

    def call(*args):
        return shown(app.Evaluate("PyConsole(" + ",".join(args) + ")"))

    mark()  # The console exists, with earlier output in it.
    try:
        yield call
    finally:
        call("FALSE", "TRUE")


def placed(window):
    return window.rect


def rounded(n):
    """Rounding to a whole pixel as PyConsole does: halves away from zero."""
    return int(math.floor(abs(n) + 0.5)) * (1 if n >= 0 else -1)


def test_show_and_hide(pyconsole, console):
    assert pyconsole("TRUE") is True
    window = console()
    assert window.title == TITLE and window.shown
    assert pyconsole("FALSE") is False
    assert not console().shown


def test_placing_applies_as_it_is_shown(pyconsole, console, desktop):
    left, top, _, _ = desktop
    pyconsole("FALSE")
    assert pyconsole("TRUE", "FALSE", str(left + 10), str(top + 20), "400", "300") is True
    assert placed(console()) == (left + 10, top + 20, 400, 300)
    # While shown, the window stays where it is.
    assert pyconsole("TRUE", "FALSE", str(left + 50), str(top + 60), "500", "400") is True
    assert placed(console()) == (left + 10, top + 20, 400, 300)
    # Leaving the place out leaves it where it was.
    pyconsole("FALSE")
    pyconsole("TRUE")
    assert placed(console()) == (left + 10, top + 20, 400, 300)
    pyconsole("FALSE")
    pyconsole("TRUE", "FALSE", str(left + 50), str(top + 60), "500", "400")
    assert placed(console()) == (left + 50, top + 60, 500, 400)


def test_placing_rounds_halves_away_from_zero(pyconsole, console, desktop):
    left, top, _, _ = desktop
    pyconsole("FALSE")
    pyconsole("TRUE", "FALSE", str(left + 10.5), str(top + 20.4), "400.5", "300.5")
    assert placed(console()) == (rounded(left + 10.5), rounded(top + 20.4), 401, 301)
    pyconsole("FALSE")
    assert pyconsole("TRUE", "FALSE", str(left - 0.4), str(top), "400", "300") is True
    pyconsole("FALSE")
    assert pyconsole("TRUE", "FALSE", str(left - 0.5), str(top), "400", "300") == E("#VALUE!")
    assert pyconsole("TRUE", "FALSE", "", "", "0.4", "300") == E("#VALUE!")
    assert pyconsole("TRUE", "FALSE", "", "", "0.5", "300") is True


@pytest.mark.parametrize(
    "args, said",
    [
        (('"show"',), "PyConsole: show must be TRUE or FALSE"),
        (("1",), "PyConsole: show must be TRUE or FALSE"),
        (("TRUE", '"yes"'), "PyConsole: clear must be TRUE or FALSE"),
        (("TRUE", "FALSE", "10"), "PyConsole: give x and y together, or neither"),
        (("TRUE", "FALSE", "", "10"), "PyConsole: give x and y together, or neither"),
        (("TRUE", "FALSE", "", "", "400"), "PyConsole: give width and height together, or neither"),
        (("TRUE", "FALSE", "", "", "", "300"), "PyConsole: give width and height together, or neither"),
        (("TRUE", "FALSE", "", "", "0", "300"), "PyConsole: width and height must be more than 0 pixels"),
        (("TRUE", "FALSE", "", "", "400", "-1"), "PyConsole: width and height must be more than 0 pixels"),
        (("TRUE", "FALSE", '"10"', "20"), "PyConsole: x must be a number"),
        (("TRUE", "FALSE", "1E+300", "20"), "PyConsole: x is far too large"),
    ],
)
def test_a_bad_argument_is_refused(pyconsole, cleared, args, said):
    assert pyconsole(*args) == E("#VALUE!")
    window = cleared("Pyinex: " + said)
    assert window and "Pyinex: " + said in window.text, window.text if window else None


def test_a_place_off_the_desktop_is_refused(pyconsole, console, mark, desktop):
    left, top, width, height = desktop
    for x, y in [(left + width, top), (left, top + height), (left - 1, top), (left, top - 1)]:
        pyconsole("FALSE", "TRUE")
        assert pyconsole("TRUE", "FALSE", str(x), str(y), "400", "300") == E("#VALUE!"), (x, y)
        said = f"Pyinex: PyConsole: x and y must lie on a monitor; {x},{y} is on none"
        assert said in mark(), (x, y)
    pyconsole("FALSE")
    # The main monitor's last pixel: the desktop's own corner may lie on no monitor, where
    # monitors of different sizes leave it bare.
    metric = ctypes.windll.user32.GetSystemMetrics
    assert pyconsole("TRUE", "FALSE", str(metric(0) - 1), str(metric(1) - 1), "400", "300") is True


def test_closing_hides_it_and_showing_places_it_again(pyconsole, console, desktop):
    left, top, _, _ = desktop
    pyconsole("TRUE", "FALSE", str(left + 10), str(top + 20), "400", "300")
    console().close()
    window = console()
    assert window is not None and not window.shown  # Hidden, not gone, and Excel carries on.
    pyconsole("TRUE", "FALSE", str(left + 50), str(top + 60), "500", "400")
    assert console().shown and placed(console()) == (left + 50, top + 60, 500, 400)


def test_showing_leaves_the_keyboard_with_excel(pyconsole, console, excel_in_front):
    pyconsole("FALSE")
    assert pyconsole("TRUE") is True
    assert console().shown
    assert foreground() == excel_in_front


def test_clearing_while_shown(pyconsole, console, mark):
    pyconsole("TRUE")
    mark()
    assert "marker" in console().text
    assert pyconsole("TRUE", "TRUE") is True
    assert console().text == "" and console().shown


def test_clearing_empties_it(pyconsole, console, mark):
    mark()
    assert pyconsole("FALSE", "FALSE") is False
    assert "marker" in console().text
    assert pyconsole("FALSE", "TRUE") is False
    assert console().text == ""


def test_output_is_kept_across_hiding_and_showing(pyconsole, console, mark):
    text = mark()
    pyconsole("TRUE")
    pyconsole("FALSE")
    pyconsole("TRUE")
    assert console().text == text


# Light and dark.


def light_colors():
    system = ctypes.windll.user32.GetSysColor  # Windows' own, for a read-only text box.
    return system(8), system(15)  # COLOR_WINDOWTEXT, COLOR_3DFACE


COLORS = {"dark": (0xDCDCDC, 0x1E1E1E)}


def shows_colors(console, book, colors):
    """The console, shown and holding solid glyphs, draws its text and background in colors."""
    book.script("blocks.py", f"def f():\n    print({BLOCKS!r})\n    return 1\n")
    assert book.formula("A1", pycall("blocks.py", "f")) == 1.0
    window = console(holding=BLOCKS)
    assert window and window.shown
    text, background = colors
    # Right after a change of mode the window may still be repainting, and PrintWindow then
    # draws nothing (all black): it gets three seconds to show its new colors.
    deadline = time.monotonic() + 3
    while True:
        counts = window.colors()
        most = max(counts, key=counts.get)
        if most == background or time.monotonic() > deadline:
            break
        time.sleep(0.1)
    assert most == background, f"most pixels are {most:06X}, not {background:06X}"
    assert counts.get(text, 0) > 100, f"few pixels of {text:06X}: {sorted(counts.items(), key=lambda c: -c[1])[:5]}"


@pytest.mark.parametrize("mode", ["light", "dark"])
def test_its_colors_follow_the_app_mode(pyconsole, console, app, book, app_mode, mode):
    app_mode(mode)
    pyconsole("TRUE")
    shows_colors(console, book, COLORS.get(mode) or light_colors())


def test_its_colors_follow_a_change_of_mode_while_open(pyconsole, console, app, book, app_mode):
    app_mode("light")
    pyconsole("TRUE")
    shows_colors(console, book, light_colors())
    app_mode("dark")
    shows_colors(console, book, COLORS["dark"])
    app_mode("light")
    shows_colors(console, book, light_colors())
