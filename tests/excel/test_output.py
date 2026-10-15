"""Output reaching Pyinex's console in real Excel (docs/architecture.md, "Console"): print(),
tracebacks, a script's own thread, C code, and any character.

Excel starts with no standard output, so everything Python writes goes through Pyinex's
pipe to the window, whether it is shown or not.
"""

from conftest import E, pycall

UNICODE = "Bücher ✓ 😀"


def test_print_reaches_the_console(book, cleared):
    book.script("talks.py", "def f():\n    print('hello from print')\n    return 1\n")
    assert book.formula("A1", pycall("talks.py", "f")) == 1.0
    window = cleared("hello from print")
    assert window and "hello from print" in window.text


def test_a_traceback_reaches_the_console(book, cleared):
    book.script("fails.py", "def inner():\n    raise ValueError('from the script')\n\ndef f():\n    inner()\n")
    assert book.formula("A1", pycall("fails.py", "f")) == E("#VALUE!")
    window = cleared("ValueError: from the script")
    text = window.text if window else ""
    assert "Traceback (most recent call last)" in text and "in inner" in text and "ValueError: from the script" in text


def test_a_scripts_thread_prints_between_calls(book, cleared):
    # The thread prints only once the test has seen the call return, so its line arrives
    # with no PyCall running, whose end would flush it.
    go = book.folder / "go"
    book.script(
        "later.py",
        "import os, threading, time\n\n"
        f"GO = {str(go)!r}\n\n"
        "def wait_then_print():\n"
        "    while not os.path.exists(GO):\n        time.sleep(0.05)\n"
        "    print('from a thread')\n\n"
        "def f():\n    threading.Thread(target=wait_then_print, daemon=True).start()\n    return 1\n",
    )
    assert book.formula("A1", pycall("later.py", "f")) == 1.0
    go.write_text("go", encoding="utf-8")
    window = cleared("from a thread")
    assert window and "from a thread" in window.text


def test_c_codes_output_reaches_the_console(book, cleared):
    # puts in the C runtime that Python and C extensions share.
    book.script("puts.py", "import ctypes\n\ndef f():\n    ctypes.cdll.ucrtbase.puts(b'from puts')\n    ctypes.cdll.ucrtbase.fflush(None)\n    return 1\n")
    assert book.formula("A1", pycall("puts.py", "f")) == 1.0
    window = cleared("from puts")
    assert window and "from puts" in window.text


def test_any_character_reaches_the_console(book, cleared):
    # 5,000 two-byte characters span several reads of the pipe, so some arrive split.
    book.script(
        "unicode.py",
        "def f():\n    print('h\\u00e9llo \\u03ba\\u03cc\\u03c3\\u03bc\\u03b5 \\U0001f600')\n"
        "    print('\\u00e9' * 5000)\n    print('before\\x00after')\n    return 1\n",
    )
    assert book.formula("A1", pycall("unicode.py", "f")) == 1.0
    window = cleared("before\ufffdafter")  # A NUL would end the text box's text there.
    text = window.text if window else ""
    assert "h\u00e9llo \u03ba\u03cc\u03c3\u03bc\u03b5 \U0001f600\r\n" in text
    assert "\u00e9" * 5000 in text
    assert "before\ufffdafter" in text


def test_a_unicode_traceback_reaches_the_console(workbook, cleared, tmp_path):
    folder = tmp_path / UNICODE
    book = workbook(folder / "book.xlsx")
    book.script(f"Skript {UNICODE}.py", "def f():\n    raise ValueError('Fehler ✓ 😀')\n")
    assert book.formula("A1", pycall(f"Skript {UNICODE}.py", "f")) == E("#VALUE!")
    window = cleared("ValueError: Fehler ✓ 😀")
    text = window.text if window else ""
    assert "ValueError: Fehler ✓ 😀" in text
    assert str(folder / f"Skript {UNICODE}.py").lower() in text.lower()
