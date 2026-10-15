"""A check of the driver: a PyCall that never returns must end in its timeout, with the
Excel it hangs ended and the run failed. Run it with a short --timeout."""


def test_hangs(book):
    book.script("hang.py", "import time\n\ndef hang():\n    while True:\n        time.sleep(1)\n")
    book.formula("A1", '=PyCall("hang.py", "hang")')
