"""The add-in releases the GIL between calls (docs/architecture.md, "Threading")."""

import time

from conftest import evaluate

TICKER = """
import threading, time
ticks = [0]
def tick():
    while True:
        ticks[0] += 1
        time.sleep(0.01)
threading.Thread(target=tick, daemon=True).start()
"""


def test_a_thread_keeps_running_between_calls(host):
    evaluate(host, "0", TICKER)
    first = int(evaluate(host, "ticks[0]"))
    time.sleep(1)
    # Ticking every 10 ms, a free-running thread manages dozens in a second. Were the GIL
    # held between calls it would manage none.
    assert int(evaluate(host, "ticks[0]")) - first > 20
