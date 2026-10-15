"""A check of the driver: a skipped test must fail its run."""

import pytest


def test_skips(app):
    pytest.skip("deliberate skip")
