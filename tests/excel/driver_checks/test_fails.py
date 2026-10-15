"""A check of the driver: a failing test must fail its run."""


def test_fails(app):
    assert False, "deliberate failure"
