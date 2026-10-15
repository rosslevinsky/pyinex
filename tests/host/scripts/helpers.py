"""Functions the fake-Excel tests call through PyCall."""

namespace = {}  # evaluate's, kept between calls


def evaluate(expression, setup=None):
    """Runs setup, then returns str() of the expression."""
    if setup is not None:
        exec(setup, namespace)
    return str(eval(expression, namespace))


def echo(value=None):
    """Returns its argument, so a value goes into Python and back out."""
    return value


def attempt(expression, setup=None):
    """Like evaluate, but an exception comes back as text, "Type: message"."""
    try:
        return evaluate(expression, setup)
    except BaseException as e:
        return f"{type(e).__name__}: {e}"
