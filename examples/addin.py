"""Output to Pyinex's console.

addin.xlsx calls this with PyCall, beside its PyInfo, PyReloadOnChange and PyConsole
formulas.
"""


def hello(name):
    """print() writes to Pyinex's console; show it with PyConsole(TRUE)."""
    print(f"Hello from Python, {name}.")
    return "printed"
