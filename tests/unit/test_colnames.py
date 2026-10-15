"""Column numbers to letters and back, for every column Excel has (A to XFD).

The reference list is built the slow, obvious way, as 2010's ExcelColnameValidation.py
built it: every one-, two- and three-letter name in order. The add-in's arithmetic must
agree with it at every column.
"""

import itertools
import string

import pytest

COLUMNS = 16384


def reference_names():
    names = []
    for length in (1, 2, 3):
        for letters in itertools.product(string.ascii_uppercase, repeat=length):
            names.append("".join(letters))
    return names[:COLUMNS]


def test_every_column_to_letters(testing):
    assert [testing.column_letters(n) for n in range(1, COLUMNS + 1)] == reference_names()


def test_every_column_to_number(testing):
    assert [testing.column_number(name) for name in reference_names()] == list(range(1, COLUMNS + 1))


# The boundaries 2010's validation script was written for: where a letter stops meaning
# "zero" and a new digit starts.
@pytest.mark.parametrize(
    "number, letters",
    [
        (1, "A"),
        (26, "Z"),
        (27, "AA"),
        (52, "AZ"),
        (53, "BA"),
        (676, "YZ"),
        (702, "ZZ"),
        (703, "AAA"),
        (704, "AAB"),
        (728, "AAZ"),
        (729, "ABA"),
        (1378, "AZZ"),
        (1379, "BAA"),
        (16384, "XFD"),
    ],
)
def test_boundaries(testing, number, letters):
    assert testing.column_letters(number) == letters
    assert testing.column_number(letters) == number


@pytest.mark.parametrize("number", [0, -1, 16385, -(2**31), 2**31 - 1])
def test_number_outside_the_grid(testing, number):
    with pytest.raises(ValueError):
        testing.column_letters(number)


@pytest.mark.parametrize("letters", ["", "XFE", "ZZZ", "AAAA", "a", "xfd", "1", "A1", "Ä", " A"])
def test_letters_not_a_column(testing, letters):
    with pytest.raises(ValueError):
        testing.column_number(letters)
