"""NumPy in real Excel: the 2010 demo (NumPyDemo), an array returned through .tolist(), and
a raw array, which Pyinex does not convert (docs/conversion.md, "Anything else").

The expected results are checked here with NumPy against the same input, so the tests check
what reaches the cells, not NumPy's arithmetic.
"""

import numpy as np
import pytest
from conftest import E, pycall, typed

SYMMETRIC = [[2.0, 1.0], [1.0, 2.0]]
INDEFINITE = [[1.0, 2.0], [2.0, 1.0]]  # Eigenvalues 3 and -1.


@pytest.fixture
def demo(book):
    book.copy("numpy_demo.py")
    return book


def put_block(book, top_left_row, rows):
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            book.sheet.Cells(top_left_row + r, 1 + c).Value2 = value


def close(got, expected):
    return all(abs(g - e) <= 1e-12 * max(1.0, abs(e)) for gr, er in zip(got, expected) for g, e in zip(gr, er))


def test_compute_eigenstructure_2010(demo):
    put_block(demo, 1, SYMMETRIC)
    demo.formula("D1", pycall("numpy_demo.py", "ComputeEigenstructure", "A1:B2"))
    got = demo.values("D1:E4")
    assert got[1] == ["", ""]  # The 2010 row of spacing: None becomes a blank cell.
    # Each column below it is a unit eigenvector for the value above it. (Checked so, since
    # another NumPy may give them in another order or sign.)
    a, vals, vecs = np.array(SYMMETRIC), np.array(got[0]), np.array(got[2:])
    for j in range(2):
        assert np.allclose(a @ vecs[:, j], vals[j] * vecs[:, j], atol=1e-12)
        assert np.isclose(np.linalg.norm(vecs[:, j]), 1.0)
    assert sorted(vals) == pytest.approx([1.0, 3.0])


def test_matrix_multiply_2010(demo):
    put_block(demo, 1, [[1.0, 2.0], [3.0, 4.0]])
    put_block(demo, 4, [[1.0], [1.0]])
    demo.formula("D1", pycall("numpy_demo.py", "MatrixMultiply", "A1:B2", "A4:A5"))
    assert typed(demo.values("D1:D2")) == typed([[3.0], [7.0]])


def test_hammer_into_positive_definite_2010(demo):
    put_block(demo, 1, INDEFINITE)
    demo.formula("D1", pycall("numpy_demo.py", "HammerIntoPositiveDefinite", "A1:B2"))
    vals, vecs = map(np.real_if_close, np.linalg.eig(np.array(INDEFINITE)))
    vals = np.array([max(v, 0.0001) for v in vals])
    expected = (vecs @ np.diag(vals) @ vecs.T).tolist()
    got = demo.values("D1:E2")
    assert close(got, expected)
    assert all(v > 0 for v in np.linalg.eigvalsh(np.array(got)))


def test_the_python_numpy_runs_in_is_excels(demo, expected):
    version = demo.formula("A1", pycall("numpy_demo.py", "PyVersion"))
    if expected is not None:
        assert version == expected["version"]


def test_a_raw_array_is_not_converted(demo, cleared):
    put_block(demo, 1, SYMMETRIC)
    assert demo.formula("D1", pycall("numpy_demo.py", "raw", "A1:B2")) == E("#VALUE!")
    said = "Pyinex: numpy_demo.py, raw: ndarray cannot be returned to Excel"
    window = cleared(said)
    assert window and said in window.text
