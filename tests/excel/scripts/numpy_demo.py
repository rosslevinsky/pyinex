"""The 2010 NumPy demo (0.83's Examples/NumPyDemo.py) without its Python 2 version check,
for tests/excel/test_numpy.py; and a function returning an array itself.

One change to 2010's code: np.linalg.eig gives complex numbers in NumPy 2.5, even for a
real symmetric matrix, and Pyinex returns no complex number, so the eigenvalues and
vectors are made real (np.real_if_close) where their imaginary parts are zero.

Not named numpy.py: a script's folder goes on sys.path, where that name would hide NumPy.
"""

import sys

import numpy as np


def PyVersion():
    return sys.version


def ComputeEigenstructure(xlA):
    a = np.array(xlA)
    (vals, vecs) = map(np.real_if_close, np.linalg.eig(a))
    # Put a row of spacing in Excel between vals and vecs
    pad = [None for number in range(len(vals))]
    retval = np.vstack((vals, pad, vecs))
    return retval.tolist()


def MatrixMultiply(xlMatrix, xlVec):
    a = np.array(xlMatrix)
    x = np.array(xlVec)
    retval = np.dot(a, x)
    return retval.tolist()


def HammerIntoPositiveDefinite(xlA):
    a = np.array(xlA)
    (vals, vecs) = map(np.real_if_close, np.linalg.eig(a))
    # Make all negative eigenvalues slightly positive
    vals = np.array([max(number, 0.0001) for number in vals])
    retval = np.dot(np.dot(vecs, np.diag(vals)), vecs.transpose())
    return retval.tolist()


def raw(xlA):
    """The array itself, which Pyinex does not convert (docs/conversion.md)."""
    return np.array(xlA)
