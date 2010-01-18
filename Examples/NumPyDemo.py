import string, sys

###############################################################################

def CheckVersion():
    # NumPy is only installed for Python 2.5 and 2.6
    version = str.split(str.split(sys.version)[0], '.')

    if not((map(int, version[0:2]) ==  [2, 5]) or
           (map(int, version[0:2]) ==  [2, 6])):
        print('Sorry - NumPy only runs on Python 2.5 and 2.6 right now.')
        return False
    return True

###############################################################################

if CheckVersion():
    import numpy as np

###############################################################################

def PyVersion():
    return sys.version

###############################################################################

def ComputeEigenstructure(xlA):

    a = np.array(xlA)

    # I had tests for squareness and symmetry in Python, but
    # NumPy has them internally (and probably faster), so I dropped 
    # them

    (vals, vecs) = np.linalg.eig(a)
    # Put a row of spacing in Excel between vals and vecs
    pad = [None for number in range(len(vals))]
    retval = np.vstack((vals, pad, vecs))
    return retval.tolist()
    
###############################################################################

def MatrixMultiply(xlMatrix, xlVec):

    a = np.array(xlMatrix)
    x = np.array(xlVec)
    retval = np.dot(a,x)
    return retval.tolist()

###############################################################################
#
# Cheesy hack - doesn't guarantee 1's on the diagonal. Fix later...
#
def HammerIntoPositiveDefinite(xlA):

    a = np.array(xlA)

    (vals, vecs) = np.linalg.eig(a)

    # Make all negative eigenvalues slightly positive
    vals = np.array([max(number,0.0001) for number in vals])

    retval = np.dot( np.dot( vecs, np.diag(vals)), vecs.transpose() )
    return retval.tolist()

###############################################################################
    
