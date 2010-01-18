# Note that no "import pyinex" statement is required (though you may
# add one if you like); the pyinex module name is added by the XLL.

import sys


###############################################################################
#
# When developing Python functions that use the pyinex namespace, one can't
# debug the code outside of Excel unless there's a stub implementation of
# some callable named 'pyinex'. A simple solution is to conditionally create 
# a dummy 'pyinex' class.

try:
    pyinex.CallerA1()
except NameError:
    class pyinex:
        @staticmethod
        def CallerA1():
            return 'Dummy CallerA1()'
        @staticmethod
        def CallerA1Full():
            return 'Dummy CallerA1Full()'
        @staticmethod
        def CallerR1C1():
            return 'Dummy CallerR1C1()'
        @staticmethod
        def CallerR1C1Full():
            return 'Dummy CallerR1C1Full()'
        @staticmethod
        def CallerSheet():
            return 'Dummy CallerSheet'

###############################################################################

def TestAll():
    print('\n')
    print(pyinex.CallerA1())
    print(pyinex.CallerA1Full())
    print(pyinex.CallerR1C1())
    print(pyinex.CallerR1C1Full())
    print(pyinex.CallerSheet())

    return pyinex.CallerA1Full()

###############################################################################

def CallerA1Test():
    return pyinex.CallerA1()

###############################################################################

def CallerA1FullTest():
    return pyinex.CallerA1Full()

###############################################################################

def CallerR1C1Test():
    return pyinex.CallerR1C1()

###############################################################################

def CallerR1C1FullTest():
    return pyinex.CallerR1C1Full()

###############################################################################

def CallerSheetTest():
    return pyinex.CallerSheet()

###############################################################################

if __name__ == '__main__':
    TestAll()
