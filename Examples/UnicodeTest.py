import sys

if sys.version[0] == '3':
    # Old-style repr is gone in 3.0; it's now "ascii". Revert to that behavior,
    # so this script will work identically for versions 2 and 3
    repr = ascii

###############################################################################

def ReturnConstString():
    uni = 'Hello...\u03c1\u03b5\u03b1\u03bb\u03bb\u03be Unicode'

    # Can't use 'u' prefix in v3 of Python; won't compile
    if sys.version[0] == '2':
        try:
            uni = uni.decode("unicode_escape")             
        except:
            uni = 'Couldn\'t convert text string to Unicode (in Python)'
    return uni

###############################################################################

def ReturnString(s):

    print('\nInput is ' +    str(type(s)) + ' - ' +  repr(s) + '\n')
    return s

###############################################################################



