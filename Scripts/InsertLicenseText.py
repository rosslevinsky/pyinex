###############################################################################
#
# Script goes through relevant .cpp and .h files, searching for the
# delimiters of the Pyinex license, and replacing everything between those
# delimiters with the latest license text in LICENSE.TXT.
#
# I could've done this whole thing with regexps, but we'd still need to count
# the number of occurrences of the delimiters to make sure that each file is
# well-formed. Seemed like a wash, so I just did it with string's find 
# functions.

import sys, glob, os

openDelim = '<PyinexLicense>'
closeDelim = '</PyinexLicense>'
licenseFile = 'LICENSE.TXT'

###############################################################################

def main():

    if os.path.basename(os.getcwd()) != 'Scripts':
        exit("Run this from the top-level Scripts directory")
    os.chdir('..')

    # Extract current Pyinex license from the LICENSE.TXT file
    with open(licenseFile) as f:
        text = f.read()
    (ldx, rdx) = DelimitedTextBoundaries( licenseFile, text )
    license = text[ldx:rdx]
    print('New license text:\n')
    print(license)

    # Get list of .cpp and .h files that we care about
    fileList = []
    for dir in ('Pyinex', 'TestHarness', 'Utils'):
        for type in ('.h', '.cpp'):
            for fn in glob.glob(dir + '/*' + type):
                fileList.append( os.path.normpath(fn) )
    print()
    print(len(fileList), 'files to change\n')

    # Munge each file
    for fn in fileList:
        with open(fn, "r") as f:
            text = f.read()
        (ldx, rdx) = DelimitedTextBoundaries( fn, text )
        newText = ''.join((text[0:ldx], license, text[rdx:]))
        print('Modifying', fn)
        with open(fn, "w") as f:
            f.write(newText)        

###############################################################################
#
# Exits if the file doesn't have exactly one openDelim before exactly one
# closeDelim.
#
# Returns the index of the start of openDelim, and the index one past the end
# of closeDelim
#

def DelimitedTextBoundaries( filename, text ):
    
    ldx = OneDelimLocation(filename, text, openDelim)
    rdx = OneDelimLocation(filename, text, closeDelim)
    if ldx >= rdx:
        exit(''.join((filename, ' is not formed correctly. ', openDelim, 
                     ' occurs after ', closeDelim)))
    return (ldx, rdx + len(closeDelim))

###############################################################################

def OneDelimLocation( filename, text, delim ):

    ldx = text.find(delim)
    if ldx == -1:
        exit(''.join((filename, ' doesn\'t contain ', delim)))
    rdx = text.rfind(delim)
    if ldx != rdx:
        exit(''.join((filename, ' has multiple copies of ', delim,
             '; only 1 allowed')))
    return ldx

###############################################################################

if __name__ == '__main__':
    main()

###############################################################################
