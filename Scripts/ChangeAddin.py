import re, os, sys
from _winreg import *

###############################################################################

def main():

    version = 12
    if len(sys.argv) > 1:
        version = sys.argv[1]

    addinKey, xllDir, xlls = GetXllReginfo(version, True)

    cInt = SelectXll(xllDir, xlls)

    UpdateAddinKey(version, addinKey, xllDir, xlls, cInt, True)

###############################################################################

def SelectXll(xllDir, xlls):

    # Choose one of the available XLLs for installation
    print("Choose an XLL from " + xllDir + ":\n")
    while True:
        for k in xlls:
            print(str(k) + "\t" + xlls[k])

        c = raw_input("\n>")
        try: 
            cInt = int(c)
            if cInt < 1 or cInt > len(xlls):
                raise ValueError
            else:
                break
        except ValueError:
            print("Enter a value between 1 and " + str(len(xlls)) + "\n")

    return cInt

###############################################################################

def GetXllReginfo(version, pr = False):

    # Get existing registered XLLs
    aReg = ConnectRegistry(None,HKEY_CURRENT_USER)

    # Raw string can't end in a backslash
    aKey = OpenKey(aReg, XLLRegPath(version))
    pyRegKeys = {}
    pattern = re.compile("Pyinex", re.UNICODE)
    for i in range(1024):                                           
        try:
            n,v,t = EnumValue(aKey,i)
            if t == 1 and pattern.search(v) != None:
                pyRegKeys[n] = v
        except EnvironmentError:                                               
            break          
    CloseKey(aKey)                                                  

    # Make sure we have only one Pyinex XLL registered, and that it's under an OPEN key
    if len(pyRegKeys) != 1:
        exit("This script only works if there's one operating Pyinex add-in already installed in Excel")

    addinKey = pyRegKeys.keys()[0]
    regVal = pyRegKeys[addinKey]

    if re.search("OPEN", addinKey, re.UNICODE) is None:
        exit("Existing Pyinex XLL isn't registered under an OPEN key")

    # Assume that all possible Pyinex XLLs are in the same directory. Parse out the path
    # to the one that's found.

    pattern = re.compile("\"(.*)\"", re.UNICODE)
    m = pattern.search(regVal)
    if m is None:
        exit("Couldn't get a file location out of the Pyinex registry value")
    regFullFile = m.group(1)
    xllDir = os.path.dirname(regFullFile)
    if pr: print("\nCurrent XLL: \n" + regFullFile + "\n")

    # Get files we can use as XLLs
    [root, dirs, files] = os.walk(xllDir).next()
    xlls = {}
    for i,v in enumerate(files):
        xlls[i+1] = v

    return addinKey, xllDir, xlls

###############################################################################

def UpdateAddinKey(version, addinKey, xllDir, xlls, cInt, pr = False):

    # Insert the new value
    newRegVal = xllDir + "\\" + xlls[cInt]
    newRegVal = "/R \"" + newRegVal + "\""
    if pr: print("Inserting: \nKey:\t" + addinKey + "\nValue:\t" + newRegVal + "\n")

    aReg = ConnectRegistry(None,HKEY_CURRENT_USER)
    aKey = OpenKey(aReg, XLLRegPath(version), 0, KEY_WRITE)
    try:   
        SetValueEx(aKey, addinKey,0, REG_SZ, newRegVal)
    except EnvironmentError:                                          
        print("Encountered problems writing into the Registry")
    CloseKey(aKey)

###############################################################################

def XLLRegPath(version):
    # Raw string can't end in a backslash
    return r'SOFTWARE\Microsoft\Office' + '\\' + str(version) + r'.0\Excel\Options'

###############################################################################

if __name__ == '__main__':
    main()

###############################################################################
