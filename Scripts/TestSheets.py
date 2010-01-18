import re, os, sys, time, win32com.client
from _winreg import *

###############################################################################

def main():

    excelVersions = (12,)
    xllVersions = ( r'C:\Documents and Settings\Ross\Desktop\Pyinex\trunk\Bin\Pyinex082-py25-vc90-mt.xll',
                    r'C:\Documents and Settings\Ross\Desktop\Pyinex\trunk\Bin\Pyinex082-py26-vc90-mt.xll',
                    r'C:\Documents and Settings\Ross\Desktop\Pyinex\trunk\Bin\Pyinex082-py31-vc90-mt.xll' )

    for excelVersion in excelVersions:
        for xllVersion in xllVersions:
            RunOneExcelVersion(excelVersion, xllVersion)

###############################################################################
#
# This code was inline in main(), but it failed to run both instances of 
# Excel unless I called del(xl) at the end of the main loop. It's unclear if
# this is a Windows issue or a win32com.client issue, but only be deleting the
# entire object reference (or enclosing it in a separate function, which has a
# separate scope that also kills the object at function end) did I get both 
# versions of Excel to run
#

def RunOneExcelVersion(excelVersion, xllVersion):

    # HackExcelCLSID(excelVersion)
    progID = 'Excel.Application'
    xl = win32com.client.Dispatch(progID)
    xl.Visible = 1
    
    # Walk the spreadsheets
    booksDir =  r'C:\Documents and Settings\Ross\Desktop\Pyinex\trunk\Examples' + '\\'
    books = ( 'NumPyDemo.xls',
               'PyinexTest.xls',
               'PythonBreakTest.xls',
               'PythonExtensionTest.xls',
               'UnicodeTest.xls' )

    for b in books:
        bookName = booksDir + b
        bookObj = xl.Workbooks.Open(bookName)
        # Calc type can only be set when a workbook is open. Otherwise, it throws an exception.
        if  excelVersion > 10: xl.Calculation = 0xFFFFEFD9 # = xlCalculationManual        
        xl.RegisterXLL(xllVersion)

        sheet = bookObj.ActiveSheet

        # Running under Parallels, the Excel home dir is funny. Change it here.
        script = sheet.Cells(1,3).Value
        script = os.path.join( booksDir, os.path.basename(script) )
        sheet.Cells(1,3).Value = script

        # Make it easier to see the console
        sheet.Cells(2,5).Formula = '=PyConsole(C2,1100,600,800,400,TRUE)'

        if excelVersion > 10: xl.Calculation = 0xFFFFEFF7 # = xlCalculationAutomatic
        xl.CalculateFull()
        
        raw_input("Hit enter to continue...")
        print

        bookObj.Close(SaveChanges=0)

    # Shut down the application (so we can change the addin)
    xl.Quit()
    del(xl)

###############################################################################
#
# For some unknown reason, MSFT recycles Excel CLSIDs amongst different versions.
#
# To test different versions, one has to manually hack the registry to point to
# the desired version.
# 

def HackExcelCLSID(excelVersion):

    if excelVersion == 10:
        val = "\"C:\Program Files\Microsoft Office\Office10\Excel.exe\" /automation"
    elif excelVersion == 12:
        val = "\"C:\Program Files\Microsoft Office\Office12\Excel.exe\" /automation"
    else:
        exit('Version must be one of (10, 12)')

    aReg = ConnectRegistry(None,HKEY_CLASSES_ROOT)
    finalPaths = ('LocalServer', 'LocalServer32')

    for f in finalPaths:
        excelKey = r'CLSID\{00024500-0000-0000-C000-000000000046}' + '\\' + f
        aKey = OpenKey(aReg, excelKey, 0, KEY_WRITE)
#        print 'BEFORE: ', EnumValue(aKey, 0)[1]
        SetValueEx(aKey, '', 0, REG_SZ, val)
#        print 'AFTER:  ', EnumValue(aKey, 0)[1]
        CloseKey(aKey)

###############################################################################

if __name__ == '__main__':
    main()
