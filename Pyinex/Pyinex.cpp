// $Id: Pyinex.cpp 182 2010-01-19 07:04:18Z Ross $

/*
<PyinexLicense>

This file is part of Pyinex, a project to embed python in Excel.
 
Copyright (c) 2010 Ross Levinsky

All rights reserved.

The Pyinex project is built using the xlw framework, found at 
http://xlw.sourceforge.net

The Pyinex license is based on the BSD license template found at
http://www.opensource.org/licenses/bsd-license.php

Redistribution and use in source and binary forms, with or without 
modification, are permitted provided that the following conditions are met:

    Redistributions of source code must retain the above copyright notice,
    this list of conditions and the following disclaimer.
    
    Redistributions in binary form must reproduce the above copyright notice, 
    this list of conditions and the following disclaimer in the documentation 
    and/or other materials provided with the distribution.
    
    Neither the name of Ross Levinsky nor the names of any other contributors 
    may be used to endorse or promote products derived from this software 
    without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDERS OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

</PyinexLicense>
*/

#include "stdafx.h"

// Force export of functions implemented in XlOpenClose.h and required by Excel
#pragma comment (linker, "/export:_xlAutoOpen")
#pragma comment (linker, "/export:_xlAutoClose")

using namespace xlw;

//////////////////////////////////////////////////////////////////////////////
//
// Before Excel 2007, Excel took up to 30 arguments, but even that seems ridiculously many, so I 
// limited it to (a still too-large) 15. Interestingly, I tried to set this to 20, and the call
// to register the function with Excel failed at his line:
//
//     int err = static_cast<int>(XlfExcel::Instance().Call4v(xlfRegister, &res, 10 + nbargs, rgx));
// 
// The +10 slots are used to pass other info about the call into Excel; I don't know if this is
// something that's unique to XLW, or generic to Excel (haven't had a chance to read the docs), but
// it seems to cause function registration failure. For now, I'll lower the input arg count to 15,
// which should be sufficient for almost anything.
//
// Reminder - anonymous namespaces are the new, C++ way to have static vars
// and functions, with linkage local to the file

namespace {
    const long g_numCMArgs = 15; 
    const long g_argcountBeyondPyArgs = 2; // Filename + function name
}

// The only prototype we need from the python add-in functions; no need for a separate header
PyMODINIT_FUNC PyInit_pyinex(void);

//////////////////////////////////////////////////////////////////////////////
//
// MSDN has a good DLLMain skeleton:
//
// http://msdn.microsoft.com/en-us/library/ms682596%28VS.85%29.aspx
//
// But I'm not including it here, to avoid the temptation to put anything
// inappropriate into the DLL_PROCESS_ATTACH call (a mistake I've made).
// Do all interesting initialization in PyinexGlobalInit(); its Factory()
// function has a static object that does all initialization. CRT guarantees
// single-threaded initialization of statics, so this will only happen once,
// in a thread-safe fashion. To be sure that initialization happens before 
// the first call into the XLL, we need to get a ref to the object in each
// function. I've wrapped this reference grabbing in the EXCEL_BEGIN_PYINEX 
// macro, which simply extends XLW's EXCEL_BEGIN macro by prepending a call
// to PyinexGlobalInit::Factory()

#define EXCEL_BEGIN_PYINEX  const PyinexGlobalInit& rGlobalInit = \
    PyinexGlobalInit::Factory(); EXCEL_BEGIN

namespace {

    class PyinexGlobalInit {

    public:
        static const PyinexGlobalInit& Factory()
        {
            static PyinexGlobalInit g_obj;
            return g_obj;
        }

        HWND ConsoleHandle() const { return m_hConsoleWindow; }

    private:
        PyinexGlobalInit()
        {
            // Get console allocated and output streams hooked up first, so we can log
            // errors in the Python initialization
            AllocConsole();
            m_hConsoleWindow = GetConsoleWindow();
            
            // Call this now so we can log error
            if (!SetupOutputStreams()) {
                ERROUT("Error setting up output streams in all CRTs...");
            }
            
            // I tried to hook the console wndproc to stop use of the console close button
            // (because it brings down Excel), but it turns out that Windows won't let us
            // do that - consoles are owned by csrss.exe, not by our proc, so we get an
            // ACCESS_DENIED if we call SetWindowLongPtr. The best we can hope for is to
            // disable the close button entirely.

            HMENU hMenu = GetSystemMenu(m_hConsoleWindow, FALSE);
            if (!ModifyMenu(hMenu, SC_CLOSE, MF_BYCOMMAND | MF_GRAYED, NULL, NULL)) {
                std::string err;
                GetWindowsErrorText(err);
                ERROUT(err.c_str());
            }

            SetConsoleTitle("Excel Python output");

            // Return code refers to previous state of windows hiding; no errors to receive, apparently.
            ShowWindow(m_hConsoleWindow, SW_HIDE);

            // Now we can bring in Python - errors in startup should be logged to console
            // Add a builtin module, before Py_Initialize
            PyImport_AppendInittab("pyinex", PyInit_pyinex);

#if PY_MAJOR_VERSION < 3    
            Py_SetProgramName("Excel");
#else
            Py_SetProgramName(L"Excel");
#endif

            // Initialize the Python interpreter.  Required.
            Py_Initialize();

            PyImport_ImportModule("pyinex");
        }

        ~PyinexGlobalInit()
        {
            Py_Finalize();

            // For reasons I don't understand, this call:
            //
            // CloseHandle(m_hConsoleWindow);
            //
            // causes an "invalid handle" exception when Excel shuts down. 
            // Maybe GetConsoleWindow() doesn't increment the handle count?
            // In any case, the problem goes away when we stop calling CloseHandle().

            FreeConsole();
        }

        HWND m_hConsoleWindow;
    };
}

//////////////////////////////////////////////////////////////////////////////

extern "C" {

//////////////////////////////////////////////////////////////////////////////

    LPXLFOPER EXCEL_EXPORT 
    xlPyCall(   XlfOper xlFilename,  
                XlfOper xlFunction,
                XlfOper xlCM1,
                XlfOper xlCM2,
                XlfOper xlCM3,
                XlfOper xlCM4,
                XlfOper xlCM5,
                XlfOper xlCM6,
                XlfOper xlCM7,
                XlfOper xlCM8,
                XlfOper xlCM9,
                XlfOper xlCM10,
                XlfOper xlCM11,
                XlfOper xlCM12,
                XlfOper xlCM13,
                XlfOper xlCM14,
                XlfOper xlCM15 )
    {
        EXCEL_BEGIN_PYINEX;
  
        // Don't execute this call from the function wizard
        if (XlfExcel::Instance().IsCalledByFuncWiz()) {
            return XlfOper(true);
        }

        // DON'T DECREMENT THE MODULE POINTER - its lifetime is managed by a separate cache object.
        PyObject* pModule = NULL, *pFunction = NULL;
        bool rc =  GetPyModuleAndFunctionObjects(  xlFilename.AsWstring(), 
                                                   xlFunction.AsString(), 
                                                   pModule, 
                                                   pFunction);
        if (!rc) {
            assert(!pModule);
            assert(!pFunction);
            return XlfOper::Error(0);
        }

        CellMatrix arrCM[] = {
               CellMatrix(xlCM1.AsCellMatrix("CM1")),
               CellMatrix(xlCM2.AsCellMatrix("CM2")),
               CellMatrix(xlCM3.AsCellMatrix("CM3")),
               CellMatrix(xlCM4.AsCellMatrix("CM4")),
               CellMatrix(xlCM5.AsCellMatrix("CM5")),
               CellMatrix(xlCM6.AsCellMatrix("CM6")),
               CellMatrix(xlCM7.AsCellMatrix("CM7")),
               CellMatrix(xlCM8.AsCellMatrix("CM8")),
               CellMatrix(xlCM9.AsCellMatrix("CM9")),
               CellMatrix(xlCM10.AsCellMatrix("CM10")),
               CellMatrix(xlCM11.AsCellMatrix("CM11")),
               CellMatrix(xlCM12.AsCellMatrix("CM12")),
               CellMatrix(xlCM13.AsCellMatrix("CM13")),
               CellMatrix(xlCM14.AsCellMatrix("CM14")),
               CellMatrix(xlCM15.AsCellMatrix("CM15"))
        };

        // Compiler doesn't complain if we have too few initializers (only if too many);
        // need to explicitly test sizing
        assert( NELEMS(arrCM)== g_numCMArgs );

        // Examine the function's PyCodeObject to see how many arguments its definition contains.
        //
        // If the function is declared with a vararg param (*varname), pass all possible params to Python,
        // whether they're empty or not. We can't say if those empty params are meaningful to a user function,
        // so we pass them all on.
        //
        // If the function does not have a vararg param, prune out all params past the number defined, and only
        // pass that defined number. To do otherwise causes Python to throw a TypeError (see err_args() in ceval.c).

        int pyCallArgcount = 0;
        if (rc) {
            assert(pFunction);
            PyCodeObject* pCO = (PyCodeObject*)PyFunction_GET_CODE(pFunction);
            assert(pCO);
            pyCallArgcount = pCO->co_argcount;

            // co_argcount refers to the number  of params in the function's opening "define" statement.
            // If it exceeds the number of params that Excel can pass in, there's no way to call this function here.
            // Default param values don't provide a loophole, as Excel has no natural way to support named params.
            if (pyCallArgcount > g_numCMArgs) {
                ERROUT("Function %s has %d arguments; this exceeds the maximum allowable number %d", 
                    xlFunction.AsString(), pyCallArgcount, g_numCMArgs);
                rc = false;
            }

            // co_flags has the CO_VARARGS bit set if the function has a vararg param in its definition.
            if (rc && (pCO->co_flags & CO_VARARGS)) {
                pyCallArgcount = g_numCMArgs; // Forces us to pass everything Excel has on to Python
            }
        }

        // Assemble the args
        long cmDx;
        PyObject* pArgs = PyTuple_New(pyCallArgcount);
        PyObject *pValue = NULL;

        for(cmDx = 0; rc && cmDx < pyCallArgcount; ++cmDx) {
            CellMatrix& rCM = arrCM[cmDx];
            rc = ConvertCellMatrixToPyObject( rCM, pValue );
            if (rc) {
                assert(pValue);
                PyTuple_SetItem(pArgs, cmDx, pValue); // pRows reference stolen here
            } else {
                // pArgs is freed below, with pFunction and pResult
                assert(!pValue);
                ERROUT("Failed to convert argument %d to a PyObject", cmDx);
            }
        }

        // Make the call
        PyObject* pResult = NULL;
        if (rc) {
            pResult = PyObject_CallObject(pFunction, pArgs);
            if (!pResult) {
                if (PyErr_Occurred()) {
                    PyErr_Print();
                }
                rc = false;
            }
        }

        // Unpack results
        CellMatrix retMatrix;
        if (rc && pResult != NULL) {
            rc = ConvertPyObjectToCellMatrix(pResult, retMatrix);
            if (!rc) {
                if (PyErr_Occurred()) {
                    PyErr_Print();
                }
            }
        }

        // Clean up
        // Because PyTuple_SetItems steals refs, the decrement of pArgs should free all contained objects
        Py_XDECREF(pArgs);  
        Py_XDECREF(pFunction);
        Py_XDECREF(pResult);

        if (rc) {
            return XlfOper(retMatrix);
        } else {
            return XlfOper::Error(0);
        }

        EXCEL_END;
    }

//////////////////////////////////////////////////////////////////////////////
//
// Console display param validation is quite long; isolate it (and refactor
// it later into a set of calls to a generic validator function?)

    bool ValidatePyConsoleParams( XlfOper& xlShowConsole,
                                  XlfOper& xlX,
                                  XlfOper& xlY,
                                  XlfOper& xlWidth,
                                  XlfOper& xlHeight,
                                  XlfOper& xlPreserveLocation,
                                  bool& showConsole,
                                  int& x,
                                  int& y,
                                  int& width,
                                  int& height,
                                  bool& preserveLocation,
                                  std::string& errTxt )
    {
        // The only required param is the showConsole boolean
        if (!xlShowConsole.IsBool()) {
            errTxt = "showConsole parameter is not TRUE/FALSE";
            return false;
        }

        showConsole = xlShowConsole.AsBool();
        // Save time by bailing out early if we're just hiding the console.
        // Caller must recognize that the other params haven't been validated.

        if (!showConsole) {
            return true;
        }
        
        // Need to determine the desktop size so we can appropriately limit console
        // placement. Do it just once, to save time.
        //
        // All of these statics are not thread safe, but since they get immmutable
        // system info, it should be fine to have multiple threads stepping on them

        static RECT desktopSize;
        static bool bGotDesktopSize = false;
        if (!bGotDesktopSize) {
            GetWindowRect( GetDesktopWindow(), &desktopSize );
            bGotDesktopSize = true;
        }

        // Default location of the upper-left corner of the console
        x = 50, y = 50;

        if (xlX.IsNumber() && !xlY.IsNumber()) {
            errTxt = "x specified, but not y. If one is present, the other must be, too.";
            return false;
        }

        if (!xlX.IsNumber() && xlY.IsNumber()) {
            errTxt = "y specified, but not x. If one is present, the other must be, too.";
            return false;
        }

        if (xlX.IsNumber() && xlY.IsNumber()) {
            x = xlX.AsInt();
            y = xlY.AsInt(); 

            // Defaults are > 0, so save time by only testing for this after using user vals
            if (x < 0) {
                errTxt = "x must be positive";
                return false;
            }

            if (y < 0) {
                errTxt = "y must be positive";
                return false;
            }
        }

        // Default locations could theoretically be too large, though, on some hard-to-envision
        // super-tiny screen. Test always.
        //
        // The -1 is because the pixels are zero-indexed, and right and bottom for the desktop look 
        // to be one past the actual max pixel count
        if (x >= desktopSize.right) {
            std::ostringstream os;
            os << "x is past the right edge of the screen - max value is " << desktopSize.right - 1;
            errTxt = os.str();
            return false;
        }

        if (y >= desktopSize.bottom) {
            std::ostringstream os;
            os << "y is past the bottom edge of the screen - max value is " << desktopSize.bottom - 1;
            errTxt = os.str();
            return false;  
        }

        // On to width/height
        //
        // Console can hang off the edge of the screen. Let the default size be smallish (VGA window).
        width= 640;
        height = 480;

        if (xlWidth.IsNumber() && !xlHeight.IsNumber()) {
            errTxt = "Width specified, but not height. If one is present, the other must be, too.";
            return false;
        }

        if (!xlWidth.IsNumber() && xlHeight.IsNumber()) {
            errTxt = "Height specified, but not width. If one is present, the other must be, too.";
            return false;
        }

        if (xlWidth.IsNumber() && xlHeight.IsNumber()) {
            width = xlWidth.AsInt();
            height = xlHeight.AsInt(); 
        }

        if (width <= 0) {
            errTxt = "Width must be positive";
            return false;
        }

        if (height <= 0) {
            errTxt = "Height must be positive";
            return false;
        }

        // Finally, the optional flag to preserve location
        if (xlPreserveLocation.IsBool()) {
            preserveLocation = xlPreserveLocation.AsBool();
        } else if (xlPreserveLocation.IsMissing()) { // not passed in
            preserveLocation = true;
        } else {
            errTxt = "preserveLocation flag is specified, but is not TRUE or FALSE";
            return false;
        }

        return true;
    }

    ////////////////////////////////////////////////////

    LPXLFOPER EXCEL_EXPORT 
    xlPyConsole(  XlfOper xlShowConsole,
                  XlfOper xlX,
                  XlfOper xlY,
                  XlfOper xlWidth,
                  XlfOper xlHeight,
                  XlfOper xlPreserveLocation )
    {
       EXCEL_BEGIN_PYINEX;

        // Don't execute this call from the function wizard
        if (XlfExcel::Instance().IsCalledByFuncWiz()) {
            return XlfOper(true);
        }

        int x,y,width,height;
        bool showConsole, preserveLocation;
        std::string errTxt;

        // If showConsole is false, this function doesn't validate the rest of the 
        // input params; it just fails out early to save time

        if (!ValidatePyConsoleParams( xlShowConsole,
                                      xlX,
                                      xlY,
                                      xlWidth,
                                      xlHeight,
                                      xlPreserveLocation,
                                      showConsole,
                                      x,
                                      y,
                                      width,
                                      height,
                                      preserveLocation,
                                      errTxt ) ) {
            return XlfOper(errTxt);
        }


        HWND hConsole = rGlobalInit.ConsoleHandle();

        if (showConsole) {
            // Always have to set the position the first time. Need to
            // think more about the thread safety of this - it seems fine,
            // as we flip the flag AFTER the display is done, and one thread
            // will always see this as the first time.

            static bool firstTime = true;
            if (firstTime) {
                MoveWindow(hConsole, x, y, width, height, TRUE);
                ShowWindow(hConsole, SW_SHOW);
                firstTime = false;
            } else { // Console has already been displayed
                // Only move the window (and attempt to reshow it) if it's 
                // not visible. The move on visibility cycling is part of
                // the design behavior of this function; the avoidiance of
                // ShowWindow() saves an unnecessary call.

                if (IsWindowVisible(hConsole) == FALSE) {
                    if ( !preserveLocation ){
                        MoveWindow(hConsole, x, y, width, height, TRUE);
                    }
                    ShowWindow(hConsole, SW_SHOW);
                }
            }    
            return XlfOper("Showing python console");
        } else {
            ShowWindow(hConsole, SW_HIDE);
            return XlfOper("Hiding python console");
        }

        EXCEL_END;
    }

//////////////////////////////////////////////////////////////////////////////

    LPXLFOPER EXCEL_EXPORT 
    xlPyClearConsole(  XlfOper xlClearConsole )
    {
        EXCEL_BEGIN_PYINEX;

        // Don't execute this call from the function wizard
        if (XlfExcel::Instance().IsCalledByFuncWiz()) {
            return XlfOper(false);
        }

        bool bClear = xlClearConsole.AsBool();
        if (bClear) {
            system("cls");
            return XlfOper("Console cleared");
        } else {
            return XlfOper("Console untouched");
        }

        EXCEL_END;
    }

//////////////////////////////////////////////////////////////////////////////

    LPXLFOPER EXCEL_EXPORT 
    xlPyVerbose(  XlfOper xlVerbosity )
    {
        EXCEL_BEGIN_PYINEX;

        // Don't execute this call from the function wizard
        if (XlfExcel::Instance().IsCalledByFuncWiz()) {
            return XlfOper(false);
        }

        int verbosity = xlVerbosity.AsInt();
        if (verbosity < 0) {
            WARNOUT("Input verbosity was %d; min value is zero", verbosity);
            verbosity = 0;
        } 

        Py_VerboseFlag =  verbosity;
        return XlfOper((double)Py_VerboseFlag); // no int c-tor for XlfOper
        
        EXCEL_END;
    }

//////////////////////////////////////////////////////////////////////////////

    LPXLFOPER EXCEL_EXPORT 
    xlPyModuleFreshnessCheck(  XlfOper xlCheckFreshness )
    {
        EXCEL_BEGIN_PYINEX;

        // Don't execute this call from the function wizard
        if (XlfExcel::Instance().IsCalledByFuncWiz()) {
            return XlfOper(false);
        }

        bool bOutput;
        if (xlCheckFreshness.IsBool()) {
            bOutput = xlCheckFreshness.AsBool();
            SetModuleFreshnessCheck( bOutput );
        } else {
            bOutput = ModuleFreshnessCheckEnabled();
        }

        return XlfOper(bOutput);
        
        EXCEL_END;
    }

//////////////////////////////////////////////////////////////////////////////

    LPXLFOPER EXCEL_EXPORT 
    xlPyLoadedLibrary(  XlfOper xlLibrary )
    {
        EXCEL_BEGIN_PYINEX;

        // Don't execute this call from the function wizard
        if (XlfExcel::Instance().IsCalledByFuncWiz()) {
            return XlfOper(false);
        }

        std::string library = xlLibrary.AsString();
        std::transform( library.begin(), library.end(), library.begin(), (int(*)(int)) tolower );

        // Could use "library" var to do the search, rather than these seeminly-superfluous vars, 
        // but this gives us more flexibility. Someday we might want to fuzzy-match the library
        // string (taking, say, "pyth" for python), and we'll need canonical search strings.
        //        
        // Plus, we need to set the extension search string, anyway (as it's not passed-in).

        std::wstring basenameSearch, extensionSearch;
        if (library == "python") {
            basenameSearch = L"python";
            extensionSearch = L"dll";
        } else if (library == "pyinex") {
            basenameSearch = L"pyinex";
            extensionSearch = L"xll";
        } else {
            return XlfOper("Library param must be \"python\" or \"pyinex\"");
        }

        std::vector<std::wstring> vecNames;
        if (!GetLoadedModuleNames(vecNames)) {
            ERROUT("GetLoadedModuleNames failed");
            return XlfOper("GetLoadedModuleNames failed"); 
        }

        std::wstring dllName;
        std::vector<std::wstring>::const_iterator it;
        for (it = vecNames.begin(); it!=vecNames.end(); ++it) {
            std::wstring path, basename, extension;            
            if (!SplitPathBasenameExtension(*it, path, basename, extension)) {
                ERROUT("Couldn't split module name %s", ASCII_REPR(*it));
                continue;
            }

            if ( _wcsnicmp(basename.c_str(), basenameSearch.c_str(), basenameSearch.length()) == 0 &&
                 _wcsicmp(extension.c_str(), extensionSearch.c_str()) == 0 )
            {
                // We have a candidate. Return the first one seen, but make sure it's unique;
                // output error messages for any competing candidates.
                if (dllName.empty()) {
                    dllName = *it;
                } else {
                    ERROUT("Multiple %s DLLs loaded. First seen: %s", ASCII_REPR(basename), ASCII_REPR(dllName));
                    ERROUT("Next %s DLL seen: %s", ASCII_REPR(basename), ASCII_REPR(*it));
                } 
            }
        }   // vecNames loop
  
        // Should not be possible to find no loaded Python or Pyinex libs (as this function lives in Pyinex)
        if (dllName.empty()) {
            dllName = L"No " + basenameSearch + L" DLL loaded; this should not be possible";
        }
      
        return XlfOper(dllName); 
        
        EXCEL_END;
    }

//////////////////////////////////////////////////////////////////////////////

} // extern "C"

//////////////////////////////////////////////////////////////////////////////

namespace {

    XLRegistration::Arg PyConsoleArgs[] = {
        { "showConsole", "Required - TRUE to show python console, FALSE to hide it", "XLF_OPER" },
        { "x", "Optional - x location of upper-left of console window, in pixels", "XLF_OPER" },
        { "y", "Optional - y location of upper-left of console window, in pixels", "XLF_OPER" },
        { "width", "Optional - width of the console window, in pixels", "XLF_OPER" },
        { "height", "Optional - height of the console window, in pixels", "XLF_OPER" },
        { "preserveLocation", "Optional - if TRUE, hiding then unhiding the console window doesn't change its size or location. "
          "If false, size and location are reset after a hide/show cycle. Default is TRUE.", "XLF_OPER" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerPyConsoleArgs(
        "xlPyConsole", "PyConsole", "Shows/hides the python console",
        "Pyinex", PyConsoleArgs, 6); 

    /******************/

    XLRegistration::Arg PyClearConsoleArgs[] = {
        { "clearConsole", "TRUE to clear python console, FALSE to do nothing", "XLF_OPER" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerPyClearConsoleArgs(
        "xlPyClearConsole", "PyClearConsole", "Clears the python console",
        "Pyinex", PyClearConsoleArgs, 1); 

    /******************/

    XLRegistration::Arg PyVerboseArgs[] = {
        { "verbosity", "Integer setting of Python verbosity flag - default is zero, and higher values are more verbose", "XLF_OPER" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerPyVerbose(
        "xlPyVerbose", "PyVerbose", "Sets Python's internal verbosity flag",
        "Pyinex", PyVerboseArgs, 1); 

    /******************/

    XLRegistration::Arg PyModuleFreshnessCheckArgs[] = {
        { "checkFreshness", "Boolean - when TRUE, Pyinex checks module file write times on every calculation, and reloads stale modules if necessary", "XLF_OPER" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerPyModuleFreshnessCheck(
        "xlPyModuleFreshnessCheck", "PyModuleFreshnessCheck", "Sets and displays the value of a flag determining whether Pyinex checks module freshness",
        "Pyinex", PyModuleFreshnessCheckArgs, 1); 

    /******************/

    XLRegistration::Arg PyLoadedLibraryArgs[] = {
        { "library", "Name of library to look up - can be one of two case-insensitive values: 'Python' or 'Pyinex'", "XLF_OPER" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerPyLoadedLibrary(
        "xlPyLoadedLibrary", "PyLoadedLibrary", "Returns location of loaded Python or Pyinex library",
        "Pyinex", PyLoadedLibraryArgs, 1); 

    /******************/

    XLRegistration::Arg PyCallArgs[] = {
        { "filename", "Python file to parse", "XLF_OPER" },
        { "function", "Function to call in the python file", "XLF_OPER" },
        { "cellMatrix1",  "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix2",  "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix3",  "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix4",  "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix5",  "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix6",  "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix7",  "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix8",  "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix9",  "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix10", "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix11", "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix12", "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix13", "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix14", "Rectangular region holding function arguments", "XLF_OPER"},
        { "cellMatrix15", "Rectangular region holding function arguments", "XLF_OPER"}
    };

    XLRegistration::XLFunctionRegistrationHelper registerPyCallArgs(
        "xlPyCall", "PyCall", "Call a function in a python file",
        "Pyinex", PyCallArgs, g_numCMArgs + g_argcountBeyondPyArgs); 

    // Compiler doesn't complain if we have too few initializers (only if too many);
    // need to explicitly test sizing. Nowhere to do this test except in the ctor of a
    // global object.

    struct RegSizeTest {
        RegSizeTest() { 
            assert( NELEMS(PyCallArgs) == (g_numCMArgs + g_argcountBeyondPyArgs) );
        }
    };

    RegSizeTest rst;
}
//////////////////////////////////////////////////////////////////////////////
