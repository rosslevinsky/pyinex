// $Id: PythonExtension.cpp 182 2010-01-19 07:04:18Z Ross $

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

using namespace xlw;

//////////////////////////////////////////////////////////////////////////////

namespace {
    
    // These are zero-indexed coming back from XLW
    // Leaving them that way, as that will be useful for A1-style cell calcs
    // R1C1 callers will increment; A1 callers will not

    void GetCallingSheetColRow( std::string* pSheet, int& rCol, int& rRow )
    {
        XlfOper caller;
        XlfExcel::Instance().Call(xlfCaller,caller, 0);
        if (!caller.IsSRef()) {
            rCol = -1;
            rRow = -1;
            if (pSheet) {
                *pSheet = "Function not called from cell; macro or toolbar call?";
            }
            return;
        }
        
        // Could use caller.AsRef() to get an XlfRef from which we can extract rows and
        // cols, but it embeds an unnecessary call to Coerce into an xltypeRef. This is
        // not needed because we have an xltypeSRef, which already has the rows and cols
        // available. Unfortunately, XLW doesn't expose SRefs (yet), so we'll have to do
        // some Excel-case-specific hackery to gain this efficiency.

        if (XlfExcel::Instance().excel12()) {
            LPXLOPER12 pX = (LPXLOPER12) caller.GetLPXLFOPER();
            assert(pX->xltype == xltypeSRef);
            rRow = pX->val.sref.ref.rwFirst;
            rCol = pX->val.sref.ref.colFirst;
        } else {
            LPXLOPER pX = (LPXLOPER) caller.GetLPXLFOPER();
            assert(pX->xltype == xltypeSRef);
            rRow = pX->val.sref.ref.rwFirst;
            rCol = pX->val.sref.ref.colFirst;
        }

        if (pSheet) {
            XlfOper sheet;
            XlfExcel::Instance().Call(xlSheetNm, sheet, 1, (LPXLFOPER) caller );
            *pSheet = sheet.AsString();
        }
    }

    //////////////////////////////////////////////////////////////////////////////
    // 
    // This algorithm is correct, but it was hard to convince myself of that fact.
    // See all of the details and test code in ExcelColnameValidation.py, in the 
    // Scripts project directory.
    //
    // Takes a zero-indexed col.

    void ConvertColnumToText( int col, std::string& colText )
    {
        // From http://msdn.microsoft.com/en-us/library/aa730921.aspx:
        //
        // The Excel 2007 "Big Grid" increases the maximum number of rows per worksheet 
        // from 65,536 to over 1 million, and the number of columns from 256 (IV) to 16,384 (XFD).

        const long int EXCEL4MAXCOL = 256;
        const long int EXCEL12MAXCOL = 16384;

        if (col < 0) {
            colText = "Col less than zero";
            ERROUT("Col %d supplied; must be greater than zero", col);
            return;
        }

        int colLimit = EXCEL4MAXCOL;
        if (XlfExcel::Instance().excel12()) {
            colLimit = EXCEL12MAXCOL;
        }
   
        if (col >= colLimit) {
            colText = "Col too large";
            ERROUT("Col %d supplied; must be less than %d", col, colLimit);
            return;
        }
        
        // This pulls out chars from the right; fill array backwards
        int remainder;
        char colName[4];
        colName[3] = 0;
        int colDx = 2;
        while( true ) {
            remainder = col % 26;
            colName[colDx] = 'A' + remainder;
            if (col < 26) {
                break;
            }
            col = ((col - remainder) / 26) - 1;
            --colDx;
        }
        assert(colDx >= 0);
        colText = colName + colDx;
    }

    //////////////////////////////////////////////////////////////////////////////

    void
    AssembleCallerName( bool R1C1,     // true = R1C1, false = A1
                        bool full,     //true = sheet name prepended, false == no sheet name
                        std::string& rCallerName )
    {
        std::string sheet;
        int col, row;
        GetCallingSheetColRow( full ? &sheet : NULL, col, row );
        
        std::ostringstream ostr;
        if (full) {
            ostr << sheet << "!";
        }

        if (R1C1) {
            ostr << "R" << (row + 1) << "C" << (col + 1);
        } else {
            std:: string colText;
            ConvertColnumToText( col, colText );
            ostr << colText << row + 1;
        }

        rCallerName = ostr.str();
    }

    //////////////////////////////////////////////////////////////////////////////

    PyObject* 
    AssembleCallerNamePyObj( bool R1C1,      // true = R1C1, false = A1
                             bool full )     //true = sheet name prepended, false == no sheet name
    {    
        std::string callerName;
        AssembleCallerName(R1C1, full, callerName);

#if PY_MAJOR_VERSION < 3
        return PyString_FromString(callerName.c_str());
#else 
        return PyUnicode_FromString(callerName.c_str());
#endif
    }

    //////////////////////////////////////////////////////////////////////////////

} // namespace

static PyObject* 
pyinex_CallerA1(PyObject *self, PyObject *args)       { return AssembleCallerNamePyObj(false, false); }

static PyObject* 
pyinex_CallerA1Full(PyObject *self, PyObject *args)   { return AssembleCallerNamePyObj(false, true); }

static PyObject* 
pyinex_CallerR1C1(PyObject *self, PyObject *args)     { return AssembleCallerNamePyObj(true, false); }

static PyObject* 
pyinex_CallerR1C1Full(PyObject *self, PyObject *args) { return AssembleCallerNamePyObj(true, true); }

static PyObject* 
pyinex_CallerSheet(PyObject *self, PyObject *args) 
{ 
    std::string sheet;
    int col, row;
    GetCallingSheetColRow( &sheet, col, row );
#if PY_MAJOR_VERSION < 3
    return PyString_FromString(sheet.c_str());
#else 
    return PyUnicode_FromString(sheet.c_str());
#endif
}

// This function takes an optional boolean from Python to determine whether or not to clear
// any observed abort request. True (from Python) = clear the break, and False = don't clear it.
// This is useful in case one has multiple functions on a sheet that will want to handle a 
// break request. 
//
// The call to Excel has the opposite semantics; the flag is "PreserveBreak," with a true 
// NOT clearing the break. I find that confusing, so I flipped the meaning for Python users.

static PyObject* 
pyinex_Break(PyObject *self, PyObject *args) 
{ 
    // Has a clear-break flag been passed in? If not, default is to NOT clear any break
    bool bClearBreak = false;

    if (PyTuple_Check(args)) {
        int len = (int)PyObject_Length(args);
        if (len == 0) {
            // Nothing was passed in; take the default action, which is to NOT clear the break
        } else if (len == 1) {
            PyObject* pElem = PySequence_GetItem(args, 0); // returns a new reference; must decrement
            if (pElem) {
                if (PyBool_Check(pElem)) {
                    bClearBreak = (pElem == Py_True);
                } else if (pElem == Py_None) {
                    // User might pass in an empty cell, which will be seen here as None. Consider that a False.
                } else {
                    ERROUT("pyinex.Break() takes a single optional boolean; instead, a variable of type %s was passed in",
                        Py_TYPE(pElem)->tp_name);
                }
                Py_DECREF(pElem); 
            }
        } else {
            ERROUT("pyinex.Break() takes a single optional boolean; instead, %d variables were passed in", len);
        }
    }

    XlfOper breakReq;
    XlfOper preserveBreak(!bClearBreak);   // as noted above, opposite semantics for Excel
    XlfExcel::Instance().Call(xlAbort, breakReq, 1, (LPXLFOPER) preserveBreak ); 
    bool bBreak = breakReq.AsBool();
    
    return PyBool_FromLong( (long)bBreak );
}

//////////////////////////////////////////////////////////////////////////////

static PyMethodDef PyinexMethods[] = {
    {"CallerA1",       pyinex_CallerA1,         METH_VARARGS, "Returns the calling cell in A1 format"},
    {"CallerA1Full",   pyinex_CallerA1Full,     METH_VARARGS, "Returns the calling cell in A1 format with sheet name prepended"},
    {"CallerR1C1",     pyinex_CallerR1C1,       METH_VARARGS, "Returns the calling cell in R1C1 format"},
    {"CallerR1C1Full", pyinex_CallerR1C1Full,   METH_VARARGS, "Returns the calling cell in R1C1 format with sheet name prepended"},
    {"CallerSheet",    pyinex_CallerSheet,      METH_VARARGS, "Returns the sheet name of the calling cell"},
    {"Break",          pyinex_Break,            METH_VARARGS, "Returns a boolean indicating whether or not the user has pressed the escape key"},
    {NULL, NULL, 0, NULL} /* Sentinel */
};

//////////////////////////////////////////////////////////////////////////////

#if PY_MAJOR_VERSION < 3

PyMODINIT_FUNC
PyInit_pyinex(void)
{
    PyObject* pRawModule =  Py_InitModule("pyinex", PyinexMethods);
    PyObject* pModuleName = PyString_FromString("pyinex");
    PyObject* pModule = PyImport_Import(pModuleName);
    Py_DECREF(pModuleName);

    // EXPERIMENTAL - this adds the pyinex module name to the topmost
    // builtin dictionary (I think; docs are not useful here). With these
    // lines, python scripts running in Excel don't have to call 
    // "import pyinex" to have access to the add-in functions

    PyObject* pBuiltinDict = PyEval_GetBuiltins();
    int res = PyDict_SetItemString(pBuiltinDict, "pyinex", pModule); 
}

#else 

static struct PyModuleDef pyinexmodule = {
   PyModuleDef_HEAD_INIT,
   "pyinex",   /* name of module */
   NULL, /* module documentation, may be NULL */
   -1,       /* size of per-interpreter state of the module,
                or -1 if the module keeps state in global variables. */
   PyinexMethods
};

PyMODINIT_FUNC
PyInit_pyinex(void)
{
    PyObject *pModule;
    pModule = PyModule_Create(&pyinexmodule);
    if (pModule == NULL)
        return NULL;

    // EXPERIMENTAL - this adds the pyinex module name to the topmost
    // builtin dictionary (I think; docs are not useful here). With these
    // lines, python scripts running in Excel don't have to call 
    // "import pyinex" to have access to the add-in functions

    PyObject* pBuiltinDict = PyEval_GetBuiltins();
    int res = PyDict_SetItemString(pBuiltinDict, "pyinex", pModule); 
    return pModule;
}

#endif

//////////////////////////////////////////////////////////////////////////////
