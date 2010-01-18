// $Id: Utils.cpp 183 2010-01-19 07:13:15Z Ross $

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

void
PrintError(  pyxErrorSeverity severity,  
             const char* file,
             const char* function,
             int line,
             const char* msg,
             ... )
{
    // Error messages are specified with a C-style format string, so until we
    // move to Boost.Format, we're stuck with printf and its derivatives.
    // Using C++ ostringstream, too, to handle the prepended info, because I don't
    // want to have to think hard about counting chars written by vsnprintf, and
    // the performance hit of this function is irrelevant.

    const long BUFLEN = 8192;
    char formattedText[BUFLEN];
    va_list argptr;
    va_start( argptr, msg);
    if (-1 ==  vsnprintf_s(formattedText, BUFLEN - 1, _TRUNCATE, msg, argptr)) {
        formattedText[BUFLEN - 1] = 0; // Don't know if truncation writes a terminating null
        printf("Truncation of vsnprintf_s buffer; extend it and recompile\n");            
    }
    va_end(argptr);

    // ostringstream is an output-only stringstream
    std::ostringstream stm;
    stm << "\n" << formattedText << "  (" << file << ":" << function << ":" << line << ")\n\n";

    // Would use cout, but I'm not sure if the various redirected CRTs will work with C++ 
    // output streams. So...

    printf("%s", stm.str().c_str());
}

//////////////////////////////////////////////////////////////////////////////
  
bool 
GetWindowsErrorText(std::string& text)
{
    DWORD dw = GetLastError();
    LPTSTR lpMsgBuf;

    FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | 
        FORMAT_MESSAGE_FROM_SYSTEM |
        FORMAT_MESSAGE_IGNORE_INSERTS |
        FORMAT_MESSAGE_MAX_WIDTH_MASK , // should suppress newline at end
        NULL,
        dw,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPTSTR) &lpMsgBuf,
        0, NULL );

    text = lpMsgBuf;
    LocalFree(lpMsgBuf);

    return true;
}

//////////////////////////////////////////////////////////////////////////////

std::string
WcharToASCIIRepr( const std::wstring& w )
{
    unsigned int wlen = w.length();
    std::string s;
    s.reserve(6*wlen + 2); // enough room if all chars are non-ASCII
    char hexBuff[7];

    for (unsigned int i = 0; i < wlen; ++i) {
        if (w[i] <= 0x7f) {
            s.push_back( (char)w[i] );
        } else {
            sprintf(hexBuff, "\\u%4x", (int) w[i]);
            s += hexBuff;
        }
    }
    s.push_back(0);
    return s;
}

//////////////////////////////////////////////////////////////////////////////

bool
SplitPathBasenameExtension( const std::wstring& filename,
                           std::wstring& path,
                           std::wstring& basename,
                           std::wstring& extension )
{
    if (filename.empty()) {
        ERROUT("Empty filename passed in");
        return false;
    }

    //
    // Separate the basename and path, looking for special cases
    //

    std::wstring::size_type found = filename.find_last_of(L"/\\");
    // Trailing directory separator = no basename (or extension)
    if (found == filename.length() - 1) {
        path = filename;
        basename = L"";
        extension = L"";
        return true;
    }

    // Does filename hae no preceding path?
    if (found == std::wstring::npos) {
        path = L"";
        basename = filename;
    } else { // Path is present
        path = filename.substr(0, found + 1);
        basename = filename.substr(found + 1);
    }

    //
    // Parse extension from basename
    //
    found = basename.rfind(L".");

    // Leading period, trailing period, or no period = no extension
    if ( found == 0                      || 
         found == basename.length() - 1  ||
         found == std::wstring::npos ) {
        extension = L"";
        return true;
    }

    extension = basename.substr(found + 1);
    basename.resize(found);
    return true;
}

//////////////////////////////////////////////////////////////////////////////
//
// Code to convert FROM Excel TO Python
// 

namespace {
    
    // Excel error codes taken from xlcall32, Microsoft's Excel C API header file
    const char* ExcelTextError(int excelCode) 
    {
        switch(excelCode) {
            case 0: // xlerrNull
                return "#NULL!";
            case 7: // xlerrDiv0
                return "#DIV/0!";
            case 15: // xlerrValue
                return "#VALUE!";
            case 23: // xlerrRef
                return "#REF!";
            case 29: // xlerrName
                return "#NAME?";
            case 36: // xlerrNum
                return "#NUM!";
            case 42: //xlerrNA
                return "#N/A";
        }
        return "Unrecognized Excel error code";
    }     

    bool
    ConvertCellValueToPyObject( const CellValue& rCV,
        PyObject*& rpObj )
    {
        rpObj = NULL;

        if (rCV.IsAString()) { // Strings from pre-2007 Excel; wstrings from 2007+

#if PY_MAJOR_VERSION < 3
            // Excel 2002 or 2003, calling into Python 2.x
            rpObj = PyString_FromString( rCV.StringValue().c_str() );                
#else 
            // Excel 2002 or 2003, calling into Python 3.x
            rpObj = PyUnicode_FromString( rCV.StringValue().c_str() );
#endif
            return (rpObj != NULL);
        }
        
        if (rCV.IsAWstring()) {
            const std::wstring& rS = rCV.WstringValue();

#if PY_MAJOR_VERSION < 3
            // Excel 2007 calling into Python 2.x
            //
            // Unicode handling in pre-3.0 Python is wonky. Make Unicode-encoded 
            // ASCII into char strings wherever possible, so that Python code will 
            // be simpler.
            //
            // I tried IsTextUnicode() first, but that wasn't reliable - it does some 
            // statistical testing. We actually know that the text coming in is Unicode,
            // and the only question is whether or not it is representable as ASCII. That 
            // test is quite trivial; just make sure that each character code is <= 0x7F.

            bool bAllASCII = true;
            std::string tmp;
            unsigned long len = rS.length();
            tmp.resize(len);
            for(unsigned long i = 0; i < len; ++i) {
                if (rS[i] > 0x7F) {
                    bAllASCII = false;
                    break;
                }
                tmp[i]= rS[i];
            }
   
            if (bAllASCII) {
                rpObj = PyString_FromString( tmp.c_str() );  
            } else {
                rpObj = PyUnicode_FromWideChar( rS.c_str(), rS.length() );
            }
#else 
            // Excel 2007 calling into Python 3.x
            rpObj = PyUnicode_FromWideChar( rS.c_str(), rS.length() );
#endif

            return (rpObj != NULL);
        }

        if (rCV.IsANumber()) {
            rpObj = PyFloat_FromDouble( rCV.NumericValue() );
            return (rpObj != NULL); 
        }

        if (rCV.IsBoolean()) {
            rpObj = PyBool_FromLong( (long) rCV.BooleanValue() );
            return (rpObj != NULL); 
        }

        if (rCV.IsXlfOper()) {
            return false; // Should never happen in Excel
        }

        if (rCV.IsError()) {
#if PY_MAJOR_VERSION < 3
            rpObj = PyString_FromString( ExcelTextError(rCV.ErrorValue()) );                
#else 
            rpObj = PyUnicode_FromString( ExcelTextError(rCV.ErrorValue()) );                 
#endif
            return (rpObj != NULL);
        }

        if (rCV.IsEmpty()) {
            rpObj = Py_None;
            Py_INCREF(rpObj); // C API manual says ref counting is required for None
            return true;
        }

        return true;
    }

}

//////////////////////////////////////////////////////////////////////////////

bool
ConvertCellMatrixToPyObject( const CellMatrix& rCM,
                             PyObject*& rpObj )
{
    bool rc = true;
    long i, j, cols, rows;
    PyObject *pCols, *pValue;

    rows = rCM.RowsInStructure();
    cols = rCM.ColumnsInStructure();
    assert(rows && cols);

    if (rows == 1 && cols == 1) {
        // Don't build a nested tuple; extract the single value
        const CellValue& rCV = rCM(0, 0);
        rc = ConvertCellValueToPyObject( rCV, rpObj );
        if (rc) {
            assert(rpObj);
        } else {
            assert(!rpObj);
            ERROUT("Failed to convert single cell to PyObject");
            rc = false;
        }
    } else if (rows == 1) { // Single horizontal rows should NOT be double-nested; they're vectors, not matrices
        assert(cols > 1);
        rpObj = PyTuple_New(cols); // Just ONE ROW here, with cols # of elements
        for (j = 0; rc && j < cols; ++j) {
            const CellValue& rCV = rCM(0, j);
            rc = ConvertCellValueToPyObject( rCV, pValue );
            if (rc) {
                assert(pValue);
                PyTuple_SetItem(rpObj, j, pValue); // pValue reference stolen here
            } else {
                assert(!pValue);
                ERROUT("Failed to convert element %d of single-row cell to PyObject", j);
                Py_DECREF(rpObj);   // Get rid of the entire row; it owns elements and will delete them
                rpObj = NULL;
                rc = false;
            }
        } // end j
    } else { // 2-D matrix
        rpObj = PyTuple_New(rows);
        for (i = 0; rc && i < rows; ++i) {
            pCols = PyTuple_New(cols);
            PyTuple_SetItem(rpObj, i, pCols); // pCols reference stolen here
            for (j = 0; rc && j < cols; ++j) {
                const CellValue& rCV = rCM(i, j);
                rc = ConvertCellValueToPyObject( rCV, pValue );
                if (rc) {
                    assert(pValue);
                    PyTuple_SetItem(pCols, j, pValue); // pValue reference stolen here
                } else {
                    assert(!pValue);
                    ERROUT("Failed to convert element %d, %d of matrix cell to PyObject", i, j);
                    Py_DECREF(rpObj);   // Get rid of the entire matrix; it owns elements and will delete them
                    rpObj = NULL;
                    rc = false;
                }
            } // end j
        } // end i
    } // end 2-D matrix code

    return rc;
}

//////////////////////////////////////////////////////////////////////////////
//
// Code to convert back FROM Python TO Excel
// 

namespace {

    bool
    ConvertPyObjectToCellValue( PyObject* pObj, CellValue& rVal ) 
    {
        const char* pTypeName = Py_TYPE(pObj)->tp_name;
        if (PyBool_Check(pObj)) {
            if (pObj == Py_True) {
                rVal = CellValue(true);
            } else {
                assert(pObj == Py_False);   
                rVal = CellValue(false);
            }
#if PY_MAJOR_VERSION < 3
        // No separate int object in 3.x; just long
        } else if (PyInt_Check(pObj)) {
            // Excel only takes doubles; CellValue c-tors are
            // sufficiently confusing that we want to be sure we're
            // passing a number and not an error code
            rVal = CellValue((double)PyInt_AsLong(pObj));
#endif
        } else if (PyLong_Check(pObj)) {
            rVal = CellValue(PyLong_AsDouble(pObj));
        } else if (PyFloat_Check(pObj)) {
            rVal = CellValue(PyFloat_AsDouble(pObj));
#if PY_MAJOR_VERSION < 3
        // No separate string object in 3.x; just unicode
        } else if (PyString_Check(pObj)) {
            rVal = CellValue(PyString_AsString(pObj));
#endif
        } else if (PyUnicode_Check(pObj)) {
            if (XlfExcel::Instance().excel12()) { 
                // Excel 2007 can handle Unicode directly - just convert to wchar_t
                Py_ssize_t len = PyUnicode_GetSize(pObj);
                std::wstring tmp;
                // Unclear if PyUnicode_AsWideChar copies in a null terminator, but the wstring
                // should do it internally. As an experiment, I tried adding in one deliberately, 
                // and it returned a garbage last character to Excel.
                tmp.resize(len);  
                PyUnicode_AsWideChar((PyUnicodeObject*)pObj, (wchar_t*) tmp.c_str(), len);
                rVal = CellValue(tmp);
            } else { 
                // Excel 2002/2003 can't handle Unicode internally; only deal with Unicode
                // strings that can be converted to ASCII
                PyObject* pAS = PyUnicode_AsASCIIString(pObj);
                if (!pAS) {
                    // Error handling here was tricky. Reading the Python docs closely: pAS is NULL when an exception has
                    // been set during the attempted conversion. If we don't handle that exception, it hangs around and
                    // screws up Python garbage collection.
                    assert(PyErr_Occurred());
                    if (PyErr_ExceptionMatches(PyExc_UnicodeEncodeError)) {
                        PyErr_Clear();
                        rVal = CellValue("Unicode string returned from Python - this version of Excel can't handle Unicode");
                    } else {
                        PyErr_Print();
                    }
                } else {
                    if (PyBytes_Check(pAS)) {
                        const char* pS = PyBytes_AsString(pAS);
                        if (pS) {
                            rVal = CellValue(pS);
                        } else {
                            // This should never happen
                            rVal = CellValue("ASCII string did not contain a valid char pointer");
                        }
                    } else {
                        // This should never happen
                        rVal = CellValue("ASCII string was not of type 'bytes'");
                    }
                    Py_DECREF(pAS);
                }
            }
        } else if (pObj == Py_None) {
            rVal = CellValue();
        } else {    
            ERROUT("Returned data structure contains type %s; this can't be returned to Excel", pTypeName);
            return false;
        }
        return true;
    }

    inline bool PyObjIsAnElementalType(const PyObject* pObj) 
    {
        return  ( PyBool_Check(pObj) || PyLong_Check(pObj) || PyFloat_Check(pObj) ||
                  PyUnicode_Check(pObj) || pObj == Py_None
                 // All ints are long in 3.0, and all strings are Unicode. Look for the older
                 // int and string types in pre-3.0 versions.
#if PY_MAJOR_VERSION < 3
                 || PyInt_Check(pObj)|| PyString_Check(pObj));
#else 
                 );
#endif
    }

    inline bool PyObjIsATupleOrList(const PyObject* pObj) 
    {
        return  (PyTuple_Check(pObj) || PyList_Check(pObj));
    }

    // One-dimensional sequence is an elemental type, or tuple or list that contains only elemental types
    inline bool
    PyObjIsOneDimensional( PyObject* pObj, bool& rOneD ) 
    {
        assert(pObj);
        rOneD = true;
        if (PyObjIsAnElementalType(pObj)) {
            return true;
        }

        // Could be something we don't handle (like a dictionary)
        if (!PyObjIsATupleOrList(pObj)) {
            ERROUT("PyObj of type %s found; we don't handle these", Py_TYPE(pObj)->tp_name);
            return false;
        }
        
        int len = (int)PyObject_Length(pObj);
        PyObject* pElem;
        for (int i = 0; rOneD && i < len; ++i) {
            pElem = PySequence_GetItem(pObj, i); // returns a new reference; must decrement
            if (pElem) {
                rOneD = PyObjIsAnElementalType(pElem);
                Py_DECREF(pElem); 
            } else {
                ERROUT("Couldn't extract element %d from the top-level PyObject", i);
                rOneD = false;
                return false;
            }
        }        

        return true;
    }

    inline bool
    OneDimensionalPyObjToCellMatrix( PyObject* pObj, CellMatrix& rMat ) 
    {
        bool rc = true;
        assert(pObj);
    
        if (PyObjIsAnElementalType(pObj)) {
            rMat = CellMatrix(1,1);
            if (!ConvertPyObjectToCellValue(pObj, rMat(0,0))) {
                ERROUT("Couldn't convert top-level singleton PyObject to cell value");
                return false;
            }    
            return true;
        }

        assert(PyObjIsATupleOrList(pObj));

        int len = (int)PyObject_Length(pObj);
        rMat = CellMatrix(1,len);
        PyObject* pElem;
        for (int i = 0; rc && i < len; ++i) {
            pElem = PySequence_GetItem(pObj, i); // returns a new reference
            // Should be no need to check for pElem != NULL; was previously checked 
            // in an outer 1-D test
            assert(pElem);
            assert(PyObjIsAnElementalType(pElem));
            CellValue& rCV = rMat(0,i);
            rc = ConvertPyObjectToCellValue(pElem, rCV);
            Py_DECREF(pElem); 
        }
        return rc;
    }
}

//////////////////////////////////////////////////////////////////////////////

bool
ConvertPyObjectToCellMatrix( PyObject* pTopObj, CellMatrix& rMat ) 
{
    if (!pTopObj) {
        ERROUT("Top-level object is NULL");
        return false;
    }

    // PyTypeObject* pTP = pTopObj->ob_type;
    const char* pTypeName = Py_TYPE(pTopObj)->tp_name;
    
    // Single element should be returned as such
    if (PyObjIsAnElementalType(pTopObj)) {
        rMat = CellMatrix(1,1);
        if (!ConvertPyObjectToCellValue(pTopObj, rMat(0,0))) {
            ERROUT("Couldn't convert top-level singleton PyObject to cell value");
            return false;
        }    
        return true;
    } 
    
    // Not a single element; either it's a sequence we can deal with (list or tuple),
    // or something else that we can't handle
    if ( !PyObjIsATupleOrList(pTopObj) ) {
        ERROUT("Couldn't convert top-level type (%s)", pTypeName);
        return false;
    }

    // Is it a single row?    
    bool oneD;
    if (!PyObjIsOneDimensional(pTopObj, oneD)) {
        ERROUT("Couldn't determine if top-level PyObj sequence is one-dimensional");
        return false; 
    }

    bool rc = true;
    if (oneD) {
        if (!OneDimensionalPyObjToCellMatrix( pTopObj, rMat )) {
            ERROUT("Couldn't convert top-level one-dimensional PyObj");
            return false; 
        }
    } else { // it's a 2-D matrix
        int topObjLen = (int)PyObject_Length(pTopObj);
        bool rowOneD;
        for (int i = 0; rc && i < topObjLen; ++i) {
            PyObject* pRowObj = PySequence_GetItem(pTopObj, i); // returns a new reference
            // Should be no need to check for pRowObj != NULL; was just checked above in the oneD test loop
            assert(pRowObj);
            if (!PyObjIsOneDimensional(pRowObj, rowOneD)) {
                ERROUT("Couldn't determine if row %d of top PyObj sequence is one-dimensional", i);
                rc = false;
            }

            if (rc  && !rowOneD) {
                ERROUT("Row %d of top PyObj sequence is not one-dimensional", i);
                rc = false;
            }    

            CellMatrix rowMat;
            if (rc && !OneDimensionalPyObjToCellMatrix( pRowObj, rowMat )) {
                ERROUT("Couldn't convert row %d of PyObj", i);
                rc = false;                
            }

            if (rc) {
                rMat.PushBottom(rowMat);
            }
            Py_DECREF(pRowObj);
        }
    }

    return rc;
}           

//////////////////////////////////////////////////////////////////////////////

void
CellMatrixDump( xlw::CellMatrix& rMat )
{
    unsigned long rows = rMat.RowsInStructure();
    unsigned long cols = rMat.ColumnsInStructure();
    unsigned long i, j;

    printf("\nCellMatrixDump:\n");
    for (i = 0; i < rows; ++i) {
        for (j = 0; j < cols; ++j) {
            if (j > 0) printf(",");
            CellValue& rCV = rMat(i, j);

            if (rCV.IsAString()) {
                printf("%s", rCV.StringValue().c_str());
            }
            if (rCV.IsAWstring()) {
                printf("%ls", rCV.WstringValue().c_str());  // l is a MSFT extension
            }
            if (rCV.IsANumber()) {
                printf("%g", rCV.NumericValue());
            }
            if (rCV.IsBoolean()) {
                printf("%s", rCV.BooleanValue()? "True" : "False");
            }
            if (rCV.IsXlfOper()) {
                printf("XLOPER?"); // Should never happen in Excel
            }
            if (rCV.IsError()) {
                printf("%s", ExcelTextError(rCV.ErrorValue()) );
            }
            if (rCV.IsEmpty()) {
                printf("(Empty)");
            }
        }
        printf("\n");
    }
    printf("\n");
}

//////////////////////////////////////////////////////////////////////////////