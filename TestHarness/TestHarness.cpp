// $Id: TestHarness.cpp 182 2010-01-19 07:04:18Z Ross $
 
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

// The purpose of this project is to serve as an area for rapid testing of ideas,
// rather than as a traditional test suite (as the name may imply). Running Excel
// from the debugger is much slower than testing in a small console application, 
// such as this one. Typically, I develop an idea in this file and then move it
// to Utils, so that it'll be accessible to the main Pyinex project.

//////////////////////////////////////////

void TestPathSplitter()
{
    const int numTests = 11;
    wchar_t* pTestCases[numTests][4] = {
        { L"C:\\Foo\\Bar.exe", L"C:\\Foo\\", L"Bar", L"exe"},
        { L"C:/Foo/Bar.exe", L"C:/Foo/", L"Bar", L"exe" },
        { L"C:/Foo/Bar/", L"C:/Foo/Bar/", L"", L"" },
        { L"C:\\Foo\\Bar.exe\\",  L"C:\\Foo\\Bar.exe\\", L"", L"" },
        { L"C:\\Foo\\Bar.",  L"C:\\Foo\\", L"Bar.", L"" },
        { L"C:\\Foo\\Bar",  L"C:\\Foo\\", L"Bar", L"" },
        { L"Foo.exe",  L"", L"Foo", L"exe" },
        { L"Foo",  L"", L"Foo", L"" },
        { L".exe",  L"", L".exe", L"" },
        { L"C:\\Foo\\.exe",  L"C:\\Foo\\", L".exe", L"" },
        { L"C:\\", L"C:\\", L"", L"" } 
    };

    for (int i = 0; i < numTests; ++i) {
        std::wstring path, basename, extension;
        if (! SplitPathBasenameExtension(pTestCases[i][0], path, basename, extension)) {
            printf("Path splitting failure on test %d\n", i);
            continue;
        }
        if (path != pTestCases[i][1]) {
            wprintf(L"Split %d failed (path): %s desired, %s returned\n", i, pTestCases[i][1], path.c_str());
        }
        if (basename != pTestCases[i][2]) {
            wprintf(L"Split %d failed (basename): %s desired, %s returned\n", i, pTestCases[i][2], basename.c_str());
        }
        if (extension != pTestCases[i][3]) {
            wprintf(L"Split %d failed (ext): %s desired, %s returned\n", i, pTestCases[i][3], extension.c_str());
        }
        wprintf(L"%-20s%-20s%-20s%-6s\n", pTestCases[i][0], path.c_str(), basename.c_str(), extension.c_str());
    }
}

//////////////////////////////////////////

bool CallPythonFunction( std::wstring& filename, std::string& function )
{   
    PyObject* pModule = NULL, *pFunc = NULL;
    bool rc = 
    GetPyModuleAndFunctionObjects(  filename, function, pModule, pFunc);
    if (!rc) return false;

    for (int i = 0; i < 3; ++i) {

        PyObject* pArg = PyTuple_New(1); 
#if PY_MAJOR_VERSION < 3       
        PyObject* pInt = PyInt_FromSize_t(i);
#else
        PyObject* pInt = PyLong_FromSize_t(i);
#endif
        PyTuple_SetItem(pArg, 0, pInt); // steals ref

        PyObject* pValue = PyObject_CallObject(pFunc, pArg);
        Py_DECREF(pArg); // decrements the contained int

        CellMatrix cm;
        if (pValue != NULL) {
            ConvertPyObjectToCellMatrix(pValue, cm);
            CellMatrixDump(cm);
            Py_DECREF(pValue);
        }
        else {
            PyErr_Print();
        }
    }

    Py_DECREF(pFunc);
    return true;
}

//////////////////////////////////////////

static PyObject *
pyinex_Thousand(PyObject *self, PyObject *args)
{
    return Py_BuildValue("i", 1000);
}

static PyMethodDef PyinexMethods[] = {
    {"Thousand", pyinex_Thousand, METH_VARARGS, "Return 1000"},
    {NULL, NULL, 0, NULL} /* Sentinel */
};

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
    PyObject* pModule;
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

//////////////////////////////////////////

int _tmain(int argc, _TCHAR* argv[])
{
    PyImport_AppendInittab("pyinex", PyInit_pyinex);

#if PY_MAJOR_VERSION < 3    
    Py_SetProgramName("Excel");
#else
    Py_SetProgramName(L"Excel");
#endif

    // Initialize the Python interpreter.  Required.
    Py_Initialize();
    bool rc = SetupOutputStreams();

    PyImport_ImportModule("pyinex");

    int n;
    while(true) {
        rc = CallPythonFunction( std::wstring(L"..\\Examples\\PyinexTest.py"), std::string("TestHarnessFunc") );
        printf("Enter 9 to stop: \n");
        scanf_s("%d", &n);
        if (n == 9) break;
    }

    Py_Finalize();
    return 0;
}