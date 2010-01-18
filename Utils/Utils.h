// $Id: Utils.h 182 2010-01-19 07:04:18Z Ross $

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

#pragma once

#include "PythonWrapper.h"
#include <string>
#include <vector>
#include <map>

namespace xlw {
    class CellValue;
    class CellMatrix;
}


// Severity codes are bitmasks; it will mildly simplify arbitrary output filtering 

enum pyxErrorSeverity
{
    pyxError    = 0x1,
    pyxWarning  = 0x2,
    pyxInfo     = 0x4
};

// PrintError adds a newline at the end of each logged message; don't put them in msg

void
PrintError(  pyxErrorSeverity severity,  
             const char* file,
             const char* function,
             int line,
             const char* msg, // C-style printf formatting string
             ... );           // C-style set of vars to print

#define ERROUT(x, ...)  (PrintError( pyxError,   __FILE__, __FUNCTION__, __LINE__, x, __VA_ARGS__ ))
#define WARNOUT(x, ...) (PrintError( pyxWarning, __FILE__, __FUNCTION__, __LINE__, x, __VA_ARGS__ ))
#define INFOUT(x, ...)  (PrintError( pyxInfo,    __FILE__, __FUNCTION__, __LINE__, x, __VA_ARGS__ ))

// Utility routine to facilitate printing of wstrings that may contain non-ASCII chars

std::string
WcharToASCIIRepr( const std::wstring& w );

#define ASCII_REPR( x ) (WcharToASCIIRepr(x).c_str())

// WinDef.h's MAX_PATH only applies to ASCII paths. MSDN docs say that wide-string verisons of file
// manipulation functions can handle paths approximately 32767 wchars long, but this doesn't seem
// to be #define'd anywhere in the Windows headers (I can't find it, anyway) 

#define UNICODE_MAX_PATH (0x7FFF)

// Nice macro stolen from Kernighan and Pike's "The Practice Of Programming"
#define NELEMS(array) ( sizeof(array) / sizeof((array[0])) )

// Multiple parts of the code need to split a filename's path and basename. This is 
// doable with wsplitpath, in the CRT, but the MSDN docs suggests that it can't deal
// with the potentially very long paths seen when using wchar_t's (i.e., it maxes
// out at MAX_PATH, and we will see things up to UNICODE_MAX_PATH). So, we roll
// our own. Some details of the contract:
//
// Splits on either / or \.
// Filenames with trailing periods incorporate the period into the basename. Extension is empty.
// Path has a trailing / or \ (whichever is there).
// Basename DOES NOT have the extension; that's split into ext.
//
// For all of the corner cases, look at the test code in TestHarness.cpp.

bool
SplitPathBasenameExtension( const std::wstring& filename,
                           std::wstring& path,
                           std::wstring& basename,
                           std::wstring& extension );

// Different versions of python are compiled/linked against different CRTs, so the
// loaded python DLL will potentially be calling printf through CRT handles that
// are different than those our XLL will be using. In particular, python 2.5
// uses the msvcr7.1 dlls, while this XLL is compiled against msvcr9.0 dlls
// (because it's developed in Visual Studio 2008). To make console output work, 
// we need to iterate through all loaded CRT dlls, extract a particular set of
// function pointers and global vars, and do some magical MSFT-prescribed setup.

bool
SetupOutputStreams();

// Returns a vector of bools specifying if the env var was originally set. 
// This is passed back to the next function, so it can reset the env vars.
// We don't return the values because we don't overwrite any already-set vals;
// just need to know what to unset. TRUE = value was already set.
//
// This hack is needed to set/reset PYTHONCASEOK, so we can handle filenames
// with non-ASCII characters.

bool
SetEnvVarsIfUnset( const char* pEnvVar, std::map<std::wstring, bool>& crtSettings );

bool
ResetEnvVars( const char* pEnvVar, std::map<std::wstring, bool>& crtSettings );

// Used to display names of loaded Python dll and Pyinex XLL
//
bool
GetLoadedModuleNames( std::vector<std::wstring>& vecNames );

// Translate error code of GetLastError(), which it calls internally
//
bool 
GetWindowsErrorText(std::string& text);

// DO NOT decrement the module reference; it's cached by this function
// DO decrement the function reference; it must be disposed
//
bool 
GetPyModuleAndFunctionObjects(  const std::wstring& filename, 
                                const std::string& function,
                                PyObject*& rpModule,
                                PyObject*& rpFunction );

// Get/set flag that turns on checking of module file write times and reloads stale modules
bool
ModuleFreshnessCheckEnabled();

void
SetModuleFreshnessCheck( bool bCheck );

//
bool
ConvertCellMatrixToPyObject( const xlw::CellMatrix& rCM,
                             PyObject*& rpObj );

// Would like for pObj to be const, but Python headers make that
// impossible (too many internal functions take a non-const ptr)
//
bool
ConvertPyObjectToCellMatrix( PyObject* pObj, 
                             xlw::CellMatrix& rMat );

// Diagnostic use only
//
void
CellMatrixDump( xlw::CellMatrix& rMat );