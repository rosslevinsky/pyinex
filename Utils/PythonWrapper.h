// $Id: PythonWrapper.h 182 2010-01-19 07:04:18Z Ross $

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

// If you don't have the python debug dll (python26_d.dll, for version 2.6), 
// and its associated static import library, you must undefine
// PYTHON_DEBUG_BUILD_PRESENT to get a debug Pyinex build, because of a peculiarity
// in the Python.h file (it demands the debug dll be linked whenenver _DEBUG is defined).
//
// The debug dll isn't included in the standard precompiled python distributions
// (neither from python.org nor from ActiveState), so you'll have to get the 
// code and build it yourself (or use the copy included in the Pyinex distribution).
// Don't be anxious about building it yourself - the python.org donwload is very
// easy (and fast).

// Best to define this in the configuration properties; in this release, it's 
// only set for Debug-26

#ifdef PYTHON_DEBUG_BUILD_PRESENT

    #include <Python.h>

#else // no debug build present

// This hack (temporarily undef'ing _DEBUG) will allow a debug build to work.

    #ifdef _DEBUG
        #undef _DEBUG
        #include <Python.h>
        #define _DEBUG
    #else
        #include <Python.h>
    #endif

#endif

//
// Python 2.5 doesn't define the following macros; 2.6 and greater do
//

#ifndef Py_TYPE
    #define Py_TYPE(ob)      (((PyObject*)(ob))->ob_type)
#endif

#ifndef PyBytes_Check
    #define PyBytes_Check    PyString_Check
#endif

// PyBytes_AsString is a function in 3.0; can't redefine it
#if PY_MAJOR_VERSION < 3
    #ifndef PyBytes_AsString 
        #define PyBytes_AsString PyString_AsString
    #endif
#endif
