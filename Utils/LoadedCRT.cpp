// $Id: LoadedCRT.cpp 182 2010-01-19 07:04:18Z Ross $

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

//////////////////////////////////////////////////////////////////////////////

namespace 
{    
    // Function pointers that we'll be calling in each loaded CRT
    //
    // There's also an _wfdopen, but the w only refers to the second param.
    // Behavior at the OS level is the same, so there's no need to get/use both.

    typedef int (*FPprintf)(const char *format, ...);
    typedef int (*FP_open_osfhandle) ( intptr_t osfhandle, int flags );
    typedef FILE* (*FP_fdopen)( int fd, const char *mode );
    typedef int (*FP_setvbuf)( FILE *stream, char *buffer, int mode, size_t size );
    typedef char* (*FPgetenv)(const char* _VarName);
    typedef int (*FP_putenv)(const char* _EnvString);

    // This function should return the CRT _iobuf array that holds stdin, stdout, and stderr
    // MSVC++ 6.0 CRT headers call it __p__iob; the 9.0 headers call it __iob_func. Experimentally,
    // both seem to work even on 9.0 CRT DLLs, so I'll call the old name in an attempt at backwards
    // compatability

    typedef FILE* (*FP__p__iob)(void);


    //////////////////////////////////////////////////////////////////////////////
    
    typedef std::pair<std::wstring, HMODULE> NameModulePair;

    bool 
    GetLoadedModules(std::vector<NameModulePair>& vecNameModulePairs)
    {
        DWORD procID = GetCurrentProcessId(); // no error state, according to the docs

        // MSFT docs aren't clear, but example on the web show that these two permissions
        // are needed to be able to enumerate process modules
        std::string errTxt;    
        HANDLE hProc = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, procID);
        if (!hProc) {
            GetWindowsErrorText(errTxt);
            ERROUT("OpenProcess failed: %s", errTxt.c_str());
            return false;
        }

        // MSFT docs say "don't call CloseHandle on any of the returned module handles"
        //
        // Very unlikely that we'll exceed 1024 modules, but check just in case
        const DWORD moduleArraySize = 1024;
        HMODULE hModules[moduleArraySize];
        DWORD numBytesNeeded;
        DWORD numModules;
        EnumProcessModules( hProc, hModules, sizeof(hModules), &numBytesNeeded);
        numModules = numBytesNeeded/sizeof(HMODULE);
        if (numModules > moduleArraySize) {
            ERROUT("%d modules loaded; exceeds %d allocated. Increase array size and recompile.", numModules, moduleArraySize);
            return false;
        }   

        vecNameModulePairs.reserve(numModules);
        wchar_t wszName[UNICODE_MAX_PATH];
        for (DWORD i = 0; i < numModules; ++i) { 
            if (!GetModuleFileNameW( hModules[i], wszName, NELEMS(wszName) )) {
                GetWindowsErrorText(errTxt);
                ERROUT("GetModuleFileNameEx failed on handle %d: %s", hModules[i], errTxt.c_str());
                swprintf(wszName, UNICODE_MAX_PATH, L"Unknown filename; module handle %d", hModules[i]);
            }

            NameModulePair p(wszName, hModules[i]);
            vecNameModulePairs.push_back(p);
        }

        CloseHandle(hProc);
        return true;
    }   

    //////////////////////////////////////////////////////////////////////////////

    class LoadedCRT 
    {
    public:
        static bool GetCRTs( std::vector<LoadedCRT>& vecCRT );

        LoadedCRT( std::wstring& rFN, HMODULE hM );
        const std::wstring& Filename() const;

        bool SetupOutputStreams();

        char* getenv( const char* _VarName );
        int putenv( const char* _EnvString );

    private:
        template<typename T> bool SetFunctionPointer( const char* functionName, T& rpFunc );
        bool SetPointers();
        bool SetOneStream( const char* handleName, DWORD handleType, FILE* pGlobalStream );


        bool IsValidCRT() const;

    private:
        std::wstring m_filename;
        HMODULE m_hModule;

        FPprintf            p_printf;
        FP_open_osfhandle   p_open_osfhandle;
        FP_fdopen           p_fdopen;
        FP_setvbuf          p_setvbuf;
        FPgetenv            p_getenv;
        FP_putenv           p_putenv;
        FP__p__iob          p__p__iob;

        FILE* m_stdout;
        FILE* m_stderr; 
    };

//////////////////////////////////////////////////////////////////////////////

    LoadedCRT::LoadedCRT(std::wstring& rFN, HMODULE hM) : 
             m_filename(rFN), 
             m_hModule(hM),

             p_printf(NULL), 
             p_open_osfhandle(NULL), 
             p_fdopen(NULL), 
             p_setvbuf(NULL), 
             p_getenv(NULL), 
             p_putenv(NULL), 
             p__p__iob(NULL), 

             m_stdout(NULL),
             m_stderr(NULL) 
    {
    }

    //////////////////////////////////////////////////////////////////////////////

    bool
    LoadedCRT::SetPointers() 
    {
        // Each sub-call will log adequate errors; no error string needed here
        bool rc = 
            SetFunctionPointer( "printf",          p_printf )         &&
            SetFunctionPointer( "_open_osfhandle", p_open_osfhandle ) &&
            SetFunctionPointer( "_fdopen",         p_fdopen )         &&
            SetFunctionPointer( "setvbuf",         p_setvbuf )        &&
            SetFunctionPointer( "getenv",          p_getenv )         &&
            SetFunctionPointer( "_putenv",         p_putenv )         &&
            SetFunctionPointer( "__p__iob",        p__p__iob );

        if (rc) {
            assert(p__p__iob);
            m_stdout = &(p__p__iob()[1]);
            if (!m_stdout) {
                std::string errTxt;
                GetWindowsErrorText(errTxt);
                ERROUT("Couldn't get stdout from candidate CRT %s; %s", ASCII_REPR(m_filename), errTxt.c_str());
                return false;
            }

            m_stderr = &(p__p__iob()[2]);
            if (!m_stderr) {
                std::string errTxt;
                GetWindowsErrorText(errTxt);
                ERROUT("Couldn't get stderr from candidate CRT %s; %s", ASCII_REPR(m_filename), errTxt.c_str());
                return false;
            }
        }

        return rc;
    }

    //////////////////////////////////////////////////////////////////////////////
    // http://support.microsoft.com/default.aspx?scid=kb;en-us;105305 shows how to
    // redirect stdout and stderr. The only cleverness here is doing it for each
    // loaded CRT, via function pointers extracted from the loaded modules.

    bool 
    LoadedCRT::SetupOutputStreams()
    {
        // Each sub-call will log adequate errors; no error string needed here
        bool rc = SetOneStream("stdout", STD_OUTPUT_HANDLE, m_stdout) &&
                  SetOneStream("stderr", STD_ERROR_HANDLE, m_stderr);
        if (rc) {
            printf("Initialized console for %s\n", ASCII_REPR(m_filename));
        }
        return rc;
    }

    //////////////////////////////////////////////////////////////////////////////

    bool 
    LoadedCRT::IsValidCRT() const
    {
        return ( (!m_filename.empty()) && 
                  m_hModule            &&
 
                  p_printf             &&
                  p_open_osfhandle     && 
                  p_fdopen             && 
                  p_setvbuf            && 
                  p_getenv             && 
                  p_putenv             && 
                  p__p__iob            &&

                  m_stdout             && 
                  m_stderr );
    }

    //////////////////////////////////////////////////////////////////////////////

    const std::wstring&
    LoadedCRT::Filename() const {return m_filename;}

    //////////////////////////////////////////////////////////////////////////////

    char* LoadedCRT::getenv( const char* _VarName )
    {
        if (p_getenv) {
            return p_getenv(_VarName);
        } else {
            return NULL;
        }
    }

    //////////////////////////////////////////////////////////////////////////////

    int LoadedCRT::putenv( const char* _EnvString )
    {
        if (p_putenv) {
            return p_putenv(_EnvString);
        } else {
            return -1;
        }
    }

    //////////////////////////////////////////////////////////////////////////////

    bool 
    LoadedCRT::GetCRTs(std::vector<LoadedCRT>& vecCRT)
    {
        std::vector<NameModulePair> vecNameModulePairs;
        if (!GetLoadedModules(vecNameModulePairs)) {
            ERROUT("GetLoadedModules failed");
            return false;
        }

        std::vector<NameModulePair>::iterator it;
        for (it = vecNameModulePairs.begin(); it != vecNameModulePairs.end(); ++it) {

            std::wstring path, basename, extension;
            if (!SplitPathBasenameExtension( it->first.c_str(), path, basename, extension )) {
                ERROUT("Couldn't split module name %s", ASCII_REPR(it->first));
                continue;
            }

            // CRT dlls should never have non-ASCII chars in them.
            // "msvcr," not "msvcrt," because the version-specific CRTs don't have the "t".
            // I.e., "msvcr90.dll"
            const wchar_t msvcr[] = L"msvcr"; int msvcrLen = NELEMS(msvcr) - 1;
            const wchar_t dll[] = L"dll"; 

            if ( _wcsnicmp( basename.c_str(), msvcr, msvcrLen ) == 0 &&
                _wcsicmp( extension.c_str(), dll ) == 0) 
            {
                // Got a potential CRT; try to dig out the necessary function pointers, and if 
                // they're not there, pass on the candidate
                LoadedCRT crt(it->first, it->second);
                if (crt.SetPointers()) {
                    assert(crt.IsValidCRT());
                    vecCRT.push_back(crt);
                }
            }
        }

        return true;
    }   

    //////////////////////////////////////////////////////////////////////////////

    template<typename T>
    bool LoadedCRT::SetFunctionPointer(const char* functionName, T& rpFunc) 
    {
        rpFunc = (T)GetProcAddress(m_hModule, functionName);
        if (!rpFunc) {
            std::string errTxt;
            GetWindowsErrorText(errTxt);
            ERROUT("Couldn't get %s from candidate CRT %s; %s", functionName, ASCII_REPR(m_filename), errTxt.c_str());
            return false;
        }
        return true;
    }

    //////////////////////////////////////////////////////////////////////////////

    bool
    LoadedCRT::SetOneStream(const char* handleName, DWORD handleType, FILE* pGlobalStream ) 
    {
        assert(pGlobalStream);

        // We probably don't need to call the CRT-specific version of this function on each loaded
        // CRT; the os filehandles seem to be fixed at 3 and 4 for stdout and stderr, respectively, 
        int hCrtFD = p_open_osfhandle(
            (intptr_t) GetStdHandle(handleType),
            _O_TEXT
            );

        if (hCrtFD == -1) {
            ERROUT("Failed to associate CRT file descriptor with os-level %s handle in %s", handleName, ASCII_REPR(m_filename));
            return false;
        }

        FILE* hf = p_fdopen( hCrtFD, "w" );
        if (!hf) {
            ERROUT("Failed to open a new stream to CRT %s file descriptor in %s", handleName, ASCII_REPR(m_filename));
            return false;
        }

        *pGlobalStream = *hf;

        // Turn off buffering
        p_setvbuf( pGlobalStream, NULL, _IONBF, 0 );
        return true;
    }

} // end anonymous namespace

//////////////////////////////////////////////////////////////////////////////
//
// This is broken out as a separate function because exposing LoadedCRT::
// GetLoadedModules would require including windows.h in the utils.h header file,
// plus putting LoadedCRT in there as well. Since this function is not
// going to be called frequently, it's fine to do this vector copying.

bool
GetLoadedModuleNames( std::vector<std::wstring>& vecNames )
{
    std::vector<NameModulePair> vecNameModulePairs;
    if (!GetLoadedModules(vecNameModulePairs)) {
        ERROUT("GetLoadedModules failed");
        return false;
    }

    // Could use a functor + std::transform, but this is less painful/more obvious
    vecNames.reserve(vecNameModulePairs.size());
    std::vector<NameModulePair>::iterator it;
    for (it = vecNameModulePairs.begin(); it != vecNameModulePairs.end(); ++it) {
        vecNames.push_back(it->first);
    }

    return true;
}

//////////////////////////////////////////////////////////////////////////////
//
// I think the load-order dependency is such that all CRTs will be loaded before our XLL
// is loaded. If that supposition is correct and this function is only called from an XLL's 
// DllMain(), it should always work.

bool
SetupOutputStreams()
{
    std::vector<LoadedCRT> vecCRT;
    if (!LoadedCRT::GetCRTs(vecCRT)) {
        ERROUT("LoadedCRT::GetCRTs failed");
        return false;
    }

    bool rc = true;
    std::vector<LoadedCRT>::iterator it;
    for (it = vecCRT.begin(); it != vecCRT.end(); ++it) {
        if (!it->SetupOutputStreams()) {
            ERROUT("SetupOutputStreams() failed on %s", ASCII_REPR(it->Filename()));
            rc = false;
        }
    }

    return rc;
}

//////////////////////////////////////////////////////////////////////////////
//
// Ugly hack. Either we pass back the CRT objects to the caller, so it can set/reset
// the same objects in a well-understood context, or we pass around maps of CRT name to
// last-retrieved variable setting. Neither is a good solution, because there's no way to 
// guarantee that the CRTs will remain loaded between set/reset, or that the variables won't be 
// changed by another thread.
//
// In practice we don't expect this to be a problem, as no one but us is likely to touch 
// PYTHONCASEOK, but it's a hole.
//
// Rather than expose the CRT machinery, I opted for the name->status mapping. This requires
// us to reload and re-iterate the CRTs in each cycle, which is slow and bad, but only expected
// to happen upon module import (which typically happens once per module per process lifetime in
// this application). Thus, it's not likely to be a performance problem.

bool
SetEnvVarsIfUnset( const char* pEnvVar, std::map<std::wstring, bool>& crtSettings ) 
{
    std::vector<LoadedCRT> vecCRT;
    if (!LoadedCRT::GetCRTs(vecCRT)) {
        ERROUT("GetCRTs failed");
        return false;
    }

    bool rc = true, bAlreadySet;
    std::string valToSet(pEnvVar);
    valToSet += "=1";

    std::vector<LoadedCRT>::iterator it;
    for (it = vecCRT.begin(); it != vecCRT.end(); ++it) {
        bAlreadySet = (it->getenv(pEnvVar) != NULL);
        crtSettings[it->Filename()] =  bAlreadySet;
        if (!bAlreadySet) {
            it->putenv(valToSet.c_str());
        }
    }

    return rc;
}

//////////////////////////////////////////////////////////////////////////////

bool
ResetEnvVars( const char* pEnvVar, std::map<std::wstring, bool>& crtSettings )
{
    std::vector<LoadedCRT> vecCRT;
    if (!LoadedCRT::GetCRTs(vecCRT)) {
        ERROUT("GetCRTs failed");
        return false;
    }

    bool rc = true;
    std::string valToSet(pEnvVar), valToReset(pEnvVar);
    valToSet += "=1";
    valToReset += "=";
    std::map<std::wstring, bool>::iterator mapIt;

    std::vector<LoadedCRT>::iterator vecIt;
    for (vecIt = vecCRT.begin(); vecIt != vecCRT.end(); ++vecIt) {
        mapIt = crtSettings.find(vecIt->Filename());
        if (mapIt == crtSettings.end()) {
            ERROUT("%s is new; wasn't seen in the original variable retrieval", ASCII_REPR(vecIt->Filename()));
            continue;
        }

        // For reasons that I don't understand, MSVCRT.DLL corrupts its heap if I try to delete a variable
        // before I've set it. No way to debug without source, which I don't have. Fortunately, overwriting
        // the existing variable causes no problem, as it's immediately deleted thereafter, but...it's
        // not comfort-inducing.
        if (!mapIt->second) {
            vecIt->putenv(valToSet.c_str());                  
            vecIt->putenv(valToReset.c_str());
        }
    }

    return rc;
}

//////////////////////////////////////////////////////////////////////////////