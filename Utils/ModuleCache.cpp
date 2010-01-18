// $Id: ModuleCache.cpp 182 2010-01-19 07:04:18Z Ross $

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
// 
// Module contains plumbing for experimental use of Windows file change
// notifications, as an attempt at an interrupt-driven means of prompting
// reloads of changed modules. Doesn't look like it's going to work, though,
// because network drives don't always provide the right signals to Windows.
//
// Plumbing was never completed, so merely defining this constant is NOT 
// sufficient to make it work. I got far enough to see that there would
// be notification problems with network drives, and then I stopped.
//
// The overall idea of watching dirs is not crazy, though - we might put in
// place a background thread that periodically polls for changes.

#undef WINDOWS_FILE_CHANGE_NOTIFICATION_EXPERIMENT

//////////////////////////////////////////////////////////////////////////////

namespace {

    bool operator>(const FILETIME& lhs, const FILETIME& rhs) 
    {
        if (lhs.dwHighDateTime > rhs.dwHighDateTime) {
            return true;
        }

        if (lhs.dwHighDateTime == rhs.dwHighDateTime) {
            if (lhs.dwLowDateTime > rhs.dwLowDateTime) {
                return true;
            }
        }

        return false;
    }

    //////////////////////////////////////////////////////////////////////////////

    class CriticalSectionWrapper 
    {
    public:
        CriticalSectionWrapper( CRITICAL_SECTION& rCS ) : m_rCS(rCS) {
            EnterCriticalSection(&m_rCS);  
        }

        ~CriticalSectionWrapper() {
            LeaveCriticalSection(&m_rCS);
        }

    private:
        CRITICAL_SECTION& m_rCS;
    };

    //////////////////////////////////////////////////////////////////////////////

    class ModuleCache
    {
    public:
        static ModuleCache& Factory();
        bool GetModule( const std::wstring& filename, PyObject*& rpModule );
        bool ModuleFreshnessCheckEnabled() const;
        void SetModuleFreshnessCheck( bool bCheck );

#ifdef WINDOWS_FILE_CHANGE_NOTIFICATION_EXPERIMENT
        void DirChanged( void* lpParameter, BOOLEAN TimerOrWaitFired );
#endif

    private:
        // These classes are just dumb containers for file/directory information. They DO NOT manage the lifecycle
        // of the open handles themselves (i.e., don't close handles in d-tors); that's done by ModuleCache.

        struct FileInfo 
        {
            FileInfo() : m_hFile(INVALID_HANDLE_VALUE), m_bClean(false), m_pModule(NULL) 
            {
                // Time members aren't so critical; save time by not nulling
            }
            std::wstring m_filename;
            HANDLE m_hFile;
            FILETIME m_creation;
            FILETIME m_lastAccess;
            FILETIME m_lastWrite;
            bool m_bClean; // true == we've loaded the latest version
            PyObject* m_pModule;
        };

        struct DirInfo 
        {
            DirInfo(): m_hDir(INVALID_HANDLE_VALUE)
            {
#ifdef WINDOWS_FILE_CHANGE_NOTIFICATION_EXPERIMENT 
                m_hDirChangeNotification = 0;
                m_hWaitObject = 0;
#endif
            }
            std::wstring m_dirName;
            HANDLE m_hDir;
#ifdef WINDOWS_FILE_CHANGE_NOTIFICATION_EXPERIMENT 
            HANDLE m_hDirChangeNotification;
            HANDLE m_hWaitObject;
#endif
            std::set<std::wstring> m_setFilesInDir;
        };

    private:
        // Both private to enforce singleton nature of this class
        ModuleCache();
        ~ModuleCache(); 

        bool GetModuleFirstTime( const std::wstring& canonicalFN, PyObject*& rpModule );

        inline static bool IsAllASCII(const std::wstring& w);

        bool GetNewFileInfo( const std::wstring& filename, ModuleCache::FileInfo& rNewInfo ) ;
        bool GetNewDirInfo( const std::wstring& path, 
                            const std::wstring& filename,
                            ModuleCache::DirInfo& rNewInfo );

        // These funcs don't need to be static, as they're now private, but if they ever move
        // to Utils.cpp, this makes it clear that they don't use any member vars in ModuleCache.

        static bool GetFullFilename( const std::wstring& filename, 
                                     bool bLongPath, // false == short path
                                     wchar_t* pFullFilename, // max necessary is size UNICODE_MAX_PATH 
                                     size_t fullFilenameSize,
                                     std::string& errTxt );

        static bool GetFullPathBasenameExtension( const std::wstring& filename, 
                                                  bool bLongPath, // false == short path
                                                  std::wstring& path, 
                                                  std::wstring& basename,
                                                  std::wstring& extension,
                                                  std::string& errTxt );

        static bool GetASCIIFullPathBasenameExtension( const std::wstring& filename, 
                                                       std::wstring& path, 
                                                       std::wstring& basename, 
                                                       std::wstring& extension,
                                                       bool& bBasenameIsShort, 
                                                       std::string& errTxt );

        bool ImportOrReload( const std::wstring& filename,
                             bool bImport,  // if true, import, else reload
                             PyObject*& rpModule );

    private:
        typedef std::map<std::wstring, std::wstring> WstringWstringMap;
        WstringWstringMap m_mapUserFNToCanonicalFN;

        typedef std::map<std::wstring, FileInfo> FileInfoMap;
        typedef std::map<std::wstring, DirInfo> DirInfoMap;
        FileInfoMap       m_mapFileInfo;
        DirInfoMap        m_mapDirInfo;

        bool m_bModuleFreshnessCheck;

        // Used by callback function to find DirInfo from dir file handle
        typedef std::map<HANDLE, std::wstring> DirNameMap;

#ifdef WINDOWS_FILE_CHANGE_NOTIFICATION_EXPERIMENT 
        DirNameMap m_mapDirNames;
#endif 
        CRITICAL_SECTION m_cs;
    };

    //////////////////////////////////////////////////////////////////////////////

    ModuleCache& 
    ModuleCache::Factory()
    {
        static ModuleCache f;
        return f;
    }

    //////////////////////////////////////////////////////////////////////////////

 #ifdef  WINDOWS_FILE_CHANGE_NOTIFICATION_EXPERIMENT 

    VOID CALLBACK 
    DirChangeCallback ( PVOID lpParameter,
                        BOOLEAN TimerOrWaitFired )
    {
        ModuleCache& rMC = ModuleCache::Factory();
        rMC.DirChanged(lpParameter, TimerOrWaitFired);
    }

    void 
    ModuleCache::DirChanged( void* lpParameter, BOOLEAN TimerOrWaitFired )
    {
        CriticalSectionWrapper csWrapper(m_cs);  // exception-safe; exits CS in d-tor

        printf("Here: %d %d %d\n", (int)lpParameter, (int)TimerOrWaitFired, GetCurrentThreadId());
        std::wstring& dirName = m_mapDirNames[(HANDLE) lpParameter];
        if (dirName.empty()) {
            ERROUT("Dir file handle %d doesn't map to a known directory name", (int)lpParameter);
        } else {
            DirInfo& rDi = m_mapDirInfo[ dirName ];
            wprintf(L"Dir: %s\n", rDi.m_dirName.c_str());
            BOOL rc2 = FindNextChangeNotification(rDi.m_hDirChangeNotification);
            BOOL rc3 = RegisterWaitForSingleObject( &rDi.m_hWaitObject, rDi.m_hDirChangeNotification, DirChangeCallback, (PVOID)rDi.m_hDir, INFINITE, WT_EXECUTEDEFAULT | WT_EXECUTEONLYONCE );
        }
    }

#endif
    
    //////////////////////////////////////////////////////////////////////////////

    bool 
    ModuleCache::ModuleFreshnessCheckEnabled() const
    {
        return m_bModuleFreshnessCheck;
    }

    //////////////////////////////////////////////////////////////////////////////

    void 
    ModuleCache::SetModuleFreshnessCheck( bool bCheck )
    {
        m_bModuleFreshnessCheck = bCheck;
    } 

    //////////////////////////////////////////////////////////////////////////////
    //
    // This is the only public function that touches anything that
    // needs synchronization; everything that requires locking is
    // called from within here

    bool 
    ModuleCache::GetModule( const std::wstring& filename, // may have relative paths
                            PyObject*& rpModule )
    {
        bool rc = true;
        CriticalSectionWrapper csWrapper(m_cs);  // exception-safe; exits CS in d-tor

        // What's this file's canonical name?
        //
        // Files are registered under their user-specified name, whether that be a short name
        // or a long name. User could perversely pass in the same file under both short and long
        // name (or worse, under some combination of short/long path/basename, which leads to more
        // than two possibilities). This could conceivabley blow up Python, which might wind up loading
        // the same module twice (or something along those lines). The fix is to always refer to files
        // by their full, canonical name internally, which means we have to do some mapping at initial
        // load time.

        std::wstring canonicalFN;
        WstringWstringMap::iterator nameIt = m_mapUserFNToCanonicalFN.find(filename);
        if (nameIt != m_mapUserFNToCanonicalFN.end()) {
            canonicalFN = nameIt->second;
        } else {
            // filename may be a synonym for a canonical filename we've seen before, or 
            // it may refer to a canonical filename we've never seen

            wchar_t pFullFilename[UNICODE_MAX_PATH];
            std::string errTxt;
            rc = GetFullFilename( filename.c_str(), true, pFullFilename, NELEMS(pFullFilename), errTxt ); 
            if (!rc) {
                ERROUT("Couldn't get canonical name of %s: %s", ASCII_REPR(filename), errTxt);
                ERROUT("Does file exist?");
                return false;
            }
            canonicalFN = pFullFilename;
            m_mapUserFNToCanonicalFN[filename] = canonicalFN;
        }

        FileInfoMap::const_iterator fileIt = m_mapFileInfo.find(canonicalFN);
        if (fileIt == m_mapFileInfo.end()) {
            rc = GetModuleFirstTime( canonicalFN, rpModule );
            if (!rc) {
                ERROUT("Failed first load of %s", ASCII_REPR(canonicalFN));
            }
        } else {
            // Already saw this file - check if it has been written since last seen.
            FileInfo oldInfo = fileIt->second;
            FileInfo newInfo = oldInfo; // Copies the all-important file handle and python module pointer

            // Optimization: we can save file operations if we know that nothing has changed (i.e., in a prod environment).
            // User has to explicitly turn off checking from the front end.
            if (!m_bModuleFreshnessCheck) {
                rpModule = oldInfo.m_pModule;
                return rc;
            }

            // XXX See comment about Parallels, above, and the need to use fresh handles on a file. 
            // File handle is opened and closed in this call.
            if ( !ModuleCache::GetNewFileInfo( newInfo.m_filename, newInfo) ) {
                std::string tmp;
                GetWindowsErrorText(tmp);
                ERROUT("GetNewFileInfo on once-loaded file %s failed: %s", ASCII_REPR(canonicalFN), tmp.c_str());
                rc = false;
            }
/*
Doesn't work in Parallels...

            if (GetFileTime(newInfo.m_hFile, &newInfo.m_creation, &newInfo.m_lastAccess, &newInfo.m_lastWrite) != TRUE) {
                std::string tmp;
                GetWindowsErrorText(tmp);
                ERROUT("GetFileTime on already-loaded file %s failed: %s", ASCII_REPR(canonicalFN), tmp.c_str());
                rc = false;
            }
*/
            if (rc) {
                if (newInfo.m_lastWrite > oldInfo.m_lastWrite) {
                    // Later write seen. Get module via Reload and update cache.
                    rc = ImportOrReload( canonicalFN, false, newInfo.m_pModule);
                    if (rc) {
                        rpModule = newInfo.m_pModule;
                        m_mapFileInfo[canonicalFN] = newInfo;
                    } else {
                        ERROUT("Failed reload of %s", ASCII_REPR(canonicalFN));
                    }
                } else {
                    // Unchanged write time. Get module from cache.
                    rpModule = oldInfo.m_pModule;
                }
            }
        } 

        return rc;
    }

    //////////////////////////////////////////////////////////////////////////////

    bool 
    ModuleCache::GetModuleFirstTime( const std::wstring& canonicalFN, 
                                     PyObject*& rpModule )
    {
        bool rc = true;

        // File has never been loaded before; get its initial info
        FileInfo newFileInfo;
        rc = GetNewFileInfo( canonicalFN, newFileInfo );
        if (!rc) {
            ERROUT("Failed to get file info for %s", ASCII_REPR(canonicalFN));
            rc = false;
        }

        // Are we already watching this directory?
        //
        // All Windows systems will have working long path names (can't be turned off)
        // so we'll register dirs under their full canonical names. This requires a call
        // to the OS, but as it's only done once per newly-seen filename (at initial load),
        // we'll take the speed hit.

        std::wstring path;
        if (rc) {
            std::wstring basename, extension;
            rc = SplitPathBasenameExtension(canonicalFN, path, basename, extension);
            if (!rc) {
                ERROUT("Failed to split filename %s", ASCII_REPR(canonicalFN));
            }
        }

        DirInfo newDirInfo;
        bool bNewDir = false;
        DirInfoMap::iterator dirIt;
        if (rc) {
            dirIt = m_mapDirInfo.find(path);
            if (dirIt == m_mapDirInfo.end()) {
                bNewDir = true;
                rc = GetNewDirInfo( path, canonicalFN, newDirInfo );
                if (!rc) {
                    ERROUT("Failed to get dir info for %s", ASCII_REPR(path));
                    rc = false;
                }
            }
        } 
       
        // Import the module
        if (rc) {
            rc = ImportOrReload( canonicalFN, true, newFileInfo.m_pModule );
            if (!rc) {
                ERROUT("Failed initial import of %s", ASCII_REPR(canonicalFN));
            }
        } 

        // Everything has succeeded, so cache the file info and cache/update the dir info
        if (rc) {
            rpModule = newFileInfo.m_pModule;
            newFileInfo.m_bClean = true;
            m_mapFileInfo[canonicalFN] = newFileInfo;
            if (bNewDir) {
                m_mapDirInfo[newDirInfo.m_dirName] = newDirInfo;
#ifdef WINDOWS_FILE_CHANGE_NOTIFICATION_EXPERIMENT 
                m_mapDirNames[newDirInfo.m_hDir] = newDirInfo.m_dirName;
#endif
            } else {
                // It's an existing dir; simply add the canonical filename to it
                dirIt->second.m_setFilesInDir.insert(canonicalFN);
            }
        }

        // Clean up any errors
        if(!rc) {
/* XXX See comments about Parallels. Can't hold file handles open.

            // If there was a failure of initial import, the FileInfo structure won't have been cached,
            // so the file handle won't be cleaned up at ModuleCache destruction. Close it here instead.
            if ( newFileInfo.m_hFile != INVALID_HANDLE_VALUE ) {
                CloseHandle(newFileInfo.m_hFile);
            }
*/
            // Ditto with the directory handle in DirInfo.
            if ( newDirInfo.m_hDir!= INVALID_HANDLE_VALUE ) {
                CloseHandle(newDirInfo.m_hDir);
            }
        }

        return rc;
    }

    //////////////////////////////////////////////////////////////////////////////

    ModuleCache::ModuleCache() 
    {
        // Didn't use spin count because many machines on which this will run
        // are single CPU (no point in spinning). I asssume Windows will avoid
        // this spinning on such machines if I call the non-spincount routine
        // (but don't really know if that's true).
        InitializeCriticalSection(&m_cs); 

        // Sensible default is to always check for module freshness
        m_bModuleFreshnessCheck = true;
    }

    //////////////////////////////////////////////////////////////////////////////

    ModuleCache::~ModuleCache() 
    {
        FileInfoMap::iterator it = m_mapFileInfo.begin();
        while (it != m_mapFileInfo.end()) {
// XXX See comments about Parallels - can't hold handle open
//          CloseHandle(it->second.m_hFile);
            Py_XDECREF(it->second.m_pModule);
            ++it;
        }
        DeleteCriticalSection(&m_cs);
    }

    //////////////////////////////////////////////////////////////////////////////
    //
    // Not wrapped in CS; caller has to lock resources
    //

    bool 
    ModuleCache::GetNewFileInfo( const std::wstring& filename,
                                 ModuleCache::FileInfo& rNewInfo ) 
    {
        // Access flags must be shared read, so Python can load the module while we hold
        // this handle open, and shared write, so we can modify the module while Excel is
        // running.
        //
        // UPDATE: holding the handle open doesn't work. When accessing files on a Mac drive
        // via Parallels 4.0, a held handle apparently caches file access time, which makes it
        // impossible to know when a file has changed. This may be generally true of various
        // network file systems; I haven't tested yet. So, we're forced to close and reopen
        // file handles on each call (if we want to poll). A better solution (background thread,
        // checking file access times every few seconds?) awaits implementation.

        bool rc = true;
        rNewInfo.m_hFile = CreateFileW( filename.c_str(), 
            GENERIC_READ, 
            FILE_SHARE_READ | FILE_SHARE_WRITE, 
            NULL,
            OPEN_EXISTING,
            0, 
            NULL );

        if (rNewInfo.m_hFile == INVALID_HANDLE_VALUE) {
            //  Don't null m_hFile; validity checks in caller will be looking for 
            // INVALID_HANDLE_VALUE, rather than zero
            std::string tmp;
            GetWindowsErrorText(tmp);
            ERROUT("CreateFile on %s failed: %s", ASCII_REPR(filename), tmp.c_str());
            rc = false;
        }

        if (rc && GetFileTime(rNewInfo.m_hFile, &rNewInfo.m_creation, &rNewInfo.m_lastAccess, &rNewInfo.m_lastWrite)!= TRUE) {
            std::string tmp;
            GetWindowsErrorText(tmp);
            ERROUT("GetFileTime on unloaded file %s failed: %s", ASCII_REPR(filename), tmp.c_str());
            rc = false;
        }

        // XXX See comment about Parallels and the need to close file handles, above.
        if (rc) {
            CloseHandle(rNewInfo.m_hFile);
        }
        // Always null this - whether we failed or not, the handle is not to be used outside this function
        rNewInfo.m_hFile = 0; 
        // XXX

        if (rc) {
            rNewInfo.m_filename = filename;
        }

        return rc;
    }

    //////////////////////////////////////////////////////////////////////////////
    //
    // Not wrapped in CS; caller has to lock resources
    //

    bool 
    ModuleCache::GetNewDirInfo( const std::wstring& path,
                                const std::wstring& filename,
                                ModuleCache::DirInfo& rNewInfo ) 
    {
        bool rc = true;
        rNewInfo.m_hDir = CreateFileW( path.c_str(), 
                                       GENERIC_READ, 
                                       FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, 
                                       NULL, 
                                       OPEN_EXISTING,
                                       FILE_FLAG_BACKUP_SEMANTICS,
                                       NULL );

        if (rNewInfo.m_hDir == INVALID_HANDLE_VALUE) {
            //  Don't null m_hDir; validity checks in caller will be looking for 
            // INVALID_HANDLE_VALUE, rather than zero
            std::string tmp;
            GetWindowsErrorText(tmp);
            ERROUT("CreateFile on %s failed: %s", ASCII_REPR(path), tmp.c_str());
            rc = false;
        }

#ifdef WINDOWS_FILE_CHANGE_NOTIFICATION_EXPERIMENT 
        if (rc) {
            // This call fails when trying to get notification from network paths. Error text is "Incorrect function," but there's no
            // more color available online. Obvious assumption is that network drives don't play according to Windows conventions in
            // this respect.
            rNewInfo.m_hDirChangeNotification = FindFirstChangeNotificationW( path.c_str(), FALSE, FILE_NOTIFY_CHANGE_LAST_WRITE );
            if (rNewInfo.m_hDirChangeNotification != INVALID_HANDLE_VALUE) {
                if (0 == RegisterWaitForSingleObject( &rNewInfo.m_hWaitObject, 
                                                      rNewInfo.m_hDirChangeNotification, 
                                                      DirChangeCallback, 
                                                      (PVOID)rNewInfo.m_hDir, 
                                                      INFINITE, 
                                                      WT_EXECUTEONLYONCE) ) 
                {
                    // MSDN docs say no error can be retrieved by GetLastError()
                    ERROUT("RegisterWaitForSingleObject() failed; no reason available");
                    rc = false;
                }
            }
        }
#endif

        if (rc) {
            rNewInfo.m_dirName = path;
            rNewInfo.m_setFilesInDir.insert(filename);
        }

        return rc;
    } 

    //////////////////////////////////////////////////////////////////////////////
    //
    // Python doesn't properly handle the use of Unicode strings in sys.path; it tries to convert them to the
    // system's default code page, which means that any extended characters are lost on a standard ASCII-centric system.
    // I didn't test it, but presumably every other language pair could be broken; Cyrillic paths might not work on a 
    // default-Greek system. See Python's import.c file for the gory details.
    //
    // The only workaround I could find was to use short pathnames in the Python search path, because those seem to always
    // contain only ASCII characters. This isn't a guaranteed fix, because there are ways to turn off the use of short path
    // names on Windows systems. I assume this is typically done for efficiency reasons.
    //
    // Thus, systems that have Unicode characters in their paths must have short pathnames enabled, or this hack will fail.
    //
    // The strategy is to make both the path and the basename ASCII strings, but do the minimal amount of munging necessary to
    // make Python work. If the path is already ASCII and only the basename non-ASCII, we combine the full long path with the 
    // short path's basename. If the path is non-ASCII and the basename ASCII, we use the short path and the long basename. If
    // both are non-ASCII, we use the short path and the short basename.
    //
    // Python currently has a problem with using the short basename, though. import.c's case_ok() only compares the submitted 
    // basename with the long basename; it doesn't bother to check the short basename. This is short-circuitable by setting the 
    // PYTHONCASEOK environment variable, which we do just before trying to import a non-ASCII basename file by its ASCII
    // short basename. We restore he variable when finished, so as not to cause problems for any other applications that
    // depend on this environment variable.
    //
    // Given that we ultimately deal with ASCII paths and basenames, we could pass them to Python as String objects, rather
    // than Unicode objects. We could also handle them in the C++ as std::strings instead of wstrings. Python can handle these
    // Unicode-represented ASCII strings, though, so I use wstrings and String objects throughout. Someday Python will handle
    // non-ASCII paths and basenames correctly, and Pyinex's plumbing will be ready.
    //
    // Efficiency note: Python 2.x will internally downcovert the Unicode strings to MBCS on each import call, so it would be 
    // a bit faster to pass it single-width strings. Pyton 3.x unavoidably downcoverts on each call, as it keeps strings as 
    // Unicode internally.

    inline bool ModuleCache::IsAllASCII(const std::wstring& w) 
    {
        unsigned long len = w.length();
        for(unsigned long i = 0; i < len; ++i) {
            if (w[i] > 0x7F) {
                return false;
            }
        }
        return true;
    }

    //////////////////////////////////////////////////////////////////////////////
    //
    // Retrieves the long/short pathnames, splits them, and checks for ASCII-ness
    //

    bool 
    ModuleCache::GetFullFilename( const std::wstring& filename, 
                                  bool bLongPath, // false == short path
                                  wchar_t* pFullFilename,
                                  size_t fullFilenameSize,
                                  std::string& errTxt )
    {
        // Must call GetFullPathName to take out . and .. shortcuts; it puts together a fully-qualified path.
        // FYI: current dir is a process-global variable, and multiple writer threads can step on it, it can
        // change during process lifetime, etc..

        assert(pFullFilename);
        pFullFilename[0] = 0;
        wchar_t pFullPathName[UNICODE_MAX_PATH];
        wchar_t* pFP;
        DWORD dw = GetFullPathNameW( filename.c_str(), NELEMS(pFullPathName), pFullPathName, &pFP); 
        if (!dw) {
            GetWindowsErrorText(errTxt);
            return false;
        }

        // Get(Long|Short)PathnameW() validates file existence
        if (bLongPath) {
            dw = GetLongPathNameW(  pFullPathName, pFullFilename, fullFilenameSize );
        } else {
            dw = GetShortPathNameW( pFullPathName, pFullFilename, fullFilenameSize );
        }
        if (!dw) {
            GetWindowsErrorText(errTxt);
            return false;
        }

        return true;
    }

    //////////////////////////////////////////////////////////////////////////////
    //
    // Retrieves the long/short pathnames, splits them, and checks for ASCII-ness
    //

    bool 
    ModuleCache::GetFullPathBasenameExtension( const std::wstring& filename, 
                                               bool bLongPath, // false == short path
                                               std::wstring& path,  
                                               std::wstring& basename,
                                               std::wstring& extension,
                                               std::string& errTxt ) 
    {
        wchar_t pFullFilename[UNICODE_MAX_PATH];
        if (!GetFullFilename(filename, bLongPath, pFullFilename, NELEMS(pFullFilename), errTxt)) {
            return false; 
        }
   
        if (!SplitPathBasenameExtension( pFullFilename, path, basename, extension )) {
            errTxt = "Couldn't split filename " + WcharToASCIIRepr(pFullFilename);
            return false;
        }

        return true;
    }

    //////////////////////////////////////////////////////////////////////////////
    // 
    // Tries to preserve as much of the original path and filename as possible, but will use short versions
    // of path and/or basename if either one is not entirely ASCII. This is purely a workaround for Python's 
    // brokenness w.r.t Unicode filenames.

    bool 
    ModuleCache::GetASCIIFullPathBasenameExtension( const std::wstring& filename, 
                                                    std::wstring& path, 
                                                    std::wstring& basename, 
                                                    std::wstring& extension,
                                                    bool& bBasenameIsShort, // Need to return this for a hack one level up (setting of PYTHONCASEOK)
                                                    std::string& errTxt ) 
    {
        if (!GetFullPathBasenameExtension( filename, true, path, basename, extension, errTxt)){
            // errTxt is already set
            return false;
        }

        // The usual case is that our path and filename are both ASCII. If they're not, we have to give the short name
        // to Python, as it doesn't properly deal with Unicode paths or filenames (and short paths are all ASCII, as far
        // as I can tell from examining Japanese filenames on my standard US Windows XP).

        bool bLongPathIsASCII = IsAllASCII(path);
        bool bLongBasenameIsASCII = IsAllASCII(basename);
        // Don't care about extension here; for Python to want to load this file, it has to be .py, .pyc, or .pyd, 
        // and that `will be checked by the caller.

        if (! (bLongPathIsASCII && bLongBasenameIsASCII) ) {			
            std::wstring shortPath, shortBasename, shortExt;
            if (!GetFullPathBasenameExtension( filename, false, shortPath, shortBasename, shortExt, errTxt)){
                // errTxt is already set
                return false;
            }

            assert(IsAllASCII(shortPath) && IsAllASCII(shortBasename)); // I think this is guaranteed with short names
            if (!bLongPathIsASCII) {
                path = shortPath;
            }
            if (!bLongBasenameIsASCII) {
                basename = shortBasename;
            }
        }

        bBasenameIsShort = !bLongBasenameIsASCII;

        return true;
    }

    //////////////////////////////////////////////////////////////////////////////

    bool 
    ModuleCache::ImportOrReload( const std::wstring& filename,
                                 bool bImport,  // if false, reload
                                 PyObject*& rpModule )   // May not be NULL - may have pointer to previous import of module
    {
        std::wstring path, basename, extension;
        bool bBasenameIsShort;
        std::string errTxt;

        bool rc = GetASCIIFullPathBasenameExtension( filename, path, basename, 
            extension, bBasenameIsShort, errTxt );
        if (!rc) {
            ERROUT("Error extracting module path and name from %s: %s", ASCII_REPR(filename), errTxt.c_str());
        }

        // Make sure it's a valid Python file extension. Could be .py or .pyc, so check to see that the extension is 
        // two or three characters, and that the first two chars match "py". 
        //
        // I'm assuming that these extension chars will always be ASCII, wo we can safely use case-insensitive _wcsnicmp.

        if (rc) {
            int extLen = extension.length();
            bool bGood = (extLen == 2 || extLen == 3) &&
                (_wcsnicmp(extension.c_str(), L"py", 2) == 0); 
            if (!bGood) {
                ERROUT("%s doesn't have a .py or .pyc extension", ASCII_REPR(filename));
                rc = false;
            }
        }

        // Only add to the path once, during import (which should only happen once).
        // If two script files are in the same directory, they'll both try to insert that
        // dir into sys.path. Keep track of what we've inserted and don't double-insert, to
        // avoid bloating sys.path.

        if ( rc && 
            bImport && 
            m_mapDirInfo.find(path) == m_mapDirInfo.end() ) 
        {
            PyObject* pPathAddition = PyUnicode_FromWideChar(path.c_str(), path.length());                 
            if (!pPathAddition) {
                ERROUT("Couldn't get python string from path %s", ASCII_REPR(path));
                rc = false;
            }

            // Python docs say PySys_GetObject returns a "borrowed reference," which we must not decrement                
            if (rc) {
                PyObject *sys_path = PySys_GetObject("path");
                if (sys_path) {
                    PyList_Insert(sys_path, 0, pPathAddition);
                } else {
                    ERROUT("Couldn't get python's module import path");
                    rc = false;
                }
            }
            Py_XDECREF(pPathAddition);
        }

        PyObject* pBasename = NULL;
        if (rc) {
            pBasename = PyUnicode_FromWideChar(basename.c_str(), basename.length());                   
            if (!pBasename) {
                ERROUT("Couldn't get python basename string from basename %s", ASCII_REPR(basename));
                rc = false;
            }
        }

        if (rc) {
            if (bImport) {
                assert(rpModule == NULL);

                // Hack to work with short names requires the PYTHONCASEOK variable be set,
                // so we'll do so if it's not already set. This has to be done for all 
                // loaded CRTs, so the job is farmed out the code that sets up the output streams.

                std::map<std::wstring, bool> crtSettings;
                if (bBasenameIsShort) {
                    if (!SetEnvVarsIfUnset( "PYTHONCASEOK", crtSettings )) {
                        // No need to fail out; the module import will almost certainly fail
                        ERROUT("Failed to temporarily set PYTHONCASEOK env var");
                    }
                }

                rpModule = PyImport_Import(pBasename);

                // Undo the case hack
                if (bBasenameIsShort) {
                    if (!ResetEnvVars( "PYTHONCASEOK", crtSettings )) {
                        // No need to fail out; the module import almost certainly failed
                        ERROUT("Failed to reset PYTHONCASEOK env var");
                    }
                }

                if (!rpModule) {
                    ERROUT("Couldn't import python module %s from path %s", ASCII_REPR(basename), ASCII_REPR(path));
                    if (PyErr_Occurred()) {
                        PyErr_Print();
                    }
                    rc = false;
                }
            } else {
                assert(rpModule);
                rpModule = PyImport_ReloadModule(rpModule);
                if (!rpModule) {
                    ERROUT("Couldn't reload python module %s from path %s", ASCII_REPR(basename), ASCII_REPR(path));
                    if (PyErr_Occurred()) {
                        PyErr_Print();
                    }
                    rc = false;
                }
            }
        }

        Py_XDECREF(pBasename);
        return rc;
    }

//////////////////////////////////////////////////////////////////////////////

} // end of anonymous namespace

//////////////////////////////////////////////////////////////////////////////

bool 
GetPyModuleAndFunctionObjects(  const std::wstring& filename, 
                              const std::string& function,
                              PyObject*& rpModule,
                              PyObject*& rpFunction ) 
{   

    bool rc;
    rpModule = rpFunction = NULL;

    rc = ModuleCache::Factory().GetModule( filename, rpModule );
    if (!rc || !rpModule) {
        ERROUT("Couldn't get python module %s", ASCII_REPR(filename));
        return false;
    }

    // rpFunction is always a new reference; easier than caching it
    rpFunction = PyObject_GetAttrString(rpModule, function.c_str());
    if (!rpFunction || !PyCallable_Check(rpFunction)) {
        if (PyErr_Occurred()) {
            PyErr_Print();
        }
        Py_XDECREF(rpFunction);
        rpModule = rpFunction = NULL; // Don't want to return any handles to the caller if there's a problem
        ERROUT("Cannot find valid function \"%s\" in %s", function.c_str(), ASCII_REPR(filename));
        return false;
    }

    return true;
}

//////////////////////////////////////////////////////////////////////////////

bool
ModuleFreshnessCheckEnabled()
{
    return ModuleCache::Factory().ModuleFreshnessCheckEnabled();
}

//////////////////////////////////////////////////////////////////////////////

void
SetModuleFreshnessCheck( bool bCheck )
{
    ModuleCache::Factory().SetModuleFreshnessCheck(bCheck);
}

//////////////////////////////////////////////////////////////////////////////
