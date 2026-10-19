// The functions pyinex_core.dll exports for pyinex.xll, which finds them with
// GetProcAddress once it has loaded Python and the core.
//
// Results are worksheet function results (excel::Value::to_excel): Excel frees them through
// the add-in's xlAutoFree12. Both DLLs build Value from the same code against the same C
// runtime, whose heap they share, so either may free what the other made.
#pragma once

#include <windows.h>

#include <XLCALL.H>

extern "C" {

// Starts Python as the chosen python.exe would start: executable is that python.exe, and
// base_executable what it reports as sys._base_executable. Called once, with Python's DLLs
// already loaded. Returns 0 on success, after which the GIL is released.
using pyinex_start_t = int(const wchar_t* executable, const wchar_t* base_executable);

// PyCall: calls function in the script file with Excel's remaining argument slots, count of
// them, and returns its result. Every failure gives #VALUE!, with the reason on the console.
using pyinex_call_t = LPXLOPER12(const XLOPER12* file, const XLOPER12* function, const XLOPER12* const* args,
                                 int count, bool reload_on_change);

// Loads a script for PyLoad or pyinex.ini's list, as PyCall loads scripts: file is its path,
// a relative one relative to the calling workbook's folder. Returns rows: first a key alone
// for each function to remove, one the script no longer has as it was registered; then, for
// each function it marks @pyinex.func, the key -1 alone for one registered already from this
// script, or a row to register: its key, name, argument count (or, when an argument takes a
// reference, a pyinex.Range, Excel's type code for each: Q for values, U for a reference),
// argument names, category and help, then each argument's help. #VALUE! means the script did
// not load, the console saying why.
using pyinex_load_t = LPXLOPER12(const XLOPER12* file, bool reload_on_change);

// What the reloads of scripts PyLoad loaded changed since the last time it was asked: rows as
// pyinex_load's, with the script's name, as PyLoad was given it, put before each; or an empty
// value (xltypeNil) for none. pyinex.xll asks after each call into the core.
using pyinex_changes_t = LPXLOPER12();

// Whether the function pyinex_load gave key registered. One that did not, or that will not
// be, frees its name; one removed from Excel, which no cell can call any more, is forgotten.
using pyinex_registered_t = void(int key, bool registered);

// A direct function: calls the function pyinex_load gave key with Excel's count argument
// slots, placed by its signature and converted by its annotations. name is its name in
// Excel, for messages.
using pyinex_function_t = LPXLOPER12(int key, const wchar_t* name, const XLOPER12* const* args, int count,
                                     bool reload_on_change);

// xlAutoClose: writes out what Python has buffered for the console. Python keeps running.
using pyinex_close_t = void();

// One of PyInfo's items that only Python knows: "version", "prefix" or "venv".
using pyinex_info_t = LPXLOPER12(const wchar_t* item);

}
