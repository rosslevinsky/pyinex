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

// xlAutoClose: writes out what Python has buffered for the console. Python keeps running.
using pyinex_close_t = void();

// One of PyInfo's items that only Python knows: "version", "prefix" or "venv".
using pyinex_info_t = LPXLOPER12(const wchar_t* item);

}
