// PyCall: loading a script by its path, calling one of its functions, and turning every
// failure into #VALUE! with the reason on the console. docs/architecture.md, "PyCall",
// "Argument rule" and "Scripts", says what each case does and why. Call these with the
// GIL held.
#pragma once

#include "core/py.h"
#include "excel/excel.h"

#include <functional>
#include <span>
#include <string>

namespace pyinex::core {

// Defines the Python code that loads and calls scripts, PyCall's and direct functions'
// (src/core/python/call.py and funcs.py, in one namespace). Call once, after
// start_conversion. False, with Python's error left set, if Python failed.
bool start_calls();

// One of the names call.py and funcs.py define, such as "func"; borrowed, or null before
// start_calls.
PyObject* helper(const char* name);

// Writes out what sys.stdout and sys.stderr have buffered. pyinex_call calls it after each
// call, so output reaches the console as the cell calculates.
void flush();

// Calls function in the script file with args, Excel's argument slots after the first two,
// and returns the result as Excel's. folder gives the calling workbook's folder, for a
// relative file; it is asked only then, and returns empty for an unsaved workbook.
excel::Value call(const XLOPER12& file, const XLOPER12& function, std::span<const XLOPER12* const> args,
                  bool reload_on_change, const std::function<std::wstring()>& folder);

// Loads a script for PyLoad or pyinex.ini's list and returns a row for each function it marks
// @pyinex.func (funcs.py, load_functions), or #VALUE! after saying why it did not load,
// made ready for Excel (Value::to_excel). folder is as for call.
LPXLOPER12 load_functions(const XLOPER12& file, bool reload_on_change, const std::function<std::wstring()>& folder);

// What reloads changed outside load_functions (funcs.py, changes), or an empty value for
// nothing, made ready for Excel (Value::to_excel).
LPXLOPER12 changes();

// Calls the direct function load_functions gave key, with Excel's argument slots, placed by
// its signature (funcs.py, call_function). name is its name in Excel, for messages.
excel::Value call_function(long key, const std::wstring& name, std::span<const XLOPER12* const> args,
                           bool reload_on_change);

}  // namespace pyinex::core
