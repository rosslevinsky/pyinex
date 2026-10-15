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

// Defines the Python code that loads and calls scripts. Call once, after
// start_conversion. False if Python failed.
bool start_calls();

// Writes out what sys.stdout and sys.stderr have buffered. pyinex_call calls it after each
// call, so output reaches the console as the cell calculates.
void flush();

// Calls function in the script file with args, Excel's argument slots after the first two,
// and returns the result as Excel's. folder gives the calling workbook's folder, for a
// relative file; it is asked only then, and returns empty for an unsaved workbook.
excel::Value call(const XLOPER12& file, const XLOPER12& function, std::span<const XLOPER12* const> args,
                  bool reload_on_change, const std::function<std::wstring()>& folder);

}  // namespace pyinex::core
