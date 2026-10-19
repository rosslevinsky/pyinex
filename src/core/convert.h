// Values crossing between Excel and Python, by the rules in docs/conversion.md, which says
// what each case does and why. Call these with the GIL held.
#pragma once

#include "core/py.h"
#include "excel/excel.h"

#include <string>

namespace pyinex::core {

// Creates the ExcelError class, which the pyinex module offers as pyinex.ExcelError. Call
// once Python has started, before the functions below. False, with Python's error left set,
// if Python failed.
bool start_conversion();

// The ExcelError class; a borrowed reference, null before start_conversion.
PyObject* excel_error_type();

// An Excel value as a Python object: a new reference. A value Pyinex does not recognize
// arrives as ExcelError("#VALUE!"), and problem says what it was (the first, if a range
// holds several). Null, with problem set, only if Python itself fails, such as by running
// out of memory.
PyObject* to_python(const XLOPER12& x, std::wstring& problem);

// A Python object as an Excel value. If any part of it cannot be returned, the whole
// result is #VALUE! and problem says why and where, e.g. "row 3, column 2: dict cannot be
// returned to Excel".
excel::Value to_excel(PyObject* object, std::wstring& problem);

}  // namespace pyinex::core
