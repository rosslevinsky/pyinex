// Pyinex's thin layer over Microsoft's Excel C API (XLCALL.H).
//
// Excel's C API passes values as XLOPER12 structs, and the hard part of using it is
// memory: who frees a value depends on who made it.
//
// - Values Excel makes, as the result of a call into Excel, are freed by calling Excel's
//   xlFree. ExcelValue does that when it goes out of scope.
// - Values the add-in makes and returns from a worksheet function are marked xlbitDLLFree;
//   Excel copies them and then hands them back to the add-in's xlAutoFree12 to free.
//   Value builds them, and Value::to_excel hands one over.
//
// Mixing the two up crashes Excel, so nothing else in Pyinex touches XLOPER12 memory.
#pragma once

#include <windows.h>

#include <XLCALL.H>

#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace pyinex::excel {

constexpr int max_rows = 1048576;  // Excel's grid.
constexpr int max_columns = 16384;
constexpr int max_text = 32767;  // Excel's limit for a cell's text, and for an XCHAR length.

// A value the add-in builds. Owns the memory behind its text and arrays.
class Value {
public:
    Value() = default;  // Empty (xltypeNil).
    static Value number(double n);
    // Text longer than Excel's 32,767-character cell limit is the caller's error to catch;
    // here it is cut to that limit rather than overflowing the length prefix.
    static Value text(std::wstring_view s);
    static Value boolean(bool b);
    static Value error(int code);  // xlerrValue and the like.
    static Value missing();        // An argument left out (xltypeMissing).
    // A rows x columns block, items in row order. items.size() must be rows * columns.
    static Value array(int rows, int columns, std::vector<Value> items);
    // A deep copy of any plain value Excel passed in. A reference becomes #VALUE!.
    static Value copy(const XLOPER12& x);

    Value(Value&& other) noexcept;
    Value& operator=(Value&& other) noexcept;
    Value(const Value&) = delete;
    Value& operator=(const Value&) = delete;
    ~Value();

    const XLOPER12& get() const { return x_; }

    // Returns the value as a worksheet function result, marked xlbitDLLFree. Excel passes
    // it to xlAutoFree12 once it has copied it. This Value is left empty.
    LPXLOPER12 to_excel();

private:
    XLOPER12 release();  // Hands over the contents; this Value is left empty.
    XLOPER12 x_{.xltype = xltypeNil};
};

// Frees a result returned by Value::to_excel. Called from xlAutoFree12.
void free_returned(LPXLOPER12 x);

// A value Excel made, freed with xlFree when this goes out of scope.
class ExcelValue {
public:
    ExcelValue() = default;
    ExcelValue(const ExcelValue&) = delete;
    ExcelValue& operator=(const ExcelValue&) = delete;
    ~ExcelValue();

    const XLOPER12& get() const { return x_; }

private:
    friend int call(int function, ExcelValue* result, std::span<const XLOPER12* const> args);
    void reset();
    XLOPER12 x_{.xltype = xltypeNil};
    bool owned_ = false;
};

// Calls an Excel C API function (xlfRegister, xlGetName, ...) through call_excel and
// returns its xlret code; xlretSuccess is 0. result may be null when the result is not
// wanted.
int call(int function, ExcelValue* result, std::span<const XLOPER12* const> args = {});
int call(int function, ExcelValue* result, std::initializer_list<const XLOPER12*> args);

// The folder of the workbook whose cell is calling the add-in, from CELL("filename"); empty
// if it has not been saved or the caller is not a cell.
std::wstring caller_folder();

// Whether Excel is calling from its function wizard, the Function Arguments dialog that
// evaluates a formula as it is typed: a window of that dialog's class (bosa_sdm_XL...) and
// title exists in Excel's process. Microsoft's sample also checks the dialog's owner is
// Excel's main window, but since Excel 2013 each workbook has a window of its own. It is
// recognized by its English title, so in Excel in other languages functions run in the
// wizard as they would in a cell.
bool in_function_wizard();

// One worksheet function for register_functions.
struct Function {
    std::wstring_view name;       // The name used in formulas.
    std::wstring_view procedure;  // The add-in's exported C function.
    // Excel's type string: the result's type, then each argument's, then flags; for
    // example L"QQQ" for a function of two values. Never "$": Pyinex runs on Excel's
    // main thread only.
    std::wstring_view type;
    std::wstring_view arguments;  // Argument names, comma separated, for the formula dialog.
    std::wstring_view help;
    std::vector<std::wstring_view> argument_help;
};

// Registers each function with Excel under the category "Pyinex". The list can be any
// length and is not tied to particular functions. Returns how many registered.
int register_functions(std::span<const Function> functions);

// Runs a worksheet function's body. No C++ exception may reach Excel, which would end
// it: any exception, such as running out of memory, gives #VALUE! from static memory,
// which needs no allocation.
template <class F>
LPXLOPER12 guarded(F body) noexcept {
    try {
        return body();
    } catch (...) {
        static XLOPER12 error{.xltype = xltypeErr};
        error.val.err = xlerrValue;
        return &error;
    }
}

}  // namespace pyinex::excel
