#include "core/convert.h"

#include <algorithm>
#include <cmath>
#include <format>
#include <iterator>
#include <memory>
#include <vector>

namespace pyinex::core {
namespace {

using excel::Value;

// Excel's error codes and the text each shows in a cell.
struct ErrorText {
    int code;
    const char* text;
};
constexpr ErrorText errors[] = {
    {xlerrNull, "#NULL!"}, {xlerrDiv0, "#DIV/0!"}, {xlerrValue, "#VALUE!"},
    {xlerrRef, "#REF!"},   {xlerrName, "#NAME?"},  {xlerrNum, "#NUM!"},
    {xlerrNA, "#N/A"},     {xlerrGettingData, "#GETTING_DATA"},
};

// Defines ExcelError, given the error texts above as a tuple.
constexpr const char* error_source = R"(
def make(texts):
    class ExcelError:
        """An Excel error, such as #N/A.

        An error in a cell arrives in Python as an ExcelError, and a function returns one
        to put that error in its cell. ExcelError('#N/A') makes one; str(e) and e.text give
        its text.
        """
        __slots__ = ('_text',)
        __module__ = 'pyinex'
        __qualname__ = 'ExcelError'

        def __init__(self, text):
            # str.__str__ gives a plain str, so a subclass's own comparison is never used.
            plain = str.__str__(text) if isinstance(text, str) else None
            if plain not in texts:
                shown = repr(plain) if plain is not None else 'a ' + type(text).__name__
                raise ValueError(f'not an Excel error: {shown}; expected one of {", ".join(texts)}')
            self._text = plain

        @property
        def text(self):
            return self._text

        def __str__(self):
            return self._text

        def __repr__(self):
            return f'ExcelError({self._text!r})'

        def __eq__(self, other):
            if isinstance(other, ExcelError):
                return self._text == other._text
            return NotImplemented

        def __hash__(self):
            return hash(self._text)

    return ExcelError
)";

PyObject* error_type;  // ExcelError, once start_conversion has made it.

int type_of(const XLOPER12& x) { return x.xltype & ~(xlbitDLLFree | xlbitXLFree); }

PyObject* make_error(const char* text) { return PyObject_CallFunction(error_type, "s", text); }

// Keeps the first problem: one line on the console is enough to show what went wrong.
void report(std::wstring& problem, std::wstring message) {
    if (problem.empty()) problem = std::move(message);
}

// One Excel value that is not an array.
PyObject* cell_to_python(const XLOPER12& x, std::wstring& problem) {
    switch (type_of(x)) {
    case xltypeNum: return PyFloat_FromDouble(x.val.num);
    case xltypeInt: return PyFloat_FromDouble(x.val.w);
    case xltypeStr: return PyUnicode_FromWideChar(x.val.str + 1, x.val.str[0]);
    case xltypeBool: return PyBool_FromLong(x.val.xbool);
    case xltypeNil:
    case xltypeMissing: return Py_NewRef(Py_None);
    case xltypeErr:
        for (const auto& e : errors)
            if (e.code == x.val.err) return make_error(e.text);
        report(problem, std::format(L"Excel error code {} is not one Pyinex knows; it arrives as #VALUE!",
                                    x.val.err));
        break;
    default:
        report(problem, std::format(L"an Excel value of type 0x{:04X} is not one Pyinex knows; it arrives as #VALUE!",
                                    type_of(x)));
        break;
    }
    return make_error("#VALUE!");
}

// A range: a single cell as a plain value, anything larger as a tuple of row tuples.
PyObject* array_to_python(const XLOPER12& x, std::wstring& problem) {
    const auto& a = x.val.array;
    if (a.rows < 1 || a.columns < 1 || !a.lparray) return cell_to_python(x, problem);
    if (a.rows == 1 && a.columns == 1) return cell_to_python(a.lparray[0], problem);
    // Convert every cell before making any tuple: making an ExcelError runs Python code,
    // which must not see, or keep, a tuple still being filled.
    const auto count = static_cast<size_t>(a.rows) * a.columns;
    std::vector<Ref> cells;
    cells.reserve(count);
    for (size_t i = 0; i < count; ++i) {
        cells.emplace_back(cell_to_python(a.lparray[i], problem));
        if (!cells.back().get()) return nullptr;
    }
    Ref rows(PyTuple_New(a.rows));
    if (!rows.get()) return nullptr;
    for (int r = 0; r < a.rows; ++r) {
        PyObject* row = PyTuple_New(a.columns);
        if (!row || PyTuple_SetItem(rows.get(), r, row) < 0) return nullptr;  // Steals row.
        for (int c = 0; c < a.columns; ++c)
            if (PyTuple_SetItem(row, c, cells[static_cast<size_t>(r) * a.columns + c].release()) < 0) return nullptr;
    }
    return rows.release();
}

bool is_sequence(PyObject* o) { return PyList_Check(o) || PyTuple_Check(o); }

// A list or tuple's items, held while they convert. A list is copied first, since code
// that runs during conversion (a subclass's methods) could otherwise change it underneath.
Ref items_of(PyObject* sequence) {
    return Ref(PyList_Check(sequence) ? PyList_GetSlice(sequence, 0, PY_SSIZE_T_MAX) : Py_NewRef(sequence));
}

Py_ssize_t size_of(PyObject* items) { return PyList_Check(items) ? PyList_Size(items) : PyTuple_Size(items); }

PyObject* item(PyObject* items, Py_ssize_t i) {  // Borrowed.
    return PyList_Check(items) ? PyList_GetItem(items, i) : PyTuple_GetItem(items, i);
}

std::wstring type_name(PyObject* o) {
    Ref name(PyObject_GetAttrString(reinterpret_cast<PyObject*>(Py_TYPE(o)), "__name__"));
    std::wstring result = L"this type";
    if (name.get() && PyUnicode_Check(name.get())) {
        Py_ssize_t size = 0;
        std::unique_ptr<wchar_t, void (*)(void*)> text(PyUnicode_AsWideCharString(name.get(), &size), PyMem_Free);
        if (text) result.assign(text.get(), static_cast<size_t>(size));
    }
    PyErr_Clear();
    return result;
}

// One Python value that is not a list or tuple. False, with why set, if it cannot be returned.
bool cell_to_excel(PyObject* o, Value& out, std::wstring& why) {
    if (PyBool_Check(o)) {  // Before int: bool is a kind of int.
        out = Value::boolean(o == Py_True);
    } else if (PyLong_Check(o)) {
        const double n = PyLong_AsDouble(o);
        if (n == -1.0 && PyErr_Occurred()) {
            PyErr_Clear();
            why = L"int too large to be a number in Excel";
            return false;
        }
        out = Value::number(n);
    } else if (PyFloat_Check(o)) {
        const double n = PyFloat_AsDouble(o);
        out = std::isfinite(n) ? Value::number(n) : Value::error(xlerrNum);  // NaN, infinity.
    } else if (PyUnicode_Check(o)) {
        Py_ssize_t size = 0;  // In UTF-16 code units, as Excel counts.
        std::unique_ptr<wchar_t, void (*)(void*)> text(PyUnicode_AsWideCharString(o, &size), PyMem_Free);
        if (!text) {
            PyErr_Clear();
            why = L"str could not be read";
            return false;
        }
        if (size > excel::max_text) {
            why = L"str longer than 32,767 UTF-16 code units, Excel's limit for a cell";
            return false;
        }
        out = Value::text({text.get(), static_cast<size_t>(size)});
    } else if (o == Py_None) {
        out = Value::text(L"");  // A truly empty result would show as 0.
    } else if (PyObject_TypeCheck(o, reinterpret_cast<PyTypeObject*>(error_type))) {
        Ref text(PyObject_GetAttrString(o, "_text"));
        for (const auto& e : errors) {
            if (text.get() && PyUnicode_Check(text.get()) && PyUnicode_CompareWithASCIIString(text.get(), e.text) == 0) {
                out = Value::error(e.code);
                return true;
            }
        }
        PyErr_Clear();
        why = L"ExcelError holding unknown text cannot be returned to Excel";
        return false;
    } else {
        why = type_name(o) + L" cannot be returned to Excel";
        return false;
    }
    return true;
}

}  // namespace

bool start_conversion() {
    Ref names(PyDict_New());
    Ref builtins(PyImport_ImportModule("builtins"));
    Ref ran(names.get() && builtins.get()
                ? PyObject_CallMethod(builtins.get(), "exec", "sO", error_source, names.get())
                : nullptr);
    PyObject* make = ran.get() ? PyDict_GetItemString(names.get(), "make") : nullptr;  // Borrowed.
    Ref texts(PyTuple_New(std::size(errors)));
    bool ok = make && texts.get();
    for (size_t i = 0; ok && i < std::size(errors); ++i) {
        PyObject* text = PyUnicode_FromString(errors[i].text);
        ok = text && PyTuple_SetItem(texts.get(), static_cast<Py_ssize_t>(i), text) == 0;
    }
    Ref type(ok ? PyObject_CallFunctionObjArgs(make, texts.get(), nullptr) : nullptr);
    if (!type.get()) {
        PyErr_Clear();
        return false;
    }
    Py_XDECREF(error_type);
    error_type = type.release();
    return true;
}

PyObject* excel_error_type() { return error_type; }

PyObject* to_python(const XLOPER12& x, std::wstring& problem) {
    PyObject* result = type_of(x) == xltypeMulti ? array_to_python(x, problem) : cell_to_python(x, problem);
    if (!result) {
        PyErr_Clear();
        problem = L"Python failed while converting an Excel value";
    }
    return result;
}

Value to_excel(PyObject* object, std::wstring& problem) {
    std::wstring why;
    Value out;
    if (!is_sequence(object)) {
        if (cell_to_excel(object, out, why)) return out;
        problem = why;
        return Value::error(xlerrValue);
    }
    const auto fail = [&](std::wstring message) {
        PyErr_Clear();
        problem = std::move(message);
        return Value::error(xlerrValue);
    };
    const auto too_big = [&](Py_ssize_t rows, Py_ssize_t columns) {
        return fail(std::format(L"a block of {} rows by {} columns is larger than Excel's grid of 1,048,576 by 16,384",
                                rows, columns));
    };

    const Ref top = items_of(object);
    if (!top.get()) return fail(L"Python failed while reading the result");
    const Py_ssize_t n = size_of(top.get());
    if (n == 0) return Value::text(L"");  // One blank cell.

    // Each element is a row if any element is a list or tuple; otherwise the elements are
    // one row. Hold every row's items, and find the widest row, before converting.
    bool has_rows = false;
    for (Py_ssize_t i = 0; i < n && !has_rows; ++i) has_rows = is_sequence(item(top.get(), i));
    std::vector<Ref> rows;
    Py_ssize_t width = 1;
    if (has_rows) {
        if (n > excel::max_rows) return too_big(n, 1);
        rows.reserve(static_cast<size_t>(n));
        for (Py_ssize_t i = 0; i < n; ++i) {
            PyObject* element = item(top.get(), i);
            rows.push_back(is_sequence(element) ? items_of(element) : Ref(PyTuple_Pack(1, element)));
            if (!rows.back().get()) return fail(L"Python failed while reading the result");
            width = std::max(width, size_of(rows.back().get()));
        }
    } else {
        rows.push_back(Ref(Py_NewRef(top.get())));
        width = n;
    }
    const auto height = static_cast<Py_ssize_t>(rows.size());
    if (height > excel::max_rows || width > excel::max_columns) return too_big(height, width);

    std::vector<Value> cells;
    cells.reserve(static_cast<size_t>(height * width));
    for (Py_ssize_t r = 0; r < height; ++r) {
        PyObject* row = rows[static_cast<size_t>(r)].get();
        const Py_ssize_t length = size_of(row);
        for (Py_ssize_t c = 0; c < width; ++c) {
            if (c >= length) {
                cells.push_back(Value::text(L""));  // Padding for a short row.
                continue;
            }
            PyObject* element = item(row, c);
            if (is_sequence(element)) why = L"a list or tuple inside a row is nested too deeply for Excel";
            else if (cell_to_excel(element, out, why)) {
                cells.push_back(std::move(out));
                continue;
            }
            return fail(std::format(L"row {}, column {}: {}", r + 1, c + 1, why));
        }
    }
    return Value::array(static_cast<int>(height), static_cast<int>(width), std::move(cells));
}

}  // namespace pyinex::core
