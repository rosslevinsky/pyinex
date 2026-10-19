// _pyinex_testing: the add-in's value conversion and column letters as a Python extension
// module, so tests/unit can test them without Excel. It is compiled from the same sources
// as pyinex_core.dll, and holds the core's Python files as the core does (sources).
// tests/unit/conftest.py describes how values are passed in and out.
#include "core/call.h"
#include "core/convert.h"
#include "core/py.h"
#include "core/python/bootstrap.py.h"
#include "core/python/call.py.h"
#include "core/python/excel_error.py.h"
#include "core/python/funcs.py.h"
#include "excel/address.h"
#include "excel/excel.h"

#include <deque>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

using namespace pyinex;
using core::Ref;

namespace {

// XLOPER12 values built from descriptions, and the text and arrays they point to.
struct Arena {
    std::deque<std::wstring> texts;  // Each starts with its length, as Excel's text does.
    std::deque<std::vector<XLOPER12>> arrays;
};

bool fail(const char* message) {
    PyErr_SetString(PyExc_ValueError, message);
    return false;
}

// Builds x from a description such as ("num", 1.5). False, with a Python error set, if the
// description is not one.
bool build(PyObject* desc, XLOPER12& x, Arena& arena) {
    if (!PyTuple_Check(desc) || PyTuple_Size(desc) < 1) return fail("a description is a tuple");
    const char* kind = PyUnicode_AsUTF8AndSize(PyTuple_GetItem(desc, 0), nullptr);
    if (!kind) return false;
    const std::string_view k = kind;
    const Py_ssize_t n = PyTuple_Size(desc);
    PyObject* a = n > 1 ? PyTuple_GetItem(desc, 1) : nullptr;
    x = {};
    if ((k == "nil" || k == "missing") && n == 1) {
        x.xltype = k == "nil" ? xltypeNil : xltypeMissing;
    } else if (k == "num" && n == 2) {
        x.xltype = xltypeNum;
        x.val.num = PyFloat_AsDouble(a);
    } else if (k == "str" && n == 2) {
        Py_ssize_t size = 0;
        wchar_t* text = PyUnicode_AsWideCharString(a, &size);
        if (!text) return false;
        if (size <= excel::max_text) {
            auto& s = arena.texts.emplace_back(1, static_cast<wchar_t>(size));
            s.append(text, static_cast<size_t>(size));
            x.xltype = xltypeStr;
            x.val.str = s.data();
        }
        PyMem_Free(text);
        if (size > excel::max_text) return fail("Excel text holds at most 32,767 code units");
    } else if (k == "bool" && n == 2) {
        x.xltype = xltypeBool;
        x.val.xbool = PyObject_IsTrue(a);
    } else if ((k == "err" || k == "int" || k == "raw") && n == 2) {
        const long v = PyLong_AsLong(a);
        x.xltype = k == "err" ? xltypeErr : k == "int" ? xltypeInt : static_cast<DWORD>(v);
        if (k == "err") x.val.err = static_cast<int>(v);
        if (k == "int") x.val.w = static_cast<int>(v);
    } else if (k == "multi" && n == 4) {
        const long rows = PyLong_AsLong(a), columns = PyLong_AsLong(PyTuple_GetItem(desc, 2));
        if (PyErr_Occurred()) return false;
        // A private copy of the list, so code run while building an item cannot change it.
        Ref items(PyList_Check(PyTuple_GetItem(desc, 3)) ? PyList_GetSlice(PyTuple_GetItem(desc, 3), 0, PY_SSIZE_T_MAX)
                                                          : nullptr);
        if (!items.get() || rows < 0 || columns < 0 || rows > excel::max_rows || columns > excel::max_columns ||
            PyList_Size(items.get()) != static_cast<Py_ssize_t>(rows) * columns)
            return fail("rows and columns within Excel's grid, and a list of rows * columns items");
        const auto count = static_cast<size_t>(rows) * static_cast<size_t>(columns);
        auto& cells = arena.arrays.emplace_back(count);
        for (size_t i = 0; i < count; ++i)
            if (!build(PyList_GetItem(items.get(), static_cast<Py_ssize_t>(i)), cells[i], arena)) return false;
        x.xltype = xltypeMulti;
        x.val.array = {.lparray = cells.empty() ? nullptr : cells.data(), .rows = static_cast<INT32>(rows),
                       .columns = static_cast<INT32>(columns)};
    } else {
        return fail("not a description of an Excel value");
    }
    return !PyErr_Occurred();
}

// The description of x; a new reference.
PyObject* describe(const XLOPER12& x) {
    switch (x.xltype & ~(xlbitDLLFree | xlbitXLFree)) {
    case xltypeNum: return Py_BuildValue("(sd)", "num", x.val.num);
    case xltypeStr: return Py_BuildValue("(sN)", "str", PyUnicode_FromWideChar(x.val.str + 1, x.val.str[0]));
    case xltypeBool: return Py_BuildValue("(sO)", "bool", x.val.xbool ? Py_True : Py_False);
    case xltypeErr: return Py_BuildValue("(si)", "err", x.val.err);
    case xltypeInt: return Py_BuildValue("(si)", "int", x.val.w);
    case xltypeNil: return Py_BuildValue("(s)", "nil");
    case xltypeMissing: return Py_BuildValue("(s)", "missing");
    case xltypeMulti: {
        const auto& a = x.val.array;
        Ref items(PyList_New(0));
        for (size_t i = 0; items.get() && i < static_cast<size_t>(a.rows) * a.columns; ++i) {
            Ref item(describe(a.lparray[i]));
            if (!item.get() || PyList_Append(items.get(), item.get()) < 0) return nullptr;
        }
        return items.get() ? Py_BuildValue("(siiO)", "multi", a.rows, a.columns, items.get()) : nullptr;
    }
    default: return Py_BuildValue("(si)", "raw", static_cast<int>(x.xltype));
    }
}

PyObject* problem_or_none(const std::wstring& problem) {
    if (problem.empty()) return Py_NewRef(Py_None);
    return PyUnicode_FromWideChar(problem.data(), static_cast<Py_ssize_t>(problem.size()));
}

// Runs a module function's body; a C++ exception becomes a Python one.
template <class F>
PyObject* guarded(F body) {
    try {
        return body();
    } catch (...) {
        PyErr_SetString(PyExc_MemoryError, "C++ exception in _pyinex_testing");
        return nullptr;
    }
}

// to_python(desc) -> (value, problem or None)
PyObject* to_python(PyObject*, PyObject* desc) {
    return guarded([&]() -> PyObject* {
        Arena arena;
        XLOPER12 x{};
        if (!build(desc, x, arena)) return nullptr;
        std::wstring problem;
        Ref value(core::to_python(x, problem));
        if (!value.get()) {
            Ref message(problem_or_none(problem));
            PyErr_SetObject(PyExc_RuntimeError, message.get());
            return nullptr;
        }
        return Py_BuildValue("(ON)", value.get(), problem_or_none(problem));
    });
}

// to_excel(value) -> (desc, problem or None)
PyObject* to_excel(PyObject*, PyObject* value) {
    return guarded([&]() -> PyObject* {
        std::wstring problem;
        const excel::Value result = core::to_excel(value, problem);
        Ref desc(describe(result.get()));
        return desc.get() ? Py_BuildValue("(ON)", desc.get(), problem_or_none(problem)) : nullptr;
    });
}

// call(file, function, *args, reload_on_change=True, folder="") -> desc: PyCall, with every
// value a description. folder is the calling workbook's folder, empty for an unsaved one.
PyObject* call(PyObject*, PyObject* args, PyObject* kwargs) {
    return guarded([&]() -> PyObject* {
        static const char* keywords[] = {"reload_on_change", "folder", nullptr};
        int reload_on_change = 1;
        PyObject* folder = nullptr;
        Ref empty(PyTuple_New(0));
        if (!empty.get() || !PyArg_ParseTupleAndKeywords(empty.get(), kwargs, "|pU", const_cast<char**>(keywords),
                                                         &reload_on_change, &folder))
            return nullptr;
        const Py_ssize_t n = PyTuple_Size(args);
        if (n < 2) {
            PyErr_SetString(PyExc_TypeError, "call needs a file and a function");
            return nullptr;
        }
        Arena arena;
        std::vector<XLOPER12> values(static_cast<size_t>(n));
        for (Py_ssize_t i = 0; i < n; ++i)
            if (!build(PyTuple_GetItem(args, i), values[static_cast<size_t>(i)], arena)) return nullptr;
        std::vector<const XLOPER12*> rest;
        for (size_t i = 2; i < values.size(); ++i) rest.push_back(&values[i]);
        std::wstring base;
        if (folder) {
            Py_ssize_t size = 0;
            wchar_t* text = PyUnicode_AsWideCharString(folder, &size);
            if (!text) return nullptr;
            base.assign(text, static_cast<size_t>(size));
            PyMem_Free(text);
        }
        const excel::Value result = core::call(values[0], values[1], rest, reload_on_change != 0, [&] { return base; });
        return describe(result.get());
    });
}

// column_letters(number) -> str
PyObject* column_letters(PyObject*, PyObject* number) {
    return guarded([&]() -> PyObject* {
        const long n = PyLong_AsLong(number);
        if (n == -1 && PyErr_Occurred()) return nullptr;
        const std::wstring letters = excel::column_letters(static_cast<int>(n));  // long is 32 bits.
        if (letters.empty()) {
            PyErr_SetString(PyExc_ValueError, "not a column number from 1 to 16384");
            return nullptr;
        }
        return PyUnicode_FromWideChar(letters.data(), static_cast<Py_ssize_t>(letters.size()));
    });
}

// column_number(letters) -> int
PyObject* column_number(PyObject*, PyObject* letters) {
    return guarded([&]() -> PyObject* {
        Py_ssize_t size = 0;
        wchar_t* text = PyUnicode_AsWideCharString(letters, &size);
        if (!text) return nullptr;
        const int n = excel::column_number({text, static_cast<size_t>(size)});
        PyMem_Free(text);
        if (n == 0) {
            PyErr_SetString(PyExc_ValueError, "not a column from A to XFD");
            return nullptr;
        }
        return PyLong_FromLong(n);
    });
}

PyMethodDef methods[] = {
    {"to_python", to_python, METH_O, "Converts a described Excel value to Python: (value, problem or None)."},
    {"to_excel", to_excel, METH_O, "Converts a Python value to Excel: (description, problem or None)."},
    {"call", reinterpret_cast<PyCFunction>(reinterpret_cast<void (*)()>(call)), METH_VARARGS | METH_KEYWORDS,
     "PyCall with described values: call(file, function, *args, reload_on_change=True, folder='')."},
    {"column_letters", column_letters, METH_O, "Column number (1 is A) to letters."},
    {"column_number", column_number, METH_O, "Column letters to number (A is 1)."},
    {},
};

PyModuleDef module = {
    PyModuleDef_HEAD_INIT, "_pyinex_testing", "The add-in's value conversion, for tests/unit.", -1, methods,
};

}  // namespace

PyMODINIT_FUNC PyInit__pyinex_testing() {
    if (!core::start_conversion() || !core::start_calls()) {
        PyErr_SetString(PyExc_ImportError, "could not start the conversion or the calls");
        return nullptr;
    }
    Ref m(PyModule_Create(&module));
    if (!m.get() || PyModule_AddObjectRef(m.get(), "ExcelError", core::excel_error_type()) < 0) return nullptr;
    // sources: each of the core's Python files, by name, as the build compiled it in.
    Ref sources(PyDict_New());
    if (!sources.get()) return nullptr;
    for (auto [name, text] : {std::pair{"bootstrap.py", core::python::bootstrap_py}, {"call.py", core::python::call_py},
                              {"excel_error.py", core::python::excel_error_py}, {"funcs.py", core::python::funcs_py}}) {
        Ref value(PyUnicode_FromString(text));
        if (!value.get() || PyDict_SetItemString(sources.get(), name, value.get()) < 0) return nullptr;
    }
    if (PyModule_AddObjectRef(m.get(), "sources", sources.get()) < 0) return nullptr;
    return m.release();
}
