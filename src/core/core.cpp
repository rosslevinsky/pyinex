// pyinex_core.dll: everything that runs Python. pyinex.xll loads it only after loading the
// chosen Python's DLLs, so its link to python3.dll resolves to that Python.

#include "common/environment.h"
#include "console/console.h"
#include "core/api.h"
#include "core/call.h"
#include "core/convert.h"
#include "core/module.h"
#include "core/py.h"
#include "core/python/bootstrap.py.h"
#include "excel/excel.h"

#include <string>

using namespace pyinex::excel;
using pyinex::core::Ref;

#define PYINEX_EXPORT extern "C" __declspec(dllexport)

namespace {

PyObject* helpers;  // The namespace of src/core/python/bootstrap.py, run at start-up.

// Holds the GIL for one call.
class Gil {
public:
    Gil() : state_(PyGILState_Ensure()) {}
    Gil(const Gil&) = delete;
    Gil& operator=(const Gil&) = delete;
    ~Gil() { PyGILState_Release(state_); }

private:
    PyGILState_STATE state_;
};

// Calls one of the bootstrap's helpers; a new reference, or null with the error cleared.
PyObject* call_helper(const char* name, PyObject* args) {
    PyObject* f = helpers ? PyDict_GetItemString(helpers, name) : nullptr;  // Borrowed.
    PyObject* result = f && args ? PyObject_CallObject(f, args) : nullptr;
    if (!result) PyErr_Clear();
    return result;
}

// A call's result for Excel: #VALUE! if Excel stopped the call (stop_code), decided
// last, so that a stop met while the result converted or the call's Python objects were
// released, by a finalizer, counts. Excel discards that result.
LPXLOPER12 stop_checked(Value value) {
    return (stop_code > 0 ? Value::error(xlerrValue) : std::move(value)).to_excel();
}

std::wstring module_path(HMODULE module) {
    std::wstring path(MAX_PATH, L'\0');
    for (;;) {
        const DWORD n = GetModuleFileNameW(module, path.data(), static_cast<DWORD>(path.size()));
        if (n < path.size()) return path.substr(0, n);
        path.resize(path.size() * 2);
    }
}

}  // namespace

PYINEX_EXPORT int pyinex_start(const wchar_t* executable, const wchar_t* base_executable) {
    try {
        // Python's start-up reads this variable, as a virtual environment's python.exe sets
        // it, and takes sys.executable from it, then finds and reads a pyvenv.cfg beside it,
        // so a virtual environment's packages and settings apply. It clears the variable
        // again. The configuration API that would do this directly is not in the stable
        // ABI, and Py_SetProgramName, besides being deprecated, is ignored by Python 3.10.
        pyinex::common::set_environment(L"__PYVENV_LAUNCHER__", executable);
        Py_InitializeEx(0);  // 0: leave Excel's signal handlers alone.
        bool ok = false;
        {  // Every reference here is released while the GIL is still held.
            const wchar_t* step = L"running bootstrap.py";
            helpers = pyinex::core::run_core_python(pyinex::core::python::bootstrap_py, "bootstrap.py");
            Ref args(Py_BuildValue("(uu)", base_executable, module_path(nullptr).c_str()));
            PyObject* fix_up = helpers ? PyDict_GetItemString(helpers, "fix_up") : nullptr;  // Borrowed.
            Ref fixed(fix_up && args.get() && (step = L"setting up sys") ? PyObject_CallObject(fix_up, args.get()) : nullptr);
            ok = fixed.get() != nullptr && (step = L"making pyinex.ExcelError", pyinex::core::start_conversion()) &&
                 (step = L"running call.py and funcs.py", pyinex::core::start_calls()) &&
                 (step = L"making the pyinex module", pyinex::core::start_module());
            // Said here, as nothing after can say why: Python cannot start again in this Excel.
            if (!ok)
                pyinex::console::note(std::wstring(L"Python started, but Pyinex could not set it up: ") + step +
                                      L" failed (" + pyinex::core::python_error() +
                                      L"). Fix the cause, then restart Excel: Python cannot start again within it.");
            PyErr_Clear();
        }
        PyEval_SaveThread();  // Release the GIL; each call takes it again.
        return ok ? 0 : 1;
    } catch (...) {
        return 1;
    }
}

PYINEX_EXPORT LPXLOPER12 pyinex_call(const XLOPER12* file, const XLOPER12* function, const XLOPER12* const* args,
                                     int count, bool reload_on_change) {
    return guarded([&] {
        Gil gil;
        const pyinex::core::ExcelCall excel_call;  // Made and ended with the GIL held.
        Value value = pyinex::core::call(*file, *function, {args, static_cast<size_t>(count)}, reload_on_change,
                                         caller_folder);
        // After the call has released its Python objects, so what their finalizers print
        // is written out too.
        pyinex::core::flush();
        return stop_checked(std::move(value));
    });
}

PYINEX_EXPORT LPXLOPER12 pyinex_load(const XLOPER12* file, bool reload_on_change) {
    return guarded([&] {
        Gil gil;
        const pyinex::core::OutsideCall outside;  // The script runs outside any call, even PyLoad within one.
        const LPXLOPER12 rows = pyinex::core::load_functions(*file, reload_on_change, caller_folder);
        pyinex::core::flush();
        return rows;
    });
}

PYINEX_EXPORT LPXLOPER12 pyinex_changes() {
    return guarded([&] {
        Gil gil;
        return pyinex::core::changes();
    });
}

PYINEX_EXPORT void pyinex_registered(int key, bool registered) {
    try {
        Gil gil;
        if (PyObject* f = pyinex::core::helper("registered"))  // Borrowed.
            Ref done(PyObject_CallFunction(f, "ii", key, registered ? 1 : 0));
        PyErr_Clear();
    } catch (...) {
    }
}

PYINEX_EXPORT LPXLOPER12 pyinex_function(int key, const wchar_t* name, const XLOPER12* const* args, int count,
                                         bool reload_on_change) {
    return guarded([&] {
        Gil gil;
        const pyinex::core::ExcelCall excel_call;  // As pyinex_call's.
        Value value =
            pyinex::core::call_function(key, name, {args, static_cast<size_t>(count)}, reload_on_change);
        pyinex::core::flush();
        return stop_checked(std::move(value));
    });
}

PYINEX_EXPORT void pyinex_close() {
    try {
        Gil gil;
        const pyinex::core::OutsideCall outside;  // Even should Excel close the add-in within a call.
        pyinex::core::flush();
    } catch (...) {
    }
}

PYINEX_EXPORT LPXLOPER12 pyinex_info(const wchar_t* item) {
    return guarded([&] {
        Gil gil;
        Ref args(Py_BuildValue("(u)", item));
        Ref result(call_helper("info", args.get()));
        std::wstring problem;
        Value value = result.get() ? pyinex::core::to_excel(result.get(), problem) : Value::error(xlerrValue);
        PyErr_Clear();
        return value.to_excel();
    });
}
