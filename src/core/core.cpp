// pyinex_core.dll: everything that runs Python. pyinex.xll loads it only after loading the
// chosen Python's DLLs, so its link to python3.dll resolves to that Python.

#include "common/environment.h"
#include "core/api.h"
#include "core/call.h"
#include "core/convert.h"
#include "core/module.h"
#include "core/py.h"
#include "excel/excel.h"

#include <string>

using namespace pyinex::excel;
using pyinex::core::Ref;

#define PYINEX_EXPORT extern "C" __declspec(dllexport)

namespace {

// Python code run once at start-up. Its namespace holds the helpers the C++ below calls.
constexpr const char* bootstrap = R"(
import os, sys

# Output goes to Pyinex's console window through a pipe, to a captured stream, or to a
# terminal: in UTF-8 whatever the code page, and a line at a time, so print() from a thread
# a script started shows without waiting for a buffer to fill.
for stream in (sys.stdout, sys.stderr):
    try:
        stream.reconfigure(encoding='utf-8', errors='backslashreplace', line_buffering=True)
    except Exception:
        pass

def fix_up(base_executable, host_executable):
    # Python sets sys.executable from __PYVENV_LAUNCHER__ but takes these two from the
    # process's own .exe (Excel): make them what the chosen python.exe would have. Excel's
    # folder stays on sys.path if it is one of Python's own or PYTHONPATH names it.
    sys._base_executable = base_executable
    def key(p):
        return os.path.normcase(os.path.abspath(p))
    host = key(os.path.dirname(host_executable))
    keep = {key(sys.prefix), key(sys.base_prefix)}
    keep.update(key(p) for p in os.environ.get('PYTHONPATH', '').split(os.pathsep) if p)
    if host not in keep:
        sys.path[:] = [p for p in sys.path if not p or key(p) != host]

def info(item):
    # PyInfo's items that only Python knows.
    if item == 'version':
        return sys.version
    if item == 'prefix':
        return sys.prefix
    return sys.prefix != sys.base_prefix  # 'venv'
)";

PyObject* helpers;  // The bootstrap's namespace.

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
    PyObject* f = PyDict_GetItemString(helpers, name);  // Borrowed.
    PyObject* result = f && args ? PyObject_CallObject(f, args) : nullptr;
    if (!result) PyErr_Clear();
    return result;
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
            helpers = PyDict_New();
            Ref builtins(PyImport_ImportModule("builtins"));
            Ref ran(builtins.get() && helpers
                        ? PyObject_CallMethod(builtins.get(), "exec", "sO", bootstrap, helpers)
                        : nullptr);
            Ref args(Py_BuildValue("(uu)", base_executable, module_path(nullptr).c_str()));
            Ref fixed(ran.get() ? call_helper("fix_up", args.get()) : nullptr);
            ok = fixed.get() != nullptr && pyinex::core::start_conversion() && pyinex::core::start_module() &&
                 pyinex::core::start_calls();
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
        return value.to_excel();
    });
}

PYINEX_EXPORT void pyinex_close() {
    try {
        Gil gil;
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
