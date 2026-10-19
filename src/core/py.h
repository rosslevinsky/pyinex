// Python's C API, as everything that links Python includes it, and a reference holder.
//
// Only the stable ABI as of Python 3.10 is used (Py_LIMITED_API, set by CMake), so one
// build serves every Python from 3.10 on.
#pragma once

// A debug build of the add-in still uses a release Python: without this, Python's headers
// would see _DEBUG, switch to the debug build's ABI and ask for python3_d.lib.
#ifdef _DEBUG
#undef _DEBUG
#include <Python.h>
#define _DEBUG
#else
#include <Python.h>
#endif

#include <string>

namespace pyinex::core {

// Owns one reference.
class Ref {
public:
    explicit Ref(PyObject* p = nullptr) : p_(p) {}
    Ref(Ref&& other) noexcept : p_(other.p_) { other.p_ = nullptr; }
    Ref(const Ref&) = delete;
    Ref& operator=(const Ref&) = delete;
    ~Ref() { Py_XDECREF(p_); }
    PyObject* get() const { return p_; }
    PyObject* release() {
        PyObject* p = p_;
        p_ = nullptr;
        return p;
    }

private:
    PyObject* p_;
};

// Runs one of the core's Python files (src/core/python/), given its text, in a new
// namespace, or in the namespace into: that namespace, a new reference, or null with
// Python's error set. It is compiled as <pyinex/name>, a file name that marks it as
// Pyinex's own, so tracebacks can leave its frames out (docs/architecture.md, "The core's
// Python files").
inline PyObject* run_core_python(const char* text, const char* name, PyObject* into = nullptr) {
    Ref code(Py_CompileString(text, (std::string("<pyinex/") + name + ">").c_str(), Py_file_input));
    Ref names(into ? Py_NewRef(into) : PyDict_New());
    if (!code.get() || !names.get() || PyDict_SetItemString(names.get(), "__builtins__", PyEval_GetBuiltins()) < 0)
        return nullptr;
    Ref ran(PyEval_EvalCode(code.get(), names.get(), names.get()));
    return ran.get() ? names.release() : nullptr;
}

// The error Python has set, as its type's name and text, which it clears; for the console.
inline std::wstring python_error() {
    PyObject *type = nullptr, *value = nullptr, *traceback = nullptr;
    PyErr_Fetch(&type, &value, &traceback);
    if (!type) return L"no Python error";
    PyErr_NormalizeException(&type, &value, &traceback);
    const Ref held_type(type), held_value(value), held_traceback(traceback);
    const auto wide = [](PyObject* s) {
        std::wstring out;
        Py_ssize_t size = 0;
        wchar_t* chars = s && PyUnicode_Check(s) ? PyUnicode_AsWideCharString(s, &size) : nullptr;
        if (chars) out.assign(chars, static_cast<size_t>(size));
        PyMem_Free(chars);
        return out;
    };
    const Ref name(PyObject_GetAttrString(type, "__name__")), text(value ? PyObject_Str(value) : nullptr);
    PyErr_Clear();
    const std::wstring said = wide(text.get());
    return wide(name.get()) + (said.empty() ? L"" : L": " + said);
}

}  // namespace pyinex::core
