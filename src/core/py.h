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

}  // namespace pyinex::core
