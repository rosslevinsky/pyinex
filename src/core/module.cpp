#include "core/module.h"

#include "core/convert.h"
#include "core/py.h"
#include "excel/address.h"
#include "excel/excel.h"

#include <optional>
#include <string>

namespace pyinex::core {
namespace {

using excel::ExcelValue;
using excel::Value;

DWORD calling_thread;  // The thread Excel is calling the add-in on; 0 between calls.

// Raises RuntimeError, and returns false, unless the module may ask Excel now. Excel's C API
// may be called only on the thread Excel called the add-in on, while it waits for the
// result; anything else can crash it.
bool may_ask_excel(const char* function) {
    if (calling_thread != 0 && calling_thread == GetCurrentThreadId()) return true;
    PyErr_Format(PyExc_RuntimeError,
                 calling_thread != 0
                     ? "pyinex.%s asks Excel, which allows that only on Excel's own thread, not one a script started"
                     : "pyinex.%s asks Excel, which allows that only while Excel is calculating a PyCall",
                 function);
    return false;
}

int type_of(const XLOPER12& x) { return x.xltype & ~(xlbitDLLFree | xlbitXLFree); }

PyObject* to_str(const std::wstring& s) { return PyUnicode_FromWideChar(s.data(), static_cast<Py_ssize_t>(s.size())); }

// The calling cell or range, as xlfCaller gives it: a reference with its sheet (xltypeRef)
// or without (xltypeSRef). Null, with RuntimeError raised, if the caller is not a cell.
const XLREF12* calling_area(const ExcelValue& caller) {
    const XLOPER12& x = caller.get();
    if (type_of(x) == xltypeSRef) return &x.val.sref.ref;
    if (type_of(x) == xltypeRef && x.val.mref.lpmref && x.val.mref.lpmref->count >= 1)
        return &x.val.mref.lpmref->reftbl[0];
    PyErr_SetString(PyExc_RuntimeError, "the PyCall running was not called from a cell");
    return nullptr;
}

// The workbook and sheet of the caller, such as [Book.xlsx]Sheet1, or nullopt with
// RuntimeError raised.
std::optional<std::wstring> sheet_of(const ExcelValue& caller) {
    ExcelValue name;
    if (excel::call(xlSheetNm, &name, {&caller.get()}) != xlretSuccess || type_of(name.get()) != xltypeStr) {
        PyErr_SetString(PyExc_RuntimeError, "Excel did not say which sheet the calling cell is on");
        return std::nullopt;
    }
    return std::wstring(name.get().val.str + 1, name.get().val.str[0]);
}

std::wstring cell(bool r1c1, int row, int column) {
    // XLREF12 counts from 0.
    if (r1c1) return L"R" + std::to_wstring(row + 1) + L"C" + std::to_wstring(column + 1);
    return excel::column_letters(column + 1) + std::to_wstring(row + 1);
}

// No C++ exception may reach Python, which would end Excel.
template <class F>
PyObject* guarded(F body) {
    try {
        return body();
    } catch (...) {
        return PyErr_NoMemory();
    }
}

PyObject* caller(PyObject*, PyObject* args, PyObject* keywords) {
    static const char* names[] = {"r1c1", "full", nullptr};
    int r1c1 = 0, full = 0;
    if (!PyArg_ParseTupleAndKeywords(args, keywords, "|pp:caller", const_cast<char**>(names), &r1c1, &full) ||
        !may_ask_excel("caller"))
        return nullptr;
    return guarded([&]() -> PyObject* {
        ExcelValue calling;
        excel::call(xlfCaller, &calling);
        const XLREF12* area = calling_area(calling);
        if (!area) return nullptr;
        std::wstring address = cell(r1c1, area->rwFirst, area->colFirst);
        if (area->rwLast != area->rwFirst || area->colLast != area->colFirst)
            address += L":" + cell(r1c1, area->rwLast, area->colLast);
        if (!full) return to_str(address);
        const auto sheet = sheet_of(calling);
        return sheet ? to_str(*sheet + L"!" + address) : nullptr;
    });
}

PyObject* sheet(PyObject*, PyObject*) {
    if (!may_ask_excel("sheet")) return nullptr;
    return guarded([&]() -> PyObject* {
        ExcelValue calling;
        excel::call(xlfCaller, &calling);
        if (!calling_area(calling)) return nullptr;
        const auto sheet = sheet_of(calling);
        return sheet ? to_str(*sheet) : nullptr;
    });
}

PyObject* esc_pressed(PyObject*, PyObject* args, PyObject* keywords) {
    static const char* names[] = {"clear", nullptr};
    int clear = 0;
    if (!PyArg_ParseTupleAndKeywords(args, keywords, "|p:esc_pressed", const_cast<char**>(names), &clear) ||
        !may_ask_excel("esc_pressed"))
        return nullptr;
    return guarded([&]() -> PyObject* {
        // xlAbort's argument is the opposite: TRUE keeps the key press.
        const Value keep = Value::boolean(!clear);
        ExcelValue pressed;
        if (excel::call(xlAbort, &pressed, {&keep.get()}) != xlretSuccess || type_of(pressed.get()) != xltypeBool) {
            PyErr_SetString(PyExc_RuntimeError, "Excel did not say whether Esc was pressed");
            return nullptr;
        }
        return PyBool_FromLong(pressed.get().val.xbool);
    });
}

// PyCFunctionWithKeywords is cast to PyCFunction, as METH_KEYWORDS requires.
template <class F>
PyCFunction as_function(F f) {
    return reinterpret_cast<PyCFunction>(reinterpret_cast<void (*)()>(f));
}

PyMethodDef methods[] = {
    {"caller", as_function(caller), METH_VARARGS | METH_KEYWORDS,
     "caller(r1c1=False, full=False)\n--\n\n"
     "The address of the cell whose PyCall is running, such as C12, or R12C3 with r1c1=True.\n"
     "With full=True the workbook and sheet come first: [Book.xlsx]Sheet1!C12."},
    {"sheet", sheet, METH_NOARGS, "sheet()\n--\n\nThe calling cell's workbook and sheet, such as [Book.xlsx]Sheet1."},
    {"esc_pressed", as_function(esc_pressed), METH_VARARGS | METH_KEYWORDS,
     "esc_pressed(clear=False)\n--\n\n"
     "Whether the user has pressed Esc during this calculation. clear=True makes Excel forget\n"
     "the key press, so other cells checking see False until Esc is pressed again."},
    {nullptr, nullptr, 0, nullptr},
};

PyModuleDef definition = {
    PyModuleDef_HEAD_INIT, "pyinex", "Pyinex's functions for scripts that Excel calls through PyCall.", -1, methods,
    nullptr, nullptr, nullptr, nullptr,
};

}  // namespace

bool start_module() {
    Ref module(PyModule_Create(&definition));
    PyObject* modules = PyImport_GetModuleDict();  // Borrowed.
    const bool ok = module.get() && PyModule_AddObjectRef(module.get(), "ExcelError", excel_error_type()) == 0 &&
                    PyDict_SetItemString(modules, "pyinex", module.get()) == 0;
    if (!ok) PyErr_Clear();
    return ok;
}

ExcelCall::ExcelCall() : outer_(calling_thread) { calling_thread = GetCurrentThreadId(); }

ExcelCall::~ExcelCall() { calling_thread = outer_; }

}  // namespace pyinex::core
