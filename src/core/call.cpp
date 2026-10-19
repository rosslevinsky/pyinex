#include "core/call.h"

#include "console/console.h"
#include "core/convert.h"
#include "core/python/call.py.h"
#include "core/python/funcs.py.h"

#include <filesystem>
#include <span>
#include <utility>
#include <vector>

namespace pyinex::core {
namespace {

PyObject* helpers;  // The namespace of src/core/python/call.py, PyCall's Python half.

int type_of(const XLOPER12& x) { return x.xltype & ~(xlbitDLLFree | xlbitXLFree); }

std::wstring text_of(const XLOPER12& x) { return {x.val.str + 1, static_cast<size_t>(x.val.str[0])}; }

PyObject* to_str(const std::wstring& s) { return PyUnicode_FromWideChar(s.data(), static_cast<Py_ssize_t>(s.size())); }

// A reference Excel passed (xltypeRef, or xltypeSRef on the sheet calculating) as a
// pyinex.Range: a new reference, or null with problem saying why. Asks Excel for the
// sheet's name, so only during a call.
PyObject* reference(const XLOPER12& x, std::wstring& problem) {
    const XLMREF12* m = type_of(x) == xltypeRef ? x.val.mref.lpmref : nullptr;
    const auto areas = type_of(x) == xltypeSRef ? std::span(&x.val.sref.ref, 1)
                                                : std::span(m ? m->reftbl : nullptr, m ? m->count : 0);
    // Its sheet's name, and for an xltypeSRef, which has none, its sheet's ID by that name.
    excel::ExcelValue name, id;
    if (areas.empty() || excel::call(xlSheetNm, &name, {&x}) != xlretSuccess || type_of(name.get()) != xltypeStr ||
        (!m && (excel::call(xlSheetId, &id, {&name.get()}) != xlretSuccess || type_of(id.get()) != xltypeRef))) {
        // A call Excel stopped says nothing: Excel calls it again (excel::stop_code).
        if (excel::stop_code <= 0) problem = L"Excel did not say which sheet the reference is on";
        return nullptr;
    }
    // Each area is made before the tuple holding them, which then fills with no Python
    // allocation: a garbage collection meanwhile must never find a tuple still being filled.
    // Room is made first, so each area is held as soon as it is made.
    std::vector<Ref> made;
    made.reserve(areas.size());
    for (const XLREF12& a : areas) {
        Ref area(Py_BuildValue("(iiii)", a.rwFirst, a.rwLast, a.colFirst, a.colLast));
        if (!area.get()) break;
        made.push_back(std::move(area));
    }
    Ref list(made.size() == areas.size() ? PyTuple_New(static_cast<Py_ssize_t>(areas.size())) : nullptr);
    bool ok = list.get() != nullptr;
    for (size_t i = 0; ok && i < areas.size(); ++i)
        ok = PyTuple_SetItem(list.get(), static_cast<Py_ssize_t>(i), made[i].release()) == 0;  // Steals it.
    Ref sheet(PyUnicode_FromWideChar(name.get().val.str + 1, name.get().val.str[0]));
    const auto sheet_id = static_cast<unsigned long long>(m ? x.val.mref.idSheet : id.get().val.mref.idSheet);
    PyObject* range = ok && sheet.get() && helper("Range")
                          ? PyObject_CallFunction(helper("Range"), "OKO", sheet.get(), sheet_id, list.get())
                          : nullptr;
    if (!range && excel::stop_code <= 0) problem = L"Python failed while making a pyinex.Range";
    return range;
}

// The arguments as a tuple, by the argument rule: trailing arguments the formula left out
// are dropped, and Python binds the rest. One left out before the last given is None, or the
// object missing where one is given, as for a direct function, whose references, Excel's
// only for a pyinex.Range parameter, become pyinex.Range. Null after saying why, each line
// starting with label.
Ref arguments(std::span<const XLOPER12* const> args, const std::wstring& label, PyObject* missing = nullptr) {
    size_t count = args.size();
    while (count > 0 && type_of(*args[count - 1]) == xltypeMissing) --count;
    // Every argument converts before the tuple is made: making an ExcelError runs Python
    // code, which must not see, or keep, a tuple still being filled.
    std::vector<Ref> converted;
    for (size_t i = 0; i < count; ++i) {
        if (missing && type_of(*args[i]) == xltypeMissing) {
            converted.emplace_back(Py_NewRef(missing));
            continue;
        }
        std::wstring problem;
        const int type = type_of(*args[i]);
        converted.emplace_back(missing && (type == xltypeRef || type == xltypeSRef) ? reference(*args[i], problem)
                                                                                     : to_python(*args[i], problem));
        if (!problem.empty()) console::note(label + L"argument " + std::to_wstring(i + 1) + L": " + problem);
        if (!converted.back().get()) return Ref();
    }
    Ref tuple(PyTuple_New(static_cast<Py_ssize_t>(count)));
    bool ok = tuple.get() != nullptr;
    for (size_t i = 0; ok && i < count; ++i)
        ok = PyTuple_SetItem(tuple.get(), static_cast<Py_ssize_t>(i), converted[i].release()) == 0;
    if (ok) return tuple;
    console::note(label + L"Python failed while passing the arguments");
    return Ref();
}

// The workbook's folder, for a relative path only, since asking costs calls into Excel; None
// for an absolute one.
Ref folder_of(const std::wstring& path, const std::function<std::wstring()>& folder) {
    return Ref(std::filesystem::path(path).is_absolute() ? Py_NewRef(Py_None) : to_str(folder()));
}

// A call's result as Excel's: #VALUE! if it failed, having said why or saying it now. A call
// Excel stopped says nothing: Excel discards its result and calls it again (excel::stop_code).
excel::Value result_of(PyObject* result, const std::wstring& label) {
    const bool stopped = excel::stop_code > 0;
    if (!result && !stopped) console::note(label + L"Python failed while calling the function");
    if (!result || result == PyDict_GetItemString(helpers, "FAILED")) {
        PyErr_Clear();
        return excel::Value::error(xlerrValue);
    }
    std::wstring problem;
    excel::Value value;
    try {
        value = to_excel(result, problem);
    } catch (const std::bad_alloc&) {  // A result far larger than memory allows.
        value = excel::Value::error(xlerrValue);
        problem = L"memory ran out while converting the result for Excel";
    }
    if (!problem.empty() && excel::stop_code <= 0) console::note(label + problem);
    return value;
}

}  // namespace

bool start_calls() {
    PyObject* names = run_core_python(python::call_py, "call.py");
    Ref funcs(names ? run_core_python(python::funcs_py, "funcs.py", names) : nullptr);
    if (!funcs.get()) {  // Python's error is left for pyinex_start to report.
        Py_XDECREF(names);
        return false;
    }
    Py_XDECREF(helpers);
    helpers = names;
    return true;
}

PyObject* helper(const char* name) { return helpers ? PyDict_GetItemString(helpers, name) : nullptr; }

void flush() {
    PyObject* f = helpers ? PyDict_GetItemString(helpers, "_flush") : nullptr;  // Borrowed.
    Ref done(f ? PyObject_CallNoArgs(f) : nullptr);
    PyErr_Clear();
}

excel::Value call(const XLOPER12& file, const XLOPER12& function, std::span<const XLOPER12* const> args,
                  bool reload_on_change, const std::function<std::wstring()>& folder) {
    const auto failed = [](const std::wstring& why) {
        if (!why.empty()) console::note(why);
        PyErr_Clear();
        return excel::Value::error(xlerrValue);
    };
    if (type_of(file) != xltypeStr || type_of(function) != xltypeStr)
        return failed(L"PyCall needs the script's path and the function's name as text");
    const std::wstring path = text_of(file), name = text_of(function);
    const std::wstring label = path + L", " + name + L": ";

    Ref tuple = arguments(args, label);
    if (!tuple.get()) return failed(L"");

    Ref base = folder_of(path, folder);
    PyObject* call_fn = helpers ? PyDict_GetItemString(helpers, "call") : nullptr;  // Borrowed.
    Ref py_path(to_str(path)), py_name(to_str(name));
    Ref result(call_fn && base.get() && py_path.get() && py_name.get()
                   ? PyObject_CallFunctionObjArgs(call_fn, py_path.get(), py_name.get(), tuple.get(),
                                                  reload_on_change ? Py_True : Py_False, base.get(), nullptr)
                   : nullptr);
    return result_of(result.get(), label);
}

LPXLOPER12 load_functions(const XLOPER12& file, bool reload_on_change, const std::function<std::wstring()>& folder) {
    if (type_of(file) != xltypeStr) {
        console::note(L"PyLoad needs the script's path as text");
        return excel::Value::error(xlerrValue).to_excel();
    }
    const std::wstring path = text_of(file);
    PyObject* load = helper("load_functions");  // Borrowed.
    Ref py_path(to_str(path)), base = folder_of(path, folder);
    Ref rows(load && py_path.get() && base.get()
                 ? PyObject_CallFunctionObjArgs(load, py_path.get(), base.get(), reload_on_change ? Py_True : Py_False,
                                                nullptr)
                 : nullptr);
    // Rows that do not reach the loader, as when memory runs out converting them or handing
    // them over, are taken back: their functions' names are freed, so the next load offers
    // them again.
    const bool offered = rows.get() && rows.get() != PyDict_GetItemString(helpers, "FAILED");
    const auto take_back = [&] {
        if (PyObject* f = offered ? helper("_unoffer") : nullptr) Ref done(PyObject_CallNoArgs(f));
        PyErr_Clear();
    };
    try {
        excel::Value value = result_of(rows.get(), path + L": ");
        if (offered && type_of(value.get()) == xltypeErr) take_back();
        return value.to_excel();
    } catch (...) {
        take_back();
        throw;
    }
}

LPXLOPER12 changes() {
    PyObject* f = helper("changes");  // Borrowed.
    Ref rows(f ? PyObject_CallNoArgs(f) : nullptr);
    std::wstring problem;
    // What does not reach the loader waits in Python for its next call, put back once.
    bool waits = false;
    const auto put_back = [&] {
        if (PyObject* g = std::exchange(waits, true) ? nullptr : helper("_restore_changes"))
            Ref done(PyObject_CallFunctionObjArgs(g, rows.get(), nullptr));
        PyErr_Clear();
    };
    const bool any = rows.get() && PyObject_Length(rows.get()) > 0;
    try {
        excel::Value value;
        if (any) value = to_excel(rows.get(), problem);
        if (any && type_of(value.get()) == xltypeErr) {
            put_back();
            value = excel::Value();
            console::note(L"Python's changes to the functions of the scripts PyLoad loaded wait: " + problem);
        }
        PyErr_Clear();
        return value.to_excel();
    } catch (...) {
        put_back();
        throw;
    }
}

excel::Value call_function(long key, const std::wstring& name, std::span<const XLOPER12* const> args,
                           bool reload_on_change) {
    const std::wstring label = name + L": ";
    // Left out before the last given, funcs.py's MISSING: its parameter's default applies.
    Ref tuple = arguments(args, label, helper("MISSING"));
    if (!tuple.get()) {
        PyErr_Clear();
        return excel::Value::error(xlerrValue);
    }
    PyObject* call_fn = helper("call_function");  // Borrowed.
    Ref py_key(PyLong_FromLong(key));
    Ref result(call_fn && py_key.get() ? PyObject_CallFunctionObjArgs(call_fn, py_key.get(), tuple.get(),
                                                                       reload_on_change ? Py_True : Py_False, nullptr)
                                       : nullptr);
    return result_of(result.get(), label);
}

}  // namespace pyinex::core
