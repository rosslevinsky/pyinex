#include "core/call.h"

#include "console/console.h"
#include "core/convert.h"

#include <filesystem>
#include <vector>

namespace pyinex::core {
namespace {

// The Python half of PyCall. Every exception, SystemExit and KeyboardInterrupt included,
// is caught here and printed with the traceback module: never with PyErr_Print, which
// exits the process on SystemExit.
constexpr const char* source = R"(
import importlib.machinery, importlib.util, os, re, sys, traceback, zlib

FAILED = object()  # call's result when it has already reported why
_scripts = {}      # os.path.normcase(path) -> (module, st_mtime_ns, st_size)
_folders = set()   # folders put on sys.path, normcased

def _say(text):
    try:
        print('Pyinex: ' + text, file=sys.stderr)
    except BaseException:
        pass

def _traceback():
    # The exception being handled, without the leading frames of this code, which the
    # user did not write.
    try:
        kind, error, tb = sys.exc_info()
        while tb is not None and tb.tb_frame.f_code.co_filename == '<pyinex>':
            tb = tb.tb_next
        traceback.print_exception(kind, error, tb)
    except BaseException:
        pass

def _flush():
    # So output reaches the console after each call, not whenever a buffer fills.
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None:
                stream.flush()
        except BaseException:
            pass

def _forget(key):
    entry = _scripts.pop(key, None)
    if entry is not None and sys.modules.get(entry[0].__name__) is entry[0]:
        del sys.modules[entry[0].__name__]

def _load(path, reload_on_change):
    # The script's module, loaded or reloaded as needed; None, having said why, if the
    # file is not there.
    key = os.path.normcase(path)
    entry = _scripts.get(key)
    if entry is not None and not reload_on_change:
        return entry[0]
    try:
        st = os.stat(path)
    except OSError:
        st = None
    if st is None or not os.path.isfile(path):
        _forget(key)
        _say(f'cannot find the script {path}')
        return None
    if entry is not None:
        if (st.st_mtime_ns, st.st_size) == entry[1:]:
            return entry[0]
        _forget(key)  # A failed reload must not leave the old version running.
    folder = os.path.dirname(path)
    if os.path.normcase(folder) not in _folders:
        sys.path.insert(0, folder)  # So helpers beside the script import.
        _folders.add(os.path.normcase(folder))
    stem = re.sub(r'\W', '_', os.path.splitext(os.path.basename(path))[0])
    name = 'pyinex_%s_%08x' % (stem, zlib.crc32(key.encode('utf-8', 'surrogatepass')))
    # The loader named, since importlib would choose one by the suffix, case-sensitively,
    # and find none for model.PY.
    loader = importlib.machinery.SourceFileLoader(name, path)
    module = importlib.util.module_from_spec(importlib.util.spec_from_file_location(name, path, loader=loader))
    # Compiled here from the source, never from a cached .pyc, whose timestamp is
    # only to the second.
    with open(path, 'rb') as f:
        code = compile(f.read(), path, 'exec', dont_inherit=True)
    sys.modules[name] = module
    try:
        exec(code, module.__dict__)
    except BaseException:
        if sys.modules.get(name) is module:
            del sys.modules[name]
        raise
    _scripts[key] = (module, st.st_mtime_ns, st.st_size)
    return module

def call(file, name, args, reload_on_change, folder):
    try:
        if folder is not None:  # file is relative: None when it is absolute.
            if not folder:
                _say(f'cannot find the script {file}: a relative path is relative to the '
                     'workbook\'s folder, and this workbook has not been saved')
                return FAILED
            file = os.path.join(folder, file)
        path = os.path.abspath(file)
        try:
            module = _load(path, reload_on_change)
        except BaseException:
            _say(f'the script {path} failed to load:')
            _traceback()
            return FAILED
        if module is None:
            return FAILED
        function = vars(module).get(name, FAILED)
        if function is FAILED:
            _say(f'the script {path} has no function {name}')
            return FAILED
        if not callable(function):
            _say(f'{name} in the script {path} is not a function')
            return FAILED
        return function(*args)
    except BaseException:
        _say(f'{name} in the script {file} raised an exception:')
        _traceback()
        return FAILED
)";

PyObject* helpers;  // The namespace source ran in.

int type_of(const XLOPER12& x) { return x.xltype & ~(xlbitDLLFree | xlbitXLFree); }

std::wstring text_of(const XLOPER12& x) { return {x.val.str + 1, static_cast<size_t>(x.val.str[0])}; }

PyObject* to_str(const std::wstring& s) { return PyUnicode_FromWideChar(s.data(), static_cast<Py_ssize_t>(s.size())); }

}  // namespace

bool start_calls() {
    // Named <pyinex>, so tracebacks can leave this code's frames out.
    Ref code(Py_CompileString(source, "<pyinex>", Py_file_input));
    Ref names(PyDict_New());
    const bool ready = code.get() && names.get() &&
                       PyDict_SetItemString(names.get(), "__builtins__", PyEval_GetBuiltins()) == 0;
    Ref ran(ready ? PyEval_EvalCode(code.get(), names.get(), names.get()) : nullptr);
    if (!ran.get()) {
        PyErr_Clear();
        return false;
    }
    Py_XDECREF(helpers);
    helpers = names.release();
    return true;
}

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

    // Trailing arguments the formula left out are dropped; Python binds the rest.
    size_t count = args.size();
    while (count > 0 && type_of(*args[count - 1]) == xltypeMissing) --count;
    // Every argument converts before the tuple is made: making an ExcelError runs Python
    // code, which must not see, or keep, a tuple still being filled.
    std::vector<Ref> converted;
    for (size_t i = 0; i < count; ++i) {
        std::wstring problem;
        converted.emplace_back(to_python(*args[i], problem));
        if (!problem.empty()) console::note(label + L"argument " + std::to_wstring(i + 1) + L": " + problem);
        if (!converted.back().get()) return failed(L"");
    }
    Ref arguments(PyTuple_New(static_cast<Py_ssize_t>(count)));
    if (!arguments.get()) return failed(label + L"Python failed while passing the arguments");
    for (size_t i = 0; i < count; ++i)
        if (PyTuple_SetItem(arguments.get(), static_cast<Py_ssize_t>(i), converted[i].release()) < 0)
            return failed(label + L"Python failed while passing the arguments");

    // The workbook's folder, for a relative path only, since asking costs calls into Excel.
    Ref base(std::filesystem::path(path).is_absolute() ? Py_NewRef(Py_None) : to_str(folder()));
    PyObject* call_fn = helpers ? PyDict_GetItemString(helpers, "call") : nullptr;  // Borrowed.
    Ref py_path(to_str(path)), py_name(to_str(name));
    Ref result(call_fn && base.get() && py_path.get() && py_name.get()
                   ? PyObject_CallFunctionObjArgs(call_fn, py_path.get(), py_name.get(), arguments.get(),
                                                  reload_on_change ? Py_True : Py_False, base.get(), nullptr)
                   : nullptr);
    if (!result.get()) return failed(label + L"Python failed while calling the function");
    if (result.get() == PyDict_GetItemString(helpers, "FAILED")) return failed(L"");
    std::wstring problem;
    excel::Value value = to_excel(result.get(), problem);
    if (!problem.empty()) console::note(label + problem);
    return value;
}

}  // namespace pyinex::core
