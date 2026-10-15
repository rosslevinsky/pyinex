#include "excel/excel.h"

#include <algorithm>
#include <cstring>
#include <cwchar>
#include <utility>

namespace pyinex::excel {
namespace {

// Calls into Excel through the MdCallBack12 that Excel exports from its own program file,
// found once (docs/architecture.md, "Calling Excel"). Fails outside Excel.
int call_excel(int function, LPXLOPER12 result, int count, LPXLOPER12* args) {
    using Callback = int(__stdcall*)(int, int, LPXLOPER12*, LPXLOPER12);
    static const auto callback = reinterpret_cast<Callback>(GetProcAddress(GetModuleHandleW(nullptr), "MdCallBack12"));
    return callback ? callback(function, count, args, result) : xlretFailed;
}

// Frees whatever x points to (text, array items), not x itself.
void free_contents(XLOPER12& x) {
    switch (x.xltype & ~(xlbitDLLFree | xlbitXLFree)) {
    case xltypeStr:
        delete[] x.val.str;
        break;
    case xltypeMulti: {
        const auto count = static_cast<size_t>(x.val.array.rows) * x.val.array.columns;
        for (size_t i = 0; i < count; ++i) free_contents(x.val.array.lparray[i]);
        delete[] x.val.array.lparray;
        break;
    }
    default:
        break;
    }
    x.xltype = xltypeNil;
}

}  // namespace

Value Value::number(double n) {
    Value v;
    v.x_.xltype = xltypeNum;
    v.x_.val.num = n;
    return v;
}

Value Value::text(std::wstring_view s) {
    const auto length = static_cast<int>(std::min<size_t>(s.size(), max_text));
    Value v;
    // Excel text is length-prefixed: element 0 is the length, and there is no terminator.
    auto* buffer = new XCHAR[length + 1];
    buffer[0] = static_cast<XCHAR>(length);
    std::memcpy(buffer + 1, s.data(), length * sizeof(XCHAR));
    v.x_.xltype = xltypeStr;
    v.x_.val.str = buffer;
    return v;
}

Value Value::missing() {
    Value v;
    v.x_.xltype = xltypeMissing;
    return v;
}

Value Value::boolean(bool b) {
    Value v;
    v.x_.xltype = xltypeBool;
    v.x_.val.xbool = b;
    return v;
}

Value Value::error(int code) {
    Value v;
    v.x_.xltype = xltypeErr;
    v.x_.val.err = code;
    return v;
}

Value Value::array(int rows, int columns, std::vector<Value> items) {
    Value v;
    auto* cells = new XLOPER12[items.size()];
    for (size_t i = 0; i < items.size(); ++i) cells[i] = items[i].release();
    v.x_.xltype = xltypeMulti;
    v.x_.val.array = {.lparray = cells, .rows = rows, .columns = columns};
    return v;
}

Value Value::copy(const XLOPER12& x) {
    switch (x.xltype & ~(xlbitDLLFree | xlbitXLFree)) {
    case xltypeNum: return number(x.val.num);
    case xltypeStr: return text({x.val.str + 1, static_cast<size_t>(x.val.str[0])});
    case xltypeBool: return boolean(x.val.xbool != 0);
    case xltypeErr: return error(x.val.err);
    case xltypeInt: return number(x.val.w);
    case xltypeMissing:
    case xltypeNil: return Value();
    case xltypeMulti: {
        const auto& a = x.val.array;
        std::vector<Value> items;
        items.reserve(static_cast<size_t>(a.rows) * a.columns);
        for (size_t i = 0; i < static_cast<size_t>(a.rows) * a.columns; ++i)
            items.push_back(copy(a.lparray[i]));
        return array(a.rows, a.columns, std::move(items));
    }
    default: return error(xlerrValue);
    }
}

Value::Value(Value&& other) noexcept : x_(std::exchange(other.x_, {.xltype = xltypeNil})) {}

Value& Value::operator=(Value&& other) noexcept {
    if (this != &other) {
        free_contents(x_);
        x_ = std::exchange(other.x_, {.xltype = xltypeNil});
    }
    return *this;
}

Value::~Value() { free_contents(x_); }

XLOPER12 Value::release() { return std::exchange(x_, {.xltype = xltypeNil}); }

LPXLOPER12 Value::to_excel() {
    auto* result = new XLOPER12(release());
    result->xltype |= xlbitDLLFree;
    return result;
}

void free_returned(LPXLOPER12 x) {
    free_contents(*x);
    delete x;
}

ExcelValue::~ExcelValue() { reset(); }

void ExcelValue::reset() {
    if (owned_) {
        XLOPER12* args[] = {&x_};
        call_excel(xlFree, nullptr, 1, args);
        owned_ = false;
    }
    x_ = {.xltype = xltypeNil};
}

int call(int function, ExcelValue* result, std::span<const XLOPER12* const> args) {
    // Excel's callback is not const-correct; Excel does not change its arguments.
    std::vector<LPXLOPER12> opers;
    opers.reserve(args.size());
    for (const auto* a : args) opers.push_back(const_cast<LPXLOPER12>(a));
    if (result) result->reset();  // A reused result frees what it held.
    XLOPER12 scratch{.xltype = xltypeNil};
    XLOPER12* out = result ? &result->x_ : &scratch;
    const int ret = call_excel(function, out, static_cast<int>(opers.size()), opers.data());
    if (ret == xlretSuccess) {
        if (result) result->owned_ = true;
        else {
            XLOPER12* free_args[] = {&scratch};
            call_excel(xlFree, nullptr, 1, free_args);
        }
    }
    return ret;
}

int call(int function, ExcelValue* result, std::initializer_list<const XLOPER12*> args) {
    return call(function, result, std::span(args.begin(), args.size()));
}

std::wstring caller_folder() {
    ExcelValue caller, name;
    // A reference with its sheet (xltypeRef) or without (xltypeSRef, the sheet calculating):
    // Excel may give either, and CELL takes both.
    if (call(xlfCaller, &caller) != xlretSuccess) return L"";
    if (const int type = caller.get().xltype & ~xlbitXLFree; type != xltypeRef && type != xltypeSRef) return L"";
    // C:\folder\[Book.xlsx]Sheet1; C:\folder\Book.xlsx when the workbook's only sheet has its
    // name; or empty text for a workbook never saved.
    const Value info = Value::text(L"filename");
    if (call(xlfCell, &name, {&info.get(), &caller.get()}) != xlretSuccess ||
        (name.get().xltype & ~xlbitXLFree) != xltypeStr)
        return L"";
    const std::wstring_view text(name.get().val.str + 1, name.get().val.str[0]);
    // The folder ends at the last \: neither a workbook's name nor a sheet's may hold one.
    const auto slash = text.rfind(L'\\');
    return slash == std::wstring_view::npos ? L"" : std::wstring(text.substr(0, slash + 1));
}

bool in_function_wizard() {
    bool found = false;
    EnumWindows(
        [](HWND window, LPARAM param) -> BOOL {
            DWORD process = 0;
            wchar_t name[32] = {}, title[64] = {};
            GetWindowThreadProcessId(window, &process);
            if (process != GetCurrentProcessId() || !GetClassNameW(window, name, 32) ||
                std::wcsncmp(name, L"bosa_sdm_XL", 11) != 0)
                return TRUE;
            GetWindowTextW(window, title, 64);
            bool& found = *reinterpret_cast<bool*>(param);
            found = std::wcsstr(title, L"Function Arguments") != nullptr;
            return !found;
        },
        reinterpret_cast<LPARAM>(&found));
    return found;
}

int register_functions(std::span<const Function> functions) {
    ExcelValue module;  // The add-in's own path, which xlfRegister needs first.
    if (call(xlGetName, &module) != xlretSuccess) return 0;

    int registered = 0;
    for (const auto& f : functions) {
        std::vector<Value> values;
        values.push_back(Value::text(f.procedure));
        values.push_back(Value::text(f.type));
        values.push_back(Value::text(f.name));
        values.push_back(Value::text(f.arguments));
        values.push_back(Value::number(1));  // Macro type 1: a worksheet function.
        values.push_back(Value::text(L"Pyinex"));
        values.push_back(Value::missing());  // No shortcut key.
        values.push_back(Value::missing());  // No help file.
        values.push_back(Value::text(f.help));
        for (const auto& h : f.argument_help) values.push_back(Value::text(h));

        std::vector<const XLOPER12*> args{&module.get()};
        for (const auto& v : values) args.push_back(&v.get());
        ExcelValue id;
        if (call(xlfRegister, &id, args) == xlretSuccess && id.get().xltype == xltypeNum)
            ++registered;
    }
    return registered;
}

}  // namespace pyinex::excel
