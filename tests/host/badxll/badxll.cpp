// A deliberately wrong add-in. Each function breaks one of Excel's memory rules, so the
// tests can check that xllhost notices (tests/host/test_memory_rules.py).
#include <windows.h>

#include <XLCALL.H>

#include <initializer_list>
#include <vector>

#define BAD_EXPORT extern "C" __declspec(dllexport)

namespace {

// Calls into Excel, as the add-in's own call_excel does (docs/architecture.md, "Calling Excel").
int excel(int function, LPXLOPER12 result, std::initializer_list<const XLOPER12*> args = {}) {
    using Callback = int(__stdcall*)(int, int, LPXLOPER12*, LPXLOPER12);
    static const auto callback = reinterpret_cast<Callback>(GetProcAddress(GetModuleHandleW(nullptr), "MdCallBack12"));
    std::vector<LPXLOPER12> opers;
    for (const auto* a : args) opers.push_back(const_cast<LPXLOPER12>(a));
    return callback ? callback(function, static_cast<int>(opers.size()), opers.data(), result) : xlretFailed;
}

void register_function(const XLOPER12& module, const wchar_t* name, const wchar_t* type_string = L"Q") {
    // Excel text is length-prefixed; these names are short ASCII.
    static XCHAR buffers[14][64];  // Two per function.
    static int used = 0;
    auto text = [&](const wchar_t* s) {
        XCHAR* b = buffers[used++];
        int n = 0;
        while (s[n]) { b[n + 1] = s[n]; ++n; }
        b[0] = static_cast<XCHAR>(n);
        XLOPER12 x{.xltype = xltypeStr};
        x.val.str = b;
        return x;
    };
    XLOPER12 procedure = text(name), type = text(type_string);
    XLOPER12 id;
    excel(xlfRegister, &id, {&module, &procedure, &type, &procedure});
}

}  // namespace

BAD_EXPORT int xlAutoOpen() {
    XLOPER12 module;
    excel(xlGetName, &module);
    register_function(module, L"BadNoDllFree");
    register_function(module, L"BadDoubleFree");
    register_function(module, L"BadLeak");
    register_function(module, L"BadNestedLeak");
    register_function(module, L"BadSwapFree");
    register_function(module, L"BadArity", L"QQ");  // One argument, but the export takes two.
    register_function(module, L"BadChangesArgument", L"QQ");
    excel(xlFree, nullptr, {&module});
    return 1;
}

LPXLOPER12 swap_result = nullptr;
char* swap_cache = nullptr;

// Frees the XLOPER12 but not its text: wrong for BadNestedLeak's and BadSwapFree's
// results. For BadSwapFree's it frees an unrelated block instead, so the number of blocks
// freed is right while the text still leaks.
BAD_EXPORT void xlAutoFree12(LPXLOPER12 x) {
    if (x == swap_result) {
        delete[] swap_cache;
        swap_result = nullptr;
    }
    delete x;
}

// Returns heap memory without xlbitDLLFree, so it is never freed.
BAD_EXPORT LPXLOPER12 BadNoDllFree() {
    auto* x = new XLOPER12{.xltype = xltypeNum};
    x->val.num = 1;
    return x;
}

// Frees a value Excel returned twice, through a copy taken before the first free. (Freeing
// the same XLOPER12 twice is harmless: xlFree sets its pointer to NULL.)
BAD_EXPORT LPXLOPER12 BadDoubleFree() {
    static XLOPER12 result{.xltype = xltypeNum};
    XLOPER12 name;
    excel(xlGetName, &name);
    XLOPER12 copy = name;
    excel(xlFree, nullptr, {&name});
    excel(xlFree, nullptr, {&copy});
    return &result;
}

// Never frees a value Excel returned.
BAD_EXPORT LPXLOPER12 BadLeak() {
    static XLOPER12 result{.xltype = xltypeNum};
    XLOPER12 name;
    excel(xlGetName, &name);
    return &result;
}

// Returns text marked xlbitDLLFree, which xlAutoFree12 above then fails to free.
BAD_EXPORT LPXLOPER12 BadNestedLeak() {
    auto* x = new XLOPER12{.xltype = xltypeStr | xlbitDLLFree};
    x->val.str = new XCHAR[3]{2, L'h', L'i'};
    return x;
}

// Returns text marked xlbitDLLFree; see xlAutoFree12 above.
BAD_EXPORT LPXLOPER12 BadSwapFree() {
    swap_cache = new char[16];
    swap_result = new XLOPER12{.xltype = xltypeStr | xlbitDLLFree};
    swap_result->val.str = new XCHAR[3]{2, L'h', L'i'};
    return swap_result;
}

// Registered as taking one argument, but reads a second.
BAD_EXPORT LPXLOPER12 BadArity(LPXLOPER12, LPXLOPER12 second) {
    static XLOPER12 result{.xltype = xltypeNum};
    result.val.num = second->xltype;
    return &result;
}

// Overwrites its read-only argument.
BAD_EXPORT LPXLOPER12 BadChangesArgument(LPXLOPER12 value) {
    static XLOPER12 result{.xltype = xltypeNum};
    value->xltype = xltypeNil;
    return &result;
}
