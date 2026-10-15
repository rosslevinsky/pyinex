// The add-in's entry points: what Excel calls when it opens and closes the add-in, frees a
// returned value, or calculates one of Pyinex's worksheet functions (docs/architecture.md,
// "Excel functions").
#include "console/console.h"
#include "excel/excel.h"
#include "loader/console_window.h"
#include "loader/python.h"

#include <cmath>
#include <cwctype>
#include <optional>
#include <string>
#include <vector>

using namespace pyinex::excel;
using pyinex::loader::python;

#define PYINEX_EXPORT extern "C" __declspec(dllexport)

namespace {

constexpr int max_arguments = 30;  // PyCall's, after the file and the function.

bool reload_on_change = true;

std::vector<std::wstring_view> pycall_help() {
    std::vector<std::wstring_view> help = {L"Path of the Python script; a relative path is relative to the workbook's folder.",
                                           L"Name of the function in the script."};
    help.insert(help.end(), max_arguments, L"An argument passed to the function.");
    return help;
}

// None is registered thread safe ($): Excel calls them only on its main thread.
const Function functions[] = {
    {.name = L"PyCall",
     .procedure = L"PyCall",
     .type = L"QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ",  // The result, file, function, 30 arguments.
     .arguments = L"file,function,arg1,arg2,arg3,arg4,arg5,arg6,arg7,arg8,arg9,arg10,arg11,arg12,arg13,arg14,"
                  L"arg15,arg16,arg17,arg18,arg19,arg20,arg21,arg22,arg23,arg24,arg25,arg26,arg27,arg28,arg29,arg30",
     .help = L"Calls a function in a Python script and returns its result.",
     .argument_help = pycall_help()},
    {.name = L"PyConsole",
     .procedure = L"PyConsole",
     .type = L"QQQQQQQ",
     .arguments = L"show,clear,x,y,width,height",
     .help = L"Shows, hides, clears and places Pyinex's console window. Returns show.",
     .argument_help = {L"TRUE to show the console, FALSE to hide it.", L"TRUE to clear its text.",
                       L"Left edge of the window, in pixels.", L"Top edge of the window, in pixels.",
                       L"Width of the window, in pixels.", L"Height of the window, in pixels."}},
    {.name = L"PyReloadOnChange",
     .procedure = L"PyReloadOnChange",
     .type = L"QQ",
     .arguments = L"on",
     .help = L"Turns reloading of changed scripts on or off. Returns the setting.",
     .argument_help = {L"TRUE to reload a script when its file changes, FALSE to load each once; "
                       L"leave out to keep the setting."}},
    {.name = L"PyInfo",
     .procedure = L"PyInfo",
     .type = L"QQ",
     .arguments = L"item",
     .help = L"Reports which Python and which add-in are loaded.",
     .argument_help = {L"status, python, pyinex, version, executable, prefix, venv or addin_version; "
                       L"leave out for all of them."}},
};

int type_of(const XLOPER12& x) { return x.xltype & ~(xlbitDLLFree | xlbitXLFree); }

bool given(const XLOPER12& x) { return type_of(x) != xltypeMissing && type_of(x) != xltypeNil; }

LPXLOPER12 blank() { return Value::text(L"").to_excel(); }

LPXLOPER12 refuse(const std::wstring& why) {
    pyinex::loader::prepare_output();  // In Excel, there is nowhere to say why until then.
    pyinex::console::note(why);
    return Value::error(xlerrValue).to_excel();
}

// One of PyConsole's pixel arguments, rounded to a whole pixel, halves away from zero.
// Returns false, with why set, if it is given and is not a number that fits.
bool pixels(const XLOPER12& x, const wchar_t* name, std::optional<int>& out, std::wstring& why) {
    if (!given(x)) return true;
    const double n = type_of(x) == xltypeNum ? std::round(x.val.num) : NAN;
    if (std::isnan(n)) why = std::wstring(L"PyConsole: ") + name + L" must be a number";
    else if (std::abs(n) > 1e9) why = std::wstring(L"PyConsole: ") + name + L" is far too large";
    else out = static_cast<int>(n);
    return why.empty();
}

// Checks where PyConsole is asked to put the window; returns why it cannot, or empty.
std::wstring check_placement(const pyinex::loader::Placement& p) {
    if (p.x.has_value() != p.y.has_value()) return L"PyConsole: give x and y together, or neither";
    if (p.width.has_value() != p.height.has_value()) return L"PyConsole: give width and height together, or neither";
    if (p.width && (*p.width < 1 || *p.height < 1))
        return L"PyConsole: width and height must be more than 0 pixels, after rounding to whole pixels";
    if (p.x) {
        // The desktop spans every monitor, so it can start left of or above the main one.
        const int left = GetSystemMetrics(SM_XVIRTUALSCREEN), top = GetSystemMetrics(SM_YVIRTUALSCREEN);
        const int right = left + GetSystemMetrics(SM_CXVIRTUALSCREEN) - 1;
        const int bottom = top + GetSystemMetrics(SM_CYVIRTUALSCREEN) - 1;
        if (*p.x < left || *p.x > right || *p.y < top || *p.y > bottom)
            return L"PyConsole: x and y must lie on the desktop, from " + std::to_wstring(left) + L"," +
                   std::to_wstring(top) + L" to " + std::to_wstring(right) + L"," + std::to_wstring(bottom);
    }
    return L"";
}

// PyInfo's items, in the order its table lists them.
constexpr std::wstring_view info_items[] = {L"status", L"python",  L"pyinex", L"version",
                                            L"executable", L"prefix", L"venv",   L"addin_version"};

Value info_item(std::wstring_view item, const pyinex::loader::Core* core) {
    if (item == L"status") return Value::text(core ? L"running" : pyinex::loader::refusal());
    if (item == L"pyinex") return Value::text(pyinex::loader::addin_path());
    if (!core) return Value::error(xlerrValue);  // The rest need Python.
    if (item == L"python") return Value::text(core->dll);
    if (item == L"executable") return Value::text(core->executable);
    if (item == L"addin_version") return Value::text(PYINEX_VERSION);
    LPXLOPER12 known = core->info(std::wstring(item).c_str());  // version, prefix, venv.
    Value value = Value::copy(*known);
    if (known->xltype & xlbitDLLFree) free_returned(known);  // Not guarded's static error.
    return value;
}

}  // namespace

PYINEX_EXPORT int xlAutoOpen() {
    try {
        pyinex::loader::forget_refusal();
        register_functions(functions);
        return 1;
    } catch (...) {
        return 0;
    }
}

// Python keeps running, and the add-in stays loaded once it has started (docs/architecture.md,
// "Lifetime").
PYINEX_EXPORT int xlAutoClose() {
    if (const auto* core = pyinex::loader::running()) core->close();
    return 1;
}

PYINEX_EXPORT void xlAutoFree12(LPXLOPER12 x) { free_returned(x); }

PYINEX_EXPORT LPXLOPER12 PyCall(LPXLOPER12 file, LPXLOPER12 function, LPXLOPER12 a1, LPXLOPER12 a2, LPXLOPER12 a3,
                                LPXLOPER12 a4, LPXLOPER12 a5, LPXLOPER12 a6, LPXLOPER12 a7, LPXLOPER12 a8,
                                LPXLOPER12 a9, LPXLOPER12 a10, LPXLOPER12 a11, LPXLOPER12 a12, LPXLOPER12 a13,
                                LPXLOPER12 a14, LPXLOPER12 a15, LPXLOPER12 a16, LPXLOPER12 a17, LPXLOPER12 a18,
                                LPXLOPER12 a19, LPXLOPER12 a20, LPXLOPER12 a21, LPXLOPER12 a22, LPXLOPER12 a23,
                                LPXLOPER12 a24, LPXLOPER12 a25, LPXLOPER12 a26, LPXLOPER12 a27, LPXLOPER12 a28,
                                LPXLOPER12 a29, LPXLOPER12 a30) {
    return guarded([&] {
        if (in_function_wizard()) return blank();
        const auto* core = python();
        if (!core) return Value::error(xlerrValue).to_excel();
        const XLOPER12* const args[max_arguments] = {a1,  a2,  a3,  a4,  a5,  a6,  a7,  a8,  a9,  a10,
                                                     a11, a12, a13, a14, a15, a16, a17, a18, a19, a20,
                                                     a21, a22, a23, a24, a25, a26, a27, a28, a29, a30};
        return core->call(file, function, args, max_arguments, reload_on_change);
    });
}

PYINEX_EXPORT LPXLOPER12 PyConsole(LPXLOPER12 show, LPXLOPER12 clear, LPXLOPER12 x, LPXLOPER12 y, LPXLOPER12 width,
                                   LPXLOPER12 height) {
    return guarded([&] {
        if (in_function_wizard()) return blank();
        if (type_of(*show) != xltypeBool) return refuse(L"PyConsole: show must be TRUE or FALSE");
        if (given(*clear) && type_of(*clear) != xltypeBool) return refuse(L"PyConsole: clear must be TRUE or FALSE");
        pyinex::loader::Placement placement;
        std::wstring why;
        if (!pixels(*x, L"x", placement.x, why) || !pixels(*y, L"y", placement.y, why) ||
            !pixels(*width, L"width", placement.width, why) || !pixels(*height, L"height", placement.height, why) ||
            !(why = check_placement(placement)).empty())
            return refuse(why);
        const bool shown = show->val.xbool != 0;
        pyinex::loader::show_console(shown, given(*clear) && clear->val.xbool != 0, placement);
        return Value::boolean(shown).to_excel();
    });
}

PYINEX_EXPORT LPXLOPER12 PyReloadOnChange(LPXLOPER12 on) {
    return guarded([&] {
        if (in_function_wizard()) return blank();
        if (type_of(*on) == xltypeBool) reload_on_change = on->val.xbool != 0;
        else if (type_of(*on) == xltypeNum) reload_on_change = on->val.num != 0;
        else if (given(*on)) return Value::error(xlerrValue).to_excel();
        return Value::boolean(reload_on_change).to_excel();
    });
}

PYINEX_EXPORT LPXLOPER12 PyInfo(LPXLOPER12 item) {
    return guarded([&] {
        if (in_function_wizard()) return blank();
        const auto* core = python();  // Every item starts Python.
        if (!given(*item)) {
            std::vector<Value> table;
            for (const auto name : info_items) {
                table.push_back(Value::text(name));
                table.push_back(info_item(name, core));
            }
            return Value::array(static_cast<int>(std::size(info_items)), 2, std::move(table)).to_excel();
        }
        if (type_of(*item) != xltypeStr) return Value::error(xlerrValue).to_excel();
        std::wstring name(item->val.str + 1, item->val.str[0]);
        for (auto& c : name) c = static_cast<wchar_t>(std::towlower(c));
        for (const auto known : info_items)
            if (name == known) return info_item(known, core).to_excel();
        return Value::error(xlerrValue).to_excel();
    });
}
