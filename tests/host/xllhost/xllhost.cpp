// xllhost: a fake Excel for testing add-ins without Excel.
//
//   xllhost.exe <request pipe handle> <response pipe handle> [--no-output]
//
// --no-output, given to xllhost_gui.exe (the same program built with a window, as Excel is)
// started with no standard handles, checks it starts as Excel does, with no standard output
// or error, so the add-in's output has nowhere to go until it makes its console window.
//
// It loads an add-in the way Excel does, answers the add-in's calls into Excel
// (MdCallBack12, which the add-in finds among this program's exports), records what the
// add-in registers, calls its functions the way Excel calculates a cell, and enforces
// Excel's memory rules, reporting each break as a "fault". tests/host/hostdriver.py drives
// it. Requests and responses travel on the two inherited pipe handles, one line each, so
// that standard output and error carry only what the add-in itself prints.
//
// Requests:
//   open <path>                  load the add-in and call xlAutoOpen
//   load <path>                  load some other DLL, as another add-in might
//   close                        call xlAutoClose
//   unload                       FreeLibrary the add-in; answers whether the add-in and
//                                pyinex_core.dll are still loaded
//   call <name> <value>...       call a registered function
//   registrations                list what xlfRegister recorded
//   caller none | <sheet> <row> <column> <rows> <columns> [sref]
//                                what xlfCaller returns (default: [Book1]Sheet1!R1C1);
//                                sref: as an xltypeSRef, with no sheet, not an xltypeRef
//   abort 0|1                    whether xlAbort reports Esc pressed
//   folder none | <path> [alone] the calling workbook's folder, for CELL("filename")
//                                (default: none, an unsaved workbook); alone: CELL gives
//                                the workbook's path without its sheet, as Excel does for
//                                a workbook whose only sheet has the workbook's name
//   wizard 0|1                   open or close a stand-in for the function wizard
//   faults                       faults since the last "faults", then forget them
//   calls                        how many calls the add-in made into Excel since the
//                                last "calls"
//   handles                      the process's open handle count
//   console                      Pyinex's console window: whether there is one, its
//                                title, left, top, width and height, its text, and the
//                                text and background colors it gives its text box
//   close-console                close that window, as its close button does
//   theme light|dark             Windows' app mode as the add-in reads it, from this
//                                process's own registry (below); and tell the console
//                                window it changed, as Windows does
//   register user|machine <version> <executable>|- <folder>|-
//                                record a Python installation as python.org's installer
//                                does (PEP 514), in this process's own HKEY_CURRENT_USER or
//                                HKEY_LOCAL_MACHINE: InstallPath's ExecutablePath value and
//                                its default value
//   write <text>                 write text to standard error, as C code's output would
// Before each request the host handles its windows' messages, as Excel does between
// calculations. From its start, the host's HKEY_CURRENT_USER and HKEY_LOCAL_MACHINE are
// keys of its own, empty but for a copy of Windows' cryptography settings, so the add-in
// reads only what a test put there.
//   quit
// Responses: "ok" and values, or "error <text>".
//
// Values, as space-separated tokens:
//   n:<number>  s:<text>  b:0|1  e:<error code>  i:<integer>  m (missing)  z (empty)
//   a:<rows>x<columns> followed by that many values, in row order
// Text (names, paths, s: values) is UTF-16 as four hex digits per code unit, so it
// carries any character, spaces and unpaired surrogates included.
#include <windows.h>

#include <XLCALL.H>

#include <fcntl.h>
#include <io.h>
#include <malloc.h>

#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <cwchar>
#include <map>
#include <sstream>
#include <stdexcept>
#include <string>
#include <vector>

namespace {

// ---------------------------------------------------------------------------------------
// Text encoding for the protocol.

std::string hex(const XCHAR* s, size_t n) {
    static const char digits[] = "0123456789abcdef";
    std::string out;
    for (size_t i = 0; i < n; ++i)
        for (int shift = 12; shift >= 0; shift -= 4) out += digits[(s[i] >> shift) & 0xf];
    return out;
}

std::string hex(const std::wstring& s) { return hex(s.data(), s.size()); }

std::wstring unhex(const std::string& h) {
    std::wstring out;
    for (size_t i = 0; i + 4 <= h.size(); i += 4)
        out += static_cast<wchar_t>(std::stoul(h.substr(i, 4), nullptr, 16));
    return out;
}

// ---------------------------------------------------------------------------------------
// Faults: breaks of Excel's rules, collected for the "faults" request.

std::vector<std::wstring> faults;

void fault(const std::wstring& message) { faults.push_back(message); }

// ---------------------------------------------------------------------------------------
// Values the host makes for itself (function arguments): plain trees, freed by
// free_tree.

void free_tree(XLOPER12& x) {
    switch (x.xltype & ~(xlbitDLLFree | xlbitXLFree)) {
    case xltypeStr: delete[] x.val.str; break;
    case xltypeRef: delete x.val.mref.lpmref; break;
    case xltypeMulti:
        for (int i = 0; i < x.val.array.rows * x.val.array.columns; ++i)
            free_tree(x.val.array.lparray[i]);
        delete[] x.val.array.lparray;
        break;
    default: break;
    }
    x.xltype = xltypeNil;
}

XLOPER12 make_text(const std::wstring& s) {
    auto* b = new XCHAR[s.size() + 1];
    b[0] = static_cast<XCHAR>(s.size());
    for (size_t i = 0; i < s.size(); ++i) b[i + 1] = s[i];
    XLOPER12 x{.xltype = xltypeStr};
    x.val.str = b;
    return x;
}

XLOPER12 make_num(double n) {
    XLOPER12 x{.xltype = xltypeNum};
    x.val.num = n;
    return x;
}

// Parses one value from the request tokens; arrays read their items recursively.
XLOPER12 parse_value(std::istringstream& in) {
    std::string token;
    if (!(in >> token)) throw std::runtime_error("missing value");
    const std::string body = token.size() > 2 ? token.substr(2) : "";
    XLOPER12 x{.xltype = xltypeNil};
    switch (token[0]) {
    case 'n': return make_num(std::strtod(body.c_str(), nullptr));
    case 's': return make_text(unhex(body));
    case 'b': x.xltype = xltypeBool; x.val.xbool = body == "1"; return x;
    case 'e': x.xltype = xltypeErr; x.val.err = std::stoi(body); return x;
    case 'i': x.xltype = xltypeInt; x.val.w = std::stoi(body); return x;
    case 'm': x.xltype = xltypeMissing; return x;
    case 'z': return x;
    case 'a': {
        const auto by = body.find('x');
        const int rows = std::stoi(body.substr(0, by)), columns = std::stoi(body.substr(by + 1));
        auto* items = new XLOPER12[static_cast<size_t>(rows) * columns];
        for (int i = 0; i < rows * columns; ++i) items[i] = parse_value(in);
        x.xltype = xltypeMulti;
        x.val.array.lparray = items;
        x.val.array.rows = rows;
        x.val.array.columns = columns;
        return x;
    }
    default: throw std::runtime_error("bad value token: " + token);
    }
}

std::string format_value(const XLOPER12& x) {
    char number[64];
    switch (x.xltype & ~(xlbitDLLFree | xlbitXLFree)) {
    case xltypeNum:
        std::snprintf(number, sizeof number, "n:%.17g", x.val.num);
        return number;
    case xltypeStr: return "s:" + hex(x.val.str + 1, x.val.str[0]);
    case xltypeBool: return x.val.xbool ? "b:1" : "b:0";
    case xltypeErr: return "e:" + std::to_string(x.val.err);
    case xltypeInt: return "i:" + std::to_string(x.val.w);
    case xltypeMissing: return "m";
    case xltypeNil: return "z";
    case xltypeMulti: {
        const auto& a = x.val.array;
        std::string out = "a:" + std::to_string(a.rows) + "x" + std::to_string(a.columns);
        for (int i = 0; i < a.rows * a.columns; ++i) out += " " + format_value(a.lparray[i]);
        return out;
    }
    default: return "e:" + std::to_string(xlerrValue);
    }
}

// ---------------------------------------------------------------------------------------
// Values Excel makes for the add-in (results of Excel12v calls). The add-in must free
// each with xlFree exactly once. They are tracked by the memory block at their root, since
// the add-in may free a copy of the XLOPER12 struct rather than the one it was given.

std::map<const void*, XLOPER12> excel_allocations;

const void* block_of(const XLOPER12& x) {
    switch (x.xltype & ~(xlbitDLLFree | xlbitXLFree)) {
    case xltypeStr: return x.val.str;
    case xltypeRef: return x.val.mref.lpmref;
    case xltypeMulti: return x.val.array.lparray;
    default: return nullptr;
    }
}

// Gives a value to the add-in, remembering it if it owns memory.
void hand_over(XLOPER12 x, LPXLOPER12 result) {
    if (!result) {  // The add-in did not want the result.
        free_tree(x);
        return;
    }
    if (const void* block = block_of(x)) excel_allocations[block] = x;
    *result = x;
}

// Like Excel's xlFree: frees the memory x points to and sets that pointer to NULL, leaving
// the rest of x alone, so freeing the same XLOPER12 twice is harmless. Freeing a copy made
// before the first free is the real double free, and a fault.
void excel_free(XLOPER12& x) {
    const void* block = block_of(x);
    if (!block) return;  // Nothing to free; xlFree accepts any value.
    const auto it = excel_allocations.find(block);
    if (it == excel_allocations.end()) {
        fault(L"xlFree of memory Excel did not allocate, or already freed");
        return;
    }
    free_tree(it->second);
    excel_allocations.erase(it);
    switch (x.xltype & ~(xlbitDLLFree | xlbitXLFree)) {
    case xltypeStr: x.val.str = nullptr; break;
    case xltypeRef: x.val.mref.lpmref = nullptr; break;
    case xltypeMulti: x.val.array.lparray = nullptr; break;
    default: break;
    }
}

// After each request: anything Excel gave the add-in that it never freed is a leak.
void check_leaks() {
    for (auto& [block, x] : excel_allocations) {
        fault(L"a value Excel returned was never freed with xlFree");
        free_tree(x);
    }
    excel_allocations.clear();
}

// ---------------------------------------------------------------------------------------
// The add-in and what it registered.

HMODULE xll = nullptr;
std::wstring xll_path;

struct Registration {
    std::wstring procedure, type, name, arguments, category, help;
    std::vector<std::wstring> argument_help;
    FARPROC address = nullptr;
    int argument_count = 0;
};
std::vector<Registration> registrations;

std::wstring text_of(const XLOPER12* x) {
    if (!x || (x->xltype & ~(xlbitDLLFree | xlbitXLFree)) != xltypeStr) return L"";
    return {x->val.str + 1, static_cast<size_t>(x->val.str[0])};
}

// Counts the arguments in an Excel type string: one result code, then one per argument;
// a trailing "!" (volatile), "#" (macro sheet equivalent) or "$" (thread safe) is a flag.
// xllhost calls only Q and U functions (XLOPER12 values and references), which is all
// Pyinex registers.
int count_arguments(const std::wstring& type) {
    std::wstring codes = type;
    while (!codes.empty() && (codes.back() == L'!' || codes.back() == L'#' || codes.back() == L'$'))
        codes.pop_back();
    if (codes.empty()) return -1;
    for (wchar_t c : codes)
        if (c != L'Q' && c != L'U') return -1;
    return static_cast<int>(codes.size()) - 1;
}

int do_register(int count, LPXLOPER12* args, LPXLOPER12 result) {
    if (count < 4) {
        fault(L"xlfRegister needs at least module, procedure, type and name");
        return xlretInvCount;
    }
    Registration r;
    if (_wcsicmp(text_of(args[0]).c_str(), xll_path.c_str()) != 0)
        fault(L"xlfRegister module is not the add-in's path: " + text_of(args[0]));
    r.procedure = text_of(args[1]);
    r.type = text_of(args[2]);
    r.name = text_of(args[3]);
    if (count > 4) r.arguments = text_of(args[4]);
    if (count > 6) r.category = text_of(args[6]);
    if (count > 9) r.help = text_of(args[9]);
    for (int i = 10; i < count; ++i) r.argument_help.push_back(text_of(args[i]));
    char procedure[256] = {};
    WideCharToMultiByte(CP_ACP, 0, r.procedure.c_str(), -1, procedure, sizeof procedure, nullptr, nullptr);
    r.address = GetProcAddress(xll, procedure);
    r.argument_count = count_arguments(r.type);
    XLOPER12 id{.xltype = xltypeErr};
    id.val.err = xlerrValue;
    if (!r.address) fault(L"xlfRegister procedure is not exported: " + r.procedure);
    else if (r.argument_count < 0) fault(L"xlfRegister type string not supported by xllhost: " + r.type);
    else if (r.argument_count > 32) fault(L"xlfRegister: more than 32 arguments: " + r.name);
    else {
        registrations.push_back(r);
        id = make_num(static_cast<double>(registrations.size()));
    }
    hand_over(id, result);
    return xlretSuccess;
}

// ---------------------------------------------------------------------------------------
// Excel's other services the add-in may call.

std::wstring caller_sheet = L"[Book1]Sheet1";
bool caller_set = true;
bool caller_sref = false;
XLREF12 caller_cell{.rwFirst = 0, .rwLast = 0, .colFirst = 0, .colLast = 0};
std::vector<std::wstring> sheets = {L"[Book1]Sheet1"};  // idSheet is index + 1.
bool escape_pressed = false;
std::wstring workbook_folder;  // Empty: the workbook has not been saved.
bool path_alone = false;       // CELL("filename") gives C:\folder\Book.xlsx, no sheet.

// Excel's main window, which xlGetHwnd returns, and a stand-in for the function wizard: a
// window of the wizard's class and title, owned by a workbook's window, as since Excel 2013
// each workbook has a window of its own. None is shown.
HWND main_window = nullptr;
HWND wizard_window = nullptr;
HWND workbook_window = nullptr;

HWND make_window(const wchar_t* window_class, const wchar_t* title, HWND owner, DWORD style) {
    WNDCLASSW wc{.lpfnWndProc = DefWindowProcW, .hInstance = GetModuleHandleW(nullptr), .lpszClassName = window_class};
    RegisterClassW(&wc);  // Fails harmlessly once registered.
    return CreateWindowExW(WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE, window_class, title, style, -32000, -32000, 100, 100,
                           owner, nullptr, GetModuleHandleW(nullptr), nullptr);
}

HWND excel_window() {
    if (!main_window) main_window = make_window(L"XLMAIN", L"Excel", nullptr, WS_OVERLAPPEDWINDOW);
    return main_window;
}

void set_wizard(bool open) {
    if (open && !wizard_window) {
        excel_window();
        if (!workbook_window) workbook_window = make_window(L"XLMAIN", L"Book1 - Excel", nullptr, WS_OVERLAPPEDWINDOW);
        wizard_window = make_window(L"bosa_sdm_XL9", L"Function Arguments", workbook_window, WS_POPUP);
    }
    if (!open && wizard_window) {
        DestroyWindow(wizard_window);
        wizard_window = nullptr;
    }
}

IDSHEET sheet_id(const std::wstring& name) {
    for (size_t i = 0; i < sheets.size(); ++i)
        if (sheets[i] == name) return static_cast<IDSHEET>(i + 1);
    sheets.push_back(name);
    return static_cast<IDSHEET>(sheets.size());
}

XLOPER12 make_ref(IDSHEET sheet, const XLREF12& area) {
    auto* m = new XLMREF12{.count = 1};
    m->reftbl[0] = area;
    XLOPER12 x{.xltype = xltypeRef};
    x.val.mref.lpmref = m;
    x.val.mref.idSheet = sheet;
    return x;
}

XLOPER12 coerce(const XLOPER12& x, int to) {
    const int type = x.xltype & ~(xlbitDLLFree | xlbitXLFree);
    if (type == xltypeRef || type == xltypeSRef) return {.xltype = xltypeNil};  // An empty sheet.
    if (to == 0 || (to & type)) {
        // A copy, through the protocol's own encoding.
        std::istringstream in(format_value(x));
        return parse_value(in);
    }
    if ((to & xltypeStr) && type == xltypeNum) {
        wchar_t buffer[64];
        swprintf(buffer, 64, L"%.15g", x.val.num);
        return make_text(buffer);
    }
    if ((to & xltypeNum) && type == xltypeStr) {
        const std::wstring s = text_of(&x);
        wchar_t* end = nullptr;
        const double n = wcstod(s.c_str(), &end);
        if (!s.empty() && end && *end == 0) return make_num(n);
    }
    XLOPER12 error{.xltype = xltypeErr};
    error.val.err = xlerrValue;
    return error;
}

int excel_calls = 0;  // Since the last "calls" request.

// Every add-in call into Excel arrives here, through the add-in's call_excel.
int excel_callback(int function, int count, LPXLOPER12* args, LPXLOPER12 result) {
    ++excel_calls;
    if (count < 0 || count > 255) return xlretInvCount;
    switch (function) {
    case xlFree:
        for (int i = 0; i < count; ++i) excel_free(*args[i]);
        return xlretSuccess;
    case xlGetName:
        hand_over(make_text(xll_path), result);
        return xlretSuccess;
    case xlfRegister:
        return do_register(count, args, result);
    case xlfCaller: {
        XLOPER12 error{.xltype = xltypeErr};
        error.val.err = xlerrRef;
        XLOPER12 sref{.xltype = xltypeSRef};
        sref.val.sref = {.count = 1, .ref = caller_cell};
        hand_over(!caller_set ? error : caller_sref ? sref : make_ref(sheet_id(caller_sheet), caller_cell), result);
        return xlretSuccess;
    }
    case xlSheetNm: {
        const int type = count ? args[0]->xltype & ~(xlbitDLLFree | xlbitXLFree) : 0;
        const IDSHEET id = type == xltypeRef ? args[0]->val.mref.idSheet : sheet_id(caller_sheet);
        if (id < 1 || id > sheets.size()) return xlretFailed;
        hand_over(make_text(sheets[id - 1]), result);
        return xlretSuccess;
    }
    case xlCoerce: {
        if (count < 1) return xlretInvCount;
        const int to = count > 1 && args[1]->xltype == xltypeInt ? args[1]->val.w
                     : count > 1 && args[1]->xltype == xltypeNum ? static_cast<int>(args[1]->val.num)
                     : 0;
        hand_over(coerce(*args[0], to), result);
        return xlretSuccess;
    }
    case xlGetHwnd: {
        XLOPER12 handle{.xltype = xltypeInt};
        handle.val.w = static_cast<int>(reinterpret_cast<INT_PTR>(excel_window()));
        hand_over(handle, result);
        return xlretSuccess;
    }
    case xlfCell: {
        // Only CELL("filename", reference): the workbook's folder and the sheet's name, or
        // empty text for an unsaved workbook. An xltypeSRef is on the sheet calculating.
        const int type = count == 2 ? args[1]->xltype & ~(xlbitDLLFree | xlbitXLFree) : 0;
        if (count != 2 || text_of(args[0]) != L"filename" || (type != xltypeRef && type != xltypeSRef)) {
            fault(L"xllhost implements only CELL(\"filename\", reference)");
            return xlretInvXlfn;
        }
        const IDSHEET id = type == xltypeRef ? args[1]->val.mref.idSheet : sheet_id(caller_sheet);
        if (id < 1 || id > sheets.size()) return xlretFailed;
        // [Book1]Sheet1, or with path_alone just Book1.
        const std::wstring& sheet = sheets[id - 1];
        const std::wstring name = path_alone ? sheet.substr(1, sheet.find(L']') - 1) : sheet;
        hand_over(make_text(workbook_folder.empty() ? L"" : workbook_folder + L"\\" + name), result);
        return xlretSuccess;
    }
    case xlAbort: {
        XLOPER12 pressed{.xltype = xltypeBool};
        pressed.val.xbool = escape_pressed;
        // xlAbort(FALSE) clears the break condition.
        if (count > 0 && args[0]->xltype == xltypeBool && !args[0]->val.xbool) escape_pressed = false;
        hand_over(pressed, result);
        return xlretSuccess;
    }
    default:
        fault(L"xllhost does not implement Excel function " + std::to_wstring(function));
        return xlretInvXlfn;
    }
}

// ---------------------------------------------------------------------------------------
// Calling a registered function, and checking what it returns.

bool in_xll_image(const void* p) {
    HMODULE owner = nullptr;
    return p && GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                                       GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                   static_cast<LPCWSTR>(p), &owner) &&
           owner == xll;
}

// Checks a returned value's shape: arrays hold plain values, never other arrays.
void check_shape(const XLOPER12& x, bool nested) {
    const int type = x.xltype & ~(xlbitDLLFree | xlbitXLFree);
    if (nested && (x.xltype & (xlbitDLLFree | xlbitXLFree)))
        fault(L"an array item carries xlbitDLLFree or xlbitXLFree; only the result itself may");
    if (type == xltypeMulti) {
        if (nested) fault(L"an array inside an array");
        else if (x.val.array.rows < 1 || x.val.array.columns < 1) fault(L"an array with no cells");
        else
            for (int i = 0; i < x.val.array.rows * x.val.array.columns; ++i)
                check_shape(x.val.array.lparray[i], true);
    }
}

// Whether a result without xlbitDLLFree lives in the add-in's static memory, the only
// memory Excel may leave unfreed.
bool all_static(const XLOPER12& x) {
    if (!in_xll_image(&x)) return false;
    const void* block = block_of(x);
    if (block && !in_xll_image(block)) return false;
    if ((x.xltype & ~(xlbitDLLFree | xlbitXLFree)) == xltypeMulti)
        for (int i = 0; i < x.val.array.rows * x.val.array.columns; ++i)
            if (const void* b = block_of(x.val.array.lparray[i]); b && !in_xll_image(b)) return false;
    return true;
}

// The memory blocks behind a value — each text and array, down the tree — with what each
// is, for fault messages.
void blocks_in(const XLOPER12& x, std::vector<std::pair<const void*, std::wstring>>& out) {
    switch (x.xltype & ~(xlbitDLLFree | xlbitXLFree)) {
    case xltypeStr: out.emplace_back(x.val.str, L"text"); break;
    case xltypeMulti:
        out.emplace_back(x.val.array.lparray, L"array");
        for (int i = 0; i < x.val.array.rows * x.val.array.columns; ++i) blocks_in(x.val.array.lparray[i], out);
        break;
    default: break;
    }
}

// Calls f for each heap block in use.
template <class F>
void each_busy_block(HANDLE heap, F f) {
    PROCESS_HEAP_ENTRY entry{};
    while (HeapWalk(heap, &entry))
        if (entry.wFlags & PROCESS_HEAP_ENTRY_BUSY) f(entry);
}

// Calls xlAutoFree12 on a result and checks it released exactly that result: the XLOPER12
// itself and every text and array below it are no longer in use afterwards, and nothing
// else was freed. It walks the C runtime heap, with the heap locked so no other thread
// changes it, and so assumes the add-in allocates each of those separately with new or
// malloc, as Pyinex does.
void free_and_check(void (*free_fn)(LPXLOPER12), LPXLOPER12 out) {
    std::vector<std::pair<const void*, std::wstring>> blocks{{out, L"XLOPER12"}};
    blocks_in(*out, blocks);
    const auto heap = reinterpret_cast<HANDLE>(_get_heap_handle());
    size_t before = 0, after = 0;
    // Nothing may allocate between the two walks, or the count is off: reserve first, and
    // record what is unfreed by pointer.
    std::vector<const std::wstring*> unfreed;
    unfreed.reserve(blocks.size());
    HeapLock(heap);
    each_busy_block(heap, [&](const PROCESS_HEAP_ENTRY&) { ++before; });
    free_fn(out);
    each_busy_block(heap, [&](const PROCESS_HEAP_ENTRY& e) {
        ++after;
        const auto* start = static_cast<const char*>(e.lpData);
        for (const auto& [block, what] : blocks)
            if (static_cast<const char*>(block) >= start && static_cast<const char*>(block) < start + e.cbData)
                unfreed.push_back(&what);
    });
    HeapUnlock(heap);
    for (const auto* what : unfreed) fault(L"xlAutoFree12 did not free the result's " + *what);
    if (before - after != blocks.size())
        fault(L"xlAutoFree12 released " + std::to_wstring(before - after) + L" heap blocks; the result had " +
              std::to_wstring(blocks.size()));
}

// Where Excel passes no argument at all — past the registered count — a function reading
// one gets whatever its register holds. The host passes a pointer to memory that cannot
// be read, so a function registered with too few arguments crashes the test instead of
// working by luck.
LPXLOPER12 unreadable() {
    static void* page = VirtualAlloc(nullptr, 4096, MEM_RESERVE | MEM_COMMIT, PAGE_NOACCESS);
    return static_cast<LPXLOPER12>(page);
}

std::string do_call(std::istringstream& in) {
    std::string name;
    in >> name;
    const std::wstring wanted = unhex(name);
    const Registration* r = nullptr;
    for (const auto& reg : registrations)
        if (_wcsicmp(reg.name.c_str(), wanted.c_str()) == 0) r = &reg;
    if (!r) return "error " + hex(L"no registered function " + wanted);

    std::vector<XLOPER12> values;
    try {
        while (in >> std::ws, !in.eof()) values.push_back(parse_value(in));
    } catch (...) {
        for (auto& v : values) free_tree(v);
        throw;
    }
    if (static_cast<int>(values.size()) > r->argument_count) {
        for (auto& v : values) free_tree(v);
        return "error " + hex(L"too many arguments for " + wanted);
    }
    // Excel passes every registered argument the formula left out as xltypeMissing.
    while (static_cast<int>(values.size()) < r->argument_count) values.push_back({.xltype = xltypeMissing});
    LPXLOPER12 a[32];
    for (int i = 0; i < 32; ++i) a[i] = i < r->argument_count ? &values[i] : unreadable();
    // Arguments belong to Excel and are read-only: keep the originals, to free those and to
    // catch a function that changes or frees what it was passed.
    const std::vector<XLOPER12> originals = values;
    std::vector<std::string> before;
    for (const auto& v : values) before.push_back(format_value(v));

    // x64 has one calling convention, in which the caller cleans up, so passing a function
    // more arguments than it declares is safe.
    using Fn = LPXLOPER12 (*)(LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                              LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                              LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                              LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                              LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                              LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12,
                              LPXLOPER12, LPXLOPER12);
    const auto fn = reinterpret_cast<Fn>(r->address);
    LPXLOPER12 out = fn(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], a[9], a[10],
                        a[11], a[12], a[13], a[14], a[15], a[16], a[17], a[18], a[19], a[20],
                        a[21], a[22], a[23], a[24], a[25], a[26], a[27], a[28], a[29], a[30],
                        a[31]);
    for (size_t i = 0; i < values.size(); ++i)
        if (std::memcmp(&values[i], &originals[i], sizeof(XLOPER12)) != 0 || format_value(originals[i]) != before[i])
            fault(L"the function changed or freed argument " + std::to_wstring(i + 1));
    for (auto v : originals) free_tree(v);

    if (!out) {
        fault(L"returned a null pointer");
        return "ok 0 e:" + std::to_string(xlerrNum);
    }
    check_shape(*out, false);
    const std::string value = format_value(*out);
    const bool dll_free = (out->xltype & xlbitDLLFree) != 0;
    if (out->xltype & xlbitXLFree) {
        // A value Excel gave the add-in, handed back for Excel to free. Excel frees what it
        // points to, not the XLOPER12 itself, so that must be static.
        if (!dll_free && !in_xll_image(out))
            fault(L"returned an xlbitXLFree value in heap memory; the XLOPER12 itself is never freed");
        excel_free(*out);
    } else if (dll_free) {
        using Free = void (*)(LPXLOPER12);
        if (const auto free_fn = reinterpret_cast<Free>(GetProcAddress(xll, "xlAutoFree12")))
            free_and_check(free_fn, out);
        else fault(L"returned xlbitDLLFree but exports no xlAutoFree12");
    } else if (!all_static(*out)) {
        fault(L"returned memory outside the add-in's static data without xlbitDLLFree; "
              L"it is never freed");
    }
    return std::string("ok ") + (dll_free ? "1 " : "0 ") + value;
}

// ---------------------------------------------------------------------------------------
// Pyinex's console window, as the "console" and "close-console" requests see it: found by
// its class among this process's windows.

HWND pyinex_console() {
    HWND found = nullptr;
    EnumWindows(
        [](HWND window, LPARAM param) -> BOOL {
            DWORD process = 0;
            wchar_t name[32] = {};
            GetWindowThreadProcessId(window, &process);
            if (process != GetCurrentProcessId() || !GetClassNameW(window, name, 32) ||
                std::wcscmp(name, L"PyinexConsole") != 0)
                return TRUE;
            *reinterpret_cast<HWND*>(param) = window;
            return FALSE;
        },
        reinterpret_cast<LPARAM>(&found));
    return found;
}

std::wstring text_of_window(HWND window) {
    std::wstring text(GetWindowTextLengthW(window) + 1, L'\0');
    text.resize(GetWindowTextW(window, text.data(), static_cast<int>(text.size())));
    return text;
}

std::string do_console() {
    const HWND window = pyinex_console();
    if (!window) return "ok b:0";
    RECT r{};
    GetWindowRect(window, &r);
    std::string out = "ok b:1 s:" + hex(text_of_window(window));
    for (const LONG n : {r.left, r.top, r.right - r.left, r.bottom - r.top}) out += " n:" + std::to_string(n);
    const HWND text_box = FindWindowExW(window, nullptr, L"EDIT", nullptr);
    out += " s:" + hex(text_box ? text_of_window(text_box) : L"");
    // Its colors, as it answers the text box's request for them while drawing.
    const HDC dc = CreateCompatibleDC(nullptr);
    SendMessageW(window, WM_CTLCOLORSTATIC, reinterpret_cast<WPARAM>(dc), reinterpret_cast<LPARAM>(text_box));
    out += " n:" + std::to_string(GetTextColor(dc)) + " n:" + std::to_string(GetBkColor(dc));
    DeleteDC(dc);
    return out;
}

// Volatile keys standing in for this process's HKEY_CURRENT_USER and HKEY_LOCAL_MACHINE,
// so that what the add-in reads there (the app mode, registered Pythons) is what a test
// wrote, and the user's own registry is neither read nor changed. forget_registry removes
// them as the host exits; keys a killed host leaves go when the user signs out, and are
// deleted first in case this host has a killed one's process ID.
HKEY user = nullptr;
HKEY machine = nullptr;
const std::wstring stand_in_key = L"Software\\PyinexHostTests\\xllhost-" + std::to_wstring(GetCurrentProcessId());

bool create_key(HKEY parent, const std::wstring& name, HKEY* key) {
    return RegCreateKeyExW(parent, name.c_str(), 0, nullptr, REG_OPTION_VOLATILE, KEY_ALL_ACCESS, nullptr, key,
                           nullptr) == ERROR_SUCCESS;
}

// Copies a key's values and subkeys into a volatile key (RegCopyTreeW makes keys that are
// not volatile, which a volatile key cannot hold).
void copy_key(HKEY from, HKEY to) {
    std::wstring name(16384, L'\0');
    std::vector<BYTE> data(65536);
    for (DWORD i = 0;; ++i) {
        DWORD name_size = static_cast<DWORD>(name.size()), size = static_cast<DWORD>(data.size()), type = 0;
        if (RegEnumValueW(from, i, name.data(), &name_size, nullptr, &type, data.data(), &size) != ERROR_SUCCESS) break;
        RegSetValueExW(to, name.c_str(), 0, type, data.data(), size);
    }
    for (DWORD i = 0;; ++i) {
        DWORD name_size = static_cast<DWORD>(name.size());
        if (RegEnumKeyExW(from, i, name.data(), &name_size, nullptr, nullptr, nullptr, nullptr) != ERROR_SUCCESS) break;
        const std::wstring child(name.c_str(), name_size);
        HKEY source = nullptr, target = nullptr;
        if (RegOpenKeyExW(from, child.c_str(), 0, KEY_READ, &source) == ERROR_SUCCESS && create_key(to, child, &target))
            copy_key(source, target);
        if (source) RegCloseKey(source);
        if (target) RegCloseKey(target);
    }
}

bool stand_in_registry() {
    HKEY root = nullptr;
    RegDeleteTreeW(HKEY_CURRENT_USER, stand_in_key.c_str());
    bool ok = create_key(HKEY_CURRENT_USER, stand_in_key, &root) && create_key(root, L"user", &user) &&
              create_key(root, L"machine", &machine);
    // Windows reads its cryptography settings from HKEY_LOCAL_MACHINE whenever a program,
    // Python's start-up among them, asks for random numbers; so the stand-in has a copy.
    for (const wchar_t* path : {L"SOFTWARE\\Microsoft\\Cryptography\\Defaults", L"SYSTEM\\CurrentControlSet\\Control\\Cryptography\\Configuration"}) {
        HKEY source = nullptr, target = nullptr;
        if (ok && RegOpenKeyExW(HKEY_LOCAL_MACHINE, path, 0, KEY_READ, &source) == ERROR_SUCCESS &&
            create_key(machine, path, &target))
            copy_key(source, target);
        if (source) RegCloseKey(source);
        if (target) RegCloseKey(target);
    }
    ok = ok && RegOverridePredefKey(HKEY_CURRENT_USER, user) == ERROR_SUCCESS &&
         RegOverridePredefKey(HKEY_LOCAL_MACHINE, machine) == ERROR_SUCCESS;
    if (root) RegCloseKey(root);
    return ok;
}

void set_theme(bool dark) {
    HKEY key = nullptr;
    if (RegCreateKeyExW(user, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 0, nullptr,
                        REG_OPTION_VOLATILE, KEY_SET_VALUE, nullptr, &key, nullptr) != ERROR_SUCCESS)
        throw std::runtime_error("cannot write the app mode");
    const DWORD light = dark ? 0 : 1;
    RegSetValueExW(key, L"AppsUseLightTheme", 0, REG_DWORD, reinterpret_cast<const BYTE*>(&light), sizeof light);
    RegCloseKey(key);
    if (const HWND window = pyinex_console())
        SendMessageW(window, WM_SETTINGCHANGE, 0, reinterpret_cast<LPARAM>(L"ImmersiveColorSet"));
}

std::string register_python(std::istringstream& in) {
    std::string hive, version, executable, folder;
    in >> hive >> version >> executable >> folder;
    const std::wstring path = L"Software\\Python\\PythonCore\\" + unhex(version) + L"\\InstallPath";
    HKEY key = nullptr;
    if (RegCreateKeyExW(hive == "machine" ? machine : user, path.c_str(), 0, nullptr, REG_OPTION_VOLATILE,
                        KEY_SET_VALUE, nullptr, &key, nullptr) != ERROR_SUCCESS)
        return "error " + hex(L"cannot record the Python");
    for (const auto& [name, text] : {std::pair<const wchar_t*, std::string>{L"ExecutablePath", executable}, {nullptr, folder}}) {
        if (text == "-") continue;
        const std::wstring value = unhex(text);
        RegSetValueExW(key, name, 0, REG_SZ, reinterpret_cast<const BYTE*>(value.c_str()),
                       static_cast<DWORD>((value.size() + 1) * sizeof(wchar_t)));
    }
    RegCloseKey(key);
    return "ok";
}

void forget_registry() {
    RegOverridePredefKey(HKEY_CURRENT_USER, nullptr);
    RegOverridePredefKey(HKEY_LOCAL_MACHINE, nullptr);
    if (user) RegCloseKey(user);
    if (machine) RegCloseKey(machine);
    RegDeleteTreeW(HKEY_CURRENT_USER, stand_in_key.c_str());
    RegDeleteKeyW(HKEY_CURRENT_USER, L"Software\\PyinexHostTests");  // Only once no other host's key is in it.
}

std::string do_registrations() {
    std::string out = "ok";
    for (const auto& r : registrations) {
        out += " a:1x" + std::to_string(6 + r.argument_help.size());
        for (const auto* s : {&r.name, &r.procedure, &r.type, &r.arguments, &r.category, &r.help})
            out += " s:" + hex(*s);
        for (const auto& h : r.argument_help) out += " s:" + hex(h);
    }
    return out;
}

int call_entry(const char* name) {
    using Entry = int (*)();
    const auto entry = reinterpret_cast<Entry>(GetProcAddress(xll, name));
    if (!entry) {
        fault(L"the add-in does not export " + std::wstring(name, name + strlen(name)));
        return 0;
    }
    return entry();
}

std::string handle(const std::string& line) {
    for (MSG message; PeekMessageW(&message, nullptr, 0, 0, PM_REMOVE);) {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
    std::istringstream in(line);
    std::string verb;
    in >> verb;
    std::string response;
    if (verb == "open") {
        std::string path;
        in >> path;
        xll_path = unhex(path);
        xll = LoadLibraryW(xll_path.c_str());
        if (!xll) return "error " + hex(L"LoadLibrary failed with error " + std::to_wstring(GetLastError()));
        response = "ok n:" + std::to_string(call_entry("xlAutoOpen"));
    } else if (verb == "load") {
        std::string path;
        in >> path;
        if (!LoadLibraryExW(unhex(path).c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH))
            return "error " + hex(L"LoadLibrary failed with error " + std::to_wstring(GetLastError()));
        return "ok";
    } else if (verb == "close") {
        response = "ok n:" + std::to_string(call_entry("xlAutoClose"));
    } else if (verb == "unload") {
        if (!xll) return "error " + hex(L"no add-in is loaded");
        FreeLibrary(xll);
        xll = nullptr;
        registrations.clear();
        const bool core = GetModuleHandleW(L"pyinex_core.dll") != nullptr;
        return std::string("ok b:") + (GetModuleHandleW(xll_path.c_str()) ? "1" : "0") + (core ? " b:1" : " b:0");
    } else if (verb == "call") {
        response = do_call(in);
    } else if (verb == "registrations") {
        return do_registrations();
    } else if (verb == "caller") {
        std::string sheet;
        in >> sheet;
        caller_set = sheet != "none";
        if (caller_set) {
            caller_sheet = unhex(sheet);
            int row, column, rows, columns;
            in >> row >> column >> rows >> columns;
            std::string kind;
            caller_sref = (in >> kind) && kind == "sref";
            caller_cell = {.rwFirst = row - 1, .rwLast = row + rows - 2,
                           .colFirst = column - 1, .colLast = column + columns - 2};
        }
        return "ok";
    } else if (verb == "abort") {
        std::string on;
        in >> on;
        escape_pressed = on == "1";
        return "ok";
    } else if (verb == "folder") {
        std::string folder;
        in >> folder;
        workbook_folder = folder == "none" ? L"" : unhex(folder);
        std::string alone;
        path_alone = (in >> alone) && alone == "alone";
        return "ok";
    } else if (verb == "wizard") {
        std::string on;
        in >> on;
        set_wizard(on == "1");
        return "ok";
    } else if (verb == "calls") {
        response = "ok n:" + std::to_string(excel_calls);
        excel_calls = 0;
        return response;
    } else if (verb == "handles") {
        DWORD count = 0;
        GetProcessHandleCount(GetCurrentProcess(), &count);
        return "ok n:" + std::to_string(count);
    } else if (verb == "console") {
        return do_console();
    } else if (verb == "close-console") {
        const HWND window = pyinex_console();
        if (!window) return "error " + hex(L"no console window");
        SendMessageW(window, WM_SYSCOMMAND, SC_CLOSE, 0);
        return "ok";
    } else if (verb == "theme") {
        std::string mode;
        in >> mode;
        set_theme(mode == "dark");
        return "ok";
    } else if (verb == "register") {
        return register_python(in);
    } else if (verb == "write") {
        std::string text;
        in >> text;
        const std::wstring wide = unhex(text);
        const HANDLE out = GetStdHandle(STD_ERROR_HANDLE);
        if (!out || out == INVALID_HANDLE_VALUE) return "error " + hex(L"no standard error");
        const int size = WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), nullptr, 0,
                                             nullptr, nullptr);
        std::string utf8(size, '\0');
        WideCharToMultiByte(CP_UTF8, 0, wide.data(), static_cast<int>(wide.size()), utf8.data(), size, nullptr,
                            nullptr);
        DWORD written = 0;
        WriteFile(out, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
        return "ok";
    } else if (verb == "faults") {
        response = "ok";
        for (const auto& f : faults) response += " s:" + hex(f);
        faults.clear();
        return response;
    } else {
        return "error " + hex(L"unknown request");
    }
    check_leaks();
    return response;
}

FILE* open_handle(const char* text, const char* mode, int flags) {
    const auto handle = reinterpret_cast<HANDLE>(std::strtoull(text, nullptr, 10));
    // Processes the add-in starts must not inherit the pipes and hold them open.
    SetHandleInformation(handle, HANDLE_FLAG_INHERIT, 0);
    const auto h = reinterpret_cast<intptr_t>(handle);
    const int fd = _open_osfhandle(h, flags);
    return fd < 0 ? nullptr : _fdopen(fd, mode);
}

}  // namespace

extern "C" __declspec(dllexport) int __stdcall MdCallBack12(int function, int count,
                                                            LPXLOPER12* args, LPXLOPER12 result) {
    return excel_callback(function, count, args, result);
}

int main(int argc, char** argv) {
    const std::string option = argc == 4 ? argv[3] : "";
    if (argc < 3 || argc > 4 || (argc == 4 && option != "--no-output")) {
        std::fputs("usage: xllhost <request pipe handle> <response pipe handle> [--no-output]\n", stderr);
        return 2;
    }
    FILE* requests = open_handle(argv[1], "rb", _O_RDONLY | _O_BINARY);
    FILE* responses = open_handle(argv[2], "wb", _O_WRONLY | _O_BINARY);
    if (!requests || !responses) {
        std::fputs("xllhost: bad pipe handles\n", stderr);
        return 2;
    }
    if (option == "--no-output") {
        // Only xllhost_gui.exe, a program with a window as Excel is, started with no standard
        // handles, begins with no output: its C runtime marks stdout and stderr as having none
        // (descriptor -2). Python's runtime is ucrtbase.dll, which a debug build of this
        // program does not link, so it is checked there by name.
        const HMODULE crt = LoadLibraryW(L"ucrtbase.dll");
        const auto stream = reinterpret_cast<FILE*(__cdecl*)(unsigned)>(GetProcAddress(crt, "__acrt_iob_func"));
        const auto descriptor_of = reinterpret_cast<int(__cdecl*)(FILE*)>(GetProcAddress(crt, "_fileno"));
        if (GetStdHandle(STD_OUTPUT_HANDLE) || descriptor_of(stream(1)) >= 0 || descriptor_of(stream(2)) >= 0)
            return 3;  // Not started as Excel is: the test would prove nothing.
    }
    if (!stand_in_registry()) {
        std::fputs("xllhost: cannot stand in for the registry\n", stderr);
        forget_registry();
        return 2;
    }
    std::string line;
    for (int c; (c = std::fgetc(requests)) != EOF;) {
        if (c != '\n') {
            line += static_cast<char>(c);
            continue;
        }
        if (line == "quit") break;
        std::string response;
        try {
            response = handle(line);
        } catch (const std::exception& e) {
            response = "error " + hex(std::wstring(e.what(), e.what() + strlen(e.what())));
        }
        std::fputs((response + "\n").c_str(), responses);
        std::fflush(responses);
        line.clear();
    }
    forget_registry();
    return 0;
}
