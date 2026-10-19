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
//   open <path>                  load the add-in and call xlAutoOpen; another path loads
//                                another copy, which Excel tells apart by its path
//   load <path>                  load some other DLL, as another add-in might
//   close                        call xlAutoClose
//   unload                       FreeLibrary the add-in; answers whether the add-in and
//                                pyinex_core.dll are still loaded
//   call <name> <value>...       call a registered function
//   registrations                list what xlfRegister recorded and is still registered
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
//   run <name>                   run a registered command, as Excel runs one in a command
//                                context (the stand-in for Application.Run); answers its
//                                result
//   workbook none|open|protected|both
//                                whether a workbook window is open, from which the add-in
//                                can get the stand-in for Excel's Application object
//                                (default open); protected: a Protected View window, whose
//                                object has no Application; both: a workbook window and an
//                                active Protected View window, while which Evaluate gives #N/A
//   busy 0|1                     whether Application.Run and Evaluate refuse, as Excel does
//                                while a cell is edited or a dialog is open
//   tick <n>                     send the add-in's queue window (class PyinexQueueWindow) its
//                                timer's message (ID 1) n times, as its timer would; answers
//                                whether the window exists, and how many times
//                                Application.Run was called and refused meanwhile
//   fail register <n>            make the add-in's next n registrations of direct functions
//                                (exports PyFunc<n>) fail, as Excel's own refusal does
//                                (xlfRegister gives #VALUE!)
//   fail command <n>             the same for its next n registrations of commands
//   fail name <name>             the same for its next registration under name, in hex
//   fail unregister <n> [<skip>] after skip more calls of xlfUnregister that succeed, make
//                                the next n fail (FALSE, the registration left as it was)
//   fail setname <n>             make the next n calls of xlfSetName fail (FALSE)
//   foreign <name>               register name for another add-in, as if it had
//   fail window <n>              make the add-in's next n attempts to create its queue window
//                                fail
//   fail timer <n>               make the add-in's next n calls of SetTimer fail
//   fail evaluate <n>            make the next n calls of xlfEvaluate fail (xlretFailed)
//   fail memory evaluate|register|unregister|now [<skip>]
//                                once the add-in's next call of xlfEvaluate, its next
//                                registration of a worksheet function that succeeds, or its
//                                next call of xlfUnregister, returns, or now, make its next
//                                allocation on this thread after skip more fail, as when
//                                memory runs out (operator new throws std::bad_alloc); with
//                                now, the core's too, pyinex_core.dll, once it is loaded
//   nest <request>[;<request>...]
//                                the next time the add-in calls xlfRegister for a direct
//                                function (an export PyFunc<n>), first carry out each request,
//                                "run <name>" or "call <name> <value>...", as if Excel had
//                                done so then, "close", calling xlAutoClose, or
//                                "memory [<skip>]", making the add-in's next
//                                allocation on this thread after skip more fail
//   nestat <procedure> <request>[;<request>...]
//                                the same, at the add-in's next xlfRegister of a function or
//                                command whose procedure (its export) starts with procedure,
//                                in hex: xlAutoRemove for the stand-in a removal registers
//   cells <sheet> <row> <column> <value>
//                                put a value, or an array's values from that cell on, into
//                                a sheet's cells, which xlCoerce and COUNTA of a reference
//                                read; rows and columns from 1; z empties a cell
//   stop uncalced|abort|none [<skip>]
//                                after skip more readings of a reference (xlCoerce or
//                                COUNTA), make the next answer xlretUncalced or xlretAbort;
//                                any call into Excel but xlFree after that answer, before
//                                the add-in's function that got it returns, is a fault
//   forget <sheet>               the sheet is deleted, or its workbook closed: xlSheetId no
//                                longer finds it, and its ID goes to the next sheet named
//   formula <sheet> <row> <column> <text>
//                                put a formula in a sheet's cell. The stand-in for Excel's
//                                Application object gives its Workbooks (each named in a
//                                sheet's name, as Book1 in [Book1]Sheet1), their Worksheets,
//                                Cells, Find and FindNext of text in formulas and in text
//                                values, a cell's Address, HasFormula and Formula2. Formula2
//                                entered again is recorded, and a formula =Name(...) with
//                                numbers and quoted texts as arguments is calculated, calling
//                                the function from that cell
//   array <sheet> <row> <column> <rows> <columns> <text>
//                                put a legacy array formula (Ctrl+Shift+Enter) in an area
//                                from that cell: each of its cells holds the formula and
//                                HasArray, CurrentArray gives the area, and only the area's
//                                FormulaArray enters it again; Formula2 entered in one of its
//                                cells fails, as Excel refuses to change part of an array
//   lock <sheet> 0|1|2           whether entering a formula on the sheet fails, as on a
//                                protected sheet; with 2, reading one fails too
//   change <request>             after each formula entered again, carry out <request>, "run
//                                <name>", as a workbook's change handler running a command
//                                would; a fault once it has run 25 times, as a repair that
//                                never ends would run it forever; no request: none
//   handler <text>               after each formula entered again, a search for text, as a
//                                workbook's change handler calling Range.Find would make,
//                                so that FindNext continues it; empty text: none
//   opened <book> [new]          Excel reports the workbook opened (WorkbookOpen), or with
//                                new made from a template (NewWorkbook), to each handler
//                                connected to its application events; answers whether there
//                                was one
//   entered                      how many times xlcCalculateNow was called, and each formula
//                                entered again: its sheet, address, formula and value; since
//                                the last "entered"
// EVALUATE (xlfEvaluate) evaluates only a name alone, as Excel does: a name a registration
// holds gives its ID, one of a few of Excel's own functions (built_in) #VALUE!, and any other
// #NAME?.
// xlfRegister, xlfUnregister and xlfSetName are allowed only in a command context
// (xlAutoOpen, xlAutoClose, or a command run by "run" or Application.Run), as in Excel;
// anywhere else they fail, and the host reports a fault.
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
//   r:<sheet>:<area>[;<area>...]  a reference (xltypeRef) to areas of the sheet named, such as
//                                [Book1]Sheet1; each area <first row>,<last row>,<first
//                                column>,<last column>, counted from 0, as XLREF12 has them
//   R:<area>                     a reference with no sheet (xltypeSRef), on the sheet
//                                calculating
// Text (names, paths, s: values) is UTF-16 as four hex digits per code unit, so it
// carries any character, spaces and unpaired surrogates included.
#include <windows.h>

#include <XLCALL.H>
#include <oleacc.h>
#include <olectl.h>

#include <fcntl.h>
#include <io.h>
#include <malloc.h>

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <cwchar>
#include <map>
#include <set>
#include <sstream>
#include <stdexcept>
#include <string>
#include <tuple>
#include <utility>
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
    case xltypeRef: ::operator delete(x.val.mref.lpmref); break;
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

IDSHEET sheet_id(const std::wstring& name);

// An area as the protocol writes it: first row, last row, first column, last column.
XLREF12 parse_area(const std::string& text) {
    XLREF12 a{};
    char comma = 0;
    std::istringstream in(text);
    if (!(in >> a.rwFirst >> comma >> a.rwLast >> comma >> a.colFirst >> comma >> a.colLast))
        throw std::runtime_error("bad area: " + text);
    return a;
}

std::string format_area(const XLREF12& a) {
    return std::to_string(a.rwFirst) + "," + std::to_string(a.rwLast) + "," + std::to_string(a.colFirst) + "," +
           std::to_string(a.colLast);
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
    case 'R':
        x.xltype = xltypeSRef;
        x.val.sref = {.count = 1, .ref = parse_area(body)};
        return x;
    case 'r': {
        const auto colon = body.find(':');
        std::vector<XLREF12> areas;
        for (size_t at = colon + 1; colon != std::string::npos && at <= body.size();) {
            const auto end = std::min(body.find(';', at), body.size());
            areas.push_back(parse_area(body.substr(at, end - at)));
            at = end + 1;
        }
        if (areas.empty()) throw std::runtime_error("bad reference: " + token);
        // XLMREF12 holds one area; more follow it in the same block, as Excel lays them out.
        auto* m = static_cast<XLMREF12*>(::operator new(sizeof(XLMREF12) + (areas.size() - 1) * sizeof(XLREF12)));
        m->count = static_cast<WORD>(areas.size());
        for (size_t i = 0; i < areas.size(); ++i) m->reftbl[i] = areas[i];
        x.xltype = xltypeRef;
        x.val.mref = {.lpmref = m, .idSheet = sheet_id(unhex(body.substr(0, colon)))};
        return x;
    }
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
    case xltypeSRef: return "R:" + format_area(x.val.sref.ref);
    case xltypeRef: {
        std::string out = "r:" + std::to_string(x.val.mref.idSheet) + ":";
        for (WORD i = 0; x.val.mref.lpmref && i < x.val.mref.lpmref->count; ++i)
            out += (i ? ";" : "") + format_area(x.val.mref.lpmref->reftbl[i]);
        return out;
    }
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
    std::wstring module, procedure, type, name, arguments, category, help;
    std::vector<std::wstring> argument_help;
    FARPROC address = nullptr;
    int argument_count = 0;
    bool registered = true;  // False once unregistered.
    bool command = false;    // Macro type 2.
};
// Excel keys a registration by its module and procedure: registering one again, under any
// name, gives the same ID, here its index plus one.
std::vector<Registration> registrations;

// Each name a formula can use, compared without regard to case, and the registration ID it
// calls; 0 once xlfSetName has deleted it. As in Excel, unregistering leaves the name.
struct NoCase {
    bool operator()(const std::wstring& a, const std::wstring& b) const { return _wcsicmp(a.c_str(), b.c_str()) < 0; }
};
std::map<std::wstring, size_t, NoCase> names;

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

// Where Excel is calling the add-in from: registering is allowed only in a command.
enum class Context { none, command, function };
Context context = Context::none;

// Sets the context for one call into the add-in, and puts the previous one back after.
class InContext {
public:
    explicit InContext(Context c) : saved_(context) { context = c; }
    InContext(const InContext&) = delete;
    InContext& operator=(const InContext&) = delete;
    ~InContext() { context = saved_; }

private:
    Context saved_;
};

bool in_command(const wchar_t* function) {
    if (context == Context::command) return true;
    fault(std::wstring(function) + (context == Context::function ? L" called from a worksheet function"
                                                                  : L" called outside any call from Excel") +
          L", which Excel refuses");
    return false;
}

int failing_registrations = 0;  // The "fail register" request.
int failing_commands = 0;       // "fail command".
std::wstring failing_name;      // "fail name".
int failing_unregisters = 0, skipped_unregisters = 0;  // "fail unregister".
int failing_set_names = 0;      // "fail setname".
int failing_evaluates = 0;      // "fail evaluate".
// "fail memory": what arms the failure, and whether the add-in's next allocation fails.
enum class Arm { none, evaluate, registration, unregister } memory_armed = Arm::none;
bool failing_allocation = false;
int skipped_allocations = 0;    // Allocations that still succeed first.
std::string nested_request;     // The "nest" request,
std::wstring nested_at;         // and the start of the procedure whose registration it awaits.
std::string carry_out(const std::string& request);

// The ID Excel gives the registration numbered n (from 1): negative, as real Excel's often
// are, so an add-in that takes an ID for positive is caught here.
double registration_id(size_t n) { return -static_cast<double>(n); }

int do_register(int count, LPXLOPER12* args, LPXLOPER12 result) {
    if (!in_command(L"xlfRegister")) return xlretInvXlfn;
    if (!nested_request.empty() && count > 1 && text_of(args[1]).starts_with(nested_at))
        carry_out(std::exchange(nested_request, ""));
    if (count < 4) {
        fault(L"xlfRegister needs at least module, procedure, type and name");
        return xlretInvCount;
    }
    Registration r;
    r.module = text_of(args[0]);
    if (_wcsicmp(r.module.c_str(), xll_path.c_str()) != 0) fault(L"xlfRegister module is not the add-in's path: " + r.module);
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
    // A command (macro type 2), such as the one Pyinex registers to clear a name, is never
    // called here, so its type string is not checked.
    r.command = count > 5 && args[5]->xltype == xltypeNum && args[5]->val.num == 2;
    r.argument_count = r.command ? 0 : count_arguments(r.type);
    XLOPER12 id{.xltype = xltypeErr};
    id.val.err = xlerrValue;
    if (r.procedure.starts_with(L"PyFunc") && failing_registrations > 0) --failing_registrations;  // Refused, as Excel refuses.
    else if (r.command && failing_commands > 0) --failing_commands;
    else if (!failing_name.empty() && r.name == failing_name) failing_name.clear();
    else if (!r.address) fault(L"xlfRegister procedure is not exported: " + r.procedure);
    else if (r.argument_count < 0) fault(L"xlfRegister type string not supported by xllhost: " + r.type);
    else if (r.argument_count > 32) fault(L"xlfRegister: more than 32 arguments: " + r.name);
    else {
        size_t i = 0;
        while (i < registrations.size() && (_wcsicmp(registrations[i].procedure.c_str(), r.procedure.c_str()) != 0 ||
                                            _wcsicmp(registrations[i].module.c_str(), r.module.c_str()) != 0))
            ++i;
        if (i == registrations.size()) registrations.push_back(r);
        else registrations[i] = r;
        names[r.name] = i + 1;
        id = make_num(registration_id(i + 1));
        if (!r.command && std::exchange(memory_armed, memory_armed == Arm::registration ? Arm::none : memory_armed) == Arm::registration)
            failing_allocation = true;
    }
    hand_over(id, result);
    return xlretSuccess;
}

// xlfUnregister(ID): TRUE if that registration was there.
int do_unregister(int count, LPXLOPER12* args, LPXLOPER12 result) {
    if (!in_command(L"xlfUnregister")) return xlretInvXlfn;
    if (count != 1 || args[0]->xltype != xltypeNum) {
        fault(L"xllhost implements only xlfUnregister(registration ID)");
        return xlretInvXlfn;
    }
    const double id = -args[0]->val.num;  // As registration_id gave it: the registration's number.
    const bool refused = failing_unregisters > 0 && skipped_unregisters-- <= 0;
    if (refused) --failing_unregisters;
    if (memory_armed == Arm::unregister) {
        memory_armed = Arm::none;
        failing_allocation = true;
    }
    const bool found = !refused && id >= 1 && id <= static_cast<double>(registrations.size()) &&
                       id == std::floor(id) && registrations[static_cast<size_t>(id) - 1].registered;
    if (found) registrations[static_cast<size_t>(id) - 1].registered = false;
    XLOPER12 done{.xltype = xltypeBool};
    done.val.xbool = found;
    hand_over(done, result);
    return xlretSuccess;
}

// xlfSetName(name), with no value: deletes the name; TRUE if there was one.
int do_set_name(int count, LPXLOPER12* args, LPXLOPER12 result) {
    if (!in_command(L"xlfSetName")) return xlretInvXlfn;
    if (count != 1 || (args[0]->xltype & ~(xlbitDLLFree | xlbitXLFree)) != xltypeStr) {
        fault(L"xllhost implements only xlfSetName(name), which deletes the name");
        return xlretInvXlfn;
    }
    const auto name = names.find(text_of(args[0]));
    XLOPER12 done{.xltype = xltypeBool};
    done.val.xbool = name != names.end() && name->second != 0 && !(failing_set_names > 0 && failing_set_names--);
    if (done.val.xbool) name->second = 0;
    hand_over(done, result);
    return xlretSuccess;
}

const wchar_t* const built_in[] = {L"SUM", L"ABS", L"NOW", L"IF", L"CONCAT"};

// EVALUATE (xlfEvaluate), for a name alone: see the header.
int do_evaluate(int count, LPXLOPER12* args, LPXLOPER12 result) {
    const std::wstring name = text_of(args[0]);
    if (count != 1 || name.empty() || name.find_first_of(L"=()\"") != std::wstring::npos) {
        fault(L"xllhost implements only EVALUATE(name)");
        return xlretInvXlfn;
    }
    if (memory_armed == Arm::evaluate) {
        memory_armed = Arm::none;
        failing_allocation = true;
    }
    if (failing_evaluates > 0) {
        --failing_evaluates;
        return xlretFailed;
    }
    XLOPER12 value{.xltype = xltypeErr};
    value.val.err = xlerrName;
    for (const wchar_t* b : built_in)
        if (_wcsicmp(b, name.c_str()) == 0) value.val.err = xlerrValue;
    if (const auto known = names.find(name); known != names.end() && known->second != 0)
        value = make_num(registration_id(known->second));
    hand_over(value, result);
    return xlretSuccess;
}

// ---------------------------------------------------------------------------------------
// Excel's other services the add-in may call.

std::wstring caller_sheet = L"[Book1]Sheet1";
bool caller_set = true;
bool caller_sref = false;
XLREF12 caller_cell{.rwFirst = 0, .rwLast = 0, .colFirst = 0, .colLast = 0};
std::vector<std::wstring> sheets = {L"[Book1]Sheet1"};  // idSheet is index + 1; empty: forgotten.
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

// The sheet's ID, or 0 for none.
IDSHEET known_sheet(const std::wstring& name) {
    const auto at = std::find(sheets.begin(), sheets.end(), name);
    return name.empty() || at == sheets.end() ? 0 : static_cast<IDSHEET>(at - sheets.begin() + 1);
}

// The sheet's ID, naming a new sheet if need be, with a forgotten sheet's ID if there is one.
IDSHEET sheet_id(const std::wstring& name) {
    if (const IDSHEET id = known_sheet(name)) return id;
    auto slot = std::find(sheets.begin(), sheets.end(), L"");
    if (slot == sheets.end()) slot = sheets.insert(slot, L"");
    *slot = name;
    return static_cast<IDSHEET>(slot - sheets.begin() + 1);
}

XLOPER12 make_ref(IDSHEET sheet, const XLREF12& area) {
    auto* m = static_cast<XLMREF12*>(::operator new(sizeof(XLMREF12)));  // As parse_value makes one.
    m->count = 1;
    m->reftbl[0] = area;
    XLOPER12 x{.xltype = xltypeRef};
    x.val.mref.lpmref = m;
    x.val.mref.idSheet = sheet;
    return x;
}

// Each sheet's cells, by sheet ID, then row and column counted from 0: plain values only.
std::map<IDSHEET, std::map<std::pair<int, int>, XLOPER12>> grid;

// Each sheet's formulas, the same way ("formula"); a cell with none holds its value in grid.
std::map<IDSHEET, std::map<std::pair<int, int>, std::wstring>> formulas;
std::map<IDSHEET, int> locked;  // "lock": 1, entering a formula there fails; 2, reading one too.
// "array": each sheet's legacy array formulas, by their first cell and their size.
std::map<IDSHEET, std::map<std::pair<int, int>, std::pair<int, int>>> arrays;
std::wstring change_search;     // "handler": what each formula entered again searches for.
std::string change_request;     // "change": what each formula entered again carries out,
int change_runs = 0;            // and how many times it has.
std::string entered;          // Each formula entered again since the last "entered", as tokens.
int entered_count = 0;
int calculations = 0;         // xlcCalculateNow, since the last "entered".

// The "stop" request: after skip more readings of a reference, answer code; stopped, the
// code given, until the add-in's function that got it returns.
int stop_with = 0, stop_skip = 0, stopped = 0;

// A reference's sheet and its one area; false, after a fault, for several areas or a
// whole column or row read cell by cell, which no add-in should ask for.
bool area_of(const XLOPER12& x, IDSHEET& sheet, XLREF12& area, bool reading) {
    const int type = x.xltype & ~(xlbitDLLFree | xlbitXLFree);
    if (type == xltypeSRef) {
        sheet = sheet_id(caller_sheet);
        area = x.val.sref.ref;
    } else {
        sheet = x.val.mref.idSheet;
        if (!x.val.mref.lpmref || x.val.mref.lpmref->count != 1) {
            fault(L"xllhost reads only a reference of one area");
            return false;
        }
        area = x.val.mref.lpmref->reftbl[0];
    }
    if (reading && (area.rwLast - area.rwFirst + 1 == 1048576 || area.colLast - area.colFirst + 1 == 16384)) {
        fault(L"read every cell of a whole column or row");
        return false;
    }
    return true;
}

// Whether the next reading of a reference stops, as "stop" asked; the code if so.
int stopping() {
    if (stop_with == 0 || stop_skip-- > 0) return 0;
    stopped = std::exchange(stop_with, 0);
    return stopped;
}

XLOPER12 coerce(const XLOPER12& x, int to) {
    const int type = x.xltype & ~(xlbitDLLFree | xlbitXLFree);
    if (type == xltypeRef || type == xltypeSRef) {  // The cells, as an array.
        IDSHEET sheet = 0;
        XLREF12 a{};
        if (!area_of(x, sheet, a, true)) {
            XLOPER12 error{.xltype = xltypeErr};
            error.val.err = xlerrValue;
            return error;
        }
        const int rows = a.rwLast - a.rwFirst + 1, columns = a.colLast - a.colFirst + 1;
        auto* items = new XLOPER12[static_cast<size_t>(rows) * columns];
        for (int r = 0; r < rows; ++r)
            for (int c = 0; c < columns; ++c) {
                const auto& cells = grid[sheet];
                const auto cell = cells.find({a.rwFirst + r, a.colFirst + c});
                items[r * columns + c] = cell == cells.end() ? XLOPER12{.xltype = xltypeNil} : coerce(cell->second, 0);
            }
        XLOPER12 multi{.xltype = xltypeMulti};
        multi.val.array = {.lparray = items, .rows = rows, .columns = columns};
        return multi;
    }
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
    if (stopped && function != xlFree)
        fault(L"called Excel (function " + std::to_wstring(function) + L") after Excel answered " +
              (stopped == xlretUncalced ? L"xlretUncalced" : L"xlretAbort") + L", which Microsoft forbids");
    const int reference = count ? args[0]->xltype & ~(xlbitDLLFree | xlbitXLFree) : 0;
    if ((function == xlCoerce || function == xlfCounta) && (reference == xltypeRef || reference == xltypeSRef))
        if (const int code = stopping()) return code;
    switch (function) {
    case xlFree:
        for (int i = 0; i < count; ++i) excel_free(*args[i]);
        return xlretSuccess;
    case xlGetName:
        hand_over(make_text(xll_path), result);
        return xlretSuccess;
    case xlfRegister:
        return do_register(count, args, result);
    case xlfUnregister:
        return do_unregister(count, args, result);
    case xlfSetName:
        return do_set_name(count, args, result);
    case xlfEvaluate:
        return do_evaluate(count, args, result);
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
        if (id < 1 || id > sheets.size() || sheets[id - 1].empty()) return xlretFailed;
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
    case xlfCounta: {
        IDSHEET sheet = 0;
        XLREF12 a{};
        if (count != 1 || (reference != xltypeRef && reference != xltypeSRef) || !area_of(*args[0], sheet, a, false)) {
            fault(L"xllhost implements only COUNTA(reference)");
            return xlretInvXlfn;
        }
        double n = 0;
        for (const auto& [at, value] : grid[sheet])
            n += at.first >= a.rwFirst && at.first <= a.rwLast && at.second >= a.colFirst && at.second <= a.colLast;
        hand_over(make_num(n), result);
        return xlretSuccess;
    }
    case xlSheetId: {
        // As Excel gives it: a reference with the sheet's ID and no areas; failing for a sheet
        // it does not have open.
        XLOPER12 id{.xltype = xltypeRef};
        id.val.mref = {.lpmref = nullptr, .idSheet = known_sheet(text_of(count ? args[0] : nullptr))};
        if (id.val.mref.idSheet == 0) return xlretFailed;
        hand_over(id, result);
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
    case xlcCalculateNow:
        if (!in_command(L"xlcCalculateNow")) return xlretInvXlfn;
        ++calculations;
        return xlretSuccess;
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
    const auto known = names.find(wanted);
    if (known == names.end()) return "error " + hex(L"no registered function " + wanted);
    // A name whose function was unregistered, or deleted, gives #VALUE!, as in Excel.
    if (known->second == 0 || !registrations[known->second - 1].registered)
        return "ok 0 e:" + std::to_string(xlerrValue);
    const Registration* r = &registrations[known->second - 1];

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
    LPXLOPER12 out;
    {
        const InContext calculating(Context::function);
        // A stop lasts for the call it came in: a call nested in this one has its own.
        const int outer = std::exchange(stopped, 0);
        out = fn(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], a[9], a[10], a[11], a[12], a[13], a[14],
                 a[15], a[16], a[17], a[18], a[19], a[20], a[21], a[22], a[23], a[24], a[25], a[26], a[27], a[28],
                 a[29], a[30], a[31]);
        stopped = outer;
    }
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
// Commands, and a stand-in for Excel's Application object, through which the add-in runs
// one as Excel's COM interface lets it (Application.Run).

// Runs a registered command in a command context; -1 if no command has the name.
int run_command(const std::wstring& name) {
    const auto known = names.find(name);
    if (known == names.end() || known->second == 0) return -1;
    const Registration& r = registrations[known->second - 1];
    if (!r.registered || !r.command || !r.address) return -1;
    const InContext running(Context::command);
    return reinterpret_cast<int (*)()>(r.address)();
}

// ---------------------------------------------------------------------------------------
// Workbooks, as far as the add-in reaches them through COM, to follow them as they open and
// to repair cells: the "formula", "lock", "opened" and "entered" requests.

// The workbooks open: each named in a sheet's name, as [Book1] in [Book1]Sheet1, in order.
std::vector<std::wstring> workbooks() {
    std::vector<std::wstring> books;
    for (const auto& s : sheets)
        if (const auto end = s.find(L']'); !s.empty() && end != std::wstring::npos &&
                                           std::find(books.begin(), books.end(), s.substr(1, end - 1)) == books.end())
            books.push_back(s.substr(1, end - 1));
    return books;
}

std::vector<IDSHEET> worksheets(const std::wstring& book) {
    std::vector<IDSHEET> ids;
    for (size_t i = 0; i < sheets.size(); ++i)
        if (sheets[i].starts_with(L"[" + book + L"]")) ids.push_back(static_cast<IDSHEET>(i + 1));
    return ids;
}

std::wstring upper(std::wstring s) {
    CharUpperBuffW(s.data(), static_cast<DWORD>(s.size()));
    return s;
}

// The cells of a sheet whose formula, or text, holds what in any case, by rows and then
// columns, as Excel's search of formulas (Range.Find, looking in formulas) finds them.
std::vector<std::pair<int, int>> search(IDSHEET sheet, const std::wstring& what) {
    std::set<std::pair<int, int>> at;
    for (const auto& [cell, formula] : formulas[sheet])
        if (upper(formula).find(upper(what)) != std::wstring::npos) at.insert(cell);
    for (const auto& [cell, value] : grid[sheet])
        if (value.xltype == xltypeStr && upper(text_of(&value)).find(upper(what)) != std::wstring::npos) at.insert(cell);
    return {at.begin(), at.end()};
}

std::wstring address(std::pair<int, int> cell, std::pair<int, int> size = {1, 1}) {
    std::wstring letters;
    for (int c = cell.second + 1; c > 0; c = (c - 1) / 26) letters.insert(letters.begin(), static_cast<wchar_t>(L'A' + (c - 1) % 26));
    const std::wstring one = L"$" + letters + L"$" + std::to_wstring(cell.first + 1);
    if (size == std::pair{1, 1}) return one;
    return one + L":" + address({cell.first + size.first - 1, cell.second + size.second - 1});
}

// The legacy array formula a cell is part of, as its first cell and size; size 0 for none.
std::pair<std::pair<int, int>, std::pair<int, int>> array_at(IDSHEET sheet, std::pair<int, int> cell) {
    for (const auto& [first, size] : arrays[sheet])
        if (cell.first >= first.first && cell.first < first.first + size.first && cell.second >= first.second &&
            cell.second < first.second + size.second)
            return {first, size};
    return {{-1, -1}, {0, 0}};
}

// Calculates a formula entered in a cell, as Excel would, when it is =Name(...) with numbers
// and texts in double quotes as arguments: calls the function as "call" does, from that
// cell. Returns the value as a token; z for a formula it does not calculate.
std::string calculate(IDSHEET sheet, std::pair<int, int> cell, const std::wstring& formula) {
    const size_t open = formula.find(L'(');
    if (formula.size() < 4 || formula[0] != L'=' || open == std::wstring::npos || formula.back() != L')') return "z";
    std::string request = hex(formula.substr(1, open - 1));
    const std::wstring inside = formula.substr(open + 1, formula.size() - open - 2);
    for (size_t at = 0; !inside.empty() && at <= inside.size();) {
        const size_t comma = std::min(inside.find(L',', at), inside.size());
        std::wstring piece = inside.substr(at, comma - at);
        at = comma + 1;
        if (piece.size() >= 2 && piece.front() == L'"' && piece.back() == L'"') {
            request += " s:" + hex(piece.substr(1, piece.size() - 2));
            continue;
        }
        wchar_t* end = nullptr;
        const double n = std::wcstod(piece.c_str(), &end);
        if (piece.empty() || *end) return "z";
        char number[64];
        std::snprintf(number, sizeof number, " n:%.17g", n);
        request += number;
    }
    const auto saved = std::tuple{caller_sheet, caller_cell, caller_set, caller_sref};
    caller_sheet = sheets[sheet - 1];
    caller_cell = {.rwFirst = cell.first, .rwLast = cell.first, .colFirst = cell.second, .colLast = cell.second};
    caller_set = true;
    caller_sref = false;
    std::istringstream in(request);
    const std::string result = do_call(in);
    std::tie(caller_sheet, caller_cell, caller_set, caller_sref) = saved;
    return result.starts_with("ok ") ? result.substr(5) : "z";
}

std::wstring finding;  // What the last Find looked for, as FindNext looks for it too.

// Excel's objects below Application, as far as the add-in uses them: Workbooks; a Workbook
// (named name); its Worksheets; a Worksheet (sheet); and a Range, every cell of a sheet
// (Cells) or one cell. Its reference count is kept, unlike Application's.
class Node final : public IDispatch {
public:
    enum Kind { books, book, sheets_of, sheet, range };
    Node(Kind kind, std::wstring name, IDSHEET id = 0, std::pair<int, int> cell = {-1, -1}, std::pair<int, int> size = {1, 1})
        : kind_(kind), name_(std::move(name)), id_(id), cell_(cell), size_(size) {}
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        *out = iid == IID_IUnknown || iid == IID_IDispatch ? this : nullptr;
        if (*out) AddRef();
        return *out ? S_OK : E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return ++count_; }
    ULONG STDMETHODCALLTYPE Release() override {
        const ULONG left = --count_;
        if (left == 0) delete this;
        return left;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* n) override {
        *n = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR* names_in, UINT count, LCID, DISPID* ids) override {
        for (DISPID i = 0; i < static_cast<DISPID>(std::size(members)) && count == 1; ++i)
            if (_wcsicmp(names_in[0], members[i]) == 0) {
                ids[0] = i;
                return S_OK;
            }
        return DISP_E_UNKNOWNNAME;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD flags, DISPPARAMS* params, VARIANT* result,
                                     EXCEPINFO* info, UINT*) override {
        VARIANT ignored;
        if (!result) result = &ignored;
        VariantInit(result);
        const UINT n = params ? params->cArgs : 0;
        // The arguments in order (COM gives them last first); null for one not given.
        const auto arg = [&](UINT i) -> const VARIANT* {
            const VARIANT* a = i < n ? &params->rgvarg[n - 1 - i] : nullptr;
            return a && !(a->vt == VT_ERROR && a->scode == DISP_E_PARAMNOTFOUND) ? a : nullptr;
        };
        const auto give = [&](Node* node) {
            result->vt = VT_DISPATCH;
            result->pdispVal = node;
            return S_OK;
        };
        const auto text = [&](const std::wstring& s) {
            result->vt = VT_BSTR;
            result->bstrVal = SysAllocStringLen(s.data(), static_cast<UINT>(s.size()));
            return S_OK;
        };
        switch (id) {
        case 0:  // Count
            result->vt = VT_I4;
            result->lVal = static_cast<LONG>(kind_ == books ? workbooks().size() : worksheets(name_).size());
            return S_OK;
        case 1: {  // Item(index), or a workbook's Item(name)
            const VARIANT* a = arg(0);
            const auto books_open = workbooks();
            const auto ids = worksheets(name_);
            const LONG size = static_cast<LONG>(kind_ == books ? books_open.size() : ids.size());
            if (!a) return DISP_E_BADPARAMCOUNT;
            if (kind_ == books && a->vt == VT_BSTR) {
                const auto at = std::find(books_open.begin(), books_open.end(), std::wstring(a->bstrVal));
                return at == books_open.end() ? DISP_E_BADINDEX : give(new Node(book, *at));
            }
            if (a->vt != VT_I4 || a->lVal < 1 || a->lVal > size) return DISP_E_BADINDEX;
            return kind_ == books ? give(new Node(book, books_open[a->lVal - 1])) : give(new Node(sheet, name_, ids[a->lVal - 1]));
        }
        case 2:  // Name
            return text(kind_ != sheet ? name_ : sheets[id_ - 1].substr(sheets[id_ - 1].find(L']') + 1));
        case 3:  // Worksheets
            return give(new Node(sheets_of, name_));
        case 4:  // Cells
            return give(new Node(range, name_, id_));
        case 5:    // Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase)
        case 6: {  // FindNext(After)
            if (id == 5) {
                const VARIANT *what = arg(0), *look_in = arg(2), *look_at = arg(3), *match_case = arg(6);
                if (!what || what->vt != VT_BSTR) return DISP_E_TYPEMISMATCH;
                // Formulas (-4123), any part (2), in any case: all the host's search does.
                if (!look_in || look_in->vt != VT_I4 || look_in->lVal != -4123 || !look_at || look_at->vt != VT_I4 ||
                    look_at->lVal != 2 || !match_case || match_case->vt != VT_BOOL || match_case->boolVal)
                    fault(L"xllhost implements only Find(text, , xlFormulas, xlPart, , , False)");
                finding = what->bstrVal;
            }
            std::pair<int, int> after{-1, -1};
            if (id == 6) {
                const VARIANT* a = arg(0);
                if (!a || a->vt != VT_DISPATCH || !a->pdispVal) return DISP_E_TYPEMISMATCH;
                after = static_cast<Node*>(a->pdispVal)->cell_;
            }
            const auto cells = search(id_, finding);
            result->vt = VT_DISPATCH;  // Nothing, if none.
            result->pdispVal = nullptr;
            if (cells.empty()) return S_OK;
            const auto next = std::upper_bound(cells.begin(), cells.end(), after);
            return give(new Node(range, name_, id_, next == cells.end() ? cells.front() : *next));
        }
        case 7: {  // Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External)
            const VARIANT* external = arg(3);
            const bool whole = external && external->vt == VT_BOOL && external->boolVal;
            return text((whole ? sheets[id_ - 1] + L"!" : L"") + address(cell_, size_));
        }
        case 8:  // HasFormula
            result->vt = VT_BOOL;
            result->boolVal = formulas[id_].contains(cell_) ? VARIANT_TRUE : VARIANT_FALSE;
            return S_OK;
        case 9:     // Formula2
        case 12: {  // FormulaArray, of the whole of a legacy array formula
            const auto refuse = [&](const wchar_t* why) {
                if (info) {
                    *info = {};
                    info->scode = static_cast<SCODE>(0x800A03EC);
                    info->bstrDescription = SysAllocString(why);
                }
                return DISP_E_EXCEPTION;
            };
            const int lock = locked.contains(id_) ? locked[id_] : 0;
            const bool whole_array = array_at(id_, cell_) == std::pair{cell_, size_};
            if (id == 12 && !whole_array) return refuse(L"Not the whole of an array formula.");
            if (!(flags & DISPATCH_PROPERTYPUT)) {
                if (lock == 2) return refuse(L"The cell cannot be read.");
                const auto f = formulas[id_].find(cell_);
                const auto v = grid[id_].find(cell_);
                return text(f != formulas[id_].end() ? f->second : v != grid[id_].end() ? text_of(&v->second) : L"");
            }
            const VARIANT* value = arg(0);
            if (!value || value->vt != VT_BSTR) return DISP_E_TYPEMISMATCH;
            if (lock) return refuse(L"The cell is on a protected sheet.");
            if (id == 9 && array_at(id_, cell_).second.first) return refuse(L"You can't change part of an array.");
            const std::wstring formula = value->bstrVal;
            for (int r = 0; r < size_.first; ++r)
                for (int c = 0; c < size_.second; ++c) formulas[id_][{cell_.first + r, cell_.second + c}] = formula;
            entered += " s:" + hex(sheets[id_ - 1]) + " s:" + hex(address(cell_, size_)) + " s:" + hex(formula) + " " +
                       (id == 9 ? calculate(id_, cell_, formula) : "z");
            ++entered_count;
            if (!change_search.empty()) finding = change_search;  // A change handler's own search.
            if (!change_request.empty()) {  // A change handler running a command.
                if (++change_runs > 25) {
                    fault(L"the change handler ran 25 times: the repair does not end");
                    change_request.clear();
                } else
                    carry_out(change_request);
            }
            return S_OK;
        }
        case 10:  // HasArray
            result->vt = VT_BOOL;
            result->boolVal = array_at(id_, cell_).second.first ? VARIANT_TRUE : VARIANT_FALSE;
            return S_OK;
        case 11: {  // CurrentArray
            const auto [first, size] = array_at(id_, cell_);
            return size.first ? give(new Node(range, name_, id_, first, size)) : DISP_E_EXCEPTION;
        }
        default:
            return DISP_E_MEMBERNOTFOUND;
        }
    }

private:
    static constexpr const wchar_t* members[] = {L"Count",    L"Item",     L"Name",         L"Worksheets", L"Cells",
                                                 L"Find",     L"FindNext", L"Address",      L"HasFormula", L"Formula2",
                                                 L"HasArray", L"CurrentArray", L"FormulaArray"};
    ULONG count_ = 1;
    Kind kind_;
    std::wstring name_;  // The workbook's.
    IDSHEET id_;
    std::pair<int, int> cell_, size_;  // A range's first cell, and its rows and columns.
};

// Excel's application events (AppEvents), to which the add-in connects its handler. A
// static object: it ignores reference counts.
const IID app_events = {0x00024413, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};
std::vector<IDispatch*> handlers;  // Each connected, by its cookie less one; null once disconnected.

class Events final : public IConnectionPointContainer, public IConnectionPoint {
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        *out = iid == IID_IUnknown || iid == IID_IConnectionPointContainer ? static_cast<void*>(static_cast<IConnectionPointContainer*>(this))
             : iid == IID_IConnectionPoint ? static_cast<void*>(static_cast<IConnectionPoint*>(this)) : nullptr;
        return *out ? S_OK : E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return 2; }
    ULONG STDMETHODCALLTYPE Release() override { return 1; }
    HRESULT STDMETHODCALLTYPE EnumConnectionPoints(IEnumConnectionPoints**) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE FindConnectionPoint(REFIID iid, IConnectionPoint** out) override {
        *out = iid == app_events ? this : nullptr;
        return *out ? S_OK : CONNECT_E_NOCONNECTION;
    }
    HRESULT STDMETHODCALLTYPE GetConnectionInterface(IID*) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetConnectionPointContainer(IConnectionPointContainer**) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE Advise(IUnknown* handler, DWORD* cookie) override {
        IDispatch* d = nullptr;
        if (!handler || FAILED(handler->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&d)))) return CONNECT_E_CANNOTCONNECT;
        handlers.push_back(d);
        *cookie = static_cast<DWORD>(handlers.size());
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Unadvise(DWORD cookie) override {
        if (cookie < 1 || cookie > handlers.size() || !handlers[cookie - 1]) return CONNECT_E_NOCONNECTION;
        std::exchange(handlers[cookie - 1], nullptr)->Release();
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE EnumConnections(IEnumConnections**) override { return E_NOTIMPL; }
} events;

// "opened": Excel reports the workbook opened (WorkbookOpen), or made from a template
// (NewWorkbook), to each handler; answers whether any was connected.
std::string do_opened(const std::wstring& book, bool made) {
    bool any = false;
    for (IDispatch* handler : std::vector(handlers))
        if (handler) {
            any = true;
            VARIANT arg;
            arg.vt = VT_DISPATCH;
            arg.pdispVal = new Node(Node::book, book);
            DISPPARAMS params{&arg, nullptr, 1, 0};
            handler->Invoke(made ? 0x61D : 0x61F, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, nullptr, nullptr, nullptr);
            arg.pdispVal->Release();
        }
    return std::string("ok b:") + (any ? "1" : "0");
}

bool busy = false;              // The "busy" request: Application.Run and Evaluate refuse.
bool protected_view = false;    // The workbook window is a Protected View window.
bool protected_active = false;  // A Protected View window is active: Evaluate gives #N/A.
int runs = 0, refused = 0;      // Calls of Application.Run, and refusals.

// What Excel's Application object offers the add-in: Application (itself, as a window's
// object gives it), Run of a command by name, Evaluate of
// PyInfo("pyinex"), and Hwnd. A static object: it ignores reference counts.
class Application final : public IDispatch {
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        if (iid == IID_IConnectionPointContainer) return events.QueryInterface(iid, out);
        *out = iid == IID_IUnknown || iid == IID_IDispatch ? this : nullptr;
        return *out ? S_OK : E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return 2; }
    ULONG STDMETHODCALLTYPE Release() override { return 1; }
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* n) override {
        *n = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR* names_in, UINT count, LCID, DISPID* ids) override {
        static const wchar_t* const members[] = {L"Application", L"Run", L"Evaluate", L"Hwnd", L"Workbooks"};
        for (DISPID i = 0; i < 5 && count == 1; ++i)
            if (_wcsicmp(names_in[0], members[i]) == 0 && !(i == 0 && protected_view)) {
                ids[0] = i;
                return S_OK;
            }
        return DISP_E_UNKNOWNNAME;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS* params, VARIANT* result,
                                     EXCEPINFO* info, UINT*) override {
        if (!result) return E_POINTER;
        VariantInit(result);
        const bool text = params && params->cArgs == 1 && params->rgvarg[0].vt == VT_BSTR;
        const std::wstring arg = text ? params->rgvarg[0].bstrVal : L"";
        runs += id == 1;
        if (busy && (id == 1 || id == 2)) {  // Excel says the same while busy and of a macro it does not have.
            refused += id == 1;
            return refuse(arg, info);
        }
        switch (id) {
        case 0:
            result->vt = VT_DISPATCH;
            result->pdispVal = this;
            return S_OK;
        case 1: {
            if (!text) return DISP_E_BADPARAMCOUNT;
            const int done = run_command(arg);
            if (done < 0) return refuse(arg, info);
            result->vt = VT_R8;
            result->dblVal = done;
            return S_OK;
        }
        case 2: {
            // PyInfo("pyinex"), calculated by the PyInfo the name reaches, as Excel would: the
            // add-in's path; #VALUE! if it reaches none; #N/A with a Protected View window active.
            if (arg != L"PyInfo(\"pyinex\")") return DISP_E_TYPEMISMATCH;
            const auto known = names.find(L"PyInfo");
            const Registration* r = known != names.end() && known->second != 0 && registrations[known->second - 1].registered
                                        ? &registrations[known->second - 1]
                                        : nullptr;
            std::wstring path;
            if (r && !protected_active && r->address) {
                wchar_t item[] = {6, L'p', L'y', L'i', L'n', L'e', L'x'};
                XLOPER12 asked{.xltype = xltypeStr};
                asked.val.str = item;
                LPXLOPER12 out;
                {
                    const InContext calculating(Context::function);
                    out = reinterpret_cast<LPXLOPER12 (*)(LPXLOPER12)>(r->address)(&asked);
                }
                if (out && (out->xltype & ~(xlbitDLLFree | xlbitXLFree)) == xltypeStr)
                    path.assign(out->val.str + 1, out->val.str[0]);
                using Free = void (*)(LPXLOPER12);
                if (out && (out->xltype & xlbitDLLFree))
                    if (const auto free_fn = reinterpret_cast<Free>(GetProcAddress(xll, "xlAutoFree12"))) free_fn(out);
            }
            result->vt = static_cast<VARTYPE>(!path.empty() ? VT_BSTR : VT_ERROR);
            if (result->vt == VT_BSTR) result->bstrVal = SysAllocString(path.c_str());
            else result->scode = static_cast<SCODE>(protected_active ? 0x800A07FA : 0x800A07DF);  // As Excel gives them.
            return S_OK;
        }
        case 3:
            result->vt = VT_I4;
            result->lVal = static_cast<LONG>(reinterpret_cast<INT_PTR>(main_window));
            return S_OK;
        case 4:
            result->vt = VT_DISPATCH;
            result->pdispVal = new Node(Node::books, L"");
            return S_OK;
        default:
            return DISP_E_MEMBERNOTFOUND;
        }
    }

private:
    // Excel's answer to a call it refuses.
    static HRESULT refuse(const std::wstring& macro, EXCEPINFO* info) {
        if (info) {
            *info = {};
            info->scode = static_cast<SCODE>(0x800A03EC);
            info->bstrDescription = SysAllocString((L"Cannot run the macro '" + macro + L"'.").c_str());
        }
        return DISP_E_EXCEPTION;
    }
} application;

// A workbook's windows, as Excel makes them: a main window (XLMAIN), its desk (XLDESK), and
// the workbook window (EXCEL7), which gives its object to AccessibleObjectFromWindow.
HWND book_window = nullptr;

LRESULT CALLBACK book_procedure(HWND window, UINT message, WPARAM w, LPARAM l) {
    if (message == WM_GETOBJECT && static_cast<DWORD>(l) == static_cast<DWORD>(OBJID_NATIVEOM))
        return LresultFromObject(IID_IDispatch, w, &application);
    return DefWindowProcW(window, message, w, l);
}

void set_workbook(const std::string& state) {
    if (book_window) DestroyWindow(book_window);
    book_window = nullptr;
    protected_view = state == "protected";
    protected_active = state == "both";
    if (state == "none") return;
    excel_window();
    if (!workbook_window) workbook_window = make_window(L"XLMAIN", L"Book1 - Excel", nullptr, WS_OVERLAPPEDWINDOW);
    WNDCLASSW book{.lpfnWndProc = book_procedure, .hInstance = GetModuleHandleW(nullptr), .lpszClassName = L"EXCEL7"};
    RegisterClassW(&book);  // Fails harmlessly once registered.
    HWND desk = FindWindowExW(workbook_window, nullptr, L"XLDESK", nullptr);
    if (!desk) desk = make_window(L"XLDESK", L"", workbook_window, WS_CHILD);
    book_window = CreateWindowExW(0, L"EXCEL7", L"Book1", WS_CHILD, 0, 0, 100, 100, desk, nullptr,
                                  GetModuleHandleW(nullptr), nullptr);
}

// The "fail window" request: a hook that stops the add-in's queue window being created.
int failing_windows = 0;

LRESULT CALLBACK refuse_window(int code, WPARAM w, LPARAM l) {
    if (code == HCBT_CREATEWND && failing_windows > 0) {
        const wchar_t* name = reinterpret_cast<CBT_CREATEWNDW*>(l)->lpcs->lpszClass;
        if (!IS_INTRESOURCE(name) && std::wcscmp(name, L"PyinexQueueWindow") == 0) {
            --failing_windows;
            return 1;
        }
    }
    return CallNextHookEx(nullptr, code, w, l);
}

// The "fail timer" and "fail memory" requests: the add-in's own calls of SetTimer and malloc
// go through these, put in its import table as it is loaded. Only this thread's allocations
// fail: the console window's thread allocates too.
int failing_timers = 0;
decltype(&SetTimer) real_set_timer = nullptr;
decltype(&malloc) real_malloc = nullptr;
const DWORD host_thread = GetCurrentThreadId();

UINT_PTR WINAPI failing_set_timer(HWND window, UINT_PTR id, UINT delay, TIMERPROC f) {
    if (failing_timers > 0) {
        --failing_timers;
        SetLastError(ERROR_NOT_ENOUGH_MEMORY);
        return 0;
    }
    return real_set_timer(window, id, delay, f);
}

void* __cdecl failing_malloc(size_t size) {
    if (failing_allocation && GetCurrentThreadId() == host_thread && skipped_allocations-- <= 0) {
        failing_allocation = false;
        return nullptr;
    }
    return real_malloc(size);
}

// Points the module's import of function, from a DLL whose name starts with dll, at hook;
// returns what it pointed at, or null if the module imports no such function.
void* hook_import(HMODULE module, const char* dll, const char* function, void* hook) {
    const auto base = reinterpret_cast<BYTE*>(module);
    const auto nt = reinterpret_cast<IMAGE_NT_HEADERS*>(base + reinterpret_cast<IMAGE_DOS_HEADER*>(base)->e_lfanew);
    const IMAGE_DATA_DIRECTORY& imports = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
    for (auto d = reinterpret_cast<IMAGE_IMPORT_DESCRIPTOR*>(base + imports.VirtualAddress); imports.Size && d->Name; ++d) {
        if (_strnicmp(reinterpret_cast<char*>(base + d->Name), dll, std::strlen(dll)) != 0) continue;
        auto name = reinterpret_cast<IMAGE_THUNK_DATA*>(base + d->OriginalFirstThunk);
        auto slot = reinterpret_cast<IMAGE_THUNK_DATA*>(base + d->FirstThunk);
        for (; name->u1.AddressOfData; ++name, ++slot) {
            if (IMAGE_SNAP_BY_ORDINAL(name->u1.Ordinal) ||
                std::strcmp(reinterpret_cast<IMAGE_IMPORT_BY_NAME*>(base + name->u1.AddressOfData)->Name, function) != 0)
                continue;
            DWORD protection = 0;
            VirtualProtect(&slot->u1.Function, sizeof slot->u1.Function, PAGE_READWRITE, &protection);
            void* const was = reinterpret_cast<void*>(slot->u1.Function);
            slot->u1.Function = reinterpret_cast<ULONG_PTR>(hook);
            VirtualProtect(&slot->u1.Function, sizeof slot->u1.Function, protection, &protection);
            return was;
        }
    }
    return nullptr;
}

std::string do_tick(int count) {
    // This process's own: another host's, idle between requests, would never answer.
    HWND queue = nullptr;
    DWORD process = 0;
    do queue = FindWindowExW(HWND_MESSAGE, queue, L"PyinexQueueWindow", nullptr);
    while (queue && (GetWindowThreadProcessId(queue, &process), process != GetCurrentProcessId()));
    const int runs_before = runs, refused_before = refused;
    for (int i = 0; i < count && IsWindow(queue); ++i) SendMessageW(queue, WM_TIMER, 1, 0);
    return std::string("ok b:") + (queue ? "1" : "0") + " n:" + std::to_string(runs - runs_before) +
           " n:" + std::to_string(refused - refused_before);
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
        if (!r.registered) continue;
        out += " a:1x" + std::to_string(6 + r.argument_help.size());
        for (const auto* s : {&r.name, &r.procedure, &r.type, &r.arguments, &r.category, &r.help})
            out += " s:" + hex(*s);
        for (const auto& h : r.argument_help) out += " s:" + hex(h);
    }
    return out;
}

// A request carried out inside the add-in's call into Excel ("nest"): only "run" and
// "call", which neither handle messages nor check for leaks, as the add-in still holds what
// Excel gave it.
int call_entry(const char* name);

std::string carry_out(const std::string& requests) {
    if (const size_t at = requests.find(';'); at != std::string::npos) {
        carry_out(requests.substr(0, at));
        return carry_out(requests.substr(at + 1));
    }
    std::istringstream in(requests);
    std::string verb, name;
    in >> verb;
    if (verb == "call") return do_call(in);
    if (verb == "close") {  // Excel closing the add-in, as by unticking it in the Add-ins dialog.
        const InContext closing(Context::command);
        call_entry("xlAutoClose");
        return "ok";
    }
    if (verb == "memory") {
        skipped_allocations = 0;
        in >> skipped_allocations;
        failing_allocation = true;
        return "ok";
    }
    if (verb != "run" || !(in >> name)) fault(L"nest carries out only run, call, close and memory");
    else if (run_command(unhex(name)) < 0) fault(L"nest: no registered command " + unhex(name));
    return "ok";
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
        // Opened again, the add-in may still be loaded with the hooks in place.
        if (void* was = hook_import(xll, "user32", "SetTimer", reinterpret_cast<void*>(failing_set_timer));
            was && was != reinterpret_cast<void*>(failing_set_timer))
            real_set_timer = reinterpret_cast<decltype(&SetTimer)>(was);
        if (void* was = hook_import(xll, "api-ms-win-crt-heap", "malloc", reinterpret_cast<void*>(failing_malloc));
            was && was != reinterpret_cast<void*>(failing_malloc))
            real_malloc = reinterpret_cast<decltype(&malloc)>(was);
        const InContext opening(Context::command);
        response = "ok n:" + std::to_string(call_entry("xlAutoOpen"));
    } else if (verb == "load") {
        std::string path;
        in >> path;
        if (!LoadLibraryExW(unhex(path).c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH))
            return "error " + hex(L"LoadLibrary failed with error " + std::to_wstring(GetLastError()));
        return "ok";
    } else if (verb == "close") {
        const InContext closing(Context::command);
        response = "ok n:" + std::to_string(call_entry("xlAutoClose"));
    } else if (verb == "run") {
        std::string name;
        in >> name;
        const int done = run_command(unhex(name));
        if (done < 0) return "error " + hex(L"no registered command " + unhex(name));
        response = "ok n:" + std::to_string(done);
    } else if (verb == "tick") {
        int count = 0;
        in >> count;
        response = do_tick(count);
    } else if (verb == "workbook") {
        std::string state;
        in >> state;
        set_workbook(state);
        return "ok";
    } else if (verb == "busy") {
        std::string on;
        in >> on;
        busy = on == "1";
        return "ok";
    } else if (verb == "fail") {
        std::string what, arg;
        in >> what >> arg;
        const int count = std::atoi(arg.c_str());
        static const HHOOK hook = SetWindowsHookExW(WH_CBT, refuse_window, nullptr, GetCurrentThreadId());
        if (what == "register") failing_registrations = count;
        else if (what == "command") failing_commands = count;
        else if (what == "name") failing_name = unhex(arg);
        else if (what == "unregister") failing_unregisters = count, in >> skipped_unregisters;
        else if (what == "setname") failing_set_names = count;
        else if (what == "window" && hook) failing_windows = count;
        else if (what == "timer" && real_set_timer) failing_timers = count;
        else if (what == "evaluate") failing_evaluates = count;
        else if (what == "memory" && real_malloc && (arg == "evaluate" || arg == "register" || arg == "unregister")) {
            memory_armed = arg == "evaluate" ? Arm::evaluate : arg == "register" ? Arm::registration : Arm::unregister;
            skipped_allocations = 0;
            in >> skipped_allocations;
        } else if (what == "memory" && real_malloc && arg == "now") {
            // The core's allocations too, once Python has started and it is loaded.
            static bool hooked = false;
            if (const HMODULE core = GetModuleHandleW(L"pyinex_core.dll"); core && !std::exchange(hooked, true))
                hook_import(core, "api-ms-win-crt-heap", "malloc", reinterpret_cast<void*>(failing_malloc));
            skipped_allocations = 0;
            in >> skipped_allocations;
            failing_allocation = true;
        }
        else return "error " + hex(L"no such failure");
        return "ok";
    } else if (verb == "foreign") {
        std::string name;
        in >> name;
        registrations.push_back({.module = L"C:\\other\\other.xll", .procedure = unhex(name), .name = unhex(name)});
        names[unhex(name)] = registrations.size();
        return "ok";
    } else if (verb == "cells") {
        std::string sheet;
        int row = 0, column = 0;
        in >> sheet >> row >> column;
        XLOPER12 value = parse_value(in);
        const bool array = value.xltype == xltypeMulti;
        const int rows = array ? value.val.array.rows : 1, columns = array ? value.val.array.columns : 1;
        auto& cells = grid[sheet_id(unhex(sheet))];
        for (int r = 0; r < rows; ++r)
            for (int c = 0; c < columns; ++c) {
                const XLOPER12& v = array ? value.val.array.lparray[r * columns + c] : value;
                const std::pair at{row - 1 + r, column - 1 + c};
                if (const auto old = cells.find(at); old != cells.end()) {
                    free_tree(old->second);
                    cells.erase(old);
                }
                if (v.xltype != xltypeNil) cells[at] = coerce(v, 0);
            }
        free_tree(value);
        return "ok";
    } else if (verb == "stop") {
        std::string code;
        in >> code;
        stop_skip = 0;
        in >> stop_skip;
        stop_with = code == "uncalced" ? xlretUncalced : code == "abort" ? xlretAbort : 0;
        return "ok";
    } else if (verb == "forget") {
        std::string sheet;
        in >> sheet;
        if (const IDSHEET id = known_sheet(unhex(sheet))) {
            for (auto& [at, value] : grid[id]) free_tree(value);
            grid.erase(id);
            sheets[id - 1].clear();
        }
        return "ok";
    } else if (verb == "nest" || verb == "nestat") {
        std::string at;
        if (verb == "nestat") in >> at;
        nested_at = verb == "nestat" ? unhex(at) : L"PyFunc";
        std::getline(in >> std::ws, nested_request);
        return "ok";
    } else if (verb == "unload") {
        if (!xll) return "error " + hex(L"no add-in is loaded");
        FreeLibrary(xll);
        xll = nullptr;
        for (auto& r : registrations)  // Excel unregisters what the add-in registered.
            if (_wcsicmp(r.module.c_str(), xll_path.c_str()) == 0) r.registered = false;
        std::erase_if(names, [](const auto& n) {  // Only this add-in's names; another copy keeps its own.
            return n.second == 0 || !registrations[n.second - 1].registered;
        });
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
    } else if (verb == "formula") {
        std::string sheet, text;
        int row = 0, column = 0;
        in >> sheet >> row >> column >> text;
        formulas[sheet_id(unhex(sheet))][{row - 1, column - 1}] = unhex(text);
        return "ok";
    } else if (verb == "array") {
        std::string sheet, text;
        int row = 0, column = 0, rows = 0, columns = 0;
        in >> sheet >> row >> column >> rows >> columns >> text;
        const IDSHEET id = sheet_id(unhex(sheet));
        arrays[id][{row - 1, column - 1}] = {rows, columns};
        for (int r = 0; r < rows; ++r)
            for (int c = 0; c < columns; ++c) formulas[id][{row - 1 + r, column - 1 + c}] = unhex(text);
        return "ok";
    } else if (verb == "lock") {
        std::string sheet;
        int level = 0;
        in >> sheet >> level;
        locked[sheet_id(unhex(sheet))] = level;
        return "ok";
    } else if (verb == "change") {
        std::getline(in >> std::ws, change_request);
        change_runs = 0;
        return "ok";
    } else if (verb == "handler") {
        std::string text;
        in >> text;
        change_search = unhex(text);
        return "ok";
    } else if (verb == "opened") {
        std::string book, made;
        in >> book >> made;
        response = do_opened(unhex(book), made == "new");
    } else if (verb == "entered") {
        response = "ok n:" + std::to_string(calculations) + " a:" + std::to_string(entered_count) + "x4" + entered;
        calculations = entered_count = 0;
        entered.clear();
        return response;
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
    // As on Excel's main thread, COM is ready, in a single-threaded apartment, and a workbook
    // is open.
    if (FAILED(OleInitialize(nullptr))) {
        std::fputs("xllhost: cannot start COM\n", stderr);
        return 2;
    }
    set_workbook("open");
    // Once, before the registry stand-in hides the machine's: the first object a window gives
    // reads how COM passes it (IDispatch's registration in HKEY_LOCAL_MACHINE), and fails
    // without it (0x80070002). Excel's own registry has it.
    IDispatch* warm = nullptr;
    AccessibleObjectFromWindow(book_window, static_cast<DWORD>(OBJID_NATIVEOM), IID_IDispatch, reinterpret_cast<void**>(&warm));
    if (warm) warm->Release();
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
