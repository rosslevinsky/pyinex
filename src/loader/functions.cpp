#include "loader/functions.h"

#include "console/console.h"
#include "excel/excel.h"
#include "loader/console_window.h"

#include <ocidl.h>
#include <oleacc.h>

#include <algorithm>
#include <climits>
#include <cmath>
#include <cstdarg>
#include <cwchar>
#include <deque>
#include <memory>
#include <set>
#include <string>
#include <type_traits>
#include <utility>
#include <vector>

using namespace pyinex::excel;

extern "C" IMAGE_DOS_HEADER __ImageBase;

#define PYINEX_EXPORT extern "C" __declspec(dllexport)

namespace pyinex::loader {

bool reload_on_change = true;

namespace {

constexpr int max_arguments = 30;
constexpr UINT_PTR queue_timer = 1;  // tests/host/xllhost sends this ID to the window's class.
constexpr UINT retry_ms = 250;
constexpr wchar_t window_class[] = L"PyinexQueueWindow";

// A function to register, from one of pyinex_load's rows, kept so it can be registered
// again; or, with remove, the function key to remove. key -1: none; kept: see below.
struct Description {
    int key = -1;
    int count = 0;  // Its arguments, all Excel is told it takes.
    bool remove = false;
    std::wstring types;  // Excel's type code for each argument: Q for values, U for a reference.
    std::wstring name, arguments, category, help, script;
    std::vector<std::wstring> argument_help;
};

// What each numbered export calls: the function in it, and Excel's registration ID.
struct Slot {
    Description f;
    double id = 0;
    bool doomed = false;  // To be removed by the queue's next run: see unoffer.
};
static_assert(std::is_nothrow_move_assignable_v<Slot>);  // See register_one.

// The key of a slot whose name could not be cleared: Excel may still send that name to it,
// so it gives #VALUE! and takes no other function until Excel restarts.
constexpr int kept = -2;
// The key of a slot while Excel is told to remove its function: a close meanwhile, from
// within Excel's call, neither removes it again nor queues it to be registered again.
constexpr int removing = -3;

std::vector<Slot> slots(PYINEX_POOL);

std::deque<Description> queued;  // Waiting for the queue's command, in order.
// Each name this copy removed a function from, with what Excel then made of it (known_name).
std::vector<std::pair<std::wstring, double>> removed_names;
std::vector<int> dropped;        // Keys dropped when Excel unloaded the add-in, for Python.
HWND window = nullptr;           // The queue's hidden window, made when work is first queued.
std::wstring command;            // The queue's command's name, this copy of Pyinex's own.
// The queue's command is running, or the add-in is opening or closing: a run started
// meanwhile, as by PyRepair within xlfRegister, does nothing, or leaves its work for later.
bool draining = false;
bool repair_wanted = false;      // PyRepair asked for, to run when the queue's run reaches it.
bool repairing = false;          // This run of the queue's has done so: asked for again, it does nothing.
bool armed = false;              // The queue's timer is set.
bool closed = false;             // From xlAutoClose until the queue's command runs.
std::wstring reported;           // Why Excel last could not run it, as reported.
// Workbooks whose PyLoad cells wait to be calculated, each counted; null: every one open.
std::vector<IDispatch*> opened;
std::vector<std::wstring> fresh;  // Names registered that Excel had not had from this copy.
bool following = false;           // Connected to Excel's events, or tried, since the add-in opened.

int type_of(const XLOPER12& x) { return x.xltype & ~(xlbitDLLFree | xlbitXLFree); }

std::wstring_view text(const XLOPER12& x) {
    if (type_of(x) != xltypeStr) return {};
    return {x.val.str + 1, static_cast<size_t>(x.val.str[0])};
}

std::wstring functions(size_t n) { return std::to_wstring(n) + (n == 1 ? L" function" : L" functions"); }

// A whole number from min to max in a cell of pyinex_load's rows, or -1.
int whole(const XLOPER12& x, int min, int max) {
    if (type_of(x) != xltypeNum || x.val.num != std::floor(x.val.num) || x.val.num < min || x.val.num > max) return -1;
    return static_cast<int>(x.val.num);
}

// Reads one of pyinex_load's rows: key, name, argument count (or each argument's type code,
// when one takes a reference), argument names, category, help, then each argument's help; or
// a key alone (the rest blank), to remove. False for one already registered (the key -1
// alone), or, after saying so, for one that is malformed.
bool describe(const XLOPER12* row, int columns, std::wstring_view script, Description& f) {
    if (type_of(row[0]) == xltypeNum && row[0].val.num == -1) return false;
    f.script = script;
    f.key = whole(row[0], 0, INT_MAX);
    if ((f.remove = f.key >= 0 && (columns == 1 || text(row[1]).empty()))) return true;
    f.key = columns >= 6 ? f.key : -1;
    f.types = text(row[2]);
    const bool codes = f.types.size() <= max_arguments && f.types.find_first_not_of(L"QU") == f.types.npos;
    f.count = f.key < 0 || !codes ? -1 : f.types.empty() ? whole(row[2], 0, max_arguments) : static_cast<int>(f.types.size());
    if (f.types.empty()) f.types.assign(std::max(f.count, 0), L'Q');
    // Checked before Excel is told: a slot must never read past what Excel was given.
    if (f.count < 0 || columns < 6 + f.count) {
        console::note(L"cannot register a function of the script " + std::wstring(script) +
                      L": its description is malformed");
        return false;
    }
    f.name = text(row[1]);
    f.arguments = text(row[3]);
    f.category = text(row[4]);
    f.help = text(row[5]);
    for (int i = 0; i < f.count; ++i) f.argument_help.emplace_back(text(row[6 + i]));
    return true;
}

// What running out of memory kept from the queue, of rows, pyinex_load's or with named
// pyinex_changes', from row from on: Python is told that each new function will not be
// registered, so its name is free and the next load offers it again; and each removal,
// which Python has already made, is marked on its slot for the queue's next run to make,
// with nothing to allocate.
void unoffer(const XLOPER12& rows, int from, bool named, const Core& core) noexcept {
    bool doomed = false;
    for (int r = from; type_of(rows) == xltypeMulti && r < rows.val.array.rows; ++r) {
        const XLOPER12* row = rows.val.array.lparray + r * rows.val.array.columns + named;
        const int columns = rows.val.array.columns - named;
        const int key = columns >= 1 ? whole(row[0], 0, INT_MAX) : -1;
        if (key < 0) continue;
        if (columns > 1 && !text(row[1]).empty()) {
            core.registered(key, false);
            continue;
        }
        for (auto& s : slots)
            if (s.f.key == key) doomed = s.doomed = true;
    }
    if (doomed && window && !armed) armed = SetTimer(window, queue_timer, retry_ms, nullptr) != 0;
}

// Adds the description of each row of pyinex_load's result, or with named of
// pyinex_changes', to out; the result is freed, whatever fails. Returns how many of the
// script's functions it gives, registered or not, or -1 for #VALUE!: the script did not load.
template <class Out>
int describe_all(Returned result, std::wstring_view script, Out& out, const Core& core, bool named = false) {
    const XLOPER12& rows = *result.get();
    int n = type_of(rows) == xltypeErr ? -1 : 0;
    int r = 0;
    try {
        for (; type_of(rows) == xltypeMulti && r < rows.val.array.rows; ++r) {
            const XLOPER12* row = rows.val.array.lparray + r * rows.val.array.columns;
            const int columns = rows.val.array.columns;
            Description f;
            const bool described = named ? columns > 1 && describe(row + 1, columns - 1, text(row[0]), f)
                                         : describe(row, columns, script, f);
            n += !f.remove;
            if (described) out.push_back(std::move(f));
        }
    } catch (...) {
        unoffer(rows, r, named, core);
        throw;
    }
    return n;
}

bool same_name(std::wstring_view a, std::wstring_view b) {
    return CompareStringOrdinal(a.data(), static_cast<int>(a.size()), b.data(), static_cast<int>(b.size()), TRUE) ==
           CSTR_EQUAL;
}

// Registers a function in the first free slot, in a command. If it cannot, says why and
// frees the function's name in Python. Once Excel has registered it, nothing can throw.
bool register_one(const Description& f) {
    int slot = 0;
    while (slot < PYINEX_POOL && slots[slot].f.key != -1) ++slot;
    double id = 0;
    // A name a slot holds now is not asked about. Nor is one Excel still makes the same of as
    // just after Pyinex removed it, or failed to (it still gives the name an ID), so long as
    // that answer was an ID and not a failed question (docs/architecture.md, "Names"). Each
    // failed removal kept a slot, with what Excel made of the name then: any may match.
    const auto named = [&](const Slot& s) { return s.f.key != -1 && same_name(s.f.name, f.name); };
    const bool held = std::ranges::any_of(slots, [&](const Slot& s) { return named(s) && s.f.key != kept; });
    const double known = held || slot == PYINEX_POOL ? 0 : known_name(f.name);
    const bool left = known != -1 && (std::ranges::any_of(removed_names, [&](const auto& r) { return same_name(r.first, f.name) && r.second == known; }) ||
                                     std::ranges::any_of(slots, [&](const Slot& s) { return named(s) && s.f.key == kept && s.id == known; }));
    if (slot == PYINEX_POOL) {
        console::note(L"cannot register " + f.name + L" from the script " + f.script + L": all " +
                      std::to_wstring(PYINEX_POOL) + L" of Pyinex's function slots are taken");
    } else if (std::isnan(known)) {  // unanswered, which no stored answer matches either.
        console::note(L"cannot register " + f.name + L" from the script " + f.script +
                      L": Excel could not say whether it already knows that name");
    } else if (known != 0 && !left) {
        console::note(L"cannot register " + f.name + L" from the script " + f.script +
                      L": Excel already knows that name, as one of its functions or another add-in's");
    } else {
        Slot filled{f};  // Copied first, so that keeping it once Excel has registered it is a move.
        const std::wstring procedure = L"PyFunc" + std::to_wstring(slot), type = L"Q" + f.types;
        Function r{.name = f.name, .procedure = procedure, .type = type, .arguments = f.arguments,
                   .help = f.help, .category = f.category};
        for (const auto& h : f.argument_help) r.argument_help.push_back(h);
        // A name removed before, kept or not, still reads as a function's in Excel: only one
        // never registered was unknown to formulas typed meanwhile (docs/architecture.md,
        // "The #NAME? repair"). It is added before Excel is told, and taken out if Excel does
        // not register the function, so nothing that can throw follows a registration. It is
        // taken out by name, the last one: a close meanwhile, within xlfRegister, empties the list.
        const bool repair = !left && std::ranges::none_of(removed_names, [&](const auto& r) { return same_name(r.first, f.name); });
        // A copy for a close within xlfRegister, which empties the list (unfollow) and keeps
        // its room: the repair is asked for again with nothing to allocate.
        std::wstring again = repair ? f.name : std::wstring();
        const auto unrepair = [&] {
            const auto last = std::find_if(fresh.rbegin(), fresh.rend(), [&](const auto& n) { return same_name(n, f.name); });
            if (repair && last != fresh.rend()) fresh.erase(std::next(last).base());
        };
        if (repair) fresh.push_back(f.name);
        try {
            filled.id = id = register_function(r);
        } catch (...) {
            unrepair();
            throw;
        }
        if (id != 0) {
            slots[slot] = std::move(filled);
            if (repair && fresh.size() < fresh.capacity() &&
                std::ranges::none_of(fresh, [&](const auto& n) { return same_name(n, again); }))
                fresh.push_back(std::move(again));
        } else {
            unrepair();
            console::note(L"Excel refused to register " + f.name + L" from the script " + f.script);
        }
    }
    if (const Core* core = running(); core && id == 0) core->registered(f.key, false);
    return id != 0;
}

// Unregisters a slot's function and clears its name, in a command. If any step fails, says
// so and keeps the slot out of use: Excel may still send the name to it. What can fail, as
// memory running out, is done before Excel is told: once it is, the slot is kept, or freed
// with what Excel then makes of the name recorded, and only the console's note follows.
void remove(Slot& s) {
    const Value name = Value::text(s.f.name);
    removed_names.reserve(removed_names.size() + 1);
    // Running out of memory before Excel is told leaves the slot as it was, for the next try.
    const int key = std::exchange(s.f.key, removing);
    bool cleared = false;
    try {
        cleared = unregister_function(s.id, s.f.name);
    } catch (...) {
        s.f.key = key;
        throw;
    }
    s.doomed = false;
    const double now = known_name(name);
    if (!cleared) {
        // Pyinex may still hold the name: what Excel makes of it now is kept as the slot's id.
        s.id = now;
        s.f.key = kept;
        console::note(L"cannot remove " + s.f.name + L" from Excel; until Excel restarts it gives #VALUE!");
        return;
    }
    const auto r = std::ranges::find_if(removed_names, [&](const auto& r) { return same_name(r.first, s.f.name); });
    if (r == removed_names.end()) removed_names.emplace_back(std::move(s.f.name), now);
    else r->second = now;
    s = Slot{};
}

// Registers a function, or removes one, in a command; returns whether it registered.
bool apply(const Description& f) {
    if (!f.remove) return register_one(f);
    for (auto& s : slots)  // A function never registered, or refused, has no slot.
        if (s.f.key == f.key) remove(s);
    // No cell can call it now, so Python forgets why it gave #VALUE!.
    if (const Core* core = running()) core->registered(f.key, false);
    return false;
}

// Unregisters every direct function and clears its name; returns what was registered.
std::vector<Description> remove_all() {
    std::vector<Description> removed;
    for (auto& s : slots)
        if (s.f.key >= 0) {
            removed.push_back(s.f);
            remove(s);
        }
    return removed;
}

// ---------------------------------------------------------------------------------------
// The queue (docs/architecture.md, "The queue").

std::wstring hex(HRESULT hr) {
    wchar_t b[16];
    swprintf(b, 16, L"0x%08lX", static_cast<unsigned long>(hr));
    return b;
}

// Calls a member of a COM object with the arguments given, in order, which it clears. result
// is initialized before anything can fail, and the caller clears it. False, with why set, if
// it failed.
bool invoke(IDispatch* object, const wchar_t* member, WORD flags, std::vector<VARIANT> args, VARIANT& result,
            std::wstring& why) {
    VariantInit(&result);
    std::ranges::reverse(args);  // COM takes them last first.
    DISPID id = 0, put = DISPID_PROPERTYPUT;
    LPOLESTR name = const_cast<LPOLESTR>(member);
    EXCEPINFO info{};
    const UINT named = flags == DISPATCH_PROPERTYPUT ? 1 : 0;  // A property's new value.
    DISPPARAMS params{args.data(), named ? &put : nullptr, static_cast<UINT>(args.size()), named};
    HRESULT hr = object->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &id);
    if (SUCCEEDED(hr)) hr = object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &params, &result, &info, nullptr);
    for (auto& a : args) VariantClear(&a);
    if (FAILED(hr))
        why = std::wstring(member) + L" failed, " + hex(hr) +
              (info.bstrDescription ? L": " + std::wstring(info.bstrDescription) : L"");
    SysFreeString(info.bstrSource);
    SysFreeString(info.bstrDescription);
    SysFreeString(info.bstrHelpFile);
    return SUCCEEDED(hr);
}

// COM's arguments: text, a number, true or false, an object, one left out.
VARIANT variant(VARTYPE type) {
    VARIANT v;
    VariantInit(&v);
    v.vt = type;
    return v;
}
VARIANT bstr(std::wstring_view s) {
    VARIANT v = variant(VT_BSTR);
    v.bstrVal = SysAllocStringLen(s.data(), static_cast<UINT>(s.size()));
    return v;
}
VARIANT number(long n) {
    VARIANT v = variant(VT_I4);
    v.lVal = n;
    return v;
}
VARIANT flag(bool b) {
    VARIANT v = variant(VT_BOOL);
    v.boolVal = b ? VARIANT_TRUE : VARIANT_FALSE;
    return v;
}
VARIANT object(IDispatch* p) {
    VARIANT v = variant(VT_DISPATCH);
    (v.pdispVal = p)->AddRef();
    return v;
}
VARIANT omitted() {
    VARIANT v = variant(VT_ERROR);
    v.scode = DISP_E_PARAMNOTFOUND;
    return v;
}

struct Release {
    void operator()(IUnknown* p) const { p->Release(); }
};
using Object = std::unique_ptr<IDispatch, Release>;

// A member that is an object; null if it fails or is Nothing.
Object member(IDispatch* from, const wchar_t* name, std::vector<VARIANT> args = {}) {
    VARIANT v;
    std::wstring why;
    if (invoke(from, name, DISPATCH_METHOD | DISPATCH_PROPERTYGET, std::move(args), v, why) && v.vt == VT_DISPATCH)
        return Object(v.pdispVal);
    VariantClear(&v);
    return nullptr;
}

// A member's value as type; VT_EMPTY if it fails. The caller clears it.
VARIANT value(IDispatch* from, const wchar_t* name, VARTYPE type, std::vector<VARIANT> args = {}) {
    VARIANT v;
    std::wstring why;
    if (!invoke(from, name, DISPATCH_PROPERTYGET, std::move(args), v, why) || FAILED(VariantChangeType(&v, &v, 0, type)))
        VariantClear(&v);
    return v;
}

// Excel's Application object: from a workbook window of this Excel, else, if anywhere, from
// the running object table if the Excel there is this one. Null, with why set, if it is not
// found.
IDispatch* application(std::wstring& why, bool anywhere) {
    why = L"no workbook window, and no Excel of this process in the running object table";
    std::vector<HWND> mains;
    EnumThreadWindows(
        GetCurrentThreadId(),
        [](HWND h, LPARAM found) -> BOOL {
            wchar_t name[8];
            if (GetClassNameW(h, name, 8) && std::wcscmp(name, L"XLMAIN") == 0)
                reinterpret_cast<std::vector<HWND>*>(found)->push_back(h);
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&mains));
    for (HWND main : mains)
        for (HWND desk = FindWindowExW(main, nullptr, L"XLDESK", nullptr); desk; desk = FindWindowExW(main, desk, L"XLDESK", nullptr))
            for (HWND book = FindWindowExW(desk, nullptr, L"EXCEL7", nullptr); book; book = FindWindowExW(desk, book, L"EXCEL7", nullptr)) {
                IDispatch* object = nullptr;
                const HRESULT hr = AccessibleObjectFromWindow(book, static_cast<DWORD>(OBJID_NATIVEOM), IID_IDispatch,
                                                              reinterpret_cast<void**>(&object));
                if (FAILED(hr) || !object) {
                    why = L"a workbook window gave no object, " + hex(hr);
                    continue;
                }
                VARIANT app;  // A Protected View window's object has no Application.
                const bool ok = invoke(object, L"Application", DISPATCH_PROPERTYGET, {}, app, why) &&
                                app.vt == VT_DISPATCH && app.pdispVal;
                object->Release();
                if (ok) return app.pdispVal;
                VariantClear(&app);
            }
    CLSID excel;
    IUnknown* running = nullptr;
    IDispatch* app = nullptr;
    if (anywhere && SUCCEEDED(CLSIDFromProgID(L"Excel.Application", &excel)) &&
        SUCCEEDED(GetActiveObject(excel, nullptr, &running))) {
        running->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&app));
        running->Release();
    }
    VARIANT hwnd;
    DWORD process = 0;
    if (app && invoke(app, L"Hwnd", DISPATCH_PROPERTYGET, {}, hwnd, why) && hwnd.vt == VT_I4)
        GetWindowThreadProcessId(reinterpret_cast<HWND>(static_cast<LONG_PTR>(hwnd.lVal)), &process);
    if (app) VariantClear(&hwnd);
    if (process == GetCurrentProcessId()) return app;
    if (app) app->Release();
    return nullptr;
}

// Whether Excel has unloaded the add-in, which unregisters its functions: PyInfo("pyinex")
// then gives #VALUE! or #NAME?. Any other answer says nothing: a path, even another copy's
// (this copy may still be loaded beside it), #N/A while a Protected View window is active,
// or a question Excel refuses.
bool unloaded(IDispatch* app) {
    VARIANT result;
    std::wstring why;
    bool gone = invoke(app, L"Evaluate", DISPATCH_METHOD, {bstr(L"PyInfo(\"pyinex\")")}, result, why);
    gone = gone && result.vt == VT_ERROR && ((result.scode & 0xFFFF) == 2015 || (result.scode & 0xFFFF) == 2029);
    VariantClear(&result);
    return gone;
}

void attempt();

// A note that may come before Python has started, when in Excel there is nowhere to say it
// until the console is made.
void say(const std::wstring& message) {
    prepare_output();
    console::note(message);
}

LRESULT CALLBACK queue_window(HWND h, UINT message, WPARAM w, LPARAM l) {
    if (message != WM_TIMER || w != queue_timer) return DefWindowProcW(h, message, w, l);
    KillTimer(h, queue_timer);
    armed = false;
    try {
        attempt();
    } catch (...) {
        if (window) SetTimer(window, queue_timer, retry_ms, nullptr);
    }
    return 0;
}

// Makes the window if need be, and sets its timer.
void arm(UINT delay) {
    const auto self = reinterpret_cast<HINSTANCE>(&__ImageBase);
    if (!window) {
        const WNDCLASSW c{.lpfnWndProc = queue_window, .hInstance = self, .lpszClassName = window_class};
        RegisterClassW(&c);
        window = CreateWindowExW(0, window_class, L"", 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, self, nullptr);
    }
    armed = window && SetTimer(window, queue_timer, delay, nullptr);
    if (!armed)
        say(L"cannot start the queue's timer (Windows error " + std::to_wstring(GetLastError()) +
            L"); what is queued waits until a PyLoad cell calculates again");
}

// Ends the window, and its timer with it.
void stop() {
    if (window) DestroyWindow(window);
    window = nullptr;
    armed = false;
    UnregisterClassW(window_class, reinterpret_cast<HINSTANCE>(&__ImageBase));
}

// ---------------------------------------------------------------------------------------
// Workbooks (docs/architecture.md, "When a workbook opens" and "The #NAME? repair").

// Excel's application events (AppEvents), of which Pyinex takes WorkbookOpen and
// NewWorkbook. Excel calls them outside any command, so they only keep the workbook for the
// queue's command. A static object: it ignores reference counts.
const IID app_events = {0x00024413, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};

class Events final : public IDispatch {
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** out) override {
        *out = iid == IID_IUnknown || iid == IID_IDispatch || iid == app_events ? this : nullptr;
        return *out ? S_OK : E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return 2; }
    ULONG STDMETHODCALLTYPE Release() override { return 1; }
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* n) override {
        *n = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS* params, VARIANT*, EXCEPINFO*, UINT*) override {
        if ((id != 0x61F && id != 0x61D) || !params || params->cArgs != 1) return S_OK;  // WorkbookOpen, NewWorkbook.
        const VARIANT& a = params->rgvarg[0];
        IDispatch* book = a.vt == VT_DISPATCH ? a.pdispVal : a.vt == (VT_DISPATCH | VT_BYREF) ? *a.ppdispVal : nullptr;
        try {
            if (book) opened.push_back(book);
        } catch (...) {
            return S_OK;
        }
        if (book) book->AddRef();
        if (!draining) arm(0);
        return S_OK;
    }
} events;

IConnectionPoint* point = nullptr;  // Excel's, while connected.
DWORD cookie = 0;

// Connects Events to Excel, once after the add-in opens, and has the command handle every
// workbook open now.
void follow(IDispatch* app) {
    following = true;
    IConnectionPointContainer* container = nullptr;
    HRESULT hr = app->QueryInterface(IID_IConnectionPointContainer, reinterpret_cast<void**>(&container));
    if (SUCCEEDED(hr)) {
        hr = container->FindConnectionPoint(app_events, &point);
        container->Release();
    }
    if (SUCCEEDED(hr)) hr = point->Advise(&events, &cookie);
    if (FAILED(hr)) {
        if (point) point->Release();
        point = nullptr;
        say(L"cannot follow workbooks as they open (" + hex(hr) +
            L"); a workbook's PyLoad cells load its scripts when they calculate, as with Ctrl+Alt+F9");
    }
    opened.push_back(nullptr);
}

// Disconnects from Excel's events and drops the workbooks waiting.
void unfollow() {
    following = false;
    if (point) {
        point->Unadvise(cookie);
        point->Release();
        point = nullptr;
    }
    for (IDispatch* book : opened)
        if (book) book->Release();
    opened.clear();
    fresh.clear();
}

// Whether formula calls name: the name whole, followed by "(", outside texts in double quotes
// and sheet names in single quotes (each doubled inside), outside brackets (a table's columns,
// where ' escapes a character, or a workbook's name), and not the end of a longer name or of
// one qualified by a sheet or workbook.
bool calls(std::wstring_view formula, std::wstring_view name) {
    const auto part = [](wchar_t c) { return IsCharAlphaNumericW(c) || std::wcschr(L"_.\\?!", c); };
    for (size_t i = 0; i < formula.size(); ++i)
        if (formula[i] == L'"' || formula[i] == L'\'') {
            if ((i = formula.find(formula[i], i + 1)) == formula.npos) return false;
        } else if (formula[i] == L'[') {  // A table's column, or a workbook's name: never a call.
            int depth = 1;
            while (depth > 0 && ++i < formula.size())
                if (formula[i] == L'\'') ++i;  // Inside brackets ' escapes the next character.
                else depth += formula[i] == L'[' ? 1 : formula[i] == L']' ? -1 : 0;
            if (depth > 0) return false;
        } else if ((i == 0 || !part(formula[i - 1])) && formula.size() - i > name.size() &&
                   formula[i + name.size()] == L'(' && same_name(formula.substr(i, name.size()), name)) {
            return true;
        }
    return false;
}

std::wstring address(IDispatch* range) {
    VARIANT where = value(range, L"Address", VT_BSTR, {omitted(), omitted(), omitted(), flag(true)});
    const std::wstring a = where.vt == VT_BSTR ? where.bstrVal : L"";
    VariantClear(&where);
    return a;
}

// Enters again the formula of a cell found by the search, if it calls name, or the whole of
// the legacy array formula (Ctrl+Shift+Enter) it is part of, once: arrays holds those done.
void enter_one(IDispatch* cell, std::wstring_view name, std::set<std::wstring>& arrays) {
    VARIANT has = value(cell, L"HasFormula", VT_BOOL), array = value(cell, L"HasArray", VT_BOOL);
    if (has.vt != VT_BOOL || !has.boolVal) return;  // The search also finds text.
    const Object whole = array.vt == VT_BOOL && array.boolVal ? member(cell, L"CurrentArray") : nullptr;
    IDispatch* target = whole ? whole.get() : cell;
    const std::wstring where = address(target);
    if (whole && !arrays.insert(where).second) return;
    // Excel without dynamic arrays has only Formula.
    const wchar_t* property = whole ? L"FormulaArray" : L"Formula2";
    VARIANT formula, done;
    VariantInit(&formula);
    VariantInit(&done);
    std::wstring why;
    const bool read = invoke(target, property, DISPATCH_PROPERTYGET, {}, formula, why) ||
                      (!whole && invoke(target, property = L"Formula", DISPATCH_PROPERTYGET, {}, formula, why));
    if (read && formula.vt == VT_BSTR && calls(formula.bstrVal, name) && whole && SysStringLen(formula.bstrVal) > 255) {
        VariantClear(&formula);  // FormulaArray takes at most 255 characters.
        why = L"a legacy array formula longer than 255 characters; enter it again with Ctrl+Shift+Enter";
    } else if (read && formula.vt == VT_BSTR && calls(formula.bstrVal, name)) {
        if (invoke(target, property, DISPATCH_PROPERTYPUT, {formula}, done, why)) why.clear();  // formula is cleared.
    } else {
        VariantClear(&formula);
        if (read) why.clear();
    }
    VariantClear(&done);
    if (!why.empty()) console::note(L"cannot enter the formula in " + where + L" again (" + why + L")");
}

// Enters again, in each worksheet of book, hidden ones too, every formula that calls name,
// found by Excel's search of formulas, so no cell is read but those it finds: Excel
// calculates each as one typed in. Every cell is found before any is entered, since entering
// one can run a workbook's change handler, whose own search FindNext would continue. Says on
// the console which it cannot.
void enter_again(IDispatch* book, const std::wstring& name) {
    const Object sheets = member(book, L"Worksheets");
    VARIANT n = sheets ? value(sheets.get(), L"Count", VT_I4) : variant(VT_EMPTY);
    for (long i = 1; n.vt == VT_I4 && i <= n.lVal; ++i) {
        const Object sheet = member(sheets.get(), L"Item", {number(i)});
        const Object cells = sheet ? member(sheet.get(), L"Cells") : nullptr;
        std::vector<Object> found;
        std::set<std::wstring> seen, arrays;  // Find goes round to the first again.
        // In formulas (-4123), as any part of one (2), in any case.
        for (Object cell = cells ? member(cells.get(), L"Find", {bstr(name + L"("), omitted(), number(-4123), number(2),
                                                                 omitted(), omitted(), flag(false)})
                                 : nullptr;
             cell && seen.insert(address(cell.get())).second;) {
            Object next = member(cells.get(), L"FindNext", {object(cell.get())});
            found.push_back(std::move(cell));
            cell = std::move(next);
        }
        for (const auto& cell : found) enter_one(cell.get(), name, arrays);
    }
}

// Calls f for book, or for each workbook open when book is null.
template <class F>
void each_book(IDispatch* app, IDispatch* book, F f) {
    if (book) return f(book);
    const Object books = member(app, L"Workbooks");
    VARIANT n = books ? value(books.get(), L"Count", VT_I4) : variant(VT_EMPTY);
    for (long i = 1; n.vt == VT_I4 && i <= n.lVal; ++i)
        if (const Object b = member(books.get(), L"Item", {number(i)})) f(b.get());
}

// Whether the queue's command has work: functions to register or remove, a repair asked for,
// workbooks to follow, or, after a close, Excel's events to connect to again should Excel
// carry on.
bool waiting() {
    return !queued.empty() || std::ranges::any_of(slots, &Slot::doomed) || repair_wanted || !opened.empty() ||
           !fresh.empty() || !following;
}

// Drops all that waits, telling Python of each function that will not be registered, and
// ends the window and its timer.
void drop_all() {
    for (const auto& f : queued) dropped.push_back(f.key);
    queued.clear();
    repair_wanted = false;
    for (IDispatch* book : opened)
        if (book) book->Release();
    opened.clear();
    fresh.clear();
    following = true;  // Nothing to follow workbooks with.
    stop();
}

// On the timer: has Excel run the queue's command, or tries again later, for as long as it
// takes. Nothing queued is dropped, unless Excel has unloaded the add-in, or refused the
// command as the add-in opened.
void attempt() {
    if (!waiting() || draining) return;
    // Inside one of the add-in's worksheet functions, reached only when a script handles
    // Windows messages: Excel may have told that call to call nothing more (excel::stop_code,
    // which the core keeps), so the command waits until the function returns.
    if (calculating > 0) return arm(retry_ms);
    if (command.empty()) {  // Never registered, so Application.Run would fail for good.
        if (!queued.empty())
            console::note(functions(queued.size()) + L" cannot be registered: Excel refused Pyinex's queue command "
                          L"as the add-in opened; close and open the add-in again");
        drop_all();
        if (const Core* core = running())
            for (const int key : std::exchange(dropped, {})) core->registered(key, false);
        return;
    }
    std::wstring why;
    bool gone = false;
    // Only functions to register look beyond workbook windows: with none open, there is no
    // workbook to follow.
    if (IDispatch* app = application(why, !queued.empty())) {
        VARIANT result;
        VariantInit(&result);
        // Excel ran the command, so whatever is left is what that run did not finish.
        if (!(gone = closed && unloaded(app)) && invoke(app, L"Run", DISPATCH_METHOD, {bstr(command)}, result, why))
            why = L"its last run of the queue stopped before the end";
        VariantClear(&result);
        app->Release();
    }
    if (gone) {
        if (!queued.empty())
            console::note(L"the add-in was unloaded, so " + functions(queued.size()) + L" will not be registered again");
        drop_all();
    } else if (waiting()) {
        if (!queued.empty() && why != reported)
            console::note(L"Excel cannot register functions yet (" + why + L"); Pyinex keeps trying");
        if (!queued.empty()) reported = why;
        arm(retry_ms);
    }
}

// Marks a run of the queue's, or the add-in opening or closing, while it lasts, restoring
// what it found: a close within a run leaves the run marked.
class Marked {
public:
    Marked() : was_(std::exchange(draining, true)) {}
    Marked(const Marked&) = delete;
    Marked& operator=(const Marked&) = delete;
    ~Marked() { draining = was_; }

private:
    bool was_;
};

// The queue's command: registers each function waiting, in order, and any queued meanwhile;
// then has the PyLoad cells of workbooks opened calculated, and repairs cells that use names
// registered for the first time, with a calculation after.
int run_queue() {
    if (draining) {
        console::note(L"the queue is already running, so this run of it does nothing");
        return 0;
    }
    if (calculating > 0) {  // Run from inside a worksheet function: never register there.
        arm(retry_ms);
        return 0;
    }
    const Marked marked;
    struct Ended {
        ~Ended() { repairing = false; }
    } ended;
    closed = false;  // Excel ran it, so it has the add-in loaded.
    reported.clear();
    std::wstring why;
    const Object app{application(why, true)};
    if (app && !following) follow(app.get());
    // A close within this run, from within a call into Excel, leaves what is left to the
    // queue's next run, once the timer has seen whether Excel unloaded the add-in.
    while (!closed) {
        if (!queued.empty()) {
            // Taken off the queue only once done: after an exception, the next run tries it
            // again. Found again by where it is: a close meanwhile puts what it removed ahead.
            const Description& f = queued.front();
            const bool done = apply(f);
            const auto at = std::ranges::find_if(queued, [&](const Description& d) { return &d == &f; });
            if (at == queued.end()) continue;
            const Description applied = std::move(*at);
            queued.erase(at);
            if (done) console::note(L"registered " + applied.name + L" from the script " + applied.script);
        } else if (const auto d = std::ranges::find_if(slots, &Slot::doomed); d != slots.end()) {
            const int key = d->f.key;
            remove(*d);
            if (const Core* core = running()) core->registered(key, false);
        } else if (repair_wanted && app) {
            // PyRepair: every workbook's PyLoad cells, then every name a direct function holds.
            // What can fail is done first, so the request is taken once and only once.
            std::vector<std::wstring> names;
            for (const auto& s : slots)
                if (s.f.key >= 0) names.push_back(s.f.name);
            opened.reserve(opened.size() + 1);
            fresh.reserve(fresh.size() + names.size());
            opened.push_back(nullptr);
            for (auto& name : names) fresh.push_back(std::move(name));
            repair_wanted = false;
            repairing = true;
        } else if (app && !opened.empty()) {
            const Object book{opened.front()};
            opened.erase(opened.begin());
            each_book(app.get(), book.get(), [](IDispatch* b) { enter_again(b, L"PyLoad"); });
        } else if (app && !fresh.empty()) {
            // Each once. The names are taken off the list only once their formulas are entered:
            // after an exception, as when memory runs out, the next run enters them again. A
            // close meanwhile empties the list, so no more than it holds are taken off.
            const size_t taken = fresh.size();
            std::vector<std::wstring> names;
            for (const auto& name : fresh)
                if (std::ranges::none_of(names, [&](const auto& n) { return same_name(n, name); }))
                    names.push_back(name);
            for (const auto& name : names) each_book(app.get(), nullptr, [&](IDispatch* b) { enter_again(b, name); });
            fresh.erase(fresh.begin(), fresh.begin() + static_cast<std::ptrdiff_t>(std::min(taken, fresh.size())));
            call(xlcCalculateNow, nullptr);
        } else {
            break;
        }
    }
    if (waiting()) arm(retry_ms);  // Workbooks to follow, waiting for Excel's Application object.
    return 1;
}

// PyRepair, a command: has every open workbook's PyLoad cells calculated, then enters again
// every formula that uses a name a direct function holds. Run while the queue runs, as by a
// workbook's change handler as the repair enters a formula, it is done once that run reaches
// it, if that run has not already; so a repair never asks for itself without end.
int repair() {
    if (calculating > 0) {  // Run from inside a worksheet function: refused, with nothing queued.
        console::note(L"PyRepair cannot run inside a worksheet function");
        return 0;
    }
    if (draining) {
        if (!repair_wanted && !repairing) {
            repair_wanted = true;
            console::note(L"the queue is already running, so PyRepair is done as that run ends");
        }
        return 0;
    }
    repair_wanted = true;
    return run_queue();
}

// Runs one of the add-in's commands. No C++ exception may reach Excel: after one, what is
// left waits for the queue's next run.
int guard(int (*body)()) {
    try {
        return body();
    } catch (...) {
        arm(retry_ms);
        return 0;
    }
}

// What every numbered export does: calls its slot's function with exactly the arguments
// Excel was told it takes, read one va_arg at a time, so no position Excel did not write is
// read (docs/architecture.md, "Direct functions").
LPXLOPER12 dispatch(int slot, LPXLOPER12 first, va_list rest) {
    return guarded([&] {
        if (in_function_wizard()) return Value::text(L"").to_excel();
        const Slot& s = slots[slot];
        const Core* core = s.f.key >= 0 ? python() : nullptr;
        if (!core) return Value::error(xlerrValue).to_excel();
        const XLOPER12* args[max_arguments];
        for (int i = 0; i < s.f.count; ++i) args[i] = i == 0 ? first : va_arg(rest, LPXLOPER12);
        Returned result(core->function(s.f.key, s.f.name.c_str(), args, s.f.count, reload_on_change));
        take_changes(*core);
        return result.release();
    });
}

}  // namespace

void open_functions() {
    closed = false;
    stop();
    unfollow();
    repair_wanted = false;
    {
        // Marked as the queue's run from the first call into Excel, so that a run started
        // meanwhile, as by PyRepair within xlfRegister, leaves its work to the run below rather
        // than taking what opening registers or removes.
        const Marked opening;
        // Named after where this copy is loaded, so that another copy of Pyinex in the same
        // Excel has a command of its own. (Run by its registration ID instead, it runs during a
        // cell edit.)
        command = L"PyinexQueue" + std::to_wstring(reinterpret_cast<uintptr_t>(&__ImageBase));
        if (!register_function({.name = command, .procedure = L"PyinexQueue", .type = L"J", .macro_type = 2})) {
            say(L"Excel refused to register Pyinex's queue command " + command +
                L", so no direct function can be registered until the add-in is opened again");
            command.clear();
        }
        if (!register_function({.name = L"PyRepair", .procedure = L"PyRepair", .type = L"J", .macro_type = 2}))
            say(L"Excel refused to register the command PyRepair");
        const Core* core = running();
        std::vector<Description> left = remove_all();  // By an xlAutoOpen with no xlAutoClose before it.
        if (!queued.empty())
            console::note(functions(queued.size()) + L" queued before the add-in closed will not be registered");
        left.insert(left.end(), queued.begin(), queued.end());
        queued.clear();
        for (const auto& f : left) dropped.push_back(f.key);
        for (const int key : dropped)
            if (core) core->registered(key, false);
        dropped.clear();
        const std::vector<std::wstring> scripts = startup_scripts();
        if (scripts.empty() || (core = python()))  // python() says why it is refused.
            for (const auto& script : scripts) {
                const Value path = Value::text(script);
                std::vector<Description> found;
                describe_all(Returned(core->load(&path.get(), reload_on_change)), script, found, *core);
                size_t i = 0;
                try {
                    for (; i < found.size(); ++i) apply(found[i]);
                } catch (...) {  // What is left is offered again when the script next loads.
                    for (; i < found.size(); ++i)
                        if (!found[i].remove) core->registered(found[i].key, false);
                    throw;
                }
            }
    }
    // Follows workbooks as they open, starting with those open now: here, a command, if a
    // workbook window gives Excel's Application object; else once the queue reaches it.
    std::wstring why;
    if (Object(application(why, false))) guard(run_queue);
    else arm(0);
}

void close_functions() {
    {
        // Marked as the queue's run while Excel is told, as opening is.
        const Marked closing;
        closed = true;
        unfollow();
        std::vector<Description> removed = remove_all();
        queued.insert(queued.begin(), removed.begin(), removed.end());  // Ahead of what waits.
    }
    arm(retry_ms);  // Even with nothing queued: the queue follows workbooks again if Excel carries on.
}

void unload_functions() { stop(); }

void take_changes(const Core& core) {
    const size_t before = queued.size();
    describe_all(Returned(core.changes()), L"", queued, core, true);
    if (queued.size() > before && !draining) arm(0);
}

LPXLOPER12 load(const Core& core, const XLOPER12& file) {
    Returned rows(core.load(&file, reload_on_change));
    // Queued after what earlier reloads changed, such as those of a call this PyLoad is
    // calculated within, so the queue keeps every change in the order the reloads happened.
    try {
        take_changes(core);
    } catch (...) {
        unoffer(*rows.get(), 0, false, core);
        throw;
    }
    const int n = describe_all(std::move(rows), text(file), queued, core);
    // Whatever this calculation found, so a timer that could not start is started again. A
    // run of the queue going on does the work before it ends.
    if (waiting() && !draining && (!armed || !queued.empty())) arm(0);
    return n < 0 ? Value::error(xlerrValue).to_excel() : Value::text(functions(n)).to_excel();
}

}  // namespace pyinex::loader

// The queue's command, which Excel runs through Application.Run.
PYINEX_EXPORT int PyinexQueue() { return pyinex::loader::guard(pyinex::loader::run_queue); }

// The repair, on demand (docs/architecture.md, "The #NAME? repair").
PYINEX_EXPORT int PyRepair() { return pyinex::loader::guard(pyinex::loader::repair); }

// The pool: PyFunc0 to PyFunc<PYINEX_POOL - 1>, generated by CMakeLists.txt. Each is
// variadic. Under the x64 convention a variadic function receives pointer arguments
// exactly as a fixed one does (rcx, rdx, r8, r9, then the stack), and its prologue only
// spills the four registers into the home area every caller reserves, so Excel may call
// one with any number of arguments.
#define PYINEX_SLOT(i)                                                          \
    PYINEX_EXPORT LPXLOPER12 PyFunc##i(LPXLOPER12 first, ...) {                 \
        va_list rest;                                                           \
        va_start(rest, first);                                                  \
        const LPXLOPER12 result = pyinex::loader::dispatch(i, first, rest);     \
        va_end(rest);                                                           \
        return result;                                                          \
    }
#include "loader/pool.inc"
