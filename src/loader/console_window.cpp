#include "loader/console_window.h"

#include "loader/python.h"

#include <windows.h>
#include <dwmapi.h>
#include <uxtheme.h>

#include <cstdio>
#include <cstring>
#include <cwchar>
#include <iterator>
#include <mutex>
#include <string>
#include <utility>

namespace pyinex::loader {
namespace {

constexpr wchar_t window_class[] = L"PyinexConsole";
constexpr UINT text_arrived = WM_APP + 1;  // Posted by the reading thread.
constexpr size_t most_text = 1'000'000;    // Characters the window keeps.

HWND window = nullptr;
HWND text_box = nullptr;
bool shown = false;  // What PyConsole last asked for, until closing the window hides it.
bool dark = false;   // Windows' app mode is dark: the window's colors follow it.
constexpr COLORREF dark_text = RGB(0xDC, 0xDC, 0xDC), dark_color = RGB(0x1E, 0x1E, 0x1E);
HBRUSH dark_brush = nullptr;

// Text read from the pipe and not yet in the window. The reading thread adds to it and
// posts text_arrived once; the window takes it on Excel's thread. Its room is reserved
// before the thread starts.
std::mutex pending_lock;
std::wstring pending;
bool posted = false;

bool usable(HANDLE h) { return h != nullptr && h != INVALID_HANDLE_VALUE; }

HMODULE this_module() {
    HMODULE module = nullptr;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                       reinterpret_cast<LPCWSTR>(&this_module), &module);
    return module;
}

// Adds the pending text at the end, with lines ending \r\n as an edit control's do, and
// drops the oldest lines beyond most_text.
void add_pending_text() {
    std::wstring text;
    {
        std::lock_guard lock(pending_lock);
        text = pending;
        pending.clear();  // Keeping its room.
        posted = false;
    }
    static wchar_t last = 0;  // The previous text's last character, for a \r\n split between two.
    std::wstring lines;
    for (wchar_t c : text) {
        if (c == L'\n' && last != L'\r') lines += L'\r';
        if (c == L'\0') c = L'\uFFFD';  // A text box's text ends at a NUL.
        lines += last = c;
    }
    if (lines.size() > most_text) lines.erase(0, lines.size() - most_text);
    const auto length = static_cast<size_t>(GetWindowTextLengthW(text_box));
    if (length + lines.size() > most_text) {
        // Drop whole lines from the start, up to the one holding the character that makes room.
        const LRESULT line = SendMessageW(text_box, EM_LINEFROMCHAR, length + lines.size() - most_text, 0);
        const LRESULT end = SendMessageW(text_box, EM_LINEINDEX, line + 1, 0);
        SendMessageW(text_box, EM_SETSEL, 0, end < 0 ? length : end);
        SendMessageW(text_box, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(L""));
    }
    const int end = GetWindowTextLengthW(text_box);
    SendMessageW(text_box, EM_SETSEL, end, end);
    SendMessageW(text_box, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(lines.c_str()));
}

// Follows Windows' app mode (Settings, Personalization, Colors): the title bar, the scroll
// bar and the text box's colors. Windows documents a dark title bar but no dark scroll bar;
// "DarkMode_Explorer" is the theme File Explorer's own uses.
void follow_app_mode() {
    DWORD light = 1, size = sizeof light;
    RegGetValueW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                 L"AppsUseLightTheme", RRF_RT_REG_DWORD, nullptr, &light, &size);
    dark = light == 0;
    const BOOL on = dark;
    DwmSetWindowAttribute(window, DWMWA_USE_IMMERSIVE_DARK_MODE, &on, sizeof on);
    SetWindowTheme(text_box, dark ? L"DarkMode_Explorer" : nullptr, nullptr);
    InvalidateRect(window, nullptr, TRUE);
}

LRESULT CALLBACK window_procedure(HWND w, UINT message, WPARAM wparam, LPARAM lparam) {
    switch (message) {
    case WM_CTLCOLORSTATIC:  // The read-only text box asks its colors as it draws.
        if (!dark) break;
        SetTextColor(reinterpret_cast<HDC>(wparam), dark_text);
        SetBkColor(reinterpret_cast<HDC>(wparam), dark_color);
        return reinterpret_cast<LRESULT>(dark_brush);
    case WM_SETTINGCHANGE:  // Sent as the user switches between light and dark.
        if (lparam && std::wcscmp(reinterpret_cast<const wchar_t*>(lparam), L"ImmersiveColorSet") == 0)
            follow_app_mode();
        break;
    case WM_CLOSE:  // Closing only hides it: PyConsole(TRUE) shows it again.
        ShowWindow(w, SW_HIDE);
        shown = false;
        return 0;
    case WM_SIZE:
        if (text_box) MoveWindow(text_box, 0, 0, LOWORD(lparam), HIWORD(lparam), TRUE);
        return 0;
    case text_arrived:
        add_pending_text();
        return 0;
    }
    return DefWindowProcW(w, message, wparam, lparam);
}

// The reading thread: carries what arrives on the pipe to the window, as UTF-16. It never
// calls Python, and never waits on the window, which only Excel's thread may touch. It
// allocates nothing, so it never changes the C runtime's heap while Excel's thread uses
// it: the fake Excel counts that heap's blocks to check xlAutoFree12 (docs/testing.md).
// For the same reason make() waits for it to start, since Windows sets up the C runtime for
// a new thread, on that heap, before the thread runs.
HANDLE reader_started = nullptr;

// How many bytes at the end of what was read begin a UTF-8 character the next read may
// finish: a lead byte, with only continuation bytes after it, fewer than it needs. Anything
// else goes now, valid or not, so a stray byte never holds back what follows it.
size_t unfinished(const char* bytes, size_t n) {
    for (size_t back = 1; back <= 3 && back <= n; ++back) {
        const auto b = static_cast<unsigned char>(bytes[n - back]);
        if ((b & 0xC0) == 0x80) continue;  // A continuation byte: its lead is further back.
        const size_t needs = (b & 0xE0) == 0xC0 ? 2 : (b & 0xF0) == 0xE0 ? 3 : (b & 0xF8) == 0xF0 ? 4 : 0;
        return back < needs ? back : 0;
    }
    return 0;
}

DWORD WINAPI read_output(void* pipe) {
    SetEvent(reader_started);
    constexpr DWORD chunk = 4096;
    char bytes[chunk + 3];  // A read, after up to 3 bytes of a character the last one cut off.
    wchar_t text[chunk + 3];
    size_t kept = 0;
    for (DWORD got = 0; ReadFile(pipe, bytes + kept, chunk, &got, nullptr) && got > 0;) {
        const size_t n = kept + got;
        // A character cut off at the end waits for the next read. Text that is not UTF-8
        // goes as it is, with replacement characters.
        const size_t use = n - unfinished(bytes, n);
        const auto length = static_cast<size_t>(
            MultiByteToWideChar(CP_UTF8, 0, bytes, static_cast<int>(use), text, static_cast<int>(std::size(text))));
        kept = n - use;
        std::memmove(bytes, bytes + use, kept);
        std::lock_guard lock(pending_lock);
        if (pending.size() + length > most_text) pending.erase(0, pending.size() + length - most_text);
        pending.append(text, length);
        if (!posted) posted = PostMessageW(window, text_arrived, 0, 0) != 0;
    }
    CloseHandle(pipe);  // Only if the output never reached the pipe, as make() then closes its end.
    return 0;
}

// Points the C runtime's standard output and error, and the process's standard handles,
// at the pipe. It is done in ucrtbase.dll, the Universal CRT that Python and C extensions
// use, so this one redirect serves them all; it is named rather than linked because a
// debug build of the add-in links the debug CRT instead, and it stays loaded, since Python
// loads it anyway and the streams live in it.
//
// In a program with a window, such as Excel, the CRT marks stdout and stderr as having no
// output (descriptor -2) when it starts, and Python asks the streams for their
// descriptors, so the streams are reopened first, then their descriptors, and 1 and 2 for
// code that writes to those directly, made copies of the pipe's.
bool send_output_to(HANDLE pipe, bool out, bool err) {
    const HMODULE crt = LoadLibraryW(L"ucrtbase.dll");
    if (!crt) return false;
    const auto stream = reinterpret_cast<FILE*(__cdecl*)(unsigned)>(GetProcAddress(crt, "__acrt_iob_func"));
    const auto reopen = reinterpret_cast<errno_t(__cdecl*)(FILE**, const wchar_t*, const wchar_t*, FILE*)>(
        GetProcAddress(crt, "_wfreopen_s"));
    const auto descriptor_of = reinterpret_cast<int(__cdecl*)(FILE*)>(GetProcAddress(crt, "_fileno"));
    const auto open_handle = reinterpret_cast<int(__cdecl*)(intptr_t, int)>(GetProcAddress(crt, "_open_osfhandle"));
    const auto dup2 = reinterpret_cast<int(__cdecl*)(int, int)>(GetProcAddress(crt, "_dup2"));
    const auto handle_of = reinterpret_cast<intptr_t(__cdecl*)(int)>(GetProcAddress(crt, "_get_osfhandle"));
    if (!stream || !reopen || !descriptor_of || !open_handle || !dup2 || !handle_of) return false;
    // Binary: the window turns \n into \r\n itself. The descriptor stays open for the
    // process's life, as the pipe does.
    const int piped = open_handle(reinterpret_cast<intptr_t>(pipe), 0x8000 /* _O_BINARY */);
    if (piped < 0) return false;
    for (const auto [number, which] : {std::pair{1u, STD_OUTPUT_HANDLE}, std::pair{2u, STD_ERROR_HANDLE}}) {
        if (!(number == 1 ? out : err)) continue;  // One the process has keeps it.
        FILE* reopened = nullptr;
        if (reopen(&reopened, L"NUL", L"w", stream(number)) != 0) continue;
        const int descriptor = descriptor_of(reopened);
        const int standard = static_cast<int>(number);
        dup2(piped, descriptor);
        if (descriptor != standard) dup2(piped, standard);
        SetStdHandle(which, reinterpret_cast<HANDLE>(handle_of(standard)));
    }
    return true;
}

// Makes the window, hidden, and if the process lacks standard output or error, the pipe and
// thread that bring what goes to either there.
void make() {
    if (window) return;
    // The window's procedure and the reading thread are in this add-in.
    pin();
    const HMODULE module = this_module();
    const WNDCLASSW wc{.lpfnWndProc = window_procedure,
                       .hInstance = module,
                       .hCursor = LoadCursorW(nullptr, IDC_ARROW),
                       .lpszClassName = window_class};
    RegisterClassW(&wc);
    // No owner: a window Excel owns closes with it, and each workbook has a window of its own.
    window = CreateWindowExW(0, window_class, L"Pyinex console", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT,
                             800, 400, nullptr, nullptr, module, nullptr);
    if (!window) return;
    text_box = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_READONLY,
                               0, 0, 0, 0, window, nullptr, module, nullptr);
    RECT inside{};
    GetClientRect(window, &inside);
    MoveWindow(text_box, 0, 0, inside.right, inside.bottom, FALSE);
    SendMessageW(text_box, EM_SETLIMITTEXT, 0, 0);  // 0: the most an edit control can hold.
    const HFONT font = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                   CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, FIXED_PITCH | FF_MODERN, L"Consolas");
    SendMessageW(text_box, WM_SETFONT, reinterpret_cast<WPARAM>(font), FALSE);
    dark_brush = CreateSolidBrush(dark_color);
    follow_app_mode();
    // A process that already has output, as the fake Excel's is captured, keeps it; one that
    // has only one of the two gets the other.
    const bool out = !usable(GetStdHandle(STD_OUTPUT_HANDLE)), err = !usable(GetStdHandle(STD_ERROR_HANDLE));
    if (!out && !err) return;
    HANDLE read = nullptr, write = nullptr;
    if (!CreatePipe(&read, &write, nullptr, 1 << 16)) return;
    pending.reserve(most_text);
    // The thread first: output sent to a pipe nobody reads would stop Excel once it filled.
    reader_started = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    const HANDLE thread = reader_started ? CreateThread(nullptr, 0, read_output, read, 0, nullptr) : nullptr;
    if (!thread) {
        if (reader_started) CloseHandle(reader_started);
        CloseHandle(read);
        CloseHandle(write);
        return;
    }
    WaitForSingleObject(reader_started, INFINITE);
    CloseHandle(reader_started);
    CloseHandle(thread);
    if (!send_output_to(write, out, err)) CloseHandle(write);  // The thread then ends.
}

}  // namespace

void prepare_output() {
    if (!usable(GetStdHandle(STD_OUTPUT_HANDLE)) || !usable(GetStdHandle(STD_ERROR_HANDLE))) make();
}

void show_console(bool show, bool clear, const Placement& placement) {
    make();
    if (!window) return;  // Windows would not make one.
    if (clear) {
        {
            std::lock_guard lock(pending_lock);
            pending.clear();
        }
        SetWindowTextW(text_box, L"");
    }
    if (show && !shown) {
        UINT flags = SWP_NOACTIVATE;
        if (!placement.x) flags |= SWP_NOMOVE;
        if (!placement.width) flags |= SWP_NOSIZE;
        // In front, but Excel keeps the keyboard.
        SetWindowPos(window, HWND_TOP, placement.x.value_or(0), placement.y.value_or(0), placement.width.value_or(0),
                     placement.height.value_or(0), flags | SWP_SHOWWINDOW);
    } else if (!show && shown) {
        ShowWindow(window, SW_HIDE);
    }
    shown = show;
}

}  // namespace pyinex::loader
