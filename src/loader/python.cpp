#include "loader/python.h"

#include "common/environment.h"
#include "console/console.h"
#include "loader/console_window.h"

#include <windows.h>

#include <psapi.h>

#include <cwchar>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <optional>
#include <string>
#include <vector>

namespace fs = std::filesystem;

namespace pyinex::loader {
namespace {

constexpr int oldest_minor = 10;  // Python 3.10
constexpr int newest_tested_minor = 15;

enum class State { untried, running, refused };
State state = State::untried;
std::wstring refusal_message;
bool started_python = false;  // Once true, a refusal is final: Python cannot be swapped.
Core core;

std::wstring getenv(const wchar_t* name) {
    std::wstring value(GetEnvironmentVariableW(name, nullptr, 0), L'\0');
    if (value.empty()) return value;
    value.resize(GetEnvironmentVariableW(name, value.data(), static_cast<DWORD>(value.size())));
    return value;
}

std::wstring lower(std::wstring s) {
    CharLowerBuffW(s.data(), static_cast<DWORD>(s.size()));
    return s;
}

std::wstring utf8_to_wide(const std::string& s) {
    std::wstring w(MultiByteToWideChar(CP_UTF8, 0, s.data(), static_cast<int>(s.size()), nullptr, 0), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.data(), static_cast<int>(s.size()), w.data(), static_cast<int>(w.size()));
    return w;
}

std::wstring trim(const std::wstring& s) {
    const auto first = s.find_first_not_of(L" \t\r\n");
    return first == s.npos ? L"" : s.substr(first, s.find_last_not_of(L" \t\r\n") - first + 1);
}

fs::path module_path(HMODULE module) {
    std::wstring path(MAX_PATH, L'\0');
    for (;;) {
        const DWORD n = GetModuleFileNameW(module, path.data(), static_cast<DWORD>(path.size()));
        if (n < path.size()) return path.substr(0, n);
        path.resize(path.size() * 2);
    }
}

HMODULE this_module() {
    HMODULE module = nullptr;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                       reinterpret_cast<LPCWSTR>(&this_module), &module);
    return module;
}

bool is_file(const fs::path& p) {
    std::error_code ec;
    return fs::is_regular_file(p, ec);
}

bool is_store_placeholder(const fs::path& p) {
    return lower(p.wstring()).find(L"\\microsoft\\windowsapps\\") != std::wstring::npos;
}

// The program a PEP 514 entry under PythonCore names: InstallPath's ExecutablePath, or else
// python.exe in the folder its default value names. Empty if it names neither.
fs::path registered_program(HKEY python_core, const std::wstring& version) {
    const std::wstring key = version + L"\\InstallPath";
    for (const wchar_t* value : {L"ExecutablePath", static_cast<const wchar_t*>(nullptr)}) {
        DWORD bytes = 0;
        if (RegGetValueW(python_core, key.c_str(), value, RRF_RT_REG_SZ, nullptr, nullptr, &bytes) != ERROR_SUCCESS) continue;
        std::wstring text(bytes / sizeof(wchar_t), L'\0');
        if (RegGetValueW(python_core, key.c_str(), value, RRF_RT_REG_SZ, nullptr, text.data(), &bytes) != ERROR_SUCCESS) continue;
        text.resize(wcsnlen(text.c_str(), text.size()));
        if (!text.empty()) return value ? fs::path(text) : fs::path(text) / L"python.exe";
    }
    return {};
}

// The newest x64 Python that python.org's installer recorded (docs/architecture.md, "Where
// Pyinex looks, in order"), the user's before all users' at the same version, or empty.
// Only a plain "3.N" names an x64 build.
fs::path registered_python() {
    fs::path newest;
    int newest_minor = -1;
    for (const HKEY root : {HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE}) {
        HKEY python_core = nullptr;
        if (RegOpenKeyExW(root, L"Software\\Python\\PythonCore", 0, KEY_READ, &python_core) != ERROR_SUCCESS) continue;
        wchar_t name[256];
        for (DWORD i = 0, size = 256; RegEnumKeyExW(python_core, i, name, &size, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS;
             ++i, size = 256) {
            const std::wstring version = name;
            if (version.size() < 3 || version.size() > 6 || !version.starts_with(L"3.") ||
                version.find_first_not_of(L"0123456789", 2) != version.npos)
                continue;
            const int minor = std::stoi(version.substr(2));
            if (minor <= newest_minor) continue;
            const fs::path program = registered_program(python_core, version);
            if (is_file(program) && !is_store_placeholder(program)) {
                newest = program;
                newest_minor = minor;
            }
        }
        RegCloseKey(python_core);
    }
    return newest;
}

// A value from pyinex.ini's [python] section, or empty. Read here, not with
// GetPrivateProfileStringW, which takes a file with no byte-order mark to be in the 8-bit
// code page, while editors save UTF-8: a path such as C:\Users\Zoë\... would come back as
// another path. A UTF-16 file, which Windows reads right, still goes to Windows.
std::wstring ini_value(const fs::path& ini, const std::wstring& key) {
    std::ifstream f(ini, std::ios::binary);
    std::string bytes{std::istreambuf_iterator<char>(f), std::istreambuf_iterator<char>()};
    if (bytes.starts_with("\xFF\xFE")) {
        std::wstring value(32768, L'\0');
        value.resize(GetPrivateProfileStringW(L"python", key.c_str(), L"", value.data(),
                                              static_cast<DWORD>(value.size()), ini.c_str()));
        return trim(value);
    }
    if (bytes.starts_with("\xEF\xBB\xBF")) bytes.erase(0, 3);
    const int size = static_cast<int>(bytes.size());
    std::wstring text = utf8_to_wide(bytes);
    if (size > 0 && MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), size, nullptr, 0) == 0) {
        text.assign(static_cast<size_t>(MultiByteToWideChar(CP_ACP, 0, bytes.data(), size, nullptr, 0)), L'\0');
        MultiByteToWideChar(CP_ACP, 0, bytes.data(), size, text.data(), static_cast<int>(text.size()));  // As Windows reads it.
    }
    // Windows' rules: lines end at \r or \n, names ignore case, ; starts a comment, only
    // the first [python] section counts, the first match in it wins, and quotes around a
    // value come off.
    bool in_python = false, seen_python = false;
    for (size_t start = 0; start < text.size();) {
        size_t end = text.find_first_of(L"\r\n", start);
        if (end == text.npos) end = text.size();
        const std::wstring line = trim(text.substr(start, end - start));
        start = end + 1;
        if (line.empty() || line.front() == L';') continue;
        if (line.front() == L'[') {
            if (seen_python) break;
            seen_python = in_python = lower(trim(line.substr(1, line.find(L']') - 1))) == L"python";
            continue;
        }
        const size_t equals = line.find(L'=');
        if (!in_python || equals == line.npos || lower(trim(line.substr(0, equals))) != key) continue;
        std::wstring value = trim(line.substr(equals + 1));
        if (value.size() >= 2 && value.front() == value.back() && (value.front() == L'"' || value.front() == L'\''))
            value = trim(value.substr(1, value.size() - 2));
        return value;
    }
    return L"";
}

// A candidate, and why it cannot be used if it cannot.
struct Choice {
    fs::path executable;
    std::wstring problem;  // A complete message when there is no executable to name.
};

// The python.exe named by PYINEX_PYTHON, pyinex.ini, PATH or the registry, in that order.
Choice choose(const fs::path& addin_folder) {
    std::wstring named = getenv(L"PYINEX_PYTHON");
    if (named.empty()) {
        const std::wstring value = ini_value(addin_folder / L"pyinex.ini", L"executable");
        // A relative path is relative to the add-in's folder, so the two can move together.
        if (!value.empty()) named = (addin_folder / value).wstring();
    }
    if (!named.empty()) {
        std::error_code ec;
        fs::path p = fs::absolute(named, ec).lexically_normal();
        if (fs::is_directory(p, ec)) p /= L"python.exe";
        return {p, L""};
    }
    std::wstring skipped;
    const std::wstring path = getenv(L"PATH");
    for (size_t start = 0; start <= path.size();) {
        size_t end = path.find(L';', start);
        if (end == path.npos) end = path.size();
        std::wstring folder = trim(path.substr(start, end - start));
        start = end + 1;
        if (folder.size() >= 2 && folder.front() == L'"' && folder.back() == L'"')
            folder = folder.substr(1, folder.size() - 2);
        if (folder.empty()) continue;
        std::error_code ec;
        const fs::path candidate = fs::absolute(fs::path(folder) / L"python.exe", ec).lexically_normal();
        if (!is_file(candidate)) continue;
        if (!is_store_placeholder(candidate)) return {candidate, L""};
        if (skipped.empty()) skipped = candidate.wstring();
    }
    if (fs::path registered = registered_python(); !registered.empty()) return {registered, L""};
    std::wstring problem = L"cannot find Python: no python.exe on PATH, and none registered";
    if (!skipped.empty()) problem += L"; skipped the Microsoft Store placeholder at " + skipped;
    return {{}, problem + L". Choose one with PYINEX_PYTHON or pyinex.ini."};
}

// The CPU a Windows program is built for (IMAGE_FILE_MACHINE_*), or 0 if it is not one.
WORD machine(const fs::path& program) {
    std::ifstream f(program, std::ios::binary);
    IMAGE_DOS_HEADER dos{};
    DWORD signature = 0;
    IMAGE_FILE_HEADER header{};
    if (!f.read(reinterpret_cast<char*>(&dos), sizeof dos) || dos.e_magic != IMAGE_DOS_SIGNATURE) return 0;
    f.seekg(dos.e_lfanew);
    if (!f.read(reinterpret_cast<char*>(&signature), sizeof signature) || signature != IMAGE_NT_SIGNATURE) return 0;
    if (!f.read(reinterpret_cast<char*>(&header), sizeof header)) return 0;
    return header.Machine;
}

// The base Python a virtual environment's pyvenv.cfg names (its "home"), or nullopt if the
// executable is not in a virtual environment. Python looks for pyvenv.cfg beside
// python.exe and in the folder above.
std::optional<fs::path> venv_home(const fs::path& executable) {
    for (const fs::path& folder : {executable.parent_path(), executable.parent_path().parent_path()}) {
        std::ifstream cfg(folder / L"pyvenv.cfg");
        if (!cfg) continue;
        for (std::string line; std::getline(cfg, line);) {
            const std::wstring wide = utf8_to_wide(line);
            const auto equals = wide.find(L'=');
            if (equals != wide.npos && lower(trim(wide.substr(0, equals))) == L"home")
                return fs::path(trim(wide.substr(equals + 1)));
        }
        return fs::path();
    }
    return std::nullopt;
}

// Why Python already in the process rules this one out, or empty if none is.
std::wstring loaded_python() {
    std::vector<HMODULE> modules(1024);
    for (DWORD needed = 0;;) {
        const auto size = static_cast<DWORD>(modules.size() * sizeof(HMODULE));
        if (!EnumProcessModules(GetCurrentProcess(), modules.data(), size, &needed))
            return L"Pyinex could not list the DLLs in Excel's process to check for another Python";
        if (needed <= size) {
            modules.resize(needed / sizeof(HMODULE));
            break;
        }
        modules.resize(needed / sizeof(HMODULE) + 64);  // More loaded meanwhile: list again.
    }
    for (HMODULE m : modules) {
        const fs::path path = module_path(m);
        const std::wstring name = lower(path.filename().wstring());
        if (name.starts_with(L"python3") && name.ends_with(L".dll"))
            return L"another add-in has already loaded Python from " + path.wstring() +
                   L"; Pyinex cannot share Excel with it";
    }
    return L"";
}

// What the chosen python.exe reports about itself when run.
struct Probe {
    int major = 0, minor = 0;
    bool free_threaded = false;
    fs::path dll;  // Its python3NN.dll.
    std::wstring base_executable;
    bool finished = false;
    DWORD exit_code = 0;
    std::wstring last_line;  // Of its output: on failure, usually the reason.

    std::wstring failure() const {
        if (!finished) return L"it did not finish starting within 60 seconds";
        std::wstring s = L"it failed to start (exit code " + std::to_wstring(exit_code) + L")";
        return last_line.empty() ? s : s + L": " + last_line;
    }
};

// Run with the same environment and current folder Python will start with in Excel, so it
// also shows that this installation can start: Python's start-up call, once made in Excel,
// ends the process if the installation is broken. It first removes the current folder,
// which `python -c` puts at the front of sys.path and the embedded Python does not, so a
// stray .py file there (beside a workbook, say) is not imported. Each line it prints starts
// on a line of its own, whatever start-up printed before it. Written to run on any Python,
// however old.
constexpr const wchar_t* probe_code =
    L"import sys;sys.path[:1]=[p for p in sys.path[:1] if p];w=lambda k,v:sys.stdout.buffer.write(('\\npyinex-probe:%s=%s\\n'%(k,v)).encode('utf-8'));"
    L"w('version','%d.%d'%sys.version_info[:2]);import sysconfig,_winapi;"
    L"w('gil_disabled',sysconfig.get_config_var('Py_GIL_DISABLED') or 0);"
    L"w('dll',_winapi.GetModuleFileName(sys.dllhandle));"
    L"w('base_executable',getattr(sys,'_base_executable',sys.executable))";

Probe probe(const fs::path& executable) {
    Probe result;
    SECURITY_ATTRIBUTES inherit{.nLength = sizeof inherit, .bInheritHandle = TRUE};
    HANDLE read = nullptr, write = nullptr;
    if (!CreatePipe(&read, &write, &inherit, 0)) {
        result.finished = true;
        result.last_line = L"Windows could not make a pipe to read its output";
        return result;
    }
    SetHandleInformation(read, HANDLE_FLAG_INHERIT, 0);

    // The program inherits only the pipe, not whatever else Excel has open.
    SIZE_T size = 0;
    InitializeProcThreadAttributeList(nullptr, 1, 0, &size);
    std::vector<char> attributes(size);
    auto* list = reinterpret_cast<LPPROC_THREAD_ATTRIBUTE_LIST>(attributes.data());
    InitializeProcThreadAttributeList(list, 1, 0, &size);
    UpdateProcThreadAttribute(list, 0, PROC_THREAD_ATTRIBUTE_HANDLE_LIST, &write, sizeof write, nullptr, nullptr);
    STARTUPINFOEXW startup{};
    startup.StartupInfo.cb = sizeof startup;
    startup.StartupInfo.dwFlags = STARTF_USESTDHANDLES;
    startup.StartupInfo.hStdOutput = write;
    startup.StartupInfo.hStdError = write;
    startup.lpAttributeList = list;

    std::wstring command = L"\"" + executable.wstring() + L"\" -c \"" + probe_code + L"\"";
    PROCESS_INFORMATION process{};
    const bool started = CreateProcessW(executable.c_str(), command.data(), nullptr, nullptr, TRUE,
                                        CREATE_NO_WINDOW | EXTENDED_STARTUPINFO_PRESENT, nullptr, nullptr,
                                        &startup.StartupInfo, &process);
    DeleteProcThreadAttributeList(list);
    CloseHandle(write);
    if (!started) {
        result.finished = true;
        result.exit_code = GetLastError();
        result.last_line = L"Windows could not run it";
        CloseHandle(read);
        return result;
    }
    // Read as it runs, so it never waits on a full pipe. Each read takes only what is there,
    // since a process the program started could hold the pipe open after it exits, and
    // only the last 64 KB are kept: the lines wanted come last.
    std::string output;
    const auto read_some = [&] {
        DWORD available = 0;
        if (!PeekNamedPipe(read, nullptr, 0, nullptr, &available, nullptr) || available == 0) return false;
        std::string chunk(available, '\0');
        DWORD got = 0;
        if (!ReadFile(read, chunk.data(), available, &got, nullptr) || got == 0) return false;
        output.append(chunk, 0, got);
        if (output.size() > 65536) output.erase(0, output.size() - 65536);
        return true;
    };
    const ULONGLONG deadline = GetTickCount64() + 60'000;
    while (!(result.finished = WaitForSingleObject(process.hProcess, 20) == WAIT_OBJECT_0) &&
           GetTickCount64() < deadline)
        read_some();
    while (GetTickCount64() < deadline && read_some()) {
    }
    if (!result.finished) TerminateProcess(process.hProcess, 1);
    GetExitCodeProcess(process.hProcess, &result.exit_code);
    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    CloseHandle(read);

    for (size_t start = 0; start < output.size();) {
        size_t end = output.find('\n', start);
        if (end == output.npos) end = output.size();
        const std::wstring line = trim(utf8_to_wide(output.substr(start, end - start)));
        start = end + 1;
        if (line.empty()) continue;
        constexpr std::wstring_view prefix = L"pyinex-probe:";
        const auto equals = line.find(L'=');
        if (!line.starts_with(prefix) || equals == line.npos) {
            result.last_line = line;
            continue;
        }
        const std::wstring key = line.substr(prefix.size(), equals - prefix.size()), value = line.substr(equals + 1);
        if (key == L"version") swscanf_s(value.c_str(), L"%d.%d", &result.major, &result.minor);
        if (key == L"gil_disabled") result.free_threaded = value != L"0";
        if (key == L"dll") result.dll = value;
        if (key == L"base_executable") result.base_executable = value;
    }
    return result;
}

// Checks the chosen python.exe; returns why it cannot be used, or empty. Fills in probed.
std::wstring check(const fs::path& executable, Probe& probed) {
    if (is_store_placeholder(executable))
        return L"it is the Microsoft Store placeholder, not an installed Python; install Python from python.org";
    if (!is_file(executable)) return L"file not found";
    const std::wstring name = lower(executable.filename().wstring());
    if (name.starts_with(L"python3") && name.ends_with(L"t.exe")) return L"free-threaded builds are not supported";
    switch (machine(executable)) {
    case IMAGE_FILE_MACHINE_AMD64:
        break;
    case IMAGE_FILE_MACHINE_I386:
        return L"it is a 32-bit (x86) build; Excel needs the x64 build";
    case IMAGE_FILE_MACHINE_ARM64:
        return L"it is an ARM64 build; Excel needs the x64 build";
    default:
        return L"it is not an x64 Windows program";
    }
    if (const auto home = venv_home(executable)) {
        std::error_code ec;
        if (home->empty()) return L"its pyvenv.cfg names no base Python";
        if (!fs::is_directory(*home, ec))
            return L"its pyvenv.cfg names base Python " + home->wstring() + L", which does not exist";
    }
    if (const std::wstring other = loaded_python(); !other.empty()) return other;
    probed = probe(executable);
    if (!probed.finished || probed.major == 0) return probed.failure();
    if (probed.major < 3 || (probed.major == 3 && probed.minor < oldest_minor))
        return L"version " + std::to_wstring(probed.major) + L"." + std::to_wstring(probed.minor) +
               L" is too old; Pyinex needs 3." + std::to_wstring(oldest_minor) + L" or later";
    if (probed.exit_code != 0 || probed.dll.empty() || probed.base_executable.empty()) return probed.failure();
    if (probed.free_threaded) return L"free-threaded builds are not supported";
    return L"";
}

// The environment Python inherits: docs/architecture.md, "Inherited environment variables".
void prepare_environment(const fs::path& addin_folder) {
    if (const std::wstring home = getenv(L"PYTHONHOME"); !home.empty()) {
        console::note(L"ignoring PYTHONHOME (" + home + L"): Pyinex finds the standard library from the Python it chose.");
        common::set_environment(L"PYTHONHOME", L"");
    }
    if (const std::wstring path = getenv(L"PYTHONPATH"); !path.empty())
        console::note(L"PYTHONPATH is set, adding to sys.path: " + path);
    if (getenv(L"PYTHONVERBOSE").empty()) {
        const std::wstring verbose = ini_value(addin_folder / L"pyinex.ini", L"verbose");
        if (!verbose.empty()) common::set_environment(L"PYTHONVERBOSE", verbose.c_str());
    }
}

std::wstring windows_error(const std::wstring& what) {
    return L"cannot load " + what + L" (Windows error " + std::to_wstring(GetLastError()) + L")";
}

// Loads the Python DLLs from the chosen installation, then the core, and starts Python.
// Returns why it failed, or empty.
std::wstring load(const fs::path& executable, const Probe& probed, const fs::path& addin_folder) {
    // python3.dll is loaded explicitly from the same folder: otherwise Windows would
    // resolve the core's link to it by searching, and could find another Python's.
    const fs::path abi = probed.dll.parent_path() / L"python3.dll";
    const fs::path core_path = addin_folder / L"pyinex_core.dll";
    // Until Python starts, a failure unloads what was loaded, so reopening the add-in can
    // try again (forget_refusal).
    std::vector<HMODULE> loaded;
    const auto fail = [&](const std::wstring& why) {
        for (auto m = loaded.rbegin(); m != loaded.rend(); ++m) FreeLibrary(*m);
        return why;
    };
    for (const fs::path& dll : {probed.dll, abi, core_path}) {
        const HMODULE module = LoadLibraryExW(dll.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH);
        if (!module) return fail(windows_error(dll.wstring()));
        loaded.push_back(module);
    }
    auto* start = reinterpret_cast<pyinex_start_t*>(GetProcAddress(loaded.back(), "pyinex_start"));
    auto* call = reinterpret_cast<pyinex_call_t*>(GetProcAddress(loaded.back(), "pyinex_call"));
    auto* info = reinterpret_cast<pyinex_info_t*>(GetProcAddress(loaded.back(), "pyinex_info"));
    auto* close = reinterpret_cast<pyinex_close_t*>(GetProcAddress(loaded.back(), "pyinex_close"));
    if (!start || !call || !info || !close) return fail(core_path.wstring() + L" is not this version's core");
    // Python and the core will hold pointers into this add-in, so Windows must never
    // unload it (docs/architecture.md, "Lifetime").
    pin();
    started_python = true;
    if (start(executable.c_str(), probed.base_executable.c_str()) != 0)
        return L"Python started, but Pyinex could not set it up";
    core = {.call = call, .info = info, .close = close, .executable = executable.wstring(), .dll = probed.dll.wstring()};
    return L"";
}

// Returns why Python cannot be used, or empty once it is running.
std::wstring start() {
    const fs::path addin_folder = module_path(this_module()).parent_path();
    prepare_environment(addin_folder);
    const Choice choice = choose(addin_folder);
    if (choice.executable.empty()) return choice.problem;
    Probe probed;
    if (const std::wstring problem = check(choice.executable, probed); !problem.empty())
        return L"cannot use Python " + choice.executable.wstring() + L": " + problem +
               L". Choose another with PYINEX_PYTHON or pyinex.ini.";
    if (probed.major > 3 || probed.minor > newest_tested_minor)
        console::note(L"Python " + std::to_wstring(probed.major) + L"." + std::to_wstring(probed.minor) +
                      L" is newer than any Pyinex was tested with (3." + std::to_wstring(newest_tested_minor) +
                      L"); using it anyway.");
    return load(choice.executable, probed, addin_folder);
}

}  // namespace

const Core* python() {
    if (state == State::untried) {
        prepare_output();  // First, so this start's own messages reach the console too.
        refusal_message = start();
        state = refusal_message.empty() ? State::running : State::refused;
        if (state == State::refused) console::note(refusal_message);
    }
    return state == State::running ? &core : nullptr;
}

const Core* running() { return state == State::running ? &core : nullptr; }

const std::wstring& refusal() { return refusal_message; }

void pin() {
    HMODULE pinned = nullptr;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN,
                       reinterpret_cast<LPCWSTR>(&this_module), &pinned);
}

std::wstring addin_path() { return module_path(this_module()).wstring(); }

void forget_refusal() {
    if (state == State::refused && !started_python) {
        state = State::untried;
        refusal_message.clear();
    }
}

}  // namespace pyinex::loader
