#include "console/console.h"

#include <windows.h>

#include <string>

namespace pyinex::console {

void note(std::wstring_view message) {
    const HANDLE out = GetStdHandle(STD_ERROR_HANDLE);
    if (out == nullptr || out == INVALID_HANDLE_VALUE) return;
    const std::wstring line = L"Pyinex: " + std::wstring(message) + L"\n";
    DWORD written = 0, mode = 0;
    if (GetConsoleMode(out, &mode)) {  // A console takes UTF-16, whatever its code page.
        WriteConsoleW(out, line.data(), static_cast<DWORD>(line.size()), &written, nullptr);
        return;
    }
    const int size = WideCharToMultiByte(CP_UTF8, 0, line.data(), static_cast<int>(line.size()),
                                         nullptr, 0, nullptr, nullptr);
    std::string utf8(size, '\0');
    WideCharToMultiByte(CP_UTF8, 0, line.data(), static_cast<int>(line.size()), utf8.data(), size,
                        nullptr, nullptr);
    WriteFile(out, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
}

}  // namespace pyinex::console
