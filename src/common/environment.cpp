#include "common/environment.h"

#include <windows.h>

namespace pyinex::common {

void set_environment(const wchar_t* name, const wchar_t* value) {
    // Windows' copy, which new processes inherit.
    SetEnvironmentVariableW(name, *value ? value : nullptr);
    // Python reads variables through the release C runtime, ucrtbase.dll, which keeps a
    // copy of its own. A release build of Pyinex shares that runtime, but a debug build has
    // its own (ucrtbased.dll), so the change is made in ucrtbase.dll by name. If ucrtbase.dll
    // is not loaded yet, it will copy Windows' environment, changed above, when it is.
    using putenv_t = errno_t(const wchar_t*, const wchar_t*);
    if (const HMODULE crt = GetModuleHandleW(L"ucrtbase.dll"))
        if (auto* put = reinterpret_cast<putenv_t*>(GetProcAddress(crt, "_wputenv_s"))) put(name, value);
}

}  // namespace pyinex::common
