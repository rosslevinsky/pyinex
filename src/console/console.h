// Pyinex's own messages: refusals, start-up notes and the like (docs/architecture.md,
// "Console"). Each goes out as one line starting "Pyinex: ".
//
// A message goes to the process's standard error: the console window in Excel, once
// loader/console_window.h has made it, a terminal's console if Excel was started from one,
// and the captured output in the fake Excel.
#pragma once

#include <string_view>

namespace pyinex::console {

void note(std::wstring_view message);

}  // namespace pyinex::console
