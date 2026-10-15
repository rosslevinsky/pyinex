// The console window, a window of Pyinex's own: docs/architecture.md, "Console" and
// "PyConsole". Both functions run on Excel's thread, which handles the window's messages.
#pragma once

#include <optional>

namespace pyinex::loader {

// Called before Python starts. When the process has no standard output, as in Excel,
// makes the window and points the standard handles and the C runtime's output at it, so
// Python's sys.stdout and sys.stderr, and C code's printf, write to it.
void prepare_output();

struct Placement {
    std::optional<int> x, y, width, height;  // x with y, width with height.
};

// PyConsole's work, once its arguments are checked: shows or hides the window, clears its
// text, and places it if it is going from hidden to shown. Makes the window if there is
// none.
void show_console(bool show, bool clear, const Placement& placement);

}  // namespace pyinex::loader
