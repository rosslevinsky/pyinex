// Direct functions: the pool of numbered exports, the slot table, registering the functions
// of the scripts pyinex.ini lists and PyLoad loads, the queue that registers them after
// start-up, loading a workbook's PyLoad scripts as it opens, and the #NAME? repair
// (docs/architecture.md, "Direct functions", "PyLoad", "The queue", "When a workbook opens"
// and "The #NAME? repair").
#pragma once

#include "loader/python.h"

namespace pyinex::loader {

// PyReloadOnChange's setting, which direct functions follow as PyCall does.
extern bool reload_on_change;

// Called by xlAutoOpen: drops what a close before left queued, removes any functions still
// registered, then, if pyinex.ini lists scripts, starts Python, loads each and registers its
// functions, saying on the console what failed; then follows workbooks as they open.
void open_functions();

// Called by xlAutoClose, a command: stops following workbooks; unregisters every direct
// function and clears its name, so that no name reaches a slot another function takes later,
// reporting each that failed; and queues each to be registered again, and Excel's events to
// be followed again, in case Excel carries on (a quit cancelled at the save prompt).
void close_functions();

// Called by DllMain as Windows unloads the add-in: ends the queue's window and its timer,
// whose code goes with the add-in. Calls no Python.
void unload_functions();

// Queues what reloads of PyLoad's scripts changed (pyinex_changes). Called after each call
// into the core that may have reloaded a script.
void take_changes(const Core& core);

// PyLoad: loads the script file names and queues its new functions to be registered.
// Returns the status, or #VALUE!.
LPXLOPER12 load(const Core& core, const XLOPER12& file);

}  // namespace pyinex::loader
