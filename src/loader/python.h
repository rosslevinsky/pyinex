// Choosing a Python, checking it, and starting it in pyinex_core.dll: docs/architecture.md,
// "Choosing a Python". Nothing here links Python, so the add-in loads without one.
#pragma once

#include "core/api.h"

#include <string>

namespace pyinex::loader {

// pyinex_core.dll's functions.
struct Core {
    pyinex_call_t* call = nullptr;
    pyinex_info_t* info = nullptr;
    pyinex_close_t* close = nullptr;
    std::wstring executable;  // The python.exe chosen.
    std::wstring dll;         // Its python3NN.dll, loaded.
};

// Chooses, checks and starts Python the first time it is called. Returns the core once
// Python is running, or null when it was refused; refusal() then says why.
const Core* python();

// The core if Python is running, without starting it.
const Core* running();

// Why Python was refused, or empty if it was not (or has not been tried).
const std::wstring& refusal();

// Keeps the add-in loaded for the rest of the process, whatever Excel does.
void pin();

// The add-in's own path.
std::wstring addin_path();

// Called by xlAutoOpen: a refused Python gets a fresh attempt when the add-in reopens.
void forget_refusal();

}  // namespace pyinex::loader
