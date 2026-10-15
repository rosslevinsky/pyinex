// The pyinex module scripts import: docs/architecture.md, "The pyinex module".
#pragma once

namespace pyinex::core {

// Makes the module and puts it in sys.modules. Call once, with the GIL held, after
// start_conversion. False if Python failed.
bool start_module();

// While one lives, Excel is calling the add-in on this thread, and the module's functions
// may ask Excel for something; at any other time, or from any other thread, they raise
// RuntimeError instead (docs/architecture.md, "When they may be called").
class ExcelCall {
public:
    ExcelCall();
    ExcelCall(const ExcelCall&) = delete;
    ExcelCall& operator=(const ExcelCall&) = delete;
    ~ExcelCall();

private:
    unsigned long outer_;  // The thread of the call this one is inside, if any.
};

}  // namespace pyinex::core
