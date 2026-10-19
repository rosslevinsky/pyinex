// The pyinex module scripts import: docs/architecture.md, "The pyinex module".
#pragma once

namespace pyinex::core {

// Makes the module and puts it in sys.modules. Call once, with the GIL held, after
// start_conversion and start_calls. False, with Python's error left set, if Python failed.
bool start_module();

// While one lives, Excel is calling the add-in on this thread, and the module's functions
// may ask Excel for something; at any other time, or from any other thread, they raise
// RuntimeError instead (docs/architecture.md, "When they may be called"). It keeps
// Microsoft's rule after xlretUncalced or xlretAbort for its length (excel::stop_code).
class ExcelCall {
public:
    ExcelCall();
    ExcelCall(const ExcelCall&) = delete;
    ExcelCall& operator=(const ExcelCall&) = delete;
    ~ExcelCall();

private:
    unsigned long outer_;  // The thread of the call this one is inside, if any.
    int outer_stop_;       // And its excel::stop_code.
};

// While one lives, the module's functions raise RuntimeError, as outside any call, even
// within one: for PyLoad, whose script runs outside any call, and for what closing flushes.
// Nothing it meets changes the stop of a call it is inside (excel::stop_code), and a call
// already stopped stays so, asking Excel nothing more.
class OutsideCall {
public:
    OutsideCall();
    OutsideCall(const OutsideCall&) = delete;
    OutsideCall& operator=(const OutsideCall&) = delete;
    ~OutsideCall();

private:
    unsigned long outer_;
    int outer_stop_;
};

}  // namespace pyinex::core
