// Changing the environment variables Python reads as it starts.
#pragma once

namespace pyinex::common {

// Sets a variable, or removes it when value is empty, both where Python reads it and where
// processes Pyinex starts inherit it.
void set_environment(const wchar_t* name, const wchar_t* value);

}  // namespace pyinex::common
