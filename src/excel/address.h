// Cell addresses: the letters Excel names its columns with.
//
// The letters are a base-26 number with no zero digit: A to Z are 1 to 26, AA follows Z,
// and AAA follows ZZ. tests/unit/test_colnames.py checks every column from A to XFD.
#pragma once

#include <string>
#include <string_view>

namespace pyinex::excel {

// Column 1 is A and column 16,384 is XFD. Empty for a number outside that range.
std::wstring column_letters(int column);

// The inverse: 0 for anything but the upper-case letters of a column from A to XFD.
int column_number(std::wstring_view letters);

}  // namespace pyinex::excel
