#include "excel/address.h"

#include "excel/excel.h"

namespace pyinex::excel {

std::wstring column_letters(int column) {
    std::wstring letters;
    if (column < 1 || column > max_columns) return letters;
    // Take the last letter off each time; subtracting 1 first is what makes A, not a
    // zero digit, follow Z.
    for (int n = column; n > 0; n = (n - 1) / 26)
        letters.insert(letters.begin(), static_cast<wchar_t>(L'A' + (n - 1) % 26));
    return letters;
}

int column_number(std::wstring_view letters) {
    if (letters.empty() || letters.size() > 3) return 0;
    int n = 0;
    for (const wchar_t c : letters) {
        if (c < L'A' || c > L'Z') return 0;
        n = n * 26 + (c - L'A' + 1);
    }
    return n <= max_columns ? n : 0;
}

}  // namespace pyinex::excel
