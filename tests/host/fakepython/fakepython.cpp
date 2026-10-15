// fakepython: stands in for a python.exe that the tests cannot install, such as one too old
// or broken. Named python.exe in a test's folder, it ignores its arguments, writes the
// environment variable FAKEPYTHON_OUTPUT to standard output, with each "|" turned into a
// line break, and exits with the code in FAKEPYTHON_EXIT (default 0).
#include <cstdio>
#include <cstdlib>
#include <string>

int main() {
    char* output = nullptr;
    char* code = nullptr;
    size_t size = 0;
    _dupenv_s(&output, &size, "FAKEPYTHON_OUTPUT");
    _dupenv_s(&code, &size, "FAKEPYTHON_EXIT");
    std::string text = output ? output : "";
    for (char& c : text)
        if (c == '|') c = '\n';
    std::fputs(text.c_str(), stdout);
    const int exit_code = code ? std::atoi(code) : 0;
    std::free(output);
    std::free(code);
    return exit_code;
}
