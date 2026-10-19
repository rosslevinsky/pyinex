# The fake Excel used by tests/host, and test programs that exercise it and the add-in.

add_executable(xllhost xllhost/xllhost.cpp)
# oleacc: a workbook window gives its stand-in for Excel's Application object as Excel's does.
target_link_libraries(xllhost PRIVATE excel_sdk_headers oleacc)
pyinex_defaults(xllhost)
# The add-in finds Excel's MdCallBack12 among the exports of the process's .exe.
set_target_properties(xllhost PROPERTIES ENABLE_EXPORTS ON)

# The same fake Excel built as a program with a window, as Excel is, for tests that start it
# as Excel starts, with no standard output or error (test_console.py).
add_executable(xllhost_gui WIN32 xllhost/xllhost.cpp)
target_link_libraries(xllhost_gui PRIVATE excel_sdk_headers oleacc)
target_link_options(xllhost_gui PRIVATE /ENTRY:mainCRTStartup)
pyinex_defaults(xllhost_gui)
set_target_properties(xllhost_gui PROPERTIES ENABLE_EXPORTS ON)

# A deliberately wrong add-in, which shows the fake Excel catches each broken rule.
add_library(badxll SHARED badxll/badxll.cpp)
set_target_properties(badxll PROPERTIES SUFFIX ".xll")
target_link_libraries(badxll PRIVATE excel_sdk_headers)
pyinex_defaults(badxll)

# A stand-in python.exe for the discovery tests.
add_executable(fakepython fakepython/fakepython.cpp)
pyinex_defaults(fakepython)
