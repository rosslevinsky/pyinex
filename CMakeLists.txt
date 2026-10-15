cmake_minimum_required(VERSION 3.26)
project(pyinex VERSION 2026.10.0 LANGUAGES CXX)

if(NOT MSVC OR NOT CMAKE_SIZEOF_VOID_P EQUAL 8)
  message(FATAL_ERROR "Pyinex builds only with MSVC for x64.")
endif()

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
# Every binary in one folder per preset, with no per-configuration subfolder: the tests
# and the CI artifacts find them there.
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY "$<1:${CMAKE_BINARY_DIR}/bin>")
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY "$<1:${CMAKE_BINARY_DIR}/bin>")

include(cmake/ExcelSdk.cmake)

# The core is compiled against Python 3.15's headers but uses only the stable ABI as of
# 3.10 (Py_LIMITED_API below), so one build runs with any Python from 3.10 on. Using the
# newest headers turns anything deprecated there into a warning, and /WX into an error.
# Set Python_ROOT_DIR to choose the installation.
find_package(Python 3.15 EXACT REQUIRED COMPONENTS Development.SABIModule)

function(pyinex_defaults target)
  target_compile_definitions(${target} PRIVATE
    PYINEX_VERSION=L"${PROJECT_VERSION}" UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
  target_compile_options(${target} PRIVATE /W4 /WX /utf-8 /permissive-)
endfunction()

# Code both DLLs below use. Neither links Python.
add_library(pyinex_common STATIC
  src/common/environment.cpp
  src/console/console.cpp
  src/excel/address.cpp
  src/excel/excel.cpp)
target_include_directories(pyinex_common PUBLIC src)
target_link_libraries(pyinex_common PUBLIC excel_sdk_headers)
pyinex_defaults(pyinex_common)

# pyinex.xll, the add-in Excel loads. It does not link Python: it chooses one when first
# needed, then loads that Python and the core (docs/architecture.md, "Choosing a Python").
add_library(pyinex SHARED
  src/loader/addin.cpp
  src/loader/console_window.cpp
  src/loader/python.cpp)
set_target_properties(pyinex PROPERTIES SUFFIX ".xll")
# dwmapi and uxtheme: the console window's dark title bar and scroll bar.
target_link_libraries(pyinex PRIVATE pyinex_common dwmapi uxtheme)
pyinex_defaults(pyinex)

# Code that links Python, shared by the core and the unit tests' extension module, so the
# tests run the same code the add-in does.
add_library(pyinex_python STATIC src/core/call.cpp src/core/convert.cpp)
target_link_libraries(pyinex_python PUBLIC pyinex_common Python::SABIModule)
target_compile_definitions(pyinex_python PUBLIC Py_LIMITED_API=0x030A0000)
pyinex_defaults(pyinex_python)

# pyinex_core.dll, everything that runs Python. It links python3.dll, which the loader has
# already loaded from the chosen installation by the time it loads the core.
add_library(pyinex_core SHARED src/core/core.cpp src/core/module.cpp)
target_link_libraries(pyinex_core PRIVATE pyinex_python)
pyinex_defaults(pyinex_core)
add_dependencies(pyinex pyinex_core)

option(PYINEX_BUILD_TESTS "Build the unit tests' extension module, the fake Excel and its test programs" ON)
if(PYINEX_BUILD_TESTS)
  add_subdirectory(tests/unit)
  add_subdirectory(tests/host)
endif()
