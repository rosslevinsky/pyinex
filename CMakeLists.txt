cmake_minimum_required(VERSION 3.26)
project(pyinex VERSION 2026.10.1 LANGUAGES CXX)

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
  src/loader/functions.cpp
  src/loader/python.cpp)
set_target_properties(pyinex PROPERTIES SUFFIX ".xll")
# The pool of numbered exports direct functions take, PyFunc0 to PyFunc999: one line each,
# PYINEX_SLOT(n), in a file generated here, so its size is set in this one place and its
# thousand lines are not Pyinex's own source (docs/architecture.md, "Direct functions").
set(PYINEX_POOL 1000)
math(EXPR last_slot "${PYINEX_POOL} - 1")
set(pool "")
foreach(slot RANGE ${last_slot})
  string(APPEND pool "PYINEX_SLOT(${slot})\n")
endforeach()
file(CONFIGURE OUTPUT ${PROJECT_BINARY_DIR}/generated/loader/pool.inc CONTENT "${pool}")
target_compile_definitions(pyinex PRIVATE PYINEX_POOL=${PYINEX_POOL})
target_include_directories(pyinex PRIVATE ${PROJECT_BINARY_DIR}/generated)
# dwmapi and uxtheme: the console window's dark title bar and scroll bar. oleacc: the queue
# reaches Excel's Application object from a workbook window (docs/architecture.md, "The queue").
target_link_libraries(pyinex PRIVATE pyinex_common dwmapi uxtheme oleacc)
pyinex_defaults(pyinex)

# Compiles src/core/python/<name>.py into target: a header in the build tree,
# generated/core/python/<name>.py.h, holds its text as <name>_py, and is written again
# whenever the file changes (docs/architecture.md, "The core's Python files"). Each file
# belongs to one target, so no two targets write the same header; targets linking it see
# the headers too.
function(pyinex_embed_python target)
  foreach(name IN LISTS ARGN)
    set(source ${PROJECT_SOURCE_DIR}/src/core/python/${name}.py)
    set(header ${PROJECT_BINARY_DIR}/generated/core/python/${name}.py.h)
    add_custom_command(OUTPUT ${header}
      COMMAND ${CMAKE_COMMAND} -DSOURCE=${source} -DOUTPUT=${header} -DNAME=${name}_py
              -P ${PROJECT_SOURCE_DIR}/cmake/EmbedPython.cmake
      DEPENDS ${source} ${PROJECT_SOURCE_DIR}/cmake/EmbedPython.cmake
      COMMENT "Embedding src/core/python/${name}.py"
      VERBATIM)
    target_sources(${target} PRIVATE ${header})
  endforeach()
  target_include_directories(${target} PUBLIC ${PROJECT_BINARY_DIR}/generated)
endfunction()

# Code that links Python, shared by the core and the unit tests' extension module, so the
# tests run the same code the add-in does.
add_library(pyinex_python STATIC src/core/call.cpp src/core/convert.cpp)
target_link_libraries(pyinex_python PUBLIC pyinex_common Python::SABIModule)
target_compile_definitions(pyinex_python PUBLIC Py_LIMITED_API=0x030A0000)
pyinex_embed_python(pyinex_python bootstrap call excel_error funcs)
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
