# Microsoft's Excel 2013 XLL SDK: XLCALL.H, the C API's types and constants, the only
# Microsoft file the build uses. Its license forbids redistributing it, so it is not in the
# repository: the build downloads Microsoft's installer, checks it, and unpacks it into
# the build tree. Nothing of the SDK's is compiled into the add-in. See sdk/README.md.
#
# Defines the interface library `excel_sdk_headers` (the XLCALL.H include directory). Set
# PYINEX_EXCEL_SDK_DIR to an unpacked SDK's Excel2013XLLSDK folder to skip the download.

set(PYINEX_EXCEL_SDK_DIR "" CACHE PATH "Unpacked Excel2013XLLSDK folder (empty: download it)")

if(PYINEX_EXCEL_SDK_DIR)
  set(_sdk "${PYINEX_EXCEL_SDK_DIR}")
else()
  set(_url "https://download.microsoft.com/download/D/A/C/DAC9C321-279C-4625-B381-BCB044EEF7B2/excelxllsdk.msi")
  set(_sha256 "1a59a4877f80881d0e8e931f39d5507b3a4eea97fa18008554c425d7df013930")
  set(_root "${CMAKE_BINARY_DIR}/_deps/excelsdk")
  set(_sdk "${_root}/unpacked/2013 Office System Developer Resources/Excel2013XLLSDK")
  if(NOT EXISTS "${_sdk}/INCLUDE/XLCALL.H")
    message(STATUS "Downloading the Excel 2013 XLL SDK from Microsoft")
    file(DOWNLOAD "${_url}" "${_root}/excelxllsdk.msi"
      EXPECTED_HASH SHA256=${_sha256} STATUS _status TLS_VERIFY ON)
    list(GET _status 0 _code)
    if(NOT _code EQUAL 0)
      message(FATAL_ERROR "Could not download ${_url}: ${_status}. "
        "Download it by hand and set PYINEX_EXCEL_SDK_DIR; see sdk/README.md.")
    endif()
    # An administrative install (/a) only unpacks the files; nothing is installed.
    file(TO_NATIVE_PATH "${_root}/excelxllsdk.msi" _msi)
    file(TO_NATIVE_PATH "${_root}/unpacked" _target)
    execute_process(
      COMMAND msiexec /a "${_msi}" /qn "TARGETDIR=${_target}"
      RESULT_VARIABLE _result)
    if(NOT _result EQUAL 0 OR NOT EXISTS "${_sdk}/INCLUDE/XLCALL.H")
      message(FATAL_ERROR "Could not unpack ${_msi} (msiexec exit code ${_result}).")
    endif()
  endif()
endif()

if(NOT EXISTS "${_sdk}/INCLUDE/XLCALL.H")
  message(FATAL_ERROR "No INCLUDE/XLCALL.H under ${_sdk}.")
endif()

add_library(excel_sdk_headers INTERFACE)
target_include_directories(excel_sdk_headers INTERFACE "${_sdk}/INCLUDE")
