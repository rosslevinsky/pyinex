# Microsoft's Excel SDK files

The add-in is compiled against one file from Microsoft's **Excel 2013 XLL SDK**, the
newest Excel C API SDK Microsoft publishes: `INCLUDE/XLCALL.H`, the C API's types and
constants. The add-in uses only those definitions. It does not compile the SDK's
`SRC/XLCALL.CPP`, whose `Excel12v` it replaces with a function of its own
([docs/architecture.md](../docs/architecture.md#calling-excel)), so the built add-in
contains no code of Microsoft's.

`XLCALL.H` is **not in this repository**. The SDK's license (`ExcelSDK_eula.rtf` inside the
installer) lets you use the SDK's files to develop your programs, but forbids distributing
them or changing them. So the build fetches them, in `cmake/ExcelSdk.cmake`:

| | |
|---|---|
| Download page | <https://www.microsoft.com/en-us/download/details.aspx?id=35567> |
| File | <https://download.microsoft.com/download/D/A/C/DAC9C321-279C-4625-B381-BCB044EEF7B2/excelxllsdk.msi> |
| SHA-256 | `1a59a4877f80881d0e8e931f39d5507b3a4eea97fa18008554c425d7df013930` |
| Size | 4,820,992 bytes |

At configure time CMake downloads the installer into `build/<preset>/_deps/excelsdk/`,
refuses it if the SHA-256 differs, and unpacks it there with `msiexec /a` (an
administrative install, which copies the files out and installs nothing). The SDK's
files never leave the build folder.

**Offline, or if the download moves:** download the installer from the page above, unpack
it with `msiexec /a excelxllsdk.msi /qn TARGETDIR=C:\excelsdk`, and configure with
`-DPYINEX_EXCEL_SDK_DIR="C:\excelsdk\2013 Office System Developer Resources\Excel2013XLLSDK"`.
