# Working on Pyinex

Instructions for anyone changing this repository, person or coding agent. What Pyinex is
and how to use it: [README.md](./README.md). How it works and why:
[docs/architecture.md](./docs/architecture.md).

Pyinex's reason to exist is to be small, simple and well documented. Prefer less code and
fewer moving parts to a feature; a change that adds C++ should say what it buys.

## Building and testing

The add-in builds and its tests run only on Windows: Visual Studio 2022 Build Tools (C++),
CMake 3.26 or later, and x64 python.org Pythons 3.10–3.15 with pytest (and openpyxl, for
the examples test, and NumPy, for the real-Excel tests, with pywin32 for the Python that
runs their driver). [docs/testing.md](./docs/testing.md)
covers each step.

On Windows, from the repository root:

```
cmake --preset x64-release
cmake --build --preset x64-release
py -3.14 -m pytest tests -q
```

Before pushing, also run `python scripts/count_lines.py` (the C++ budget) and
`python -m compileall -q .`. CI (`.github/workflows/ci.yml`) builds once and runs
`tests/unit` and `tests/host` under each Python on every push. The real-Excel tests
(`tests/excel`) run at a Windows desktop with Excel, under every Python, through
`py -3.14 scripts\excel_tests.py <pyinex.xll>`, which needs no one beyond starting it.

`tests/unit/test_examples.py` needs only openpyxl and also runs on macOS or Linux.

## Layout

| Path | What it holds |
|---|---|
| `src/loader/` | `pyinex.xll`, the DLL Excel loads: entry points, the four functions, choosing a Python, the console window. Links no Python. |
| `src/core/` | `pyinex_core.dll`: starting Python, PyCall's calls and script cache (`call.cpp`), value conversion (`convert.cpp`), the `pyinex` module (`module.cpp`). Uses only the 3.10 stable ABI. |
| `src/excel/` | The thin wrapper over Microsoft's Excel SDK: values, registration, calling Excel, column letters. |
| `src/console/`, `src/common/` | Writing `Pyinex:` messages; environment variables in both Windows' and the C runtime's copies. |
| `cmake/ExcelSdk.cmake`, `sdk/README.md` | Fetching Microsoft's SDK at configure time. Its files are never committed. |
| `tests/unit/` | Tests of the core's code through `_pyinex_testing.pyd`, built from the same sources; and the examples test. |
| `tests/host/` | Tests of the real add-in in `xllhost.exe`, a fake Excel; `hostdriver.py` drives it. |
| `tests/excel/` | Tests of the add-in in real Excel, run by `scripts/excel_tests.py`; never collected without `--excel`. |
| `scripts/` | `count_lines.py`, `make_examples.py`, `excel_tests.py` (the real-Excel driver), and `check_excel_driver.py` (its check). |
| `examples/` | Example workbooks and their scripts. The `.xlsx` files are generated: change `scripts/make_examples.py` and rerun it, never edit them in Excel. |
| `docs/` | Architecture, value conversion, testing. |

## Conventions

- **Documents first.** `docs/conversion.md` is the single source for value rules and
  `docs/architecture.md` for everything else; code and tests follow them. Change the
  document in the same commit as the behavior, and name the test that checks a rule.
- **Test first** for new behavior: a failing test in `tests/unit` or `tests/host`, then the
  code. A defect fix comes with a test that failed before it.
- **C++:** C++20, MSVC, `/W4 /WX`. The core compiles against the newest Python's headers
  with `Py_LIMITED_API=0x030A0000`: no API newer than 3.10's stable ABI, and nothing
  deprecated. Pyinex's own C++ under `src/` stays under 2,500 counted lines
  (`scripts/count_lines.py`).
- **Never call Python** from `DllMain`, a static object's constructor or destructor, or a
  thread Excel did not call in on. Never finalize Python.
- **Match the surrounding code:** its naming, its comment density, and comments that say
  why rather than what.
- **Writing:** US spelling, in code, comments and documents. Plain sentences, one idea
  each; say what a thing is and why before how.
- **Commits:** a subject line and a body saying why; no trailers.

## Do not

- Commit Microsoft's SDK files (their license forbids it) or build output.
- Edit `examples/*.xlsx` by hand.
- Add a dependency, a build step or a framework without a reason written in
  `docs/architecture.md`.
