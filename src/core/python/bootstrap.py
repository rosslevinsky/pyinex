# Run once when Python starts (src/core/core.cpp, pyinex_start). Its namespace holds the
# helpers the core calls.

import os, sys

# Output goes to Pyinex's console window through a pipe, to a captured stream, or to a
# terminal: in UTF-8 whatever the code page, and a line at a time, so print() from a thread
# a script started shows without waiting for a buffer to fill.
for stream in (sys.stdout, sys.stderr):
    try:
        stream.reconfigure(encoding='utf-8', errors='backslashreplace', line_buffering=True)
    except Exception:
        pass

def fix_up(base_executable, host_executable):
    # Python sets sys.executable from __PYVENV_LAUNCHER__ but takes these two from the
    # process's own .exe (Excel): make them what the chosen python.exe would have. Excel's
    # folder stays on sys.path if it is one of Python's own or PYTHONPATH names it.
    sys._base_executable = base_executable
    def key(p):
        return os.path.normcase(os.path.abspath(p))
    host = key(os.path.dirname(host_executable))
    keep = {key(sys.prefix), key(sys.base_prefix)}
    keep.update(key(p) for p in os.environ.get('PYTHONPATH', '').split(os.pathsep) if p)
    if host not in keep:
        sys.path[:] = [p for p in sys.path if not p or key(p) != host]

def info(item):
    # PyInfo's items that only Python knows.
    if item == 'version':
        return sys.version
    if item == 'prefix':
        return sys.prefix
    return sys.prefix != sys.base_prefix  # 'venv'
