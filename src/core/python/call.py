# The Python half of PyCall (src/core/call.cpp). Every exception, SystemExit and
# KeyboardInterrupt included, is caught here and printed with the traceback module: never
# with PyErr_Print, which exits the process on SystemExit.

import importlib.machinery, importlib.util, os, re, sys, traceback, zlib

FAILED = object()  # call's result when it has already reported why
_scripts = {}      # os.path.normcase(path) -> (module, st_mtime_ns, st_size)
_folders = set()   # folders put on sys.path, normcased
_module_names = {} # os.path.normcase(path) -> its module's name in sys.modules, each its own

def _say(text):
    try:
        print('Pyinex: ' + text, file=sys.stderr)
    except BaseException:
        pass

def _traceback():
    # The exception being handled, without the frames of Pyinex's own files, compiled under
    # names such as <pyinex/call.py>, which the user did not write: wherever they are, as
    # pyinex.ExcelError's own frame after the user's, and in chained exceptions too.
    try:
        # compact, as print_exception's: an exception both a cause and a context is kept.
        shown = traceback.TracebackException(*sys.exc_info(), compact=True)
        _drop_own_frames(shown)
        print(''.join(shown.format()), file=sys.stderr, end='')
    except BaseException:
        pass

def _drop_own_frames(shown):
    # Through causes, contexts and an exception group's exceptions (from 3.11), each once,
    # with no recursion, however long the chain.
    waiting, done = [shown], set()
    while waiting:
        shown = waiting.pop()
        if shown is None or id(shown) in done:
            continue
        done.add(id(shown))
        shown.stack = traceback.StackSummary.from_list(
            [f for f in shown.stack if not f.filename.startswith('<pyinex/')])
        waiting += [shown.__cause__, shown.__context__, *(getattr(shown, 'exceptions', None) or ())]

def _flush():
    # So output reaches the console after each call, not whenever a buffer fills.
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None:
                stream.flush()
        except BaseException:
            pass

def _forget(key):
    entry = _scripts.pop(key, None)
    if entry is not None and sys.modules.get(entry[0].__name__) is entry[0]:
        del sys.modules[entry[0].__name__]

def _loaded(path, module):
    # Called after a script has run again (funcs.py replaces it, to follow the scripts PyLoad
    # loaded).
    pass

def _load(path, reload_on_change, follow=True):
    # The script's module, loaded or reloaded as needed; None, having said why, if the
    # file is not there. follow: tell _loaded when the script has run again.
    key = os.path.normcase(path)
    entry = _scripts.get(key)
    if entry is not None and not reload_on_change:
        return entry[0]
    try:
        st = os.stat(path)
    except OSError:
        st = None
    if st is None or not os.path.isfile(path):
        _forget(key)
        _say(f'cannot find the script {path}')
        return None
    if entry is not None:
        if (st.st_mtime_ns, st.st_size) == entry[1:]:
            return entry[0]
        _forget(key)  # A failed reload must not leave the old version running.
    folder = os.path.dirname(path)
    if os.path.normcase(folder) not in _folders:
        sys.path.insert(0, folder)  # So helpers beside the script import.
        _folders.add(os.path.normcase(folder))
    name = _module_names.get(key)
    if name is None:  # Named once, by its path, and never as another path's module is named.
        stem = re.sub(r'\W', '_', os.path.splitext(os.path.basename(path))[0])
        name = base = 'pyinex_%s_%08x' % (stem, zlib.crc32(key.encode('utf-8', 'surrogatepass')))
        taken = set(_module_names.values())
        for n in range(2, 1000000):
            if name not in taken:
                break
            name = f'{base}_{n}'
        _module_names[key] = name
    # The loader named, since importlib would choose one by the suffix, case-sensitively,
    # and find none for model.PY.
    loader = importlib.machinery.SourceFileLoader(name, path)
    module = importlib.util.module_from_spec(importlib.util.spec_from_file_location(name, path, loader=loader))
    # Compiled here from the source, never from a cached .pyc, whose timestamp is
    # only to the second.
    with open(path, 'rb') as f:
        code = compile(f.read(), path, 'exec', dont_inherit=True)
    sys.modules[name] = module
    try:
        exec(code, module.__dict__)
    except BaseException:
        if sys.modules.get(name) is module:
            del sys.modules[name]
        raise
    _scripts[key] = (module, st.st_mtime_ns, st.st_size)
    if follow:
        _loaded(path, module)
    return module

def _path(file, folder):
    # The script's full path; None, having said why, for a relative path in a workbook with no
    # folder. folder is the workbook's, for a relative file, and None for an absolute one: empty
    # for a workbook never saved, or ? and the reason there is none (excel::caller_folder).
    if folder is not None:
        if not folder or folder.startswith('?'):
            why = folder[1:] if folder else 'this workbook has not been saved'
            _say(f'cannot find the script {file}: a relative path is relative to the '
                 f'workbook\'s folder, and {why}; give the script\'s full path instead')
            return None
        file = os.path.join(folder, file)
    return os.path.abspath(file)

def _function(path, name, reload_on_change):
    # The function name in the script, loaded or reloaded as needed; FAILED, having said why.
    try:
        module = _load(path, reload_on_change)
    except BaseException:
        if not _stopped():  # Excel discards the result: nothing to say.
            _say(f'the script {path} failed to load:')
            _traceback()
        return FAILED
    if module is None:
        return FAILED
    function = vars(module).get(name, FAILED)
    if function is FAILED:
        _say(f'the script {path} has no function {name}')
    elif not callable(function):
        _say(f'{name} in the script {path} is not a function')
        return FAILED
    return function

def call(file, name, args, reload_on_change, folder):
    try:
        path = _path(file, folder)
        function = FAILED if path is None else _function(path, name, reload_on_change)
        return FAILED if function is FAILED else function(*args)
    except BaseException:
        if not _stopped():  # Excel discards the result: nothing to say.
            _say(f'{name} in the script {file} raised an exception:')
            _traceback()
        return FAILED

def _stopped():
    # Whether Excel stopped the call running (funcs.py has the same; this one serves when
    # call.py runs alone, as in tests).
    module = sys.modules.get('pyinex')
    try:
        return bool(module is not None and module._stopped())
    except BaseException:
        return False
