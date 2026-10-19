# The Python half of direct functions (docs/architecture.md, "Direct functions", "PyLoad",
# "Reloading a script" and "Names"): the @pyinex.func decorator, the functions a loaded
# script declares, read from their signatures and docstrings, which are registered, what a
# reload changed, the rules for names, how a call's arguments are placed and converted
# ("Signatures"; docs/conversion.md, "Type annotations"), and pyinex.Range ("pyinex.Range").
# Run in call.py's namespace (src/core/call.cpp, start_calls), so _load, _function, _path,
# call, _say and _traceback are call.py's, and _loaded here replaces call.py's.

import functools, inspect, itertools, math, os, re, sys, types, typing

MAX_ARGUMENTS = 30  # Excel's positions after PyCall's file and function; the same here.
LIMIT = 255         # Excel cuts or refuses longer names, argument names and help texts,
                    # counted as Excel counts text: in UTF-16 code units (_units).
# Pyinex's own names: its worksheet functions, and the commands it registers.
OWN = {n.casefold() for n in ('PyCall', 'PyConsole', 'PyReloadOnChange', 'PyInfo', 'PyLoad', 'PyRepair', 'xlAutoRemove')}
_live = {}          # key -> (path, attribute, name, row): registered, or waiting to be
_names = {}         # name.casefold() -> key, for each in _live
_gone = {}          # key -> why a call of a function no longer in _live gives #VALUE!
_scripts_named = {} # os.path.normcase(path) -> the script as PyLoad named it, for each it loaded
_pending = []       # what reloads outside PyLoad changed, for the loader: [script] + row
_offered = None     # (script, rows) load_functions last gave, until they reach the loader
_keys = itertools.count()
_calls = {}         # key -> (function, then _plan's parameters and converters), as last called
MISSING = object()  # An argument left out before the last one given (src/core/call.cpp).

def func(function=None, *, name=None, category=None, help=None):
    """Makes a function a worksheet function under its own name, when its script loads.

    @pyinex.func marks it; @pyinex.func(name=..., category=..., help=...) chooses the name
    Excel shows, its category in Insert Function, and its help text, which otherwise
    come from the function's name, "Pyinex" and its docstring's first paragraph. The
    function is returned unchanged.
    """
    # The module applying the decorator: only a script's own functions are its own.
    caller = sys._getframe(1).f_globals
    module = caller.get('__name__')
    def mark(f):
        try:
            f.__pyinex_func__ = {'module': module, 'name': name, 'category': category, 'help': help}
        except (AttributeError, TypeError):  # A built-in function, which takes no attribute.
            what = 'a built-in function' if isinstance(f, types.BuiltinFunctionType) else 'it'
            _say(f'{getattr(f, "__name__", f)} in the script {caller.get("__file__")} is not registered: '
                 f'{what} cannot be marked; wrap it in a function of your own')
        return f
    return mark if function is None else mark(function)

def _docstring(f, parameters):
    # Its first paragraph, and the help of each parameter from an "Args:" section.
    text = inspect.getdoc(f) or ''
    first = re.split(r'\n\s*\n', text.strip(), maxsplit=1)[0]
    found, current, header, entries = {}, None, None, None
    for line in text.splitlines():
        indent = len(line) - len(line.lstrip())
        if header is None:
            if re.fullmatch(r'\s*(Args|Arguments):\s*', line):
                header = indent
            continue
        if not line.strip():
            continue
        if indent <= header:
            break
        entry = re.fullmatch(r'\s*\*{0,2}(\w+)\s*(?:\([^)]*\))?\s*:\s*(.*)', line)
        if entry and indent <= (entries if entries is not None else indent):
            entries, current = indent, entry.group(1)
            found[current] = entry.group(2).strip()
        elif current is not None:  # A continuation, indented further.
            found[current] = (found[current] + ' ' + line.strip()).strip()
    return ' '.join(first.split()), [found.get(p, '') for p in parameters]

def _units(text):
    # Text's length as Excel counts it: a character beyond U+FFFF takes two UTF-16 code units.
    return len(text) + sum(ord(c) > 0xFFFF for c in text)

def _cut(text, limit=LIMIT):
    # Text cut to at most limit UTF-16 code units, never within a character.
    if _units(text) <= limit:
        return text
    used = 0
    for i, c in enumerate(text):
        used += 2 if ord(c) > 0xFFFF else 1
        if used > limit:
            return text[:i]
    return text

def _name(f, attribute):
    return f.__pyinex_func__['name'] or getattr(f, '__name__', attribute)

class _NotRegistered(Exception):
    pass  # Why a function cannot be registered.

def _text(e):
    # An exception as text, even one whose own text fails.
    try:
        return str(e)
    except BaseException:
        return type(e).__name__

def _description(f, attribute, path, namespace):
    # What Excel is told of the function, the row the loader registers without its key, and a
    # note for each annotation left alone, for the console should the row be new. Raises
    # _NotRegistered, saying why, for a function that cannot be registered. namespace: the
    # script's, for annotations.
    options = vars(f)['__pyinex_func__']
    try:
        fixed, var, converters, notes = _plan(f, namespace)
    except BaseException as e:  # SystemExit from an annotation's code too: this function only.
        raise _NotRegistered(f'its signature cannot be read ({_text(e)})') from None
    if len(fixed) + (var is not None) > MAX_ARGUMENTS:
        raise _NotRegistered(f'it has {len(fixed)} parameters{"" if var is None else " and *" + var.name}, '
                             f'and Excel passes at most {MAX_ARGUMENTS}')
    name = str(_name(f, attribute))
    spread = 0 if var is None else MAX_ARGUMENTS - len(fixed)  # *args' positions.
    # Excel's type code for each position: U for a reference, Q for values.
    codes = ''.join('U' if converters[p.name] is _reference else 'Q' for p in fixed + [var] * spread)
    names = [p.name for p in fixed]
    help_text, argument_help = _docstring(f, names + [var.name] * spread if spread else names)
    names += [f'{var.name}{i}' for i in range(1, spread + 1)]
    if options['help'] is not None:
        help_text = str(options['help'])
    category = 'Pyinex' if options['category'] is None else str(options['category'])
    text = ','.join(names)
    if _units(text) > LIMIT:  # Cut after the last whole name that fits; none, if none does.
        text = text[:max(text.rfind(',', 0, len(_cut(text, LIMIT + 1))), 0)]
    # The count, or with a reference among them the type codes, whose length is the count.
    return ([name, codes if 'U' in codes else len(names), text, _cut(category), _cut(help_text)]
            + [_cut(h) for h in argument_help]), notes

def _refusal(name):
    # Why a name cannot be registered, or None. Excel refuses these itself, saying only that
    # it failed; Pyinex's own names it would take, hiding Pyinex's functions.
    if not re.fullmatch(r'[^\W\d][\w.]*', name):
        return ('Excel cannot take that name: a name is letters, digits, underscores and periods, '
                'starting with a letter or underscore')
    if _units(name) > LIMIT:
        return f'Excel cannot take that name: a name is at most {LIMIT} characters, a character beyond U+FFFF counting as two'
    cell = re.fullmatch(r'([a-z]{1,3})(\d+)', name, re.IGNORECASE)
    column = cell and sum(26 ** i * (ord(c) - 64) for i, c in enumerate(reversed(cell[1].upper())))
    r1c1 = re.fullmatch(r'r\d*(c\d*)?|c\d*', name, re.IGNORECASE)
    if (cell and column <= 16384 and 1 <= int(cell[2]) <= 1048576) or r1c1:
        return 'Excel cannot take that name: it is a cell reference'
    if name.casefold() in ('true', 'false'):
        return 'Excel cannot take that name: it is a logical value'
    if name.casefold() in OWN or name.casefold().startswith('pyinexqueue'):
        return "it is one of Pyinex's own names"
    return None

def _sync(path, module):
    # Brings what is registered from the script at path in line with its module: the rows for
    # the loader, removals first, so a name is cleared before it is registered again; then,
    # for each of the script's functions in order, [-1] for one registered already and
    # unchanged, or a new function's row.
    here = os.path.normcase(path)
    mine = {name: key for name, key in _names.items() if os.path.normcase(_live[key][0]) == here}
    for key in mine.values():  # Read again at the next call: an object marked again, such as
        _calls.pop(key, None)  # a partial from a helper, is the same object changed.
    rows, removals, seen, done = [], [], set(), set()
    for attribute, f in list(vars(module).items()):
        try:
            # Its own mark: not one a class it derives from, or the class of an instance, has.
            options = vars(f).get('__pyinex_func__')
            if not isinstance(options, dict) or options.get('module') != module.__name__ or id(f) in seen:
                continue
        except BaseException:
            continue
        seen.add(id(f))
        name = folded = None
        try:  # One function's failure refuses only that function, SystemExit too.
            name = str(_name(f, attribute))
            folded, why = name.casefold(), _refusal(name)
            if why is None and folded in done:
                why = f'the script declares {_live[_names[folded]][2]} already'
            elif why is None and folded in _names and folded not in mine:
                why = f'the script {_live[_names[folded]][0]} registered that name first'
            if why is not None:
                _say(f'{name} in the script {path} is not registered: {why}')
                continue
            description, notes = _description(f, attribute, path, vars(module))
        except BaseException as e:
            why = _text(e) if isinstance(e, _NotRegistered) else f'{type(e).__name__}: {_text(e)}'
            _say(f'{attribute if name is None else name} in the script {path} is not registered: {why}')
            key = mine.pop(folded, None) if folded is not None else None
            if key is not None:  # Registered before: its calls say why, not that it was removed.
                removals.append([key])
                _drop(key, f'{_live[key][2]} is not registered: {why}')
            continue
        done.add(folded)
        key = mine.pop(folded, None)
        if key is not None and _live[key][3] == description:
            _live[key] = (path, attribute, name, description)  # Its body may have changed.
            rows.append([-1])
            continue
        if key is not None:  # Changed: its old registration's calls must not run the new code.
            removals.append([key])
            _drop(key, f'{_live[key][2]} is being registered again, since its script changed; '
                       'until then it gives #VALUE!')
        for note in notes:  # Said once, as the function is registered, not at each load.
            _say(f'{name} in the script {path}: {note}')
        key = next(_keys)
        _live[key] = (path, attribute, name, description)
        _names[folded] = key
        rows.append([key] + description)
    for name, key in mine.items():  # Removed or renamed.
        removals.append([key])
        _drop(key, f'{_live[key][2]} was removed from the script {path}; it gives #VALUE! until it is unregistered')
    return removals + rows

def _drop(key, why=None):
    path, attribute, name, description = _live.pop(key)
    _calls.pop(key, None)
    if _names.get(name.casefold()) == key:
        del _names[name.casefold()]
    _gone[key] = why or f'{name} is not registered'

def _loaded(path, module):
    # A script that has run again: if PyLoad loaded it, what changed is the loader's to do.
    script = _scripts_named.get(os.path.normcase(path))
    if script is not None:
        _pending.extend([script] + row for row in _sync(path, module) if row[0] != -1)

def load_functions(file, folder=None, reload_on_change=True):
    # Loads the script as PyCall does and returns the rows for the loader (_sync); FAILED,
    # having said why, if the script did not load: a script that raised registers nothing.
    # folder is the workbook's, for a relative file, as for call.
    global _offered
    _offered = None
    path = _path(file, folder)
    if path is None:
        return FAILED
    try:
        module = _load(path, reload_on_change, follow=False)  # Its rows are returned here.
    except BaseException:
        _say(f'the script {path} failed to load:')
        _traceback()
        return FAILED
    if module is None:
        return FAILED
    script = _scripts_named.setdefault(os.path.normcase(path), file)
    rows = _sync(path, module)
    _offered = (script, rows)
    return rows

def _unoffer():
    # The rows load_functions last gave did not reach the loader, as when memory ran out
    # converting them: each new function's name is freed, so the next load offers it again,
    # and each removal waits for the loader's next call of changes.
    global _offered
    if _offered is not None:
        script, rows = _offered
        _offered = None
        for row in rows:
            if len(row) == 1 and row[0] >= 0:
                _pending.append([script, row[0]])
            elif row[0] != -1:
                registered(row[0], False)

def changes():
    # What reloads changed since the loader last asked: [script] + a row, for each.
    rows = _pending[:]
    del _pending[:]
    return rows

def _restore_changes(rows):
    # What changes gave did not reach the loader: it waits for the next call, ahead of what
    # came since.
    _pending[:0] = rows

def registered(key, ok):
    # The loader's word on a function load_functions or changes gave: one that did not
    # register, or was dropped from the queue, frees its name; and on one removed from Excel,
    # which no cell can call any more, so why it gave #VALUE! is forgotten.
    if not ok:
        if key in _live:
            _drop(key)
        _gone.pop(key, None)

def call_function(key, args, reload_on_change):
    # A registered function's call: PyCall's, with its script and function fixed. A script
    # that changed is reloaded first, and a function the reload changed or removed is not
    # called: its arguments were read for the registration Excel still has.
    entry = _live.get(key)
    try:
        module = None if entry is None else _load(entry[0], reload_on_change)
    except BaseException:
        if not _stopped():  # Excel discards the result: nothing to say.
            _say(f'the script {entry[0]} failed to load:')
            _traceback()
        return FAILED
    if key not in _live:
        _say(_gone.get(key, 'that function is not registered'))
        return FAILED
    path, attribute, name = _live[key][:3]
    if module is None:
        return FAILED
    try:
        function = _function(path, attribute, False)  # Loaded just now.
        if function is FAILED:
            return FAILED
        plan = _calls.get(key)
        if plan is None or plan[0] is not function:  # Read again only when the function changed.
            plan = _calls[key] = (function,) + _plan(function, vars(module))[:3]
        positional, keywords = _arguments(name, *plan[1:], args)
    except _Refused as e:
        _say(str(e))
        return FAILED
    except BaseException:
        _say(f'{name}: Python failed while reading its arguments:')
        _traceback()
        return FAILED
    try:
        result = function(*positional, **keywords)
    except BaseException:
        if _stopped():  # Excel discards the result: nothing to say (src/core/call.cpp).
            return FAILED
        _say(f'{attribute} in the script {path} raised an exception:')
        _traceback()
        return FAILED
    return FAILED if _stopped() else result

# Signatures and annotations (docs/architecture.md, "Signatures"; docs/conversion.md, "Type
# annotations (direct functions)").

class _Refused(Exception):
    pass  # A value its annotation does not take, saying why.

_TAKES = {int: 'a whole number', float: 'a number', str: 'text', bool: 'TRUE or FALSE'}
_EMPTY = {int: 0, float: 0.0, str: '', bool: False}  # What Excel makes of an empty cell.

def _home(f, namespace):
    # The names an annotation of f is evaluated among: those of the module that defined the
    # function underneath (a partial's, a wrapper's, a class's), else the script's.
    for _ in range(100):
        if isinstance(f, functools.partial):
            f = f.func
        elif hasattr(f, '__wrapped__'):
            f = f.__wrapped__
        elif isinstance(f, types.MethodType):
            f = f.__func__
        elif isinstance(f, type):
            # The method that gives a class its signature, in the class that defined it, which
            # may be in another module: the first __new__ or __init__ along its bases.
            method = next((vars(b)[m] for b in f.__mro__ if b is not object
                           for m in ('__new__', '__init__') if m in vars(b)), None)
            if method is None:
                break
            f = method.__func__ if isinstance(method, (staticmethod, classmethod)) else method
        elif not isinstance(f, (types.FunctionType, types.BuiltinFunctionType)) and callable(f):
            call = getattr(type(f), '__call__', None)  # An instance: its class's __call__.
            if not isinstance(call, types.FunctionType):
                break
            f = call
        else:
            break
    if isinstance(getattr(f, '__globals__', None), dict):
        return f.__globals__
    module = sys.modules.get(getattr(f, '__module__', None) or '')
    return vars(module) if module is not None else namespace

def _plan(f, namespace):
    # f's parameters in Excel's positions, positional then keyword-only; its *args parameter,
    # or None; each parameter's converter (_converter) by name; and a note for each
    # annotation left alone. namespace: the script's, used only when f's own module is unknown.
    namespace = _home(f, namespace)
    unevaluated = None
    if sys.version_info >= (3, 14):
        # From 3.14 annotations are evaluated only when asked for: reading them as forward
        # references keeps a name imported only for type checkers from refusing the function.
        import annotationlib
        try:
            parameters = inspect.signature(f, annotation_format=annotationlib.Format.FORWARDREF).parameters.values()
        except Exception as e:
            # An annotation whose code raises, even one Pyinex ignores, as a return annotation:
            # Python evaluates a function's annotations together or not at all, so a function
            # is read without them, and its arguments arrive unconverted. The function read is
            # the one whose signature inspect gives, the innermost a wrapper names (__wrapped__).
            inner = inspect.unwrap(f, stop=lambda g: hasattr(g, '__signature__'))
            if type(inner) is not types.FunctionType:
                raise
            bare = types.FunctionType(inner.__code__, inner.__globals__, inner.__name__, inner.__defaults__,
                                      inner.__closure__)
            bare.__kwdefaults__ = inner.__kwdefaults__
            parameters = inspect.signature(bare).parameters.values()
            unevaluated = e
    else:
        parameters = inspect.signature(f).parameters.values()
    fixed = [p for p in parameters if p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)]
    fixed += [p for p in parameters if p.kind is p.KEYWORD_ONLY]
    var = next((p for p in parameters if p.kind is p.VAR_POSITIONAL), None)
    converters, notes = {}, []
    if unevaluated is not None:
        notes.append(f'its annotations cannot be evaluated ({type(unevaluated).__name__}: {unevaluated}), '
                     'so its arguments arrive unconverted')
    for p in fixed + ([] if var is None else [var]):
        converters[p.name] = None
        try:
            converters[p.name] = _converter(p.annotation, namespace)
        except _Unevaluated as e:
            notes.append(f'the annotation of {p.name}, {e}, cannot be evaluated '
                         f'({type(e.__cause__).__name__}: {e.__cause__}), so {p.name} arrives unconverted')
        except Exception:
            notes.append(f'Pyinex does not convert the annotation of {p.name}, {p.annotation!r}, '
                         f'so {p.name} arrives unconverted')
    return fixed, var, converters, notes

class _Unevaluated(Exception):
    pass  # An annotation written as text that cannot be evaluated; its text, as repr.

def _evaluated(a, namespace):
    # Text, or a forward reference, evaluated until it is neither: "int" under from __future__
    # import annotations is text within text. A few times at most, since text can name itself.
    for _ in range(5):
        if not (isinstance(a, str) or hasattr(a, '__forward_arg__')):
            break
        text = a if isinstance(a, str) else a.__forward_arg__
        try:
            a = eval(text, namespace)
        except BaseException as e:  # SystemExit too: the annotation is left alone, nothing escapes.
            raise _Unevaluated(repr(text)) from e
    return a

def _converter(a, namespace, element=False):
    # What converts a value for the annotation a: a function, or None to leave the value as it
    # arrives. element: a list of rows' element, only a single value's annotation. Raises for
    # an annotation Pyinex does not convert.
    a = _evaluated(a, namespace)
    if a is inspect.Parameter.empty or a is object or a is typing.Any:
        return None
    args = typing.get_args(a)
    if typing.get_origin(a) in (typing.Union, types.UnionType) and len(args) == 2 and type(None) in args:
        inner = _converter(args[args[0] is type(None)], namespace, element)
        if inner is None or inner is _reference:  # A reference is never None.
            return inner
        return lambda v: None if v is None else inner(v)
    kind = next((t for t in _TAKES if a is t), None)
    if kind is not None:
        def scalar(v):
            if v is None:
                return _EMPTY[kind]
            if type(v) is (float if kind is int else kind) and (kind is not int or v.is_integer()):
                return kind(v)
            raise _Refused(f'{kind.__name__} takes {_TAKES[kind]}, not {_shown(v)}')
        return scalar
    if element:
        raise TypeError(a)
    if a is Range:
        return _reference
    if a is list or typing.get_origin(a) is list:
        rows = _evaluated(args[0], namespace) if args else list
        if not (rows is list or typing.get_origin(rows) is list):
            raise TypeError(a)  # list[float]: not rows.
        cell = typing.get_args(rows)
        cell = _converter(cell[0], namespace, True) if cell else None
        return lambda v: _cells(v, cell)
    if getattr(a, '__module__', None) == 'numpy' and getattr(a, '__name__', None) == 'ndarray':
        numpy = sys.modules['numpy']  # Imported by the script that named it.
        def number(x):
            if x is None or type(x) is float:
                return math.nan if x is None else x
            raise _Refused(f'numpy.ndarray takes numbers and empty cells, not {_shown(x)}')
        return lambda v: numpy.array(_cells(v, number), dtype=float)
    raise TypeError(a)

def _cells(v, convert):
    # A value as a list of rows, each element converted by convert, if not None.
    rows = v if isinstance(v, tuple) else ((v,),)
    out = []
    for r, row in enumerate(rows, 1):
        out.append([])
        for c, x in enumerate(row, 1):
            try:
                out[-1].append(x if convert is None else convert(x))
            except _Refused as e:
                raise _Refused(f'row {r}, column {c}: {e}') from None
    return out

def _shown(v):
    # A value as the console names it.
    if v is None:
        return 'an empty cell'
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, str):
        return f'the text {v!r}'
    if isinstance(v, float):
        return repr(v)
    if isinstance(v, tuple):
        return f'a range of {len(v)} × {len(v[0]) if v else 0} cells'
    return f'the error {v}'

def _arguments(name, fixed, var, converters, args):
    # The positional and keyword arguments for a call with Excel's args, its arguments left out
    # at the end dropped and any before the last given MISSING. Raises _Refused, saying why,
    # for a value its annotation does not take.
    positional, keywords = [], {}
    def converted(i, p, label):
        try:
            return args[i] if converters[p.name] is None else converters[p.name](args[i])
        except _Refused as e:
            raise _Refused(f'{name}: argument {i + 1} ({label}): {e}') from None
    for i, p in enumerate(fixed[:len(args)]):  # After the last given, Python applies defaults.
        if args[i] is not MISSING:
            value = converted(i, p, p.name)
        elif p.default is not p.empty:
            if p.kind is p.KEYWORD_ONLY:
                continue
            value = p.default
        else:
            value = None
        if p.kind is p.KEYWORD_ONLY:
            keywords[p.name] = value
        else:
            positional.append(value)
    for i in range(len(fixed), len(args) if var is not None else 0):
        positional.append(None if args[i] is MISSING else converted(i, var, f'{var.name}{i - len(fixed) + 1}'))
    return positional, keywords

# pyinex.Range (docs/architecture.md, "pyinex.Range"; docs/conversion.md, "References").

ROWS, COLUMNS = 1048576, 16384  # Excel's grid.

def _reference(v):
    # The converter for a parameter annotated pyinex.Range: Excel passed a reference, which
    # the core made a Range, or a value, which is refused.
    if isinstance(v, Range):
        return v
    raise _Refused(f'pyinex.Range takes a reference, not {_shown(v)}')

def _column(n):
    # Column n, counted from 1, in letters: A to Z, then AA; base 26 with no zero digit.
    letters = ''
    while n:
        n, d = divmod(n - 1, 26)
        letters = chr(65 + d) + letters
    return letters

def _cell(row, column, r1c1):
    return f'R{row + 1}C{column + 1}' if r1c1 else f'{_column(column + 1)}{row + 1}'

def _last_used(count, first, last):
    # The last of the rows (or columns) first to last that holds anything, or first if none
    # does, by halving: count(p) is how many cells hold anything from p to last.
    low, high = first + 1, last + 1  # The first from which nothing is held is in low..high.
    while low < high:
        middle = (low + high) // 2
        if count(middle):
            low = middle + 1
        else:
            high = middle
    return max(low - 1, first)

class Range:
    """A reference to cells on one sheet, as a direct function's parameter annotated
    pyinex.Range receives it: its workbook, sheet, address and size, and, during a call,
    its cells' values. It reads; Excel lets a worksheet function change no cell."""

    __slots__ = ('_sheet', '_id', '_areas')

    def __init__(self, sheet, id, areas):
        # Made by Pyinex from what Excel passed: sheet as Excel names it, [Book.xlsx]Sheet1;
        # Excel's ID for it; and each area as (first row, last row, first column, last
        # column), counted from 0.
        self._sheet, self._id, self._areas = sheet, id, tuple(tuple(a) for a in areas)

    @property
    def workbook(self):
        return self._sheet[1:].partition(']')[0]

    @property
    def sheet(self):
        return self._sheet[1:].partition(']')[2]

    @property
    def areas(self):
        return tuple(Range(self._sheet, self._id, [a]) for a in self._areas)

    @property
    def rows(self):
        r1, r2, c1, c2 = self._area()
        return r2 - r1 + 1

    @property
    def columns(self):
        r1, r2, c1, c2 = self._area()
        return c2 - c1 + 1

    def address(self, r1c1=False, full=False):
        """B2:C3, or R2C2:R3C3 with r1c1=True; with full=True the workbook and sheet come
        first, as in [Book.xlsx]Sheet1!B2:C3. Several areas are joined by commas."""
        def one(r1, r2, c1, c2):
            text = _cell(r1, c1, r1c1)
            return text if (r1, c1) == (r2, c2) else text + ':' + _cell(r2, c2, r1c1)
        text = ','.join(one(*a) for a in self._areas)
        return f'{self._sheet}!{text}' if full else text

    def values(self):
        """The cells' values, converted as an argument's are: one cell a plain value, more a
        tuple of rows, an empty cell None. A whole column or row gives its used part."""
        # Excel may give a deleted sheet's ID, or a closed workbook's sheet's, to another
        # sheet: the sheet is found by its name again, so a range kept from another call
        # never reads another sheet.
        if sys.modules['pyinex']._sheet_id(self._sheet) != self._id:
            raise RuntimeError(f'pyinex.Range: Excel no longer has the sheet {self._sheet} open; it was deleted '
                               'or renamed, or its workbook closed')
        r1, r2, c1, c2 = self._used()
        return sys.modules['pyinex']._range_values(self._id, r1, r2, c1, c2)

    def empty(self):
        """Whether each cell is empty, in the shape values() gives: empty text is not."""
        v = self.values()
        return tuple(tuple(x is None for x in row) for row in v) if isinstance(v, tuple) else v is None

    def __repr__(self):
        return f'pyinex.Range({self.address(full=True)!r})'

    def _area(self):
        if len(self._areas) != 1:
            raise ValueError(f'{self.address()} has {len(self._areas)} areas: use its areas')
        return self._areas[0]

    def _used(self):
        # The area, cut to its used part where it spans every row or every column, found by
        # Excel's counting rather than by reading a million cells.
        r1, r2, c1, c2 = self._area()
        count = sys.modules['pyinex']._range_count
        whole = (r1 == 0 and r2 == ROWS - 1) or (c1 == 0 and c2 == COLUMNS - 1)
        if whole and not count(self._id, r1, r2, c1, c2):  # Nothing used: its first cell.
            return r1, r1, c1, c1
        if r1 == 0 and r2 == ROWS - 1:
            r2 = _last_used(lambda r: count(self._id, r, r2, c1, c2), r1, r2)
        if c1 == 0 and c2 == COLUMNS - 1:
            c2 = _last_used(lambda c: count(self._id, r1, r2, c, c2), c1, c2)
        return r1, r2, c1, c2

Range.__module__ = 'pyinex'
