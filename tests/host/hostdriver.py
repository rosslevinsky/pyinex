"""Drive xllhost.exe, the fake Excel, from tests.

    with Host(xllhost, xll) as host:
        host.open()
        assert host.call("PyInfo", "status") == "running"

Values go to and come back from Excel as Python values: ``float`` for numbers, ``str``,
``bool``, ``Err`` for Excel errors, ``MISSING`` for an argument left out, ``None`` for an
empty value, and a tuple of row tuples for an array. ``Int`` sends Excel's rarely used
xltypeInt, and ``Ref`` and ``SRef`` the references Excel passes a parameter that takes one,
which read the cells ``set_cells`` puts in the host's sheets.

After every request the driver asks the host for faults — breaks of Excel's memory rules
— and raises ``HostFault`` if there are any, unless the request said it expects them.
The add-in's own output (``print()``, tracebacks) is collected separately, in ``stdout``
and ``stderr``. A request with no answer within ``timeout`` seconds (an add-in that hangs)
stops the host and raises ``HostError``.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from dataclasses import dataclass
from pathlib import Path

if sys.platform == "win32":
    import msvcrt

# Excel's error codes (xlerr* in XLCALL.H).
NULL, DIV0, VALUE, REF, NAME, NUM, NA, GETTING_DATA = 0, 7, 15, 23, 29, 36, 42, 43


@dataclass(frozen=True)
class Err:
    code: int


@dataclass(frozen=True)
class Int:
    value: int


def _area(address: str) -> str:
    """An address, B2 or B2:C3 (a whole column D:D or row 5:5 too), as the host's area:
    first row, last row, first column, last column, counted from 0."""
    import re

    def corner(text, last):
        letters, digits = re.fullmatch(r"([A-Z]*)(\d*)", text).groups()
        column = 0
        for c in letters:
            column = column * 26 + ord(c) - 64
        row = int(digits) if digits else (1048576 if last else 1)
        return row - 1, (column or (16384 if last else 1)) - 1

    first, _, last = address.upper().partition(":")
    (r1, c1), (r2, c2) = corner(first, False), corner(last or first, True)
    return f"{r1},{r2},{c1},{c2}"


@dataclass(frozen=True)
class Ref:
    """A reference Excel passes for a parameter that takes one (xltypeRef): its sheet, as
    Excel names it ("[Book1]Sheet1"), and one or more areas, as addresses."""

    sheet: str
    areas: tuple

    def __init__(self, sheet: str, *areas: str):
        object.__setattr__(self, "sheet", sheet)
        object.__setattr__(self, "areas", areas)


@dataclass(frozen=True)
class SRef:
    """A reference with no sheet (xltypeSRef), on the sheet calculating: one area."""

    area: str


class _Missing:
    def __repr__(self) -> str:
        return "MISSING"


MISSING = _Missing()


class HostFault(AssertionError):
    """The add-in broke one of Excel's rules."""


class HostError(RuntimeError):
    """The host refused a request, or stopped."""


def _hex(s: str) -> str:
    units = s.encode("utf-16-le", "surrogatepass")
    return "".join(f"{units[i + 1]:02x}{units[i]:02x}" for i in range(0, len(units), 2))


def _unhex(h: str) -> str:
    data = bytes.fromhex(h)
    return b"".join(data[i + 1 : i + 2] + data[i : i + 1] for i in range(0, len(data), 2)).decode(
        "utf-16-le", "surrogatepass"
    )


def encode(value) -> str:
    if value is MISSING:
        return "m"
    if value is None:
        return "z"
    if isinstance(value, bool):
        return f"b:{int(value)}"
    if isinstance(value, Err):
        return f"e:{value.code}"
    if isinstance(value, Int):
        return f"i:{value.value}"
    if isinstance(value, Ref):
        return f"r:{_hex(value.sheet)}:" + ";".join(_area(a) for a in value.areas)
    if isinstance(value, SRef):
        return f"R:{_area(value.area)}"
    if isinstance(value, (int, float)):
        return f"n:{float(value)!r}"
    if isinstance(value, str):
        return f"s:{_hex(value)}"
    if isinstance(value, (list, tuple)):
        rows = [list(r) for r in value]
        columns = len(rows[0])
        assert all(len(r) == columns for r in rows), "arrays sent to the host must be rectangular"
        return " ".join([f"a:{len(rows)}x{columns}"] + [encode(v) for r in rows for v in r])
    raise TypeError(f"cannot send {type(value).__name__} to the host")


def decode(tokens: list[str]):
    """Decodes one value from the front of tokens, consuming it."""
    token = tokens.pop(0)
    kind, _, body = token.partition(":")
    if kind == "n":
        return float("nan") if "nan" in body else float(body)
    if kind == "s":
        return _unhex(body)
    if kind == "b":
        return body == "1"
    if kind == "e":
        return Err(int(body))
    if kind == "i":
        return Int(int(body))
    if kind == "m":
        return MISSING
    if kind == "z":
        return None
    if kind == "a":
        rows, columns = (int(n) for n in body.split("x"))
        return tuple(tuple(decode(tokens) for _ in range(columns)) for _ in range(rows))
    raise HostError(f"bad value token from host: {token}")


@dataclass(frozen=True)
class Console:
    """Pyinex's console window: its title, place and size, its text, with lines ending in a
    plain newline, and the colors of its text and background, as Windows COLORREFs."""

    title: str
    left: int
    top: int
    width: int
    height: int
    text: str
    text_color: int
    background: int


@dataclass
class Registration:
    name: str
    procedure: str
    type: str
    arguments: str
    category: str
    help: str
    argument_help: tuple[str, ...]


class Host:
    def __init__(
        self,
        xllhost: Path,
        xll: Path,
        *,
        env: dict[str, str] | None = None,
        detached: bool = True,
        timeout: float = 60,
        output: str = "captured",
    ):
        """output: "captured", the add-in's output collected in stdout and stderr; or
        "none", a host with a window and no standard output or error, as Excel starts
        (xllhost_gui.exe, beside xllhost)."""
        self.xll = Path(xll)
        self.timeout = timeout
        request_read, request_write = os.pipe()
        response_read, response_write = os.pipe()
        try:
            handles = []
            for fd in (request_read, response_write):
                h = msvcrt.get_osfhandle(fd)
                os.set_handle_inheritable(h, True)
                handles.append(h)
            startup = subprocess.STARTUPINFO()
            startup.lpAttributeList = {"handle_list": handles}
            # Excel starts with no console, and Pyinex treats a process that has one
            # differently (docs/architecture.md, "Console"), so the host starts without one.
            flags = subprocess.DETACHED_PROCESS if detached else 0
            program = Path(xllhost).with_name("xllhost_gui.exe") if output == "none" else xllhost
            streams = None if output == "none" else subprocess.PIPE
            self._process = subprocess.Popen(
                [str(program), str(handles[0]), str(handles[1])]
                + {"captured": [], "none": ["--no-output"]}[output],
                stdin=None if output == "none" else subprocess.DEVNULL,
                stdout=streams,
                stderr=streams,
                close_fds=True,
                startupinfo=startup,
                creationflags=flags,
                env=env,
            )
        except BaseException:
            for fd in (request_read, request_write, response_read, response_write):
                os.close(fd)
            raise
        os.close(request_read)
        os.close(response_write)
        self._requests = os.fdopen(request_write, "w", encoding="ascii", newline="\n")
        self._responses = os.fdopen(response_read, "r", encoding="ascii", newline="\n")
        self._stdout: list[bytes] = []
        self._stderr: list[bytes] = []
        # Responses are read on a thread too, so a request can wait with a timeout.
        self._answers: queue.Queue[str] = queue.Queue()
        # Each stream the driver reads, with the thread reading it; None: written, not read.
        self._streams = [(self._requests, None), (self._responses, threading.Thread(target=self._read_responses))]
        for stream, into in ((self._process.stdout, self._stdout), (self._process.stderr, self._stderr)):
            if stream is not None:  # None: a host with no output.
                self._streams.append((stream, threading.Thread(target=self._drain, args=(stream, into))))
        self._readers = [reader for _, reader in self._streams if reader is not None]
        for t in self._readers:
            t.daemon = True
            t.start()

    @staticmethod
    def _drain(stream, into: list[bytes]) -> None:
        for chunk in iter(lambda: stream.read1(65536), b""):
            into.append(chunk)

    def _read_responses(self) -> None:
        for line in self._responses:
            self._answers.put(line)
        self._answers.put("")  # The host closed its end: it has stopped.

    @property
    def stdout(self) -> str:
        return b"".join(self._stdout).decode("utf-8", "replace")

    @property
    def stderr(self) -> str:
        return b"".join(self._stderr).decode("utf-8", "replace")

    def _request(self, line: str) -> list[str]:
        # The write happens on a thread of its own, because it blocks if the host stops
        # reading; the one timeout covers sending and answering.
        def send() -> None:
            try:
                self._requests.write(line + "\n")
                self._requests.flush()
            except (OSError, ValueError):
                pass  # The host has stopped; the empty answer below says so.

        threading.Thread(target=send, daemon=True).start()
        try:
            response = self._answers.get(timeout=self.timeout)
        except queue.Empty:
            self._process.kill()
            self.quit()
            raise HostError(f"xllhost did not answer {line.split()[0]!r} within {self.timeout} s; stopped it") from None
        if not response:
            code = self._process.wait(timeout=10)
            raise HostError(f"xllhost stopped (exit code {code:#x}); stderr:\n{self.stderr}")
        status, *tokens = response.split()
        if status == "error":
            raise HostError(_unhex(tokens[0]) if tokens else "unknown error")
        return tokens

    def faults(self) -> list[str]:
        tokens = self._request("faults")
        return [decode(tokens) for _ in range(len(tokens))]

    def _checked(self, line: str, expect_faults: bool) -> tuple[list[str], list[str]]:
        tokens = self._request(line)
        faults = self.faults()
        if faults and not expect_faults:
            raise HostFault("; ".join(faults))
        return tokens, faults

    def load(self, dll: Path) -> None:
        """Loads another DLL into the host, as another add-in might."""
        self._request(f"load {_hex(str(dll))}")

    def open(self, xll: Path | None = None) -> float:
        """Loads the add-in and calls its xlAutoOpen; xll: another copy of it instead."""
        tokens, _ = self._checked(f"open {_hex(str(xll or self.xll))}", False)
        return decode(tokens)

    def close(self) -> float:
        tokens, _ = self._checked("close", False)
        return decode(tokens)

    def unload(self) -> tuple[bool, bool]:
        """FreeLibrary the add-in, as Excel does after closing it; returns whether the add-in
        and pyinex_core.dll are still loaded."""
        tokens = self._request("unload")
        return decode(tokens), decode(tokens)

    def excel_calls(self) -> int:
        """How many calls the add-in made into Excel since the last time this was asked."""
        return int(decode(self._request("calls")))

    def handles(self) -> int:
        return int(decode(self._request("handles")))

    def console(self) -> Console | None:
        """Pyinex's console window, or None if it has made none."""
        tokens = self._request("console")
        if not decode(tokens):
            return None
        title, left, top, width, height, text, text_color, background = (decode(tokens) for _ in range(8))
        return Console(
            title, int(left), int(top), int(width), int(height), text.replace("\r\n", "\n"), int(text_color), int(background)
        )

    def set_theme(self, mode: str) -> None:
        """Windows' app mode, "light" or "dark", as the add-in sees it; the user's own
        setting is left alone. A console window already made is told of the change."""
        self._request(f"theme {mode}")

    def register_python(self, hive: str, version: str, executable=None, folder=None) -> None:
        """Records a Python installation as python.org's installer does (PEP 514), in this
        host's own stand-in for HKEY_CURRENT_USER ("user") or HKEY_LOCAL_MACHINE
        ("machine"): executable as InstallPath's ExecutablePath, folder as its default
        value. The machine's own registry is neither read nor changed."""
        values = " ".join(_hex(str(v)) if v is not None else "-" for v in (executable, folder))
        self._request(f"register {hive} {_hex(version)} {values}")

    def close_console(self) -> None:
        """Closes the console window, as its close button does."""
        self._request("close-console")

    def write_output(self, text: str) -> None:
        """Writes text to the host's standard error, as C code's output would."""
        self._request(f"write {_hex(text)}")

    def call(self, name: str, *args):
        """Calls a registered function and returns its result."""
        return self.call_detail(name, *args)[0]

    def call_detail(self, name: str, *args, expect_faults: bool = False):
        """Returns (result, whether it carried xlbitDLLFree, faults)."""
        line = " ".join(["call", _hex(name)] + [encode(a) for a in args])
        tokens, faults = self._checked(line, expect_faults)
        dll_free = tokens.pop(0) == "1"
        return decode(tokens), dll_free, faults

    def registrations(self) -> list[Registration]:
        tokens = self._request("registrations")
        out = []
        while tokens:
            (row,) = decode(tokens)
            out.append(Registration(*row[:6], argument_help=tuple(row[6:])))
        return out

    def set_caller(
        self, sheet: str | None, row: int = 1, column: int = 1, rows: int = 1, columns: int = 1, sref: bool = False
    ):
        """What xlfCaller returns: the calling cell or range, or None for no cell. sref: as an
        xltypeSRef, a reference with no sheet, rather than an xltypeRef."""
        if sheet is None:
            self._request("caller none")
        else:
            self._request(f"caller {_hex(sheet)} {row} {column} {rows} {columns}" + (" sref" if sref else ""))

    def set_cells(self, sheet: str, cell: str, value) -> None:
        """Puts a value, or an array's values from cell on, into a sheet's cells, which a
        reference's xlCoerce and COUNTA read; None empties a cell."""
        row, _, column, _ = (int(n) for n in _area(cell).split(","))
        self._request(f"cells {_hex(sheet)} {row + 1} {column + 1} {encode(value)}")

    def stop(self, code: str, skip: int = 0) -> None:
        """After skip more readings of a reference, Excel answers code, "uncalced" or "abort"
        (xlretUncalced, xlretAbort), as for a cell not yet calculated or an interrupted
        calculation; "none" undoes it. Any later call into Excel but xlFree, before the
        function that met it returns, is a fault."""
        self._request(f"stop {code} {skip}")

    def forget(self, sheet: str) -> None:
        """The sheet is deleted, or its workbook closed: Excel no longer finds it by its name
        (xlSheetId), and gives its ID to the next sheet named, as Excel may."""
        self._request(f"forget {_hex(sheet)}")

    def set_formula(self, sheet: str, cell: str, formula: str) -> None:
        """Puts a formula in a sheet's cell, which the stand-in for Excel's Application object
        finds through its Workbooks, as [Book1]Sheet1 is a sheet of the workbook Book1."""
        row, _, column, _ = (int(n) for n in _area(cell).split(","))
        self._request(f"formula {_hex(sheet)} {row + 1} {column + 1} {_hex(formula)}")

    def set_array(self, sheet: str, area: str, formula: str) -> None:
        """Puts a legacy array formula (Ctrl+Shift+Enter) in an area of a sheet, such as A1:B2,
        which only the area's FormulaArray enters again."""
        r1, r2, c1, c2 = (int(n) for n in _area(area).split(","))
        self._request(f"array {_hex(sheet)} {r1 + 1} {c1 + 1} {r2 - r1 + 1} {c2 - c1 + 1} {_hex(formula)}")

    def lock(self, sheet: str, locked: bool = True, unreadable: bool = False) -> None:
        """Whether entering a formula on the sheet fails, as on a protected sheet; with
        unreadable, reading one fails too."""
        self._request(f"lock {_hex(sheet)} {2 if unreadable else int(locked)}")

    def set_handler(self, search: str) -> None:
        """After each formula entered again, Excel's search is for search, as when a
        workbook's change handler calls Range.Find; empty, none."""
        self._request(f"handler {_hex(search)}")

    def set_change(self, verb: str = "", name: str = "") -> None:
        """After each formula entered again, the host runs the command name (verb "run"), as
        a workbook's change handler running a command would; no verb, none. The host reports
        a fault once it has run it 25 times."""
        self._request(f"change {verb} {_hex(name)}" if verb else "change")

    def opened(self, book: str, new: bool = False) -> bool:
        """Excel reports the workbook opened (WorkbookOpen), or with new made from a template
        (NewWorkbook); returns whether a handler of Excel's application events was connected
        to hear it."""
        tokens, _ = self._checked(f"opened {_hex(book)}" + (" new" if new else ""), False)
        return decode(tokens)

    def entered(self) -> tuple[list[tuple], int]:
        """Each formula entered again since the last call, as (sheet, address, formula,
        value; None where the host does not calculate it), and how many times the add-in
        asked Excel to calculate (xlcCalculateNow) meanwhile."""
        tokens, _ = self._checked("entered", False)
        calculations = int(decode(tokens))
        return [tuple(row) for row in decode(tokens)], calculations

    def set_escape(self, pressed: bool) -> None:
        self._request(f"abort {int(pressed)}")

    def set_folder(self, folder: str | Path | None, alone: bool = False) -> None:
        """The calling workbook's folder, which CELL("filename") reports; None: unsaved.

        alone: CELL gives the workbook's path without its sheet, as Excel does for a
        workbook whose only sheet has the workbook's name.
        """
        if folder is None:
            self._request("folder none")
        else:
            self._request(f"folder {_hex(str(folder))}" + (" alone" if alone else ""))

    def set_wizard(self, open: bool) -> None:
        """Opens or closes a stand-in for Excel's function wizard."""
        self._request(f"wizard {int(open)}")

    def run(self, name: str) -> float:
        """Runs a registered command as Excel runs one in a command context, the host's
        stand-in for Application.Run; returns its result."""
        tokens, _ = self._checked(f"run {_hex(name)}", False)
        return decode(tokens)

    def tick(self, count: int = 1) -> tuple[bool, int, int]:
        """Sends the add-in's queue window its timer's message count times; returns whether
        the window exists, and how many times the add-in called Application.Run and how
        many of those the host refused meanwhile."""
        tokens, _ = self._checked(f"tick {count}", False)
        return decode(tokens), int(decode(tokens)), int(decode(tokens))

    def set_workbook(self, state: str) -> None:
        """Whether a workbook window is open: "none", "open" (the default),
        "protected", a Protected View window, whose object has no Application, or "both", a
        workbook window and an active Protected View window, while which Evaluate gives #N/A."""
        self._request(f"workbook {state}")

    def set_busy(self, busy: bool) -> None:
        """Whether Application.Run and Evaluate refuse, as Excel does while a cell is edited."""
        self._request(f"busy {int(busy)}")

    def fail_registrations(self, count: int) -> None:
        """Makes the add-in's next count registrations of direct functions fail."""
        self._request(f"fail register {count}")

    def fail_commands(self, count: int) -> None:
        """Makes the add-in's next count registrations of commands fail."""
        self._request(f"fail command {count}")

    def fail_name(self, name: str) -> None:
        """Makes the add-in's next registration under name fail."""
        self._request(f"fail name {_hex(name)}")

    def fail_unregisters(self, count: int, skip: int = 0) -> None:
        """After skip more calls of xlfUnregister that succeed, makes the next count fail."""
        self._request(f"fail unregister {count} {skip}")

    def fail_set_names(self, count: int) -> None:
        """Makes the next count calls of xlfSetName fail, as Excel's FALSE says it failed."""
        self._request(f"fail setname {count}")

    def foreign(self, name: str) -> None:
        """Registers name as if another add-in had."""
        self._request(f"foreign {_hex(name)}")

    def fail_windows(self, count: int) -> None:
        """Makes the add-in's next count attempts to create its queue window fail."""
        self._request(f"fail window {count}")

    def fail_timers(self, count: int) -> None:
        """Makes the add-in's next count calls of SetTimer fail."""
        self._request(f"fail timer {count}")

    def fail_evaluates(self, count: int) -> None:
        """Makes the next count calls of xlfEvaluate fail, as a question Excel cannot answer."""
        self._request(f"fail evaluate {count}")

    def fail_memory(self, after: str, skip: int = 0) -> None:
        """Once the add-in's next xlfEvaluate (after "evaluate"), its next registration of a
        worksheet function that succeeds ("register"), or its next xlfUnregister that fails
        ("unregister"), returns, or at once ("now"), makes its next allocation after skip
        more fail, as when memory runs out; with "now", the core's too, once Python has
        started."""
        self._request(f"fail memory {after} {skip}")

    def nest(self, verb: str, name: str, *args) -> None:
        """The next time the add-in registers a direct function, the host first runs
        the command name (verb "run") or calls the function name with args ("call")."""
        self._request(" ".join(["nest", verb, _hex(name)] + [encode(a) for a in args]))

    def nest_all(self, *requests: str) -> None:
        """The next time the add-in registers a direct function, the host first carries out
        each request in turn, a line of nest's: "run <hex name>", "call ...", or
        "memory <skip>", making the add-in's next allocation after skip more fail."""
        self._request("nest " + ";".join(requests))

    def nest_at(self, procedure: str, *requests: str) -> None:
        """As nest_all, at the add-in's next registration of a function or command whose
        procedure starts with procedure (xlAutoRemove: the stand-in a removal registers)."""
        self._request(f"nestat {_hex(procedure)} " + ";".join(requests))

    def quit(self) -> int:
        try:
            if self._process.poll() is None:
                try:
                    self._requests.write("quit\n")
                    self._requests.flush()
                except OSError:
                    pass
            try:
                code = self._process.wait(timeout=30)
            except subprocess.TimeoutExpired:
                self._process.kill()
                code = self._process.wait()
            for t in self._readers:
                t.join(timeout=10)
            return code
        finally:
            # A stream whose reader is still blocked (a process the add-in started may hold
            # the pipe open) is left to the garbage collector: closing it would wait for the
            # reader.
            for stream, reader in self._streams:
                if reader is None or not reader.is_alive():
                    try:
                        stream.close()
                    except (OSError, ValueError):
                        pass

    def __enter__(self) -> Host:
        return self

    def __exit__(self, *exc) -> None:
        self.quit()
