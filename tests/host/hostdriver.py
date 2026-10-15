"""Drive xllhost.exe, the fake Excel, from tests.

    with Host(xllhost, xll) as host:
        host.open()
        assert host.call("PyInfo", "status") == "running"

Values go to and come back from Excel as Python values: ``float`` for numbers, ``str``,
``bool``, ``Err`` for Excel errors, ``MISSING`` for an argument left out, ``None`` for an
empty value, and a tuple of row tuples for an array. ``Int`` sends Excel's rarely used
xltypeInt.

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

    def open(self) -> float:
        tokens, _ = self._checked(f"open {_hex(str(self.xll))}", False)
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
