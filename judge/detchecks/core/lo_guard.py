"""Memory safety for every LibreOffice run in grading (Patrick, 2026-10-05).

    "For production runs, every attempt must be graded. Doesn't matter what size."

So no file is refused for its size.  Instead every LibreOffice run of a grading - the detchecks
recalculation (core/recalc.libreoffice_recalc), the answer check's recalculation
(utils/answer_check.py) and the judge's --run-calculation re-save (utils/excel_utils.recalculate_xlsx)
- goes through run_guarded():

  1. ONE LibreOffice at a time on the machine.  An exclusive lock: fcntl.flock on `lock_path`
     (default /tmp/mbabench_libreoffice.lock - the system temp dir, the same file for every
     process whatever its TMPDIR; DETCHECKS_LO_LOCK or the setting overrides it) plus a
     process-wide threading lock, so worker threads of one driver queue too.  The lock is held for
     the whole run; a waiting grader logs who holds it.  flock is released by the kernel when the
     holder dies, and the lock file descriptor is never inherited by LibreOffice (O_CLOEXEC).
     The wait for the lock is capped at `max_lock_wait_s` (3 hours; Patrick 2026-10-05: every
     attempt graded - one stuck file cannot stall a run indefinitely): after that the run fails
     loudly - LibreOfficeUnavailable (retry_later), like the other LibreOffice failures.
  2. Memory first.  Holding the lock, wait until the machine's free memory (macOS
     `memory_pressure -Q`, "System-wide memory free percentage"; Linux /proc/meminfo MemAvailable /
     MemTotal) is at least `min_free_pct`, logging while it waits.  After `max_wait_s` the run fails
     loudly - LibreOfficeUnavailable (retry_later): the attempt is not graded now and is re-run
     later.  Where free memory cannot be read at all the run starts (logged once).
  3. Retry instead of failing.  A run that fails (crash, no copy, timeout) is tried again, up to
     `retries` more times, the timeout doubled each time; every try takes the lock and waits for
     memory again.  After the last try LibreOfficeUnavailable lists every try's error.

The watchdog / reaper of core/recalc.py is unchanged: LibreOffice never outlives its grading.
Standard library only.
"""
from __future__ import annotations

import os
import re
import socket
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass
from typing import Callable, Optional

try:
    import fcntl
except ImportError:  # pragma: no cover - Windows: the process-wide lock only
    fcntl = None

from ..errors import GradingError, LibreOfficeUnavailable

LOCK_FILENAME = "mbabench_libreoffice.lock"
LOCK_ENV = "DETCHECKS_LO_LOCK"
FLOCK_POLL_S = 0.25                 # how often a waiting process retries the machine-wide lock
MAX_ERROR_CHARS = 300               # per try, in the final message


def default_lock_path() -> str:
    """The machine-wide lock file: $DETCHECKS_LO_LOCK, else /tmp/mbabench_libreoffice.lock (POSIX: one file
    for every process on the machine, whatever its TMPDIR), else the platform temp dir."""
    env = (os.environ.get(LOCK_ENV) or "").strip()
    if env:
        return env
    base = "/tmp" if os.name == "posix" and os.path.isdir("/tmp") else tempfile.gettempdir()
    return os.path.join(base, LOCK_FILENAME)


@dataclass(frozen=True)
class GuardSettings:
    lock_path: Optional[str] = None      # None / "": default_lock_path()
    min_free_pct: float = 25.0           # free memory (%) required before LibreOffice starts; 0 = no wait
    max_wait_s: float = 3600.0           # longest wait for that memory, then LibreOfficeUnavailable
    max_lock_wait_s: float = 3 * 3600.0  # longest wait for the machine-wide lock, then LibreOfficeUnavailable
                                         # (Patrick 2026-10-05: every attempt graded - a stuck file cannot stall a run)
    retries: int = 3                     # extra tries after a failed run, the timeout doubled each time
    poll_s: float = 15.0                 # memory re-read interval while waiting
    log_every_s: float = 60.0            # a waiting grader logs at most this often
    # test hooks
    memory_reader: Optional[Callable[[], Optional[float]]] = None
    sleep: Optional[Callable[[float], None]] = None
    clock: Optional[Callable[[], float]] = None

    def lock(self) -> str:
        return self.lock_path or default_lock_path()

    def problems(self) -> list:
        """Invalid values, as messages (empty when the settings are usable)."""
        out = []
        if not isinstance(self.retries, int) or isinstance(self.retries, bool) or self.retries < 0:
            out.append(f"retries must be a whole number >= 0, got {self.retries!r}")
        if not 0 <= float(self.min_free_pct) < 100:
            out.append(f"min_free_pct must be in [0, 100), got {self.min_free_pct!r}")
        if not float(self.max_wait_s) >= 0:
            out.append(f"max_wait_s must be >= 0, got {self.max_wait_s!r}")
        if not float(self.max_lock_wait_s) >= 0:
            out.append(f"max_lock_wait_s must be >= 0, got {self.max_lock_wait_s!r}")
        if not float(self.poll_s) > 0:
            out.append(f"poll_s must be > 0, got {self.poll_s!r}")
        return out


def _stderr_log(msg: str) -> None:
    sys.stderr.write(f"[LibreOffice guard] {msg}\n")
    sys.stderr.flush()


# ---------------------------------------------------------------------------- free memory
_MP_RX = re.compile(r"free percentage:\s*(\d+(?:\.\d+)?)\s*%")
_MEMORY_PRESSURE = "/usr/bin/memory_pressure"


def free_memory_pct() -> Optional[float]:
    """The machine's free memory in percent, or None when it cannot be read.  macOS: `memory_pressure -Q`
    ('System-wide memory free percentage: 63%'); Linux: MemAvailable /
    MemTotal from /proc/meminfo."""
    if sys.platform == "darwin":
        exe = _MEMORY_PRESSURE if os.path.exists(_MEMORY_PRESSURE) else "memory_pressure"
        try:
            out = subprocess.run([exe, "-Q"], capture_output=True, text=True, timeout=30).stdout
        except (OSError, subprocess.SubprocessError):
            return None
        m = _MP_RX.search(out or "")
        return float(m.group(1)) if m else None
    try:
        info = {}
        with open("/proc/meminfo", encoding="ascii", errors="replace") as fh:
            for line in fh:
                k, _, rest = line.partition(":")
                parts = rest.split()
                if parts:
                    info[k.strip()] = float(parts[0])
    except (OSError, ValueError):
        return None
    total = info.get("MemTotal")
    avail = info.get("MemAvailable")
    if avail is None:
        avail = info.get("MemFree", 0.0) + info.get("Buffers", 0.0) + info.get("Cached", 0.0)
    return 100.0 * avail / total if total else None


_UNKNOWN_MEMORY_LOGGED = {"done": False}


def wait_for_memory(settings: GuardSettings, *, what: str, log: Optional[Callable] = None) -> float:
    """Block until at least settings.min_free_pct of the machine's memory is free; returns the seconds
    waited.  Raises LibreOfficeUnavailable after settings.max_wait_s (the attempt is re-run later).
    Called with the machine-wide lock held, so no other grading starts LibreOffice meanwhile."""
    log = log or _stderr_log
    need = float(settings.min_free_pct or 0)
    if need <= 0:
        return 0.0
    reader = settings.memory_reader or free_memory_pct
    clock = settings.clock or time.monotonic
    sleep = settings.sleep or time.sleep
    t0 = clock()
    last_log = None
    while True:
        pct = reader()
        waited = clock() - t0
        if pct is None:
            if not _UNKNOWN_MEMORY_LOGGED["done"]:
                _UNKNOWN_MEMORY_LOGGED["done"] = True
                log(f"free memory cannot be read on this machine: {what} starts without the memory wait")
            return waited
        if pct >= need:
            if last_log is not None:
                log(f"{pct:.0f}% of memory free (>= {need:g}%): {what} starts after waiting {waited:.0f} s")
            return waited
        if waited >= settings.max_wait_s:
            raise LibreOfficeUnavailable(
                f"LibreOffice was not started to {what}: the machine did not have {need:g}% of its memory free "
                f"within {settings.max_wait_s / 60:g} min (last reading {pct:.0f}%) - not graded now; re-run this "
                f"attempt when the machine has memory to spare")
        if last_log is None or clock() - last_log >= settings.log_every_s:
            log(f"waiting for memory before LibreOffice runs to {what}: {pct:.0f}% free < {need:g}% "
                f"(waited {waited:.0f} s of at most {settings.max_wait_s:.0f} s)")
            last_log = clock()
        sleep(max(0.0, min(float(settings.poll_s), settings.max_wait_s - waited)))


# ---------------------------------------------------------------------------- the machine-wide lock
_PROCESS_LOCK = threading.Lock()     # worker threads of one process queue here first
_HELD = threading.local()            # per thread: nesting depth (a nested guarded run never self-deadlocks)


def _holder(path: str) -> str:
    try:
        with open(path, encoding="utf-8", errors="replace") as fh:
            text = fh.read(300).strip()
    except OSError:
        text = ""
    return text or "another process"


class machine_lock:
    """Context manager: hold the machine-wide LibreOffice lock (module doc, point 1).  Re-entrant within a
    thread.  `waited_s` is the time spent waiting for it.  max_wait_s: the longest wait for the lock (the
    process-wide lock and the lock file together) before LibreOfficeUnavailable (retry_later); None = no
    limit.  clock / sleep: test hooks (fakes) for that wait."""

    def __init__(self, path: Optional[str] = None, *, what: str = "run LibreOffice", log: Optional[Callable] = None,
                 log_every_s: float = 60.0, max_wait_s: Optional[float] = None,
                 clock: Optional[Callable[[], float]] = None, sleep: Optional[Callable[[float], None]] = None):
        self.path = path or default_lock_path()
        self.what = what
        self.log = log or _stderr_log
        self.log_every_s = log_every_s
        self.max_wait_s = max_wait_s
        self.clock = clock or time.monotonic
        self.sleep = sleep or time.sleep
        self.waited_s = 0.0
        self._fd = None
        self._nested = False

    def _give_up(self, t0: float, held_by: str):
        """LibreOfficeUnavailable once the lock has been waited for longer than max_wait_s (Patrick 2026-10-05:
        every attempt graded - a stuck LibreOffice run elsewhere cannot stall this grading indefinitely)."""
        waited = self.clock() - t0
        if self.max_wait_s is None or waited < self.max_wait_s:
            return
        raise LibreOfficeUnavailable(
            f"LibreOffice was not started to {self.what}: the machine-wide LibreOffice lock {self.path} stayed busy "
            f"({held_by}) for {waited / 60:.0f} min, longer than the maximum wait of {self.max_wait_s / 60:g} min - "
            f"not graded now; re-run this attempt later")

    def __enter__(self):
        depth = getattr(_HELD, "depth", 0)
        if depth:
            _HELD.depth = depth + 1
            self._nested = True
            return self
        t0 = self.clock()
        if not _PROCESS_LOCK.acquire(blocking=False):
            self.log(f"waiting for the LibreOffice lock to {self.what}: another grading thread of this process "
                     f"is running LibreOffice")
            if self.max_wait_s is None:
                while not _PROCESS_LOCK.acquire(timeout=self.log_every_s):
                    self.log(f"still waiting for the LibreOffice lock to {self.what} ({self.clock() - t0:.0f} s)")
            else:
                logged_at = self.clock()
                while not _PROCESS_LOCK.acquire(blocking=False):
                    self._give_up(t0, "another grading thread of this process")
                    now = self.clock()
                    if now - logged_at >= self.log_every_s:
                        self.log(f"still waiting for the LibreOffice lock to {self.what} ({now - t0:.0f} s)")
                        logged_at = now
                    self.sleep(FLOCK_POLL_S)
        try:
            self._fd = self._open()
            if fcntl is not None:
                logged_at = None
                while True:
                    try:
                        fcntl.flock(self._fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
                        break
                    except OSError:
                        now = self.clock()
                        if self.max_wait_s is not None and now - t0 >= self.max_wait_s:
                            self._give_up(t0, f"held by {_holder(self.path)}")
                        if logged_at is None or now - logged_at >= self.log_every_s:
                            held_by = _holder(self.path)
                            # the holder writes who it is just after taking the lock: give it a moment before
                            # the first log line says "another process"
                            if not (logged_at is None and held_by == "another process" and now - t0 < 1.0):
                                self.log(f"waiting for the machine-wide LibreOffice lock {self.path} to {self.what} "
                                         f"(held by {held_by}; waited {now - t0:.0f} s)")
                                logged_at = now
                        self.sleep(FLOCK_POLL_S)
            self._mark()
        except BaseException:
            self._close()
            _PROCESS_LOCK.release()
            raise
        _HELD.depth = 1
        self.waited_s = self.clock() - t0
        return self

    def _open(self) -> int:
        d = os.path.dirname(self.path)
        if d:
            os.makedirs(d, exist_ok=True)
        flags = os.O_CREAT | getattr(os, "O_CLOEXEC", 0)
        try:
            return os.open(self.path, flags | os.O_RDWR, 0o666)
        except PermissionError:                     # another user's lock file: locking needs no write access
            return os.open(self.path, flags | os.O_RDONLY)

    def _mark(self):
        """Write who holds the lock into the file (for waiting graders' logs; best effort)."""
        try:
            os.ftruncate(self._fd, 0)
            os.lseek(self._fd, 0, os.SEEK_SET)
            os.write(self._fd, f"pid {os.getpid()} on {socket.gethostname()} to {self.what}, since "
                               f"{time.strftime('%Y-%m-%d %H:%M:%S')}\n".encode("utf-8", "replace"))
        except OSError:
            pass

    def _close(self):
        if self._fd is not None:
            try:
                os.close(self._fd)
            except OSError:
                pass
            self._fd = None

    def __exit__(self, *exc):
        if self._nested:
            _HELD.depth -= 1
            return False
        _HELD.depth = 0
        try:
            if self._fd is not None:
                try:
                    os.ftruncate(self._fd, 0)
                except OSError:
                    pass
                if fcntl is not None:
                    try:
                        fcntl.flock(self._fd, fcntl.LOCK_UN)
                    except OSError:
                        pass
        finally:
            self._close()
            _PROCESS_LOCK.release()
        return False


# ---------------------------------------------------------------------------- one guarded LibreOffice job
def _short(e: BaseException) -> str:
    text = str(e) if isinstance(e, GradingError) else f"{type(e).__name__}: {e}"
    text = " ".join(text.split())
    return text if len(text) <= MAX_ERROR_CHARS else text[:MAX_ERROR_CHARS - 3] + "..."


def run_guarded(run_once: Callable[[float], object], *, timeout_s: float, settings: Optional[GuardSettings] = None,
                what: str = "run LibreOffice", log: Optional[Callable] = None):
    """run_once(timeout) under the guard (module doc): for each try - the first with `timeout_s`, each retry
    with twice the previous timeout - take the machine-wide lock, wait for memory, run.  Returns run_once's
    result.  A try fails when run_once raises an Exception (KeyboardInterrupt / SystemExit pass straight
    through).  Raises LibreOfficeUnavailable when the lock wait (max_lock_wait_s) or the memory wait times out,
    or every try failed."""
    s = settings or GuardSettings()
    bad = s.problems()
    if bad:
        raise ValueError(f"invalid LibreOffice guard settings: {'; '.join(bad)}")
    log = log or _stderr_log
    tries = s.retries + 1
    errors = []
    for k in range(tries):
        t = float(timeout_s) * (2 ** k)
        with machine_lock(s.lock(), what=what, log=log, log_every_s=s.log_every_s, max_wait_s=s.max_lock_wait_s,
                          clock=s.clock, sleep=s.sleep):
            wait_for_memory(s, what=what, log=log)
            t0 = time.monotonic()
            try:
                return run_once(t)
            except LibreOfficeUnavailable:
                raise
            except Exception as e:  # noqa: BLE001 - crash, no copy, timeout: retried below
                errors.append(f"try {k + 1}/{tries} (timeout {t:.0f} s, after {time.monotonic() - t0:.0f} s): {_short(e)}")
                if k + 1 < tries:
                    log(f"LibreOffice failed to {what} (try {k + 1} of {tries}): {_short(e)}; retrying with a "
                        f"{2 * t:.0f} s timeout once the lock is free and memory allows")
    raise LibreOfficeUnavailable(
        f"LibreOffice could not {what}: all {tries} tries failed ({'; '.join(errors)}) - not graded now; re-run "
        f"this attempt when the machine has memory to spare")
