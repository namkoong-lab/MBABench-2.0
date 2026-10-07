"""LibreOffice watchdog: a LibreOffice conversion must never outlive the grading that started it.

core/recalc.libreoffice_recalc does not start soffice itself. It starts this file as a tiny wrapper,

    <python> -I -S lo_watchdog.py <grader pid> <poll seconds> -- <soffice> <arguments ...>

which runs soffice as its own child, in the grader's process group (so a signal sent to the
grader's group - Ctrl-C, a process-group kill - reaches soffice too), and passes soffice's exit
code through.  Every <poll seconds> it checks that the grader is still its parent: when the grader
is gone (killed with SIGKILL, crashed - anything that runs no clean-up), the wrapper is re-parented,
notices, and kills soffice and everything soffice started (kill_tree), then exits.  SIGTERM, SIGINT
and SIGHUP sent to the wrapper do the same.  The grader's own side (a SIGTERM handler, the timeout
kill, the profile sweep) is in core/recalc.py.

Standard library only: the wrapper runs with -I -S, outside the package.  core/recalc.py imports
process_table / kill_tree from here, so there is one tree-kill for both sides.

Exit codes of the wrapper: soffice's own; 128 + N when soffice was ended by signal N (and a line
on stderr says so); 125 when the grader was already gone (soffice then never ran or was killed).
"""
from __future__ import annotations

import os
import signal
import subprocess
import sys

GRADER_GONE = 125
_SIGSTOP = getattr(signal, "SIGSTOP", None)
_SIGKILL = getattr(signal, "SIGKILL", signal.SIGTERM)


def process_table(timeout: float = 10.0):
    """[(pid, ppid, command line)] of every process, from `ps -ww -Ao pid=,ppid=,command=` (full
    command lines); None when ps cannot be run."""
    try:
        out = subprocess.run(["ps", "-ww", "-Ao", "pid=,ppid=,command="], capture_output=True, text=True,
                             timeout=timeout).stdout
    except (OSError, subprocess.SubprocessError):
        return None
    rows = []
    for line in out.splitlines():
        parts = line.strip().split(None, 2)
        if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
            rows.append((int(parts[0]), int(parts[1]), parts[2] if len(parts) > 2 else ""))
    return rows


def descendants(root: int, table) -> list:
    """Every process below `root` in `table` (children, grandchildren, ...)."""
    kids: dict = {}
    for pid, ppid, _cmd in table:
        kids.setdefault(ppid, []).append(pid)
    out, stack, seen = [], [root], {root}
    while stack:
        for child in kids.get(stack.pop(), ()):
            if child not in seen:
                seen.add(child)
                out.append(child)
                stack.append(child)
    return out


def ancestors(table) -> set:
    """This process and its ancestors (never to be killed by kill_tree)."""
    parent = {pid: ppid for pid, ppid, _cmd in table or ()}
    out, pid = set(), os.getpid()
    while pid and pid not in out:
        out.add(pid)
        pid = parent.get(pid, 0)
    out.add(os.getppid())
    return out


def _send(pid: int, sig) -> bool:
    try:
        os.kill(pid, sig)
        return True
    except OSError:
        return False


def kill_tree(root: int, rounds: int = 6) -> list:
    """SIGKILL `root` and every process below it.  The tree is frozen first (SIGSTOP on each
    process, repeated until a fresh process table shows no new child), so nothing forks away and
    escapes while it is being killed.  Never touches this process or its ancestors.  Returns the
    pids signalled."""
    table = process_table()
    protected = ancestors(table)
    if root in protected:
        return []
    frozen = [root]
    if _SIGSTOP is not None:
        _send(root, _SIGSTOP)
    for _ in range(rounds):
        if table is None:
            break
        new = [p for p in descendants(root, table) if p not in frozen and p not in protected]
        if not new:
            break
        for p in new:
            if _SIGSTOP is not None:
                _send(p, _SIGSTOP)
            frozen.append(p)
        table = process_table()
    return [p for p in frozen if _send(p, _SIGKILL)]


def _main(argv: list) -> int:
    if len(argv) < 4 or argv[2] != "--":
        sys.stderr.write("usage: lo_watchdog.py <grader pid> <poll seconds> -- <command ...>\n")
        return 2
    grader, poll, cmd = int(argv[0]), float(argv[1]), argv[3:]
    held: dict = {}

    def _stop(signum, _frame):
        child = held.get("child")
        if child is not None:
            kill_tree(child.pid)
            try:
                child.wait(timeout=5)
            except subprocess.TimeoutExpired:
                pass
        os._exit(128 + signum)

    for name in ("SIGTERM", "SIGINT", "SIGHUP"):
        sig = getattr(signal, name, None)
        if sig is not None:
            signal.signal(sig, _stop)
    if os.getppid() != grader:
        return GRADER_GONE
    held["child"] = child = subprocess.Popen(cmd)        # same stdio, environment and process group
    while True:
        try:
            rc = child.wait(timeout=poll)
        except subprocess.TimeoutExpired:
            if os.getppid() != grader:                    # the grader died: we were re-parented
                kill_tree(child.pid)
                try:
                    child.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    pass
                return GRADER_GONE
            continue
        if rc < 0:
            sys.stderr.write(f"[lo_watchdog] {os.path.basename(cmd[0])} ended by signal {-rc}\n")
            return 128 - rc
        return rc


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
