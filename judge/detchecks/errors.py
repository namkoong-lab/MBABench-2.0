"""The exception types of the package.

Patrick's ruling (2026-10-02): NO FALLBACK. When a check cannot grade a file
(unreadable or unsupported format, a value it needs is missing or untrusted, an
internal error) it raises GradingError with a clear message. It never quietly
passes, fails, or defers to the LLM.
"""
from __future__ import annotations


class GradingError(Exception):
    """A check could not grade a file.

    Attributes (all optional, filled in by the engine when it knows them):
      check       -- rubric key of the check that failed ("Potential Dangers/No hidden sheets")
      path        -- workbook path
      failures    -- {check key: message} for every check that failed in one grade() call
      verdicts    -- {check key: verdict} for the checks that did finish in that call
                     (for inspection only; the call as a whole failed)
      retry_later -- True when the cause is the machine, not the file: LibreOffice could not run
                     now (LibreOfficeUnavailable).  The attempt is not graded and must be re-run
                     later, when the machine has memory to spare.
    """

    def __init__(self, message: str, *, check: str | None = None, path: str | None = None,
                 failures: dict | None = None, verdicts: dict | None = None, retry_later: bool = False):
        super().__init__(message)
        self.check = check
        self.path = path
        self.failures = failures or {}
        self.verdicts = verdicts or {}
        self.retry_later = bool(retry_later)


class LibreOfficeUnavailable(GradingError):
    """LibreOffice could not recalculate a file NOW (Patrick 2026-10-05, core/lo_guard.py): the
    machine never had enough free memory within the maximum wait, or every try failed (crash, no
    copy, timeout - the timeout doubled on each retry).  A property of the machine, not of the
    file: the grading fails loudly (no fallback, no verdict), and the attempt is to be re-run later,
    when there is memory to spare (retry_later is always True)."""

    def __init__(self, message: str, **kw):
        kw["retry_later"] = True
        super().__init__(message, **kw)
