"""Recalculation pipeline: where a workbook's formula VALUES come from (Patrick's design,
2026-10-04; docs/recalc.md).

    from detchecks.core.recalc import RecalcPolicy, ensure_values
    plan = ensure_values(path, workdir, RecalcPolicy())
    plan.source       # 'cache' | 'libreoffice' | 'excel'
    plan.value_path   # None for 'cache', else the recalculated copy (values only)

1. The delivered file was saved by Excel (values.detect_writer == 'excel'): Excel's own caches
   are the display -> source 'cache', no recalculation.
2. Otherwise LibreOffice recalculates a copy (headless, private profile, threaded calculation
   and OpenCL off, OOXMLRecalcMode = always, per-file timeout), whatever the file's size, under
   the memory guard of core/lo_guard.py (Patrick 2026-10-05): one LibreOffice at a time on the
   machine, started only once enough memory is free, a failed run retried with the timeout
   doubled, then LibreOfficeUnavailable (re-run the attempt later).  The copy is cached in
   `workdir/<sha256 of the delivered file>/` so a file is recalculated once.
3. "LibreOffice gaps": formula cells whose LibreOffice value is #NAME? / #VALUE! although Excel
   would compute them (a function LibreOffice does not implement, a spill reference A1#, a
   LAMBDA name, an array-evaluation difference, ...; see `classify_lo_error`).  A GENUINE
   error (misspelled function CONUTIFS, undefined name, "abc"+1) is not a gap.
4. Any gap -> the file needs a recalculation by real Excel (`excel_recalc`: AppleScript on
   macOS, COM via PowerShell on Windows).  Behind `policy.excel_allowed` and an availability
   probe; when Excel is needed and not allowed / not available -> GradingError naming the gap
   cells (no fallback).

The copy supplies VALUES only; structure and styles always come from the delivered file.
"""
from __future__ import annotations

import hashlib
import json
import os
import platform
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from ..errors import GradingError
from . import formula as F
from . import lo_guard
from . import lo_watchdog
from .package import Package
from .sheet import ExcelError, SheetStream
from .values import ValueSource, detect_provenance, make_context

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_WORKDIR = os.path.join(os.path.dirname(HERE), "out", "recalc_cache")
def _default_soffice() -> str:
    """The LibreOffice binary for the standalone tools: $DETCHECKS_SOFFICE, else `soffice` on
    PATH, else the macOS app bundle. The judge passes its own paths.libreoffice_path explicitly."""
    import shutil
    return (os.environ.get("DETCHECKS_SOFFICE") or shutil.which("soffice")
            or "/Applications/LibreOffice.app/Contents/MacOS/soffice")


DEFAULT_SOFFICE = _default_soffice()
MAX_GAPS_LISTED = 25

# Excel functions LibreOffice (25.8, the build on this Mac) cannot evaluate.  A LibreOffice
# #NAME? / #VALUE! on a formula that uses one of these says nothing about Excel's display.
# Verified against the corpus LibreOffice outputs (SCAN, LAMBDA, REDUCE, MAKEARRAY, MAP, BYCOL
# -> #NAME?; spill references A1# / INDEX(B20#,..) -> #VALUE!) and the prototype's probes
# (clusters/errors/formula_tools.py: LO 25.8 computes XLOOKUP, IFS, LET, TEXTBEFORE, VSTACK).
LO_UNSUPPORTED_ALWAYS = frozenset("""
LAMBDA MAP REDUCE SCAN MAKEARRAY BYROW BYCOL ISOMITTED GROUPBY PIVOTBY PERCENTOF TRIMRANGE
IMAGE ANCHORARRAY SINGLE STOCKHISTORY FIELDVALUE ARRAYTOTEXT VALUETOTEXT PY
""".split())
# added to LibreOffice in 24.8 / 25.8: unsupported on older builds
LO_ADDED_24_8 = frozenset("XLOOKUP XMATCH FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY LET".split())
LO_ADDED_25_8 = frozenset("""CHOOSECOLS CHOOSEROWS DROP EXPAND HSTACK VSTACK TAKE TEXTAFTER TEXTBEFORE
TEXTSPLIT TOCOL TOROW WRAPCOLS WRAPROWS REGEXTEST REGEXEXTRACT REGEXREPLACE""".split())
# functions that return arrays: in a plain (non-array) formula Excel's legacy evaluation (take
# the top-left / intersect) and LibreOffice's (whole block -> Err:502/504 = #VALUE!) differ
ARRAY_FUNCS = frozenset("""SEQUENCE TRANSPOSE MMULT MINVERSE MDETERM MUNIT FREQUENCY TREND GROWTH
LINEST LOGEST FILTER SORT SORTBY UNIQUE RANDARRAY CHOOSECOLS CHOOSEROWS TAKE DROP VSTACK HSTACK
TOCOL TOROW WRAPROWS WRAPCOLS EXPAND TEXTSPLIT OFFSET INDEX INDIRECT""".split())
GAP_ERRORS = ("#NAME?", "#VALUE!")

# Every Excel worksheet function name (Microsoft 365, 2026), from the function reference.  A
# called name outside this set is a misspelling -> a genuine #NAME? in Excel too.
EXCEL_FUNCTIONS = frozenset("""
ABS ACCRINT ACCRINTM ACOS ACOSH ACOT ACOTH ADDRESS AGGREGATE AMORDEGRC AMORLINC ANCHORARRAY AND
ARABIC AREAS ARRAYTOTEXT ASC ASIN ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE AVERAGEA AVERAGEIF
AVERAGEIFS BAHTTEXT BASE BESSELI BESSELJ BESSELK BESSELY BETA.DIST BETA.INV BETADIST BETAINV
BIN2DEC BIN2HEX BIN2OCT BINOM.DIST BINOM.DIST.RANGE BINOM.INV BINOMDIST BITAND BITLSHIFT BITOR
BITRSHIFT BITXOR BYCOL BYROW CALL CEILING CEILING.MATH CEILING.PRECISE CELL CHAR CHIDIST CHIINV
CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHITEST CHOOSE CHOOSECOLS CHOOSEROWS
CLEAN CODE COLUMN COLUMNS COMBIN COMBINA COMPLEX CONCAT CONCATENATE CONFIDENCE CONFIDENCE.NORM
CONFIDENCE.T CONVERT CORREL COS COSH COT COTH COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS COUPDAYBS
COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD COVAR COVARIANCE.P COVARIANCE.S CRITBINOM CSC CSCH
CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE
CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE DAVERAGE DAY DAYS DAYS360 DB DBCS DCOUNT DCOUNTA DDB DEC2BIN
DEC2HEX DEC2OCT DECIMAL DEGREES DELTA DETECTLANGUAGE DEVSQ DGET DISC DMAX DMIN DOLLAR DOLLARDE
DOLLARFR DPRODUCT DROP DSTDEV DSTDEVP DSUM DURATION DVAR DVARP EDATE EFFECT ENCODEURL EOMONTH ERF
ERF.PRECISE ERFC ERFC.PRECISE ERROR.TYPE EUROCONVERT EVEN EXACT EXP EXPAND EXPON.DIST EXPONDIST
F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FACT FACTDOUBLE FALSE FDIST FIELDVALUE FILTER FILTERXML FIND
FINDB FINV FISHER FISHERINV FIXED FLOOR FLOOR.MATH FLOOR.PRECISE FORECAST FORECAST.ETS
FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT
FREQUENCY FTEST FV FVSCHEDULE GAMMA GAMMA.DIST GAMMA.INV GAMMADIST GAMMAINV GAMMALN
GAMMALN.PRECISE GAUSS GCD GEOMEAN GESTEP GETPIVOTDATA GROUPBY GROWTH HARMEAN HEX2BIN HEX2DEC
HEX2OCT HLOOKUP HOUR HSTACK HYPERLINK HYPGEOM.DIST HYPGEOMDIST IF IFERROR IFNA IFS IMABS IMAGE
IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMCOSH IMCOT IMCSC IMCSCH IMDIV IMEXP IMLN IMLOG10 IMLOG2
IMPOWER IMPRODUCT IMREAL IMSEC IMSECH IMSIN IMSINH IMSQRT IMSUB IMSUM IMTAN INDEX INDIRECT INFO INT
INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA ISNONTEXT ISNUMBER
ISO.CEILING ISODD ISOMITTED ISOWEEKNUM ISPMT ISREF ISTEXT JIS KURT LAMBDA LARGE LCM LEFT LEFTB LEN
LENB LET LINEST LN LOG LOG10 LOGEST LOGINV LOGNORM.DIST LOGNORM.INV LOGNORMDIST LOOKUP LOWER
MAKEARRAY MAP MATCH MAX MAXA MAXIFS MDETERM MDURATION MEDIAN MID MIDB MIN MINA MINIFS MINUTE MINVERSE
MIRR MMULT MOD MODE MODE.MULT MODE.SNGL MONTH MROUND MULTINOMIAL MUNIT N NA NEGBINOM.DIST
NEGBINOMDIST NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV
NORMDIST NORMINV NORMSDIST NORMSINV NOT NOW NPER NPV NUMBERVALUE OCT2BIN OCT2DEC OCT2HEX ODD
ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD OFFSET OR PDURATION PEARSON PERCENTILE PERCENTILE.EXC
PERCENTILE.INC PERCENTOF PERCENTRANK PERCENTRANK.EXC PERCENTRANK.INC PERMUT PERMUTATIONA PHI
PHONETIC PI PIVOTBY PMT POISSON POISSON.DIST POWER PPMT PRICE PRICEDISC PRICEMAT PROB PRODUCT PROPER
PV PY QUARTILE QUARTILE.EXC QUARTILE.INC QUOTIENT RADIANS RAND RANDARRAY RANDBETWEEN RANK RANK.AVG
RANK.EQ RATE RECEIVED REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST REGISTER.ID REPLACE REPLACEB REPT
RIGHT RIGHTB ROMAN ROUND ROUNDDOWN ROUNDUP ROW ROWS RRI RSQ RTD SCAN SEARCH SEARCHB SEC SECH SECOND
SEQUENCE SERIESSUM SHEET SHEETS SIGN SIN SINGLE SINH SKEW SKEW.P SLN SLOPE SMALL SORT SORTBY SQL.REQUEST
SQRT SQRTPI STANDARDIZE STDEV STDEV.P STDEV.S STDEVA STDEVP STDEVPA STEYX STOCKHISTORY SUBSTITUTE
SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 SWITCH SYD T T.DIST T.DIST.2T
T.DIST.RT T.INV T.INV.2T T.TEST TAKE TAN TANH TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT TEXTAFTER
TEXTBEFORE TEXTJOIN TEXTSPLIT TIME TIMEVALUE TINV TOCOL TODAY TOROW TRANSLATE TRANSPOSE TREND TRIM
TRIMMEAN TRIMRANGE TRUE TRUNC TTEST TYPE UNICHAR UNICODE UNIQUE UPPER VALUE VALUETOTEXT VAR VAR.P
VAR.S VARA VARP VARPA VDB VLOOKUP VSTACK WEBSERVICE WEEKDAY WEEKNUM WEIBULL WEIBULL.DIST WORKDAY
WORKDAY.INTL WRAPCOLS WRAPROWS XIRR XLOOKUP XMATCH XNPV XOR YEAR YEARFRAC YIELD YIELDDISC YIELDMAT
Z.TEST ZTEST
""".split())


def lo_unsupported(lo_version: Optional[tuple]) -> frozenset:
    """Functions the given LibreOffice version cannot evaluate."""
    v = lo_version or (0, 0)
    out = set(LO_UNSUPPORTED_ALWAYS)
    if v < (24, 8):
        out |= LO_ADDED_24_8
    if v < (25, 8):
        out |= LO_ADDED_25_8
    return frozenset(out)


def lo_version_of(app: Optional[str]) -> Optional[tuple]:
    m = re.match(r"LibreOffice(?:Dev)?[/ ](\d+)\.(\d+)", app or "")
    return (int(m.group(1)), int(m.group(2))) if m else None


# ---------------------------------------------------------------------------- policy / result
@dataclass
class RecalcPolicy:
    workdir: str = DEFAULT_WORKDIR
    libreoffice_path: str = DEFAULT_SOFFICE
    lo_timeout_s: float = 600.0         # the first try's timeout; doubled on every retry
    excel_allowed: bool = False         # Patrick 2026-10-04 (evening): LibreOffice only for now; the Excel step stays
                                        # built but OFF (Mac sandbox prompts per file, 70-155 s per file). A file
                                        # LibreOffice cannot compute fails loudly on value-based checks.
    excel_timeout_s: float = 600.0
    use_cache: bool = True              # reuse a copy made earlier for the same file hash
    # Memory safety (Patrick 2026-10-05, core/lo_guard.py): no size limit; one LibreOffice at a time on the
    # machine (lock file), started only once lo_min_free_pct of memory is free (at most lo_max_wait_s of
    # waiting), a failed run retried lo_retries times with the timeout doubled; then LibreOfficeUnavailable.
    lo_retries: int = 3
    lo_min_free_pct: float = 25.0
    lo_max_wait_s: float = 3600.0
    lo_max_lock_wait_s: float = 3 * 3600.0   # longest wait for the machine-wide lock, then LibreOfficeUnavailable
                                             # (retry later; Patrick 2026-10-05: every attempt graded)
    lo_lock_path: Optional[str] = None  # None: lo_guard.default_lock_path() (/tmp/mbabench_libreoffice.lock)
    lo_log: Optional[Callable] = None   # where the guard logs its waits and retries (None: stderr)
    # test hooks (fakes): callables (src_path, out_dir, policy) -> path of the copy
    lo_runner: Optional[Callable] = None
    excel_runner: Optional[Callable] = None
    excel_available: Optional[bool] = None   # None = probe the machine
    lo_memory_reader: Optional[Callable] = None   # () -> free memory % (None: lo_guard.free_memory_pct)

    def guard_settings(self) -> lo_guard.GuardSettings:
        return lo_guard.GuardSettings(lock_path=self.lo_lock_path, min_free_pct=self.lo_min_free_pct,
                                      max_wait_s=self.lo_max_wait_s, max_lock_wait_s=self.lo_max_lock_wait_s,
                                      retries=self.lo_retries, memory_reader=self.lo_memory_reader)


@dataclass
class Gap:
    sheet: str
    ref: str
    value: str                 # LibreOffice's error value
    formula: str
    functions: list            # the functions / features that make LibreOffice unreliable here
    why: str

    def as_dict(self) -> dict:
        return {"sheet": self.sheet, "ref": self.ref, "value": self.value, "formula": self.formula[:200],
                "functions": self.functions, "why": self.why}


@dataclass
class ValuePlan:
    source: str                        # cache | libreoffice | excel
    value_path: Optional[str]          # None for 'cache'
    writer: str                        # writer of the delivered file
    file_hash: str
    errors_vetted: bool = False        # copy: its #NAME?/#VALUE! are genuine (no gaps) or Excel's own
    n_gaps: int = 0
    gaps: list = field(default_factory=list)       # [Gap] (first MAX_GAPS_LISTED)
    gap_functions: list = field(default_factory=list)
    n_lo_errors: int = 0               # #NAME?/#VALUE! cells in the LibreOffice copy
    n_genuine_lo_errors: int = 0
    timings: dict = field(default_factory=dict)    # libreoffice_s, excel_s, gap_scan_s, from_cache
    lo_version: Optional[str] = None
    notes: list = field(default_factory=list)

    def summary(self) -> dict:
        return {"source": self.source, "writer": self.writer, "value_path": self.value_path,
                "errors_vetted": self.errors_vetted, "n_gaps": self.n_gaps,
                "gap_functions": self.gap_functions, "gaps": [g.as_dict() for g in self.gaps],
                "n_lo_errors": self.n_lo_errors, "n_genuine_lo_errors": self.n_genuine_lo_errors,
                "timings": self.timings, "lo_version": self.lo_version, "notes": self.notes}

    def value_source(self) -> Optional[ValueSource]:
        if self.value_path is None:
            return None
        return ValueSource(self.value_path, errors_vetted=self.errors_vetted)


def file_hash(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:24]


# ---------------------------------------------------------------------------- LibreOffice
_LO_REGISTRY = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Calculation"><prop oor:name="UseThreadedCalculationForFormulaGroups" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="UseOpenCL" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="FirstRun" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>
</oor:items>
"""


def file_url(path: str) -> str:
    """`path` as a percent-encoded file URL (Path.as_uri of the absolute path).  LibreOffice reads
    its profile (-env:UserInstallation) and --outdir as URLs: a raw 'file://' + path with a space
    aborts it (task 28 "FruitJuice_3-Statement-Model - v2"), and a '%41' in a plain --outdir path
    is decoded, so the copy lands in another folder ('PctATask') - every path handed to soffice
    goes through here."""
    return Path(os.path.abspath(path)).as_uri()


# ---- LibreOffice processes never outlive their grading ------------------------------------------
# 1. soffice runs in the grader's process group (no new session): a signal to the grader's group
#    (Ctrl-C, a process-group kill) reaches it.
# 2. soffice runs under core/lo_watchdog.py, which kills it (and everything it started) when the
#    grader disappears, SIGKILL included.
# 3. SIGTERM / SIGHUP to the grader: a handler (install_termination_reaper, main thread, only where
#    the signal had its default action) kills this process's running conversions, then the process
#    dies of the signal exactly as before.
# 4. A timeout, an exception (KeyboardInterrupt too) and the end of every run kill the run's tree and
#    sweep its private profile; utils/det_checks also sweeps a grading's whole profile folder.
WATCHDOG_SCRIPT = os.path.join(HERE, "lo_watchdog.py")
WATCHDOG_POLL_S = 0.5
_ACTIVE: dict = {}        # run id -> {"profile": path, "pid": watchdog/soffice pid}: this process's conversions
_TERM_HANDLER = {"installed": False}


def _watchdog_cmd(cmd: list) -> list:
    """`cmd` (soffice ...) run under the watchdog; the bare command when no Python interpreter is
    known (then only the handler, the timeout kill and the profile sweeps apply)."""
    if not (sys.executable and os.path.exists(WATCHDOG_SCRIPT)):
        return cmd
    return [sys.executable, "-I", "-S", WATCHDOG_SCRIPT, str(os.getpid()), str(WATCHDOG_POLL_S), "--"] + cmd


def _profile_markers(path: str, subtree: bool) -> tuple:
    """The strings a command line carries for a profile `path` (raw path and file URL).  subtree:
    match every profile below the folder `path` (a trailing separator) instead of `path` itself."""
    raw, url = os.path.abspath(path), file_url(path)
    if subtree:
        raw, url = os.path.join(raw, ""), url.rstrip("/") + "/"
    return raw, url


def reap_profiles(path: str, subtree: bool = True, table=None) -> list:
    """Kill - whole process trees - every process whose command line names the LibreOffice profile
    `path` (subtree=True: any profile below the folder `path`), raw or as a file URL.  Profiles are
    private to one grading (workdir/_lo_profiles/lo_*), so other jobs' LibreOffice is never touched;
    this process and its ancestors never are.  Returns the pids signalled."""
    markers = _profile_markers(path, subtree)
    table = lo_watchdog.process_table() if table is None else table
    if table is None:
        return []
    protected = lo_watchdog.ancestors(table)
    killed: list = []
    for pid, _ppid, cmd in table:
        if pid in protected or pid in killed or cmd.startswith("ps "):
            continue
        if any(m in cmd for m in markers):
            killed += lo_watchdog.kill_tree(pid)
    return sorted(set(killed))


def kill_active_libreoffice() -> list:
    """Kill every LibreOffice conversion this process is running (trees, then a profile sweep)."""
    try:
        runs = list(_ACTIVE.values())
    except RuntimeError:                       # changed size during the copy: once more
        runs = list(dict(_ACTIVE).values())
    killed: list = []
    for run in runs:
        if run.get("pid"):
            killed += lo_watchdog.kill_tree(run["pid"])
    if runs:
        table = lo_watchdog.process_table(timeout=5)
        for run in runs:
            killed += reap_profiles(run["profile"], subtree=False, table=table)
    return sorted(set(killed))


def _reap_and_die(signum, _frame):
    try:
        kill_active_libreoffice()
    finally:
        signal.signal(signum, signal.SIG_DFL)
        os.kill(os.getpid(), signum)


def install_termination_reaper() -> bool:
    """SIGTERM / SIGHUP: kill this process's running LibreOffice conversions first, then die of the
    signal as before.  Installed once, from the main thread only (Python's rule), and only for a
    signal whose action is still the default (an application's own handler, or an ignored signal,
    is left alone).  Called by libreoffice_recalc and by utils/det_checks (startup and every
    grading), so drivers that run the checks in worker threads have it from their main thread."""
    if _TERM_HANDLER["installed"] or threading.current_thread() is not threading.main_thread():
        return _TERM_HANDLER["installed"]
    for name in ("SIGTERM", "SIGHUP"):
        sig = getattr(signal, name, None)
        try:
            if sig is not None and signal.getsignal(sig) == signal.SIG_DFL:
                signal.signal(sig, _reap_and_die)
        except (ValueError, OSError):
            pass
    _TERM_HANDLER["installed"] = True
    return True


def _kill_tree(proc: subprocess.Popen):
    """Kill `proc` and every process below it (watchdog, soffice, anything soffice started)."""
    lo_watchdog.kill_tree(proc.pid)
    try:
        proc.wait(timeout=5)
    except subprocess.TimeoutExpired:
        pass
    for fh in (proc.stdout, proc.stderr):
        try:
            if fh is not None:
                fh.close()
        except OSError:
            pass


def libreoffice_recalc(src: str, out_dir: str, policy: RecalcPolicy) -> str:
    """Recalculate `src` headless and write `<out_dir>/<stem>.xlsx`, under the memory guard (Patrick
    2026-10-05, core/lo_guard.py: one LibreOffice at a time on the machine, started only when memory is
    free, a failed run retried policy.lo_retries times with the timeout doubled).  Each try: private
    profile (threaded calculation off, OpenCL off, recalculate always on load), one process under the
    watchdog, in the grader's process group, timeout; paths passed as encoded file URLs (file_url).
    Returns the copy's path.  GradingError when the binary is missing (not retried);
    LibreOfficeUnavailable (retry_later) when the memory wait times out or every try failed."""
    soffice = policy.libreoffice_path
    if not (soffice and os.path.exists(soffice)):
        raise GradingError(f"LibreOffice not found at {soffice!r} (set DETCHECKS_SOFFICE)")
    install_termination_reaper()
    return lo_guard.run_guarded(lambda timeout_s: _libreoffice_once(src, out_dir, policy, timeout_s),
                                timeout_s=policy.lo_timeout_s, settings=policy.guard_settings(),
                                what=f"recalculate {os.path.basename(src)}", log=policy.lo_log)


def _libreoffice_once(src: str, out_dir: str, policy: RecalcPolicy, timeout_s: float) -> str:
    """One LibreOffice try (libreoffice_recalc's doc); GradingError on timeout / no copy."""
    soffice = policy.libreoffice_path
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(src))[0]
    out = os.path.join(out_dir, stem + ".xlsx")
    if os.path.exists(out):
        os.remove(out)
    prof_root = os.path.join(policy.workdir, "_lo_profiles")
    os.makedirs(prof_root, exist_ok=True)
    profile = os.path.abspath(tempfile.mkdtemp(prefix="lo_", dir=prof_root))
    run_id = object()
    _ACTIVE[run_id] = {"profile": profile, "pid": None}
    try:
        os.makedirs(os.path.join(profile, "user"), exist_ok=True)
        with open(os.path.join(profile, "user", "registrymodifications.xcu"), "w") as fh:
            fh.write(_LO_REGISTRY)
        cmd = [soffice, f"-env:UserInstallation={file_url(profile)}", "--headless", "--norestore", "--nologo",
               "--nofirststartwizard", "--calc", "--convert-to", "xlsx", "--outdir", file_url(out_dir),
               file_url(src)]
        env = dict(os.environ, SAL_USE_VCLPLUGIN="svp", OMP_NUM_THREADS="1")
        t0 = time.perf_counter()
        proc = subprocess.Popen(_watchdog_cmd(cmd), stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE, env=env)
        _ACTIVE[run_id]["pid"] = proc.pid
        try:
            so, se = proc.communicate(timeout=timeout_s)
        except subprocess.TimeoutExpired:
            _kill_tree(proc)
            raise GradingError(f"LibreOffice recalculation of {src} timed out after {timeout_s:.0f} s")
        except BaseException:                 # KeyboardInterrupt, SystemExit, ...: LibreOffice goes too
            _kill_tree(proc)
            raise
        secs = time.perf_counter() - t0
        if not os.path.exists(out):
            msg = (se or b"").decode("utf-8", "replace")[-400:] + (so or b"").decode("utf-8", "replace")[-200:]
            raise GradingError(f"LibreOffice produced no copy for {src} (exit {proc.returncode}; {secs:.1f} s): {msg.strip()}")
        return out
    finally:
        _ACTIVE.pop(run_id, None)
        reap_profiles(profile, subtree=False)  # nothing may keep running on this profile
        shutil.rmtree(profile, ignore_errors=True)


# ---------------------------------------------------------------------------- Excel
_APPLESCRIPT = r'''
on run argv
    set srcPath to item 1 of argv
    set dstPath to item 2 of argv
    set srcFile to (POSIX file srcPath) as text
    set dstFile to (POSIX file dstPath) as text
    tell application "Microsoft Excel"
        set display alerts to false
        set wb to open workbook workbook file name srcFile read only true update links do not update links
        calculate full rebuild
        save workbook as wb filename dstFile file format Excel XML file format with overwrite
        close wb saving no
    end tell
    return "ok"
end run
'''

_POWERSHELL = r'''
param([string]$src, [string]$dst)
$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.AskToUpdateLinks = $false
  $wb = $xl.Workbooks.Open($src, 0, $true)       # UpdateLinks=0, ReadOnly
  $xl.CalculateFullRebuild()
  $wb.SaveAs($dst, 51)                            # 51 = xlOpenXMLWorkbook (.xlsx)
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Write-Output "ok"
'''

_MAC_EXCEL_APPS = ("/Applications/Microsoft Excel.app", os.path.expanduser("~/Applications/Microsoft Excel.app"))


def excel_available() -> bool:
    """Is Microsoft Excel installed on this machine?  macOS: the app bundle; Windows: the
    Excel.Application COM class in the registry (UNTESTED); elsewhere: never."""
    if sys.platform == "darwin":
        return any(os.path.isdir(p) for p in _MAC_EXCEL_APPS)
    if sys.platform.startswith("win"):
        try:
            import winreg  # type: ignore
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
                return True
        except OSError:
            return False
    return False


def _excel_running_mac() -> Optional[bool]:
    try:
        r = subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True, timeout=10)
        return r.returncode == 0
    except (OSError, subprocess.TimeoutExpired):
        return None


def excel_recalc(src: str, out_dir: str, policy: RecalcPolicy) -> str:
    """Open `src` in Microsoft Excel (read-only, links not updated, alerts off), run a full
    recalculation (CalculateFullRebuild), save a copy as <out_dir>/<stem>.xlsx, close without
    saving the original.  macOS: osascript (tested here 2026-10-04); Windows: PowerShell + COM
    (UNTESTED).  If Excel was not running before, it is quit afterwards.  GradingError on
    timeout / failure / unsupported platform."""
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(src))[0]
    out = os.path.join(out_dir, stem + ".xlsx")
    if os.path.exists(out):
        os.remove(out)
    src = os.path.abspath(src)
    t0 = time.perf_counter()
    if sys.platform == "darwin":
        was_running = _excel_running_mac()
        script = _APPLESCRIPT
        if was_running is False:
            script = script.replace("        close wb saving no\n", "        close wb saving no\n        quit\n")
        cmd = ["osascript", "-"] + [src, out]
        proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                start_new_session=True)
        try:
            so, se = proc.communicate(script.encode(), timeout=policy.excel_timeout_s)
        except subprocess.TimeoutExpired:
            _kill_tree(proc)
            if was_running is False:
                subprocess.run(["pkill", "-x", "Microsoft Excel"], capture_output=True)
            raise GradingError(f"Excel recalculation of {src} timed out after {policy.excel_timeout_s:.0f} s")
        err = (se or b"").decode("utf-8", "replace").strip()
    elif sys.platform.startswith("win"):
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False) as fh:
            fh.write(_POWERSHELL)
            ps1 = fh.name
        try:
            proc = subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1, src, out],
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            try:
                so, se = proc.communicate(timeout=policy.excel_timeout_s)
            except subprocess.TimeoutExpired:
                proc.kill()
                raise GradingError(f"Excel recalculation of {src} timed out after {policy.excel_timeout_s:.0f} s")
            err = (se or b"").decode("utf-8", "replace").strip()
        finally:
            os.unlink(ps1)
    else:
        raise GradingError(f"Excel recalculation is not possible on {sys.platform}")
    secs = time.perf_counter() - t0
    if not os.path.exists(out):
        raise GradingError(f"Excel produced no copy for {src} (exit {proc.returncode}; {secs:.1f} s): {err[-400:]}")
    return out


# ---------------------------------------------------------------------------- gap detection
def classify_lo_error(text: Optional[str], value: str, defined: dict, tables: frozenset,
                      unsupported: frozenset, is_array: bool) -> tuple[bool, list, str]:
    """(gap?, functions involved, why) for a formula cell LibreOffice shows as #NAME? / #VALUE!.

    gap = Excel may well compute this cell: the formula (or a defined name it uses) calls a
    function LibreOffice lacks, a spill reference (A1#), a LAMBDA name, or - for #VALUE! on a
    plain formula - a range / array function / name whose legacy evaluation differs between
    the engines, or the text cannot be parsed.  genuine = Excel shows the error too: an unknown
    function name (misspelling), an undefined name (#NAME?), or a #VALUE! on a formula whose
    operands are single cells / literals and whose functions are all scalar.
    `defined` maps UPPER name -> definition text; `tables` holds UPPER table names."""
    if text is None:
        return True, [], "formula text unavailable (data table / empty marker): Excel's display unknown"
    try:
        f = F.parse(text)
    except GradingError as e:
        return True, [], f"formula cannot be parsed ({str(e)[:120]}): Excel's display unknown"
    bad_fn = sorted({c.name for c in f.functions if c.builtin and not c.local and c.name in unsupported})
    feats = list(bad_fn)
    if any(o.spill for o in f.operands) or any(c.spilled for c in f.functions):
        feats.append("spill reference (#)")
    for c in f.functions:
        if not c.builtin and not c.local and c.name in defined:
            feats.append(f"name {c.name} (LAMBDA)")
    seen = set()
    for o in f.operands:
        for nm in o.names:
            u = (nm or "").upper()
            if u in defined and u not in seen:
                seen.add(u)
                try:
                    fn = F.parse(defined[u]).function_names()
                except GradingError:
                    fn = set()
                sub = sorted(fn & unsupported)
                if sub or "LAMBDA" in fn:
                    feats.append(f"name {u} uses {','.join(sub or ['LAMBDA'])}")
    if feats:
        return True, feats, "LibreOffice cannot evaluate: " + ", ".join(feats)
    # a called name that is neither an Excel function (parser: builtin) nor a defined name is a
    # misspelling (CONUTIFS): #NAME? in Excel too
    unknown_fn = sorted({c.name for c in f.functions if not c.local and (
        (c.builtin and c.name not in EXCEL_FUNCTIONS)
        or (not c.builtin and (c.name or "").upper() not in defined and (c.name or "").upper() not in tables))})
    undefined = sorted({(nm or "").upper() for o in f.operands for nm in o.names
                        if (nm or "").upper() not in defined})
    undefined = [u for u in undefined if u and u not in tables]
    if value == "#NAME?":
        if unknown_fn:
            return False, unknown_fn, f"genuine #NAME?: unknown function {', '.join(unknown_fn)}"
        if undefined:
            return False, undefined, f"genuine #NAME?: undefined name {', '.join(undefined)}"
        return True, sorted(f.function_names()), ("#NAME? although every function is a valid Excel function " +
                                                 "and every name is defined: LibreOffice cannot evaluate it")
    # #VALUE!
    if unknown_fn:
        return False, unknown_fn, f"genuine: unknown function {', '.join(unknown_fn)}"
    if is_array:
        return True, sorted(f.function_names()), "#VALUE! on an array formula: engines evaluate arrays differently"
    arrayish = [o.kind for o in f.operands if o.kind in ("range", "whole_column", "whole_row", "trimmed_range",
                                                         "structured", "name", "array")]
    arr_fn = sorted(f.function_names() & ARRAY_FUNCS)
    if arrayish or arr_fn:
        what = sorted(set(arrayish)) + arr_fn
        return True, what, ("#VALUE! on a plain formula with " + ", ".join(what) +
                            ": Excel's legacy array evaluation (implicit intersection / top-left) may differ")
    return False, [], "genuine #VALUE!: scalar operands and scalar functions only"


def find_gaps(pkg: Package, source: ValueSource, lo_version: Optional[tuple]) -> tuple[list, int, int]:
    """Scan every formula cell of `pkg` against the LibreOffice copy.  Returns
    (gaps [Gap], n_lo_errors, n_genuine)."""
    unsupported = lo_unsupported(lo_version)
    defined = {}
    for d in pkg.defined_names:
        if d.name and not d.builtin:
            defined[d.name.upper()] = d.text or ""
    tables = set()
    for info in pkg.sheets:
        if info.kind == "worksheet" and info.part:
            try:
                for t in pkg.tables(info):
                    if t.name:
                        tables.add(t.name.upper())
            except Exception:  # noqa: BLE001 - a broken table part does not stop the gap scan
                pass
    tables = frozenset(tables)
    gaps: list = []
    n_err = n_gen = 0
    anchors_done: set = set()
    for info in pkg.sheets:
        if info.kind not in ("worksheet", "macrosheet", "dialogsheet") or not info.part:
            continue
        ss = SheetStream(pkg, info)
        try:
            ss.read_head()
            ctx = make_context(pkg, info, source, needs_values=True)
            state = {"n": 0}

            def on_cell(cell, _sheet=info.name):
                if not cell.is_formula_result:
                    return
                v = cell.unverified_value if cell.value_source == "recalc" else None
                if not isinstance(v, ExcelError) or str(v) not in GAP_ERRORS:
                    return
                state["n"] += 1
                if cell.array is not None:                 # member: judged once, at its anchor
                    key = (_sheet, cell.array.anchor_ref)
                    if key in anchors_done:
                        return
                    anchors_done.add(key)
                    text, is_array, ref = cell.array.formula.text if cell.array.formula else None, True, cell.array.anchor_ref
                else:
                    f = cell.formula
                    text = cell.formula_text if cell.has_formula else None
                    is_array = bool(f is not None and f.kind in ("array", "dataTable"))
                    ref = cell.ref
                    if f is not None and f.kind == "array":
                        anchors_done.add((_sheet, cell.ref))
                gap, fns, why = classify_lo_error(text, str(v), defined, tables, unsupported, is_array)
                if gap:
                    gaps.append(Gap(_sheet, ref, str(v), text or "", fns, why))
                else:
                    state["gen"] = state.get("gen", 0) + 1

            ss.read_body(None, on_cell, ctx)
            n_err += state["n"]
            n_gen += state.get("gen", 0)
        finally:
            ss.close()
    return gaps, n_err, n_gen


# ---------------------------------------------------------------------------- the pipeline
def _meta_path(cache_dir: str) -> str:
    return os.path.join(cache_dir, "meta.json")


def _load_plan(cache_dir: str) -> Optional[ValuePlan]:
    p = _meta_path(cache_dir)
    if not os.path.exists(p):
        return None
    try:
        with open(p) as fh:
            d = json.load(fh)
        if d.get("value_path") and not os.path.exists(d["value_path"]):
            return None
        plan = ValuePlan(d["source"], d.get("value_path"), d["writer"], d["file_hash"], d.get("errors_vetted", False),
                         d.get("n_gaps", 0), [Gap(**g) for g in d.get("gaps", [])], d.get("gap_functions", []),
                         d.get("n_lo_errors", 0), d.get("n_genuine_lo_errors", 0), dict(d.get("timings", {})),
                         d.get("lo_version"), list(d.get("notes", [])))
        plan.timings["from_cache"] = True
        return plan
    except (OSError, ValueError, KeyError, TypeError):
        return None


def _save_plan(cache_dir: str, plan: ValuePlan):
    os.makedirs(cache_dir, exist_ok=True)
    d = plan.summary()
    d["file_hash"] = plan.file_hash
    d["timings"] = {k: v for k, v in plan.timings.items() if k != "from_cache"}
    with open(_meta_path(cache_dir), "w") as fh:
        json.dump(d, fh, indent=1)


def _excel_needed_error(path: str, plan: ValuePlan, why: str) -> GradingError:
    cells = "; ".join(f"{g.sheet}!{g.ref} {g.value} ({', '.join(g.functions) or g.why[:60]})" for g in plan.gaps[:8])
    more = f" (+{plan.n_gaps - 8} more)" if plan.n_gaps > 8 else ""
    return GradingError(f"Excel recalculation required: {plan.n_gaps} LibreOffice gap cell(s) in {path} - "
                        f"{cells}{more} - {why}")


def ensure_values(path: str, workdir: Optional[str] = None, policy: Optional[RecalcPolicy] = None,
                  pkg: Optional[Package] = None) -> ValuePlan:
    """Decide and produce the value source for `path` (module doc).  Raises GradingError when
    the file needs Excel and Excel is not allowed / not available, or when a recalculation
    fails or times out."""
    policy = policy or RecalcPolicy()
    workdir = workdir or policy.workdir
    own = pkg is None
    if own:
        pkg = Package.open(path)
    try:
        if not pkg.is_spreadsheetml:
            raise GradingError(f"cannot recalculate a {pkg.format!r} file ({pkg.unsupported_reason})")
        prov = pkg.provenance or detect_provenance(pkg)
        h = file_hash(path)
        if prov.writer == "excel":
            plan = ValuePlan("cache", None, prov.writer, h, errors_vetted=True)
            plan.notes.append("saved by Excel: its cached values are the display")
            return plan
        cache_dir = os.path.join(workdir, h)
        plan = None
        if policy.use_cache:
            cached = _load_plan(cache_dir)
            if cached is not None and cached.writer == prov.writer:
                if cached.source == "excel" or (cached.source == "libreoffice" and not cached.n_gaps
                                                and cached.errors_vetted):
                    return cached                              # final answer from an earlier run
                if cached.source == "libreoffice" and cached.n_gaps:
                    plan = cached                              # LibreOffice copy + gap list reusable; Excel still needed
        if plan is None:
            # --- LibreOffice copy
            run_lo = policy.lo_runner or libreoffice_recalc
            t0 = time.perf_counter()
            lo_path = run_lo(path, os.path.join(cache_dir, "libreoffice"), policy)
            lo_s = round(time.perf_counter() - t0, 2)
            source = ValueSource(lo_path, errors_vetted=False)
            try:
                missing = [s.name for s in pkg.sheets if s.kind == "worksheet" and s.name not in source.sheet_names()]
                lo_app = source.pkg.app.application
                lo_ver = lo_version_of(lo_app)
                t1 = time.perf_counter()
                gaps, n_err, n_gen = find_gaps(pkg, source, lo_ver)
                scan_s = round(time.perf_counter() - t1, 2)
            finally:
                source.close()
            plan = ValuePlan("libreoffice", lo_path, prov.writer, h, errors_vetted=False, n_gaps=len(gaps),
                             gaps=gaps[:MAX_GAPS_LISTED],
                             gap_functions=sorted({fn for g in gaps for fn in g.functions}),
                             n_lo_errors=n_err, n_genuine_lo_errors=n_gen,
                             timings={"libreoffice_s": lo_s, "gap_scan_s": scan_s},
                             lo_version=lo_app)
            if missing:
                plan.notes.append(f"sheets missing from the LibreOffice copy: {missing[:5]}")
            if not gaps:
                plan.errors_vetted = True
                _save_plan(cache_dir, plan)
                return plan
        # --- LibreOffice gaps
        if not policy.excel_allowed:
            # Patrick 2026-10-04: LibreOffice only, as judge v12 effectively does. The cells LibreOffice
            # could not compute keep their LibreOffice error value and are used AS DISPLAYED: the value
            # checks see an error value (not a number, not a zero, not a negative) and skip the cell.
            # The gap list stays in the plan / stats so anyone can see which cells those were.
            plan.errors_vetted = True
            plan.notes.append(f"LibreOffice could not compute {plan.n_gaps} formula cell(s) "
                              f"({', '.join(plan.gap_functions) or 'see gaps'}); their LibreOffice error values are "
                              f"used as displayed (judge v12 behaviour); Excel recalculation is off")
            _save_plan(cache_dir, plan)
            return plan
        avail = policy.excel_available if policy.excel_available is not None else excel_available()
        if not avail:
            _save_plan(cache_dir, plan)
            raise _excel_needed_error(path, plan, "Excel is not available on this machine")
        run_xl = policy.excel_runner or excel_recalc
        t2 = time.perf_counter()
        try:
            xl_path = run_xl(path, os.path.join(cache_dir, "excel"), policy)
        except GradingError as e:
            _save_plan(cache_dir, plan)
            raise GradingError(f"Excel recalculation required ({plan.n_gaps} LibreOffice gap cell(s): "
                               f"{plan.gap_functions}) but it failed: {e}") from None
        plan.timings.pop("from_cache", None)
        plan.timings["excel_s"] = round(time.perf_counter() - t2, 2)
        xs = ValueSource(xl_path, errors_vetted=True)
        try:
            xw = xs.writer
            xmissing = [s.name for s in pkg.sheets if s.kind == "worksheet" and s.name not in xs.sheet_names()]
        finally:
            xs.close()
        if xw != "excel":
            plan.notes.append(f"Excel copy writer detected as {xw!r}")
        if xmissing:
            plan.notes.append(f"sheets missing from the Excel copy: {xmissing[:5]}")
        plan.source, plan.value_path, plan.errors_vetted = "excel", xl_path, True
        _save_plan(cache_dir, plan)
        return plan
    finally:
        if own:
            pkg.close()
