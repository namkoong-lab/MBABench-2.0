"""Formula tokenization, transformation, and evaluation."""
import ast
import math
import statistics
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string


def _is_cell_ref(tok: str) -> bool:
    import re
    return re.fullmatch(r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}", tok) is not None


def _tokenize(expr: str) -> List[str]:
    tokens: List[str] = []
    i = 0
    n = len(expr)
    while i < n:
        ch = expr[i]
        if ch.isspace():
            i += 1
            continue
        if ch in ('"', "'"):
            quote = ch
            j = i + 1
            buf = [ch]
            while j < n:
                c = expr[j]
                buf.append(c)
                if c == quote and expr[j - 1] != '\\':
                    j += 1
                    break
                j += 1
            tokens.append(''.join(buf))
            i = j
            continue
        if ch.isdigit() or (ch == '.' and i + 1 < n and expr[i + 1].isdigit()):
            j = i + 1
            while j < n and (expr[j].isdigit() or expr[j] == '.'):
                j += 1
            tokens.append(expr[i:j])
            i = j
            continue
        if ch.isalpha() or ch in ('$', '_'):
            j = i + 1
            while j < n and (expr[j].isalnum() or expr[j] in ('_', '$')):
                j += 1
            tokens.append(expr[i:j])
            i = j
            continue
        if i + 1 < n and expr[i:i + 2] in ('<=', '>=', '<>', '=='):
            tokens.append(expr[i:i + 2])
            i += 2
            continue
        if ch in ('+', '-', '*', '/', '^', '(', ')', ',', ':', '<', '>', '='):
            tokens.append(ch)
            i += 1
            continue
        tokens.append(ch)
        i += 1
    return tokens


def _transform_formula(expr: str) -> str:
    expr = expr.strip()
    if expr.startswith('='):
        expr = expr[1:]
    tokens = _tokenize(expr)

    tokens = ['**' if t == '^' else t for t in tokens]
    tokens = ['!=' if t == '<>' else t for t in tokens]

    out: List[str] = []
    i = 0
    while i < len(tokens):
        t = tokens[i]
        if i + 2 < len(tokens) and _is_cell_ref(tokens[i]) and tokens[i + 1] == ':' and _is_cell_ref(tokens[i + 2]):
            a, b = tokens[i], tokens[i + 2]
            out.append(f'RNG("{a}","{b}")')
            i += 3
            continue
        out.append(t)
        i += 1

    final: List[str] = []
    for t in out:
        if _is_cell_ref(t):
            final.append(f'REF("{t}")')
        else:
            final.append(t)

    converted: List[str] = []
    for idx, t in enumerate(final):
        if t == '=':
            converted.append('==')
        else:
            converted.append(t)

    return ''.join(converted)


def _eval_formula(expr: str, wb: Workbook, ws: Worksheet, _cache: Optional[dict] = None) -> Optional[Any]:
    """Evaluate a formula with support for references, ranges, and common functions."""
    if not isinstance(expr, str) or not expr.lstrip().startswith('='):
        return None
    py_expr = _transform_formula(expr)

    def ref(addr: str):
        a = addr.replace('$', '')
        try:
            cell = ws[a]
            key = (ws.title, a)
            if _cache is not None and key in _cache:
                return _cache[key]
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                v = _eval_formula(val, wb, ws, _cache)
            else:
                v = val
            if _cache is not None:
                _cache[key] = v
            return v
        except Exception:
            return None

    def rng(a: str, b: str):
        a_clean, b_clean = a.replace('$', ''), b.replace('$', '')
        (col_a, row_a) = coordinate_from_string(a_clean)
        (col_b, row_b) = coordinate_from_string(b_clean)
        c1 = column_index_from_string(col_a)
        c2 = column_index_from_string(col_b)
        r1, r2 = int(row_a), int(row_b)
        values = []
        for r in range(min(r1, r2), max(r1, r2) + 1):
            for c in range(min(c1, c2), max(c1, c2) + 1):
                coord = f"{openpyxl.utils.get_column_letter(c)}{r}"
                values.append(ref(coord))
        return values

    def _flatten(args):
        for x in args:
            if isinstance(x, (list, tuple)):
                for y in _flatten(x):
                    yield y
            else:
                yield x

    def _nums(iterable):
        for x in iterable:
            if isinstance(x, (int, float)) and not isinstance(x, bool):
                yield x

    def SUM(*args):
        return sum(_nums(_flatten(args)))

    def AVERAGE(*args):
        arr = list(_nums(_flatten(args)))
        return (sum(arr) / len(arr)) if arr else None

    def MEAN(*args):
        return AVERAGE(*args)

    def MIN(*args):
        arr = list(_nums(_flatten(args)))
        return min(arr) if arr else None

    def MAX(*args):
        arr = list(_nums(_flatten(args)))
        return max(arr) if arr else None

    def COUNT(*args):
        return sum(1 for _ in _nums(_flatten(args)))

    def COUNTA(*args):
        return sum(1 for x in _flatten(args) if x not in (None, ''))

    def STDEV(*args):
        arr = list(_nums(_flatten(args)))
        if len(arr) >= 2:
            return statistics.stdev(arr)
        return None

    def STDEVP(*args):
        arr = list(_nums(_flatten(args)))
        if len(arr) >= 1:
            return statistics.pstdev(arr)
        return None

    def MEDIAN(*args):
        arr = list(_nums(_flatten(args)))
        return statistics.median(arr) if arr else None

    def SUMPRODUCT(*args):
        arrays = [list(_nums(a if isinstance(a, (list, tuple)) else [a])) for a in args]
        if not arrays:
            return 0
        length = min(len(a) for a in arrays)
        total = 0
        for i in range(length):
            prod = 1
            for arr in arrays:
                prod *= arr[i]
            total += prod
        return total

    def ABS(x):
        return abs(x)

    def SQRT(x):
        return math.sqrt(x)

    def POWER(x, y):
        return math.pow(x, y)

    def ROUND(x, n=0):
        try:
            return round(x, int(n))
        except Exception:
            return round(x)

    def LOG(x, base=10):
        return math.log(x, base)

    def LN(x):
        return math.log(x)

    def EXP(x):
        return math.exp(x)

    def IF(cond, a, b):
        return a if bool(cond) else b

    def AND(*args):
        return all(bool(x) for x in args)

    def OR(*args):
        return any(bool(x) for x in args)

    env = {
        'REF': ref,
        'RNG': rng,
        'SUM': SUM, 'AVERAGE': AVERAGE, 'MEAN': MEAN, 'MIN': MIN, 'MAX': MAX,
        'COUNT': COUNT, 'COUNTA': COUNTA,
        'STDEV': STDEV, 'STDEVP': STDEVP, 'MEDIAN': MEDIAN,
        'SUMPRODUCT': SUMPRODUCT,
        'ABS': ABS, 'SQRT': SQRT, 'POWER': POWER, 'ROUND': ROUND,
        'LOG': LOG, 'LN': LN, 'EXP': EXP,
        'IF': IF, 'AND': AND, 'OR': OR,
    }
    env.update({k.lower(): v for k, v in list(env.items())})

    allowed_nodes = (
        ast.Expression, ast.BinOp, ast.UnaryOp, ast.BoolOp, ast.Compare,
        getattr(ast, 'Num', ast.Constant), ast.Constant, ast.Call, ast.Name,
        ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Pow, ast.FloorDiv, ast.Mod, ast.USub, ast.UAdd,
        ast.And, ast.Or, ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE,
        ast.Load,
    )

    node = ast.parse(py_expr, mode='eval')

    def _safe(n):
        if not isinstance(n, allowed_nodes):
            raise ValueError('unsupported expression')
        for child in ast.iter_child_nodes(n):
            _safe(child)

    _safe(node)

    def _eval(n):
        if isinstance(n, ast.Expression):
            return _eval(n.body)
        if isinstance(n, ast.Constant):
            return n.value
        if hasattr(ast, 'Num') and isinstance(n, ast.Num):
            return n.n
        if isinstance(n, ast.UnaryOp):
            val = _eval(n.operand)
            if isinstance(n.op, ast.UAdd):
                return +val
            if isinstance(n.op, ast.USub):
                return -val
            raise ValueError('unsupported unary')
        if isinstance(n, ast.BinOp):
            left, right = _eval(n.left), _eval(n.right)
            if isinstance(n.op, ast.Add):
                return left + right
            if isinstance(n.op, ast.Sub):
                return left - right
            if isinstance(n.op, ast.Mult):
                return left * right
            if isinstance(n.op, ast.Div):
                return left / right
            if isinstance(n.op, ast.Pow):
                return left ** right
            if isinstance(n.op, ast.FloorDiv):
                return left // right
            if isinstance(n.op, ast.Mod):
                return left % right
            raise ValueError('unsupported op')
        if isinstance(n, ast.BoolOp):
            vals = [_eval(v) for v in n.values]
            if isinstance(n.op, ast.And):
                return all(vals)
            if isinstance(n.op, ast.Or):
                return any(vals)
            raise ValueError('unsupported boolop')
        if isinstance(n, ast.Compare):
            left = _eval(n.left)
            result = True
            for op, comp in zip(n.ops, n.comparators):
                right = _eval(comp)
                if isinstance(op, ast.Eq):
                    ok = left == right
                elif isinstance(op, ast.NotEq):
                    ok = left != right
                elif isinstance(op, ast.Lt):
                    ok = left < right
                elif isinstance(op, ast.LtE):
                    ok = left <= right
                elif isinstance(op, ast.Gt):
                    ok = left > right
                elif isinstance(op, ast.GtE):
                    ok = left >= right
                else:
                    raise ValueError('unsupported compare')
                if not ok:
                    result = False
                    break
                left = right
            return result
        if isinstance(n, ast.Call):
            if isinstance(n.func, ast.Name):
                fname = n.func.id
                fn = env.get(fname) or env.get(fname.upper()) or env.get(fname.lower())
                if not fn:
                    raise ValueError(f'unsupported function: {fname}')
                args = [_eval(a) for a in n.args]
                return fn(*args)
            raise ValueError('unsupported call')
        if isinstance(n, ast.Name):
            raise ValueError('unsupported name')
        raise ValueError('unsupported expression')

    cache = _cache if _cache is not None else {}
    return _eval(node)
