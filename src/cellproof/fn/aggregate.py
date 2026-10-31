"""AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, PRODUCT, MEDIAN, SUMPRODUCT, SUBTOTAL.

Like SUM: in a range only numbers count (text, TRUE/FALSE and empty cells are skipped); a value typed
straight into the arguments is converted (AVERAGE(1, "3") = 2). An error in a range is the result.

An array argument (SUM(A1:A3*B1:B3), SUMPRODUCT((A1:A9>5)*C1:C9)) is one relation with the cells table's
columns, a row per element position, which the aggregates read like a range: modern Excel evaluates the
arithmetic element by element, where a text element is #VALUE!, an empty one is 0 and the first error wins.
"""
from __future__ import annotations

from dataclasses import dataclass, replace

from cellproof.expr import (BLANK, CELLS, NO_ERR, Expr, errs, lit_str, num, num_text, parse_num, where_area,
                            _text_bool)
from cellproof.functions import arity, note_text_numbers, register
from cellproof.model import Unsupported
from cellproof.parser import Binary, Call, Num, RefNode, Unary
from cellproof.refs import Area, resolve


@dataclass(frozen=True)
class Rng:
    area: Area
    keys: object  # compiler.Keys
    skip: tuple = ()  # cells left out (SUBTOTAL leaves out other SUBTOTALs)

    def where(self, alias: str = "") -> str:
        p = f"{alias}." if alias else ""
        return where_area(self.area, alias) + "".join(
            f" AND NOT ({p}row_num = {r} AND {p}col_num = {col})" for _, r, col in self.skip)

    def agg(self, fn: str, kinds: str = "('n')") -> str:
        return f"(SELECT {fn} FROM {CELLS} WHERE {self.where()} AND kind IN {kinds})"

    def err(self, c) -> str:
        if self.skip:
            if not any(c.may_err(k) for k in self.keys if k not in self.skip):
                return NO_ERR
        elif not self.keys.may_err:
            return NO_ERR
        return (f"(SELECT txt FROM {CELLS} WHERE {self.where()} AND kind = 'e' "
                f"ORDER BY row_num, col_num LIMIT 1)")


@dataclass(frozen=True)
class Src:
    """One aligned range an array reads, aliased w1, w2... in the array's element SQL."""
    area: Area
    keys: object  # compiler.Keys
    alias: str


@dataclass
class Arr:
    """An element-wise computed array as the aggregates read it: sources (aligned ranges, aliased w<i>) and the
    SQL of one element's kind, number, text and error over those aliases. Scalars broadcast (no sources); sources
    of different sizes make the whole array #VALUE! (bad: an empty relation)."""
    srcs: list
    kind: str
    num: str
    txt: str
    err: str
    shape: tuple | None
    bad: bool = False

    def rel(self) -> str:
        """One subquery with the cells table's columns - row_num, col_num, kind, num, txt, err - per element."""
        if self.bad:
            return ("(SELECT * FROM (SELECT CAST(NULL AS BIGINT) AS row_num, CAST(NULL AS BIGINT) AS col_num, "
                    "CAST(NULL AS VARCHAR) AS kind, CAST(NULL AS DOUBLE) AS num, CAST(NULL AS VARCHAR) AS txt, "
                    "CAST(NULL AS VARCHAR) AS err) WHERE FALSE)")
        if not self.srcs:  # a scalar operand: one element
            return (f"(SELECT CAST(NULL AS BIGINT) AS row_num, CAST(NULL AS BIGINT) AS col_num, "
                    f"{self.kind} AS kind, {self.num} AS num, {self.txt} AS txt, {self.err} AS err)")
        first = self.srcs[0].area
        grid = (f"SELECT row_num, col_num FROM range({first.r1}, {first.r2 + 1}) AS gr(row_num), "
                f"range({first.c1}, {first.c2 + 1}) AS gc(col_num)")
        joins = "".join(
            f" LEFT JOIN {CELLS} {s.alias} ON {s.alias}.sheet = {lit_str(s.area.sheet).sql} "
            f"AND {s.alias}.row_num = g.row_num + {s.area.r1 - first.r1} "
            f"AND {s.alias}.col_num = g.col_num + {s.area.c1 - first.c1}"
            for s in self.srcs)
        return (f"(SELECT g.row_num AS row_num, g.col_num AS col_num, {self.kind} AS kind, "
                f"{self.num} AS num, {self.txt} AS txt, {self.err} AS err FROM ({grid}) g{joins})")

    def agg(self, fn: str, kinds: str = "('n')") -> str:
        return f"(SELECT {fn} FROM {self.rel()} WHERE kind IN {kinds})"

    def err_sql(self) -> str:
        """The earliest element error by position, or the whole array's #VALUE!."""
        if self.bad:
            return "'#VALUE!'"
        return f"(SELECT err FROM {self.rel()} WHERE err IS NOT NULL ORDER BY row_num, col_num LIMIT 1)"


def _base(c, node, sheet: str, fname: str) -> Arr:
    area = resolve(node.text, sheet, c.book)
    if c.rowrefs is not None and node.text in c.rowrefs:
        raise Unsupported(f"{fname} array over the region's own rows not supported yet")
    keys = c.ready(area)
    alias = c.walias()
    return Arr([Src(area, keys, alias)], f"{alias}.kind", f"{alias}.num", f"{alias}.txt",
               f"CASE WHEN {alias}.kind = 'e' THEN {alias}.txt END",
               (area.r2 - area.r1, area.c2 - area.c1))


def _merged(l: Arr, r: Arr, kind: str, num: str, txt: str, err: str) -> Arr:
    srcs = l.srcs + list(r.srcs)
    shapes = {(s.area.r2 - s.area.r1, s.area.c2 - s.area.c1) for s in srcs}
    if l.bad or r.bad or (srcs and len(shapes) != 1):  # no sources: a scalar broadcast, never a shape clash
        return Arr(srcs, kind, num, txt, err, None, True)
    return Arr(srcs, kind, num, txt, err, shapes.pop() if srcs else None)


def _propagated(l: Arr, r: Arr) -> str:
    return (f"CASE WHEN {l.err} IS NOT NULL THEN {l.err} WHEN {r.err} IS NOT NULL THEN {r.err} END")


def _element_number(a: Arr) -> str:
    """One element as Excel's arithmetic reads it: numeric text converts ("3" -> 3), other text and errors are
    NULL (the #VALUE! then comes from the error CASE), empty is 0."""
    return (f"CASE WHEN {a.kind} = 's' THEN {parse_num(f'COALESCE({a.txt}, ' + chr(39) * 2 + ')')} "
            f"WHEN {a.kind} = 'e' THEN NULL ELSE COALESCE({a.num}, 0) END")


def _arith(l: Arr, r: Arr, op: str) -> Arr:
    lnum, rnum = _element_number(l), _element_number(r)
    err = (f"CASE WHEN {l.err} IS NOT NULL THEN {l.err} WHEN {r.err} IS NOT NULL THEN {r.err} "
           f"WHEN {lnum} IS NULL THEN '#VALUE!' WHEN {rnum} IS NULL THEN '#VALUE!'"
           + (f" WHEN COALESCE({rnum}, 0) = 0 THEN '#DIV/0!'" if op == "/" else "") + " END")
    val = f"({lnum} / NULLIF({rnum}, 0))" if op == "/" else f"({lnum} {op} {rnum})"
    num = f"CASE WHEN {err} IS NOT NULL THEN NULL ELSE {val} END"
    return _merged(l, r, "'n'", num, "CAST(NULL AS VARCHAR)", err)


def _compare(l: Arr, r: Arr, op: str) -> Arr:
    """Excel's comparison per element: text case-insensitively, and across types numbers < text < TRUE/FALSE."""
    err = _propagated(l, r)
    cmp = (f"CASE WHEN {l.kind} = 's' AND {r.kind} = 's' "
           f"THEN (LOWER(COALESCE({l.txt}, '')) {op} LOWER(COALESCE({r.txt}, ''))) "
           f"WHEN {l.kind} = 's' AND {r.kind} = 'b' THEN (1 {op} 2) "
           f"WHEN {l.kind} = 'b' AND {r.kind} = 's' THEN (2 {op} 1) "
           f"WHEN {l.kind} = 's' OR {r.kind} = 's' THEN (1 {op} 0) "  # text against a number or an empty cell
           f"ELSE (COALESCE({l.num}, 0) {op} COALESCE({r.num}, 0)) END")
    num = f"CASE WHEN {err} IS NOT NULL THEN NULL WHEN {cmp} THEN 1 ELSE 0 END"
    return _merged(l, r, "'b'", num, "CAST(NULL AS VARCHAR)", err)


def _power_array(l: Arr, r: Arr) -> Arr:
    """Excel's ^ per element, with its guards: the real cube root of a negative base, #NUM! for the rest of the
    fractional powers of one, #DIV/0! for 0 to a negative, #NUM! for 0^0 (compiler.binary ports the same rules)."""
    lnum, rnum = _element_number(l), _element_number(r)
    root = f"(1 / NULLIF({rnum}, 0))"
    odd_root = (f"({lnum} < 0 AND {rnum} <> FLOOR({rnum}) AND abs({root} - round({root})) < 1e-9 "
                f"AND abs(round({root})) % 2 = 1)")
    err = (f"CASE WHEN {l.err} IS NOT NULL THEN {l.err} WHEN {r.err} IS NOT NULL THEN {r.err} "
           f"WHEN {lnum} IS NULL THEN '#VALUE!' WHEN {rnum} IS NULL THEN '#VALUE!' "
           f"WHEN {lnum} = 0 AND {rnum} < 0 THEN '#DIV/0!' "
           f"WHEN {lnum} = 0 AND {rnum} = 0 THEN '#NUM!' "
           f"WHEN {lnum} < 0 AND {rnum} <> FLOOR({rnum}) AND NOT {odd_root} THEN '#NUM!' END")
    val = f"(CASE WHEN {odd_root} THEN -POWER(-{lnum}, {rnum}) ELSE POWER({lnum}, {rnum}) END)"
    return _merged(l, r, "'n'", f"CASE WHEN {err} IS NOT NULL THEN NULL ELSE {val} END",
                   "CAST(NULL AS VARCHAR)", err)


def _element_text(a: Arr) -> str:
    """One element as Excel writes it in a concatenation: TRUE/FALSE, numbers by the 15-digit rule, empty ''."""
    return (f"CASE WHEN {a.kind} = 's' THEN COALESCE({a.txt}, '') "
            f"WHEN {a.kind} = 'b' THEN (CASE WHEN COALESCE({a.num}, 0) <> 0 THEN 'TRUE' ELSE 'FALSE' END) "
            f"WHEN {a.kind} IS NULL THEN '' ELSE {num_text(f'COALESCE({a.num}, 0)')} END")


def _concat_array(l: Arr, r: Arr) -> Arr:
    return _merged(l, r, "'s'", "CAST(NULL AS DOUBLE)",
                   f"({_element_text(l)} || {_element_text(r)})", _propagated(l, r))


def _iferror_array(c, args, sheet: str, fname: str, na_only: bool) -> Arr:
    """IFERROR/IFNA over arrays: the fallback element replaces every erroring element (IFNA: only #N/A)."""
    if len(args) != 2 or args[1] is None:
        raise Unsupported(f"{fname} needs two arguments")
    value = _array(c, args[0], sheet, fname, "raw")
    fallback = _array(c, args[1], sheet, fname, "raw")
    shapes = {a.shape for a in (value, fallback) if a.shape is not None}
    if len(shapes) > 1:
        raise Unsupported(f"{fname} arrays of different sizes not supported yet")
    swap = f"{value.err} = '#N/A'" if na_only else f"{value.err} IS NOT NULL"
    kind = f"CASE WHEN {swap} THEN {fallback.kind} ELSE {value.kind} END"
    num = f"CASE WHEN {swap} THEN {fallback.num} ELSE {value.num} END"
    txt = f"CASE WHEN {swap} THEN {fallback.txt} ELSE {value.txt} END"
    return _merged(value, fallback, kind, num, txt, f"CASE WHEN {swap} THEN {fallback.err} ELSE NULL END")


ARRAY_OPS = {"+", "-", "*", "/", "=", "<>", "<", ">", "<=", ">="}


def _scalar_element(e: Expr) -> Arr:
    """One element from a plain value, keeping its kind (comparisons rank by type; arithmetic converts below)."""
    if e.type == "blank":
        return Arr([], "CAST(NULL AS VARCHAR)", "CAST(NULL AS DOUBLE)", "CAST(NULL AS VARCHAR)", e.err, None)
    if e.type == "txt":
        return Arr([], "'s'", "CAST(NULL AS DOUBLE)", e.sql, e.err, None)
    if e.type == "bool":
        return Arr([], "'b'", f"CAST({e.sql} AS DOUBLE)", "CAST(NULL AS VARCHAR)", e.err, None)
    if e.type == "err":
        return Arr([], "'e'", "CAST(NULL AS DOUBLE)", "CAST(NULL AS VARCHAR)", e.err, None)
    if e.type == "any":
        return Arr([], e.kind, e.sql, e.txt, e.err, None)
    return Arr([], "'n'", e.sql, "CAST(NULL AS VARCHAR)", e.err, None)


def _array(c, node, sheet: str, fname: str, leaf: str = "num") -> Arr:
    """An AST node as an element-wise array: *, +, -, /, comparisons, unary minus and IF over aligned ranges,
    array constants and scalars. Anything else is Unsupported - the cell is UNRESOLVED, never intersected.
    A scalar leaf keeps its kind where the value passes through (comparisons rank by type, IF returns it) and
    is converted where arithmetic reads it ("3" -> 3)."""
    if isinstance(node, Binary):
        if node.op == "&":
            return _concat_array(_array(c, node.left, sheet, fname, "raw"),
                                 _array(c, node.right, sheet, fname, "raw"))
        if node.op == "^":
            return _power_array(_array(c, node.left, sheet, fname, "num"),
                                _array(c, node.right, sheet, fname, "num"))
        if node.op not in ARRAY_OPS:
            raise Unsupported(f"{fname} cannot apply '{node.op}' to arrays yet")
        if node.op in ("=", "<>", "<", ">", "<=", ">="):
            return _compare(_array(c, node.left, sheet, fname, "raw"),
                            _array(c, node.right, sheet, fname, "raw"), node.op)
        return _arith(_array(c, node.left, sheet, fname, "num"),
                      _array(c, node.right, sheet, fname, "num"), node.op)
    if isinstance(node, Call):
        if node.name == "IF":
            return _if_array(c, node.args, sheet, fname)
        if node.name in ("IFERROR", "IFNA"):
            return _iferror_array(c, node.args, sheet, fname, node.name == "IFNA")
        raise Unsupported(f"{fname} cannot apply {node.name} to arrays yet")
    if isinstance(node, Unary) and node.op == "-":
        zero = Arr([], "'n'", "CAST(0 AS DOUBLE)", "CAST(NULL AS VARCHAR)", "NULL", None)
        return _arith(zero, _array(c, node.operand, sheet, fname, "num"), "-")
    if isinstance(node, RefNode) and not resolve(node.text, sheet, c.book).is_single:
        return _base(c, node, sheet, fname)
    e = c.expr(node, sheet)  # a scalar: a number, one cell or an expression of them, one element for every position
    return _scalar_element(e) if leaf == "raw" else _scalar_element(num(e))


def _truth(kind: str, num: str, txt: str, err: str) -> tuple[str, str]:
    """(condition, error) of one element as Excel's IF reads it: TRUE/FALSE text, numbers <> 0, empty FALSE."""
    b = _text_bool(f"COALESCE({txt}, '')")
    cond = (f"CASE WHEN {err} IS NOT NULL THEN NULL WHEN {kind} = 's' THEN {b} "
            f"WHEN {kind} IS NULL THEN FALSE ELSE ({num} <> 0) END")
    cerr = f"CASE WHEN {err} IS NOT NULL THEN {err} WHEN {kind} = 's' AND {b} IS NULL THEN '#VALUE!' END"
    return cond, cerr


def _if_array(c, args, sheet: str, fname: str) -> Arr:
    """IF over arrays: the condition picks one element per position. The untaken branch's error never shows
    (only computed away); a missing third argument gives FALSE elements, which MAX and friends skip."""
    if not 2 <= len(args) <= 3:
        raise Unsupported(f"{fname} needs two or three arguments")
    cond = _array(c, args[0], sheet, fname, "raw")
    yes = _array(c, args[1], sheet, fname, "raw")
    no = (_array(c, args[2], sheet, fname, "raw") if len(args) == 3
          else Arr([], "'b'", "CAST(0 AS DOUBLE)", "CAST(NULL AS VARCHAR)", "NULL", None))
    shapes = {a.shape for a in (cond, yes, no) if a.shape is not None}
    if len(shapes) > 1:
        raise Unsupported(f"{fname} arrays of different sizes not supported yet")
    taken, cerr = _truth(cond.kind, cond.num, cond.txt, cond.err)
    kind = f"CASE WHEN {cerr} IS NOT NULL THEN 'e' WHEN {taken} THEN {yes.kind} ELSE {no.kind} END"
    num = f"CASE WHEN {cerr} IS NOT NULL THEN NULL WHEN {taken} THEN {yes.num} ELSE {no.num} END"
    txt = f"CASE WHEN {cerr} IS NOT NULL THEN NULL WHEN {taken} THEN {yes.txt} ELSE {no.txt} END"
    err = f"CASE WHEN {cerr} IS NOT NULL THEN {cerr} WHEN {taken} THEN {yes.err} ELSE {no.err} END"
    return _merged(_merged(cond, yes, kind, num, txt, err), no, kind, num, txt, err)


def _rng(c, node, sheet: str, skip_subtotals: bool = False) -> Rng:
    area = resolve(node.text, sheet, c.book)
    keys = c.ready(area)
    note_text_numbers(c, keys)
    skip = ()
    if skip_subtotals:
        skip = tuple(k for k in keys if "SUBTOTAL(" in (c.book.cells[k].formula or "").upper())
    return Rng(area, keys, skip)


def _split(c, args, sheet: str, fname: str, skip_subtotals: bool = False) -> tuple[list, list, list]:
    """The arguments as (ranges, arrays, values). A range arithmetic or an IF over ranges is an array; a
    calculation over plain values stays a value, evaluated before the cells are read (MAX(H3, 1/0) is #DIV/0!)."""
    from cellproof.fn.dynamic import DYNAMIC, as_array
    ranges, arrays, values = [], [], []
    for a in args:
        if isinstance(a, Call) and a.name in DYNAMIC:  # SUM(FILTER(...)): the relation, read like an array
            arrays.append(as_array(c, a, sheet))
        elif isinstance(a, RefNode):
            ranges.append(_rng(c, a, sheet, skip_subtotals))
        elif isinstance(a, Binary) or (isinstance(a, Unary) and a.op == "-") or (
                isinstance(a, Call) and a.name in ("IF", "IFERROR", "IFNA")):
            arr = _array(c, a, sheet, fname)
            if arr.srcs:
                arrays.append(arr)
            else:
                values.append(num(c.expr(a, sheet)))
        else:
            values.append(num(c.expr(a, sheet) if a is not None else BLANK))
    return ranges, arrays, values


def _errs(c, ranges, arrays, values) -> str:
    # Real Excel: MAX(H3, 1/0) is #DIV/0! although H3 holds #N/A. Typed-in expressions are evaluated before the
    # function reads any cell, so their errors come first; then the arrays' elements, then the ranges', in order.
    return errs(*(v.err for v in values), *(a.err_sql() for a in arrays), *(r.err(c) for r in ranges))


def _average(c, ranges, arrays, values) -> Expr:
    total = " + ".join([*(r.agg("COALESCE(SUM(num), 0)") for r in ranges),
                        *(f"COALESCE({a.agg('SUM(num)')}, 0)" for a in arrays),
                        *(v.sql for v in values)]) or "0"
    count = " + ".join([*(r.agg("COUNT(*)") for r in ranges),
                        *(a.agg("COUNT(*)") for a in arrays), *("1" for _ in values)]) or "0"
    return Expr(f"(({total}) / NULLIF({count}, 0))", "num",
                errs(_errs(c, ranges, arrays, values), f"(CASE WHEN ({count}) = 0 THEN '#DIV/0!' END)"))


def _extreme(c, ranges, arrays, values, how: str) -> Expr:
    parts = [*(r.agg(f"{how}(num)") for r in ranges), *(a.agg(f"{how}(num)") for a in arrays),
             *(v.sql for v in values)]
    fn = "LEAST" if how == "MIN" else "GREATEST"  # both skip NULLs in DuckDB; no numbers at all -> 0, like Excel
    return Expr(f"COALESCE({fn}({', '.join(parts)}), CAST(0 AS DOUBLE))", "num", _errs(c, ranges, arrays, values))


def _count(c, ranges, arrays, values) -> Expr:
    if values:
        raise Unsupported("COUNT of values typed into the formula not supported yet")
    parts = [*(r.agg("CAST(COUNT(*) AS DOUBLE)") for r in ranges),
             *(a.agg("CAST(COUNT(*) AS DOUBLE)") for a in arrays)]
    return Expr("(" + " + ".join(parts) + ")", "num")


def _counta(c, ranges, arrays, values) -> Expr:
    parts = [*(r.agg("CAST(COUNT(*) AS DOUBLE)", "('n', 's', 'b', 'e')") for r in ranges),
             *(a.agg("CAST(COUNT(*) AS DOUBLE)", "('n', 's', 'b', 'e')") for a in arrays)]
    return Expr("CAST((" + " + ".join(parts + ["1" for _ in values]) + ") AS DOUBLE)", "num")


def _product(c, ranges, arrays, values) -> Expr:
    parts = [*(f"COALESCE({r.agg('product(num)')}, 1)" for r in ranges),
             *(f"COALESCE({a.agg('product(num)')}, 1)" for a in arrays), *(v.sql for v in values)]
    count = " + ".join([*(r.agg("COUNT(*)") for r in ranges), *(a.agg("COUNT(*)") for a in arrays),
                        *("1" for _ in values)])
    return Expr(f"(CASE WHEN ({count}) = 0 THEN CAST(0 AS DOUBLE) ELSE {' * '.join(parts)} END)", "num",
                _errs(c, ranges, arrays, values))


def _sum(c, ranges, arrays, values) -> Expr:
    parts = [*(r.agg("COALESCE(SUM(num), 0)") for r in ranges),
             *(f"COALESCE({a.agg('SUM(num)')}, 0)" for a in arrays), *(v.sql for v in values)]
    return Expr("(" + " + ".join(parts) + ")" if parts else "CAST(0 AS DOUBLE)", "num",
                _errs(c, ranges, arrays, values))


def _plain(name, how):
    def fn(c, args, sheet):
        if not args:
            raise Unsupported(f"{name} needs arguments")
        ranges, arrays, values = _split(c, args, sheet, name)
        result = how(c, ranges, arrays, values)
        if name not in ("COUNT", "COUNTA", "SUM"):  # SUM's error order is already the rule inside _sum
            result = replace(result, err=errs(_errs(c, ranges, arrays, values), result.err))
        markers = [a.unchecked() for a in arrays if hasattr(a, "unchecked")]  # FILTER/SORT/UNIQUE: never hidden
        if markers:
            result = replace(result, err=errs(*markers, result.err))
        return result
    register(name)(fn)


_plain("AVERAGE", _average)
_plain("MIN", lambda c, r, a, v: _extreme(c, r, a, v, "MIN"))
_plain("MAX", lambda c, r, a, v: _extreme(c, r, a, v, "MAX"))
_plain("COUNT", _count)
_plain("COUNTA", _counta)
_plain("PRODUCT", _product)
_plain("SUM", _sum)  # the error order inside _sum is already the rule: typed, then arrays, then ranges


@register("COUNTBLANK")
def countblank(c, args, sheet):
    arity("COUNTBLANK", args, 1, 1)
    if not isinstance(args[0], RefNode):
        raise Unsupported("COUNTBLANK needs a cell range")
    r = _rng(c, args[0], sheet)
    a = r.area
    size = (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1)
    filled = (f"(SELECT COUNT(*) FROM {CELLS} WHERE {r.where()} "
              f"AND NOT (kind = 's' AND COALESCE(txt, '') = ''))")  # "" counts as blank here, like Excel
    return Expr(f"CAST({size} - {filled} AS DOUBLE)", "num")


@register("MEDIAN")
def median(c, args, sheet):
    if not args or not all(isinstance(a, RefNode) for a in args):
        raise Unsupported("MEDIAN of values typed into the formula not supported yet")
    ranges = [_rng(c, a, sheet) for a in args]
    rows = " UNION ALL ".join(f"SELECT num FROM {CELLS} WHERE {r.where()} AND kind = 'n'" for r in ranges)
    values = f"FROM ({rows}) median_values"
    return Expr(f"(SELECT CAST(median(num) AS DOUBLE) {values})", "num",
                errs(_errs(c, ranges, [], []), f"(CASE WHEN (SELECT COUNT(*) {values}) = 0 THEN '#NUM!' END)"))


@register("MODE")
@register("MODE.SNGL")
def mode(c, args, sheet):
    """The number that occurs most often; on a tie the one seen first (argument order, then row by row); #N/A when
    no number repeats. Text, TRUE/FALSE and empty cells are skipped."""
    if not args or not all(isinstance(a, RefNode) for a in args):
        raise Unsupported("MODE of values typed into the formula not supported yet")
    ranges = [_rng(c, a, sheet) for a in args]
    rows = " UNION ALL ".join(
        f"SELECT num, CAST({i} AS BIGINT) * 1099511627776 + CAST(row_num AS BIGINT) * 16384 + col_num AS pos "
        f"FROM {CELLS} WHERE {r.where()} AND kind = 'n'" for i, r in enumerate(ranges))
    counted = f"(SELECT num, COUNT(*) AS n, MIN(pos) AS first FROM ({rows}) mode_values GROUP BY num HAVING COUNT(*) > 1)"
    found = f"(SELECT arg_max(num, (n, -first)) FROM {counted} mode_counts)"
    return Expr(found, "num", errs(_errs(c, ranges, [], []), f"(CASE WHEN {found} IS NULL THEN '#N/A' END)"))


@register("SUMPRODUCT")
def sumproduct(c, args, sheet):
    if not args:
        raise Unsupported("SUMPRODUCT needs arguments")
    if all(isinstance(a, RefNode) for a in args):
        ranges = [_rng(c, a, sheet) for a in args]
        first = ranges[0].area
        shape = (first.r2 - first.r1, first.c2 - first.c1)
        if any((r.area.r2 - r.area.r1, r.area.c2 - r.area.c1) != shape for r in ranges):
            return Expr("NULL", "num", errs(_errs(c, ranges, [], []), "'#VALUE!'"))  # an error inside first, then sizes
        joins = "".join(
            f" JOIN {CELLS} a{i} ON a{i}.sheet = {lit_str(r.area.sheet).sql} "
            f"AND a{i}.row_num = a0.row_num + {r.area.r1 - first.r1} AND a{i}.col_num = a0.col_num + {r.area.c1 - first.c1} "
            f"AND a{i}.kind = 'n'" for i, r in enumerate(ranges[1:], 1))
        product = " * ".join(f"a{i}.num" for i in range(len(ranges)))  # blanks, text and TRUE/FALSE count as 0
        return Expr(f"(SELECT COALESCE(SUM({product}), CAST(0 AS DOUBLE)) FROM {CELLS} a0{joins} "
                    f"WHERE {where_area(first, 'a0')} AND a0.kind = 'n')", "num", _errs(c, ranges, [], []))
    # array arguments (SUMPRODUCT((A1:A9>5)*B1:B9), mixed with ranges): one relation per argument, aligned
    rels, errs_parts, shape = [], [], None
    for a in args:
        arr = _array(c, a, sheet, "SUMPRODUCT") if not isinstance(a, RefNode) else _base(c, a, sheet, "SUMPRODUCT")
        if shape is None:
            shape = arr.shape
        elif arr.shape is not None and arr.shape != shape:
            shape = "mismatch"
        rels.append(arr)
        errs_parts.append(arr.err_sql())
    if shape == "mismatch":
        return Expr("NULL", "num", errs(*errs_parts, "'#VALUE!'"))
    terms = " * ".join(
        f"({arr.num})" if not arr.srcs else f"CASE WHEN p{i}.kind = 'n' THEN p{i}.num ELSE 0 END"
        for i, arr in enumerate(rels))
    joins = "".join(f" LEFT JOIN {rels[i].rel()} p{i} ON p{i}.row_num = p0.row_num AND p{i}.col_num = p0.col_num"
                    for i in range(1, len(rels)))
    return Expr(f"(SELECT COALESCE(SUM({terms}), CAST(0 AS DOUBLE)) FROM {rels[0].rel()} p0{joins})",
                "num", errs(*errs_parts))


SUBTOTALS = {1: _average, 2: _count, 3: _counta, 4: lambda c, r, a, v: _extreme(c, r, a, v, "MAX"),
             5: lambda c, r, a, v: _extreme(c, r, a, v, "MIN"), 6: _product, 9: _sum}


@register("SUBTOTAL")
def subtotal(c, args, sheet):
    if len(args) < 2 or not isinstance(args[0], Num):
        raise Unsupported("SUBTOTAL needs a function number and ranges")
    code = int(args[0].value)
    if code > 100:
        raise Unsupported("SUBTOTAL that ignores hidden rows not supported yet")
    if code not in SUBTOTALS:
        raise Unsupported(f"SUBTOTAL function {code} not supported yet")
    if not all(isinstance(a, RefNode) for a in args[1:]):
        raise Unsupported("SUBTOTAL needs cell ranges")
    ranges = [_rng(c, a, sheet, skip_subtotals=True) for a in args[1:]]
    return SUBTOTALS[code](c, ranges, [], [])
