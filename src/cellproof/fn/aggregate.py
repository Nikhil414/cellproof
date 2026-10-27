"""AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, PRODUCT, MEDIAN, SUMPRODUCT, SUBTOTAL.

Like SUM: in a range only numbers count (text, TRUE/FALSE and empty cells are skipped); a value typed
straight into the arguments is converted (AVERAGE(1, "3") = 2). An error in a range is the result.

An array argument (SUM(A1:A3*B1:B3), SUMPRODUCT((A1:A9>5)*C1:C9)) is one relation with the cells table's
columns, a row per element position, which the aggregates read like a range: modern Excel evaluates the
arithmetic element by element, where a text element is #VALUE!, an empty one is 0 and the first error wins.
"""
from __future__ import annotations

from dataclasses import dataclass, replace

from cellproof.expr import BLANK, CELLS, NO_ERR, Expr, errs, lit_str, num, where_area
from cellproof.functions import arity, note_text_numbers, register
from cellproof.model import Unsupported
from cellproof.parser import Binary, Num, RefNode, Unary
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

    def numeric(self) -> bool:
        """Errors ride in the element error; text and TRUE/FALSE need Excel's collation per element, not built yet."""
        return self.keys.kinds <= {"n", "blank", "e"}


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


def _arith(l: Arr, r: Arr, op: str) -> Arr:
    err = (f"CASE WHEN {l.err} IS NOT NULL THEN {l.err} WHEN {r.err} IS NOT NULL THEN {r.err} "
           f"WHEN {l.kind} = 's' OR {r.kind} = 's' THEN '#VALUE!'"
           + (f" WHEN COALESCE({r.num}, 0) = 0 THEN '#DIV/0!'" if op == "/" else "") + " END")
    lnum, rnum = f"COALESCE({l.num}, 0)", f"COALESCE({r.num}, 0)"
    val = f"({lnum} / NULLIF({rnum}, 0))" if op == "/" else f"({lnum} {op} {rnum})"
    num = f"CASE WHEN {err} IS NOT NULL THEN NULL ELSE {val} END"
    return _merged(l, r, "'n'", num, "CAST(NULL AS VARCHAR)", err)


def _compare(l: Arr, r: Arr, op: str, fname: str) -> Arr:
    if not all(s.numeric() for s in l.srcs + r.srcs):
        raise Unsupported(f"{fname} array comparisons of text or TRUE/FALSE not supported yet")
    err = _propagated(l, r)
    num = (f"CASE WHEN {err} IS NOT NULL THEN NULL "
           f"WHEN COALESCE({l.num}, 0) {op} COALESCE({r.num}, 0) THEN 1 ELSE 0 END")
    return _merged(l, r, "'b'", num, "CAST(NULL AS VARCHAR)", err)


ARRAY_OPS = {"+", "-", "*", "/", "=", "<>", "<", ">", "<=", ">="}


def _array(c, node, sheet: str, fname: str, compare: bool = False) -> Arr:
    """An AST node as an element-wise array: *, +, -, /, comparisons and unary minus over aligned ranges, array
    constants and scalars. Anything else is Unsupported - the cell is UNRESOLVED, never silently intersected."""
    if isinstance(node, Binary):
        if node.op == "&":
            raise Unsupported(f"{fname} cannot concatenate arrays yet")
        if node.op not in ARRAY_OPS:
            raise Unsupported(f"{fname} cannot apply '{node.op}' to arrays yet")
        if node.op in ("=", "<>", "<", ">", "<=", ">="):
            return _compare(_array(c, node.left, sheet, fname, True),
                            _array(c, node.right, sheet, fname, True), node.op, fname)
        return _arith(_array(c, node.left, sheet, fname), _array(c, node.right, sheet, fname), node.op)
    if isinstance(node, Unary) and node.op == "-":
        zero = Arr([], "'n'", "CAST(0 AS DOUBLE)", "CAST(NULL AS VARCHAR)", "NULL", None)
        return _arith(zero, _array(c, node.operand, sheet, fname, compare), "-")
    if isinstance(node, RefNode) and not resolve(node.text, sheet, c.book).is_single:
        return _base(c, node, sheet, fname)
    e = c.expr(node, sheet)  # a scalar: a number, one cell or an expression of them, one element for every position
    if compare and e.type not in ("num", "blank"):
        raise Unsupported(f"{fname} array comparisons of text or TRUE/FALSE not supported yet")
    if e.type == "blank":
        return Arr([], "'n'", "CAST(NULL AS DOUBLE)", "CAST(NULL AS VARCHAR)", e.err, None)
    n = num(e)
    return Arr([], "'n'", n.sql, "CAST(NULL AS VARCHAR)", n.err, None)


def _rng(c, node, sheet: str, skip_subtotals: bool = False) -> Rng:
    area = resolve(node.text, sheet, c.book)
    keys = c.ready(area)
    note_text_numbers(c, keys)
    skip = ()
    if skip_subtotals:
        skip = tuple(k for k in keys if "SUBTOTAL(" in (c.book.cells[k].formula or "").upper())
    return Rng(area, keys, skip)


def _split(c, args, sheet: str, fname: str, skip_subtotals: bool = False) -> tuple[list, list, list]:
    """The arguments as (ranges, arrays, values). A range arithmetic is an array; a calculation over plain
    values stays a value, evaluated before the cells are read (MAX(H3, 1/0) is #DIV/0!)."""
    ranges, arrays, values = [], [], []
    for a in args:
        if isinstance(a, RefNode):
            ranges.append(_rng(c, a, sheet, skip_subtotals))
        elif isinstance(a, Binary) or (isinstance(a, Unary) and a.op == "-"):
            arr = _array(c, a, sheet, fname)
            if arr.srcs:
                arrays.append(arr)
            else:
                values.append(Expr(arr.num, "num", arr.err))
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
