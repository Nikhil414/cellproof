"""AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, PRODUCT, MEDIAN, SUMPRODUCT, SUBTOTAL.

Like SUM: in a range only numbers count (text, TRUE/FALSE and empty cells are skipped); a value typed
straight into the arguments is converted (AVERAGE(1, "3") = 2). An error in a range is the result.
"""
from __future__ import annotations

from dataclasses import dataclass, replace

from cellproof.expr import BLANK, CELLS, NO_ERR, Expr, errs, lit_str, num, where_area
from cellproof.functions import arity, note_text_numbers, register
from cellproof.model import Unsupported
from cellproof.parser import Num, RefNode
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


def _rng(c, node, sheet: str, skip_subtotals: bool = False) -> Rng:
    area = resolve(node.text, sheet, c.book)
    keys = c.ready(area)
    note_text_numbers(c, keys)
    skip = ()
    if skip_subtotals:
        skip = tuple(k for k in keys if "SUBTOTAL(" in (c.book.cells[k].formula or "").upper())
    return Rng(area, keys, skip)


def _split(c, args, sheet: str, skip_subtotals: bool = False) -> tuple[list[Rng], list[Expr]]:
    ranges, values = [], []
    for a in args:
        if isinstance(a, RefNode):
            ranges.append(_rng(c, a, sheet, skip_subtotals))
        else:
            values.append(num(c.expr(a, sheet) if a is not None else BLANK))
    return ranges, values


def _errs(c, ranges, values) -> str:
    # Real Excel: MAX(H3, 1/0) is #DIV/0! although H3 holds #N/A. Typed-in expressions are evaluated before the
    # function reads any cell, so their errors come first; errors inside referenced cells follow, in order.
    return errs(*(v.err for v in values), *(r.err(c) for r in ranges))


def _average(c, ranges, values) -> Expr:
    total = " + ".join([*(r.agg("COALESCE(SUM(num), 0)") for r in ranges), *(v.sql for v in values)]) or "0"
    count = " + ".join([*(r.agg("COUNT(*)") for r in ranges), *("1" for _ in values)]) or "0"
    return Expr(f"(({total}) / NULLIF({count}, 0))", "num",
                errs(_errs(c, ranges, values), f"(CASE WHEN ({count}) = 0 THEN '#DIV/0!' END)"))


def _extreme(c, ranges, values, how: str) -> Expr:
    parts = [*(r.agg(f"{how}(num)") for r in ranges), *(v.sql for v in values)]
    fn = "LEAST" if how == "MIN" else "GREATEST"  # both skip NULLs in DuckDB; no numbers at all -> 0, like Excel
    return Expr(f"COALESCE({fn}({', '.join(parts)}), CAST(0 AS DOUBLE))", "num", _errs(c, ranges, values))


def _count(c, ranges, values) -> Expr:
    if values:
        raise Unsupported("COUNT of values typed into the formula not supported yet")
    return Expr("(" + " + ".join(r.agg("CAST(COUNT(*) AS DOUBLE)") for r in ranges) + ")", "num")


def _counta(c, ranges, values) -> Expr:
    parts = [r.agg("CAST(COUNT(*) AS DOUBLE)", "('n', 's', 'b', 'e')") for r in ranges] + ["1" for _ in values]
    return Expr("CAST((" + " + ".join(parts) + ") AS DOUBLE)", "num")


def _product(c, ranges, values) -> Expr:
    parts = [f"COALESCE({r.agg('product(num)')}, 1)" for r in ranges] + [v.sql for v in values]
    count = " + ".join([*(r.agg("COUNT(*)") for r in ranges), *("1" for _ in values)])
    return Expr(f"(CASE WHEN ({count}) = 0 THEN CAST(0 AS DOUBLE) ELSE {' * '.join(parts)} END)", "num",
                _errs(c, ranges, values))


def _sum(c, ranges, values) -> Expr:
    parts = [*(r.agg("COALESCE(SUM(num), 0)") for r in ranges), *(v.sql for v in values)]
    return Expr("(" + " + ".join(parts) + ")" if parts else "CAST(0 AS DOUBLE)", "num", _errs(c, ranges, values))


def _plain(name, how):
    def fn(c, args, sheet):
        if not args:
            raise Unsupported(f"{name} needs arguments")
        ranges, values = _split(c, args, sheet)
        result = how(c, ranges, values)
        if name not in ("COUNT", "COUNTA"):
            result = replace(result, err=errs(_errs(c, ranges, values), result.err))
        return result
    register(name)(fn)


_plain("AVERAGE", _average)
_plain("MIN", lambda c, r, v: _extreme(c, r, v, "MIN"))
_plain("MAX", lambda c, r, v: _extreme(c, r, v, "MAX"))
_plain("COUNT", _count)
_plain("COUNTA", _counta)
_plain("PRODUCT", _product)


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
                errs(_errs(c, ranges, []), f"(CASE WHEN (SELECT COUNT(*) {values}) = 0 THEN '#NUM!' END)"))


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
    return Expr(found, "num", errs(_errs(c, ranges, []), f"(CASE WHEN {found} IS NULL THEN '#N/A' END)"))


@register("SUMPRODUCT")
def sumproduct(c, args, sheet):
    if not args or not all(isinstance(a, RefNode) for a in args):
        raise Unsupported("SUMPRODUCT of array expressions not supported yet")
    ranges = [_rng(c, a, sheet) for a in args]
    first = ranges[0].area
    shape = (first.r2 - first.r1, first.c2 - first.c1)
    if any((r.area.r2 - r.area.r1, r.area.c2 - r.area.c1) != shape for r in ranges):
        return Expr("NULL", "num", errs(_errs(c, ranges, []), "'#VALUE!'"))  # Excel: an error inside first, then sizes
    joins = "".join(
        f" JOIN {CELLS} a{i} ON a{i}.sheet = {lit_str(r.area.sheet).sql} "
        f"AND a{i}.row_num = a0.row_num + {r.area.r1 - first.r1} AND a{i}.col_num = a0.col_num + {r.area.c1 - first.c1} "
        f"AND a{i}.kind = 'n'" for i, r in enumerate(ranges[1:], 1))
    product = " * ".join(f"a{i}.num" for i in range(len(ranges)))  # blanks, text and TRUE/FALSE count as 0
    return Expr(f"(SELECT COALESCE(SUM({product}), CAST(0 AS DOUBLE)) FROM {CELLS} a0{joins} "
                f"WHERE {where_area(first, 'a0')} AND a0.kind = 'n')", "num", _errs(c, ranges, []))


SUBTOTALS = {1: _average, 2: _count, 3: _counta, 4: lambda c, r, v: _extreme(c, r, v, "MAX"),
             5: lambda c, r, v: _extreme(c, r, v, "MIN"), 6: _product, 9: _sum}


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
    return SUBTOTALS[code](c, ranges, [])
