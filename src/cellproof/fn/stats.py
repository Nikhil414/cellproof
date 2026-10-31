"""LARGE, SMALL, RANK(.EQ), STDEV(.S/.P/P), VAR(.S/.P/P), PERCENTILE(.INC), QUARTILE(.INC).

In a range only numbers count (text, TRUE/FALSE and empty cells are skipped) and an error in it is the result, as
in MAX. STDEV and VAR also take values typed into the formula (converted, TRUE is 1) and element-wise arrays, like
AVERAGE. A k that is not whole, and RANK over a range holding an error, are UNCHECKED (reported as UNRESOLVED).
"""
from __future__ import annotations

from cellproof.expr import CELLS, UNCHECKED, Expr, errs, num
from cellproof.fn.aggregate import _errs, _rng, _split
from cellproof.functions import arity, register
from cellproof.model import Unsupported
from cellproof.parser import RefNode


def _range_arg(c, node, sheet: str, name: str):
    if not isinstance(node, RefNode):
        raise Unsupported(f"{name} of values typed into the formula not supported yet")
    return _rng(c, node, sheet)


def _sorted(rng, order: str = "ASC") -> str:
    """The range's numbers as a sorted list (NULL when there are none)."""
    return f"(SELECT list(num ORDER BY num {order}) FROM {CELLS} WHERE {rng.where()} AND kind = 'n')"


def _kth(name: str, order: str):
    def fn(c, args, sheet):
        arity(name, args, 2, 2)
        rng = _range_arg(c, args[0], sheet, name)
        k = num(c.expr(args[1], sheet))
        xs = _sorted(rng, order)
        ok = f"({k.sql} = floor({k.sql}) AND {k.sql} >= 1 AND {k.sql} <= COALESCE(len({xs}), 0))"
        bad = (f"(CASE WHEN {k.sql} <> floor({k.sql}) THEN '{UNCHECKED}' "
               f"WHEN NOT {ok} THEN '#NUM!' END)")
        return Expr(f"(CASE WHEN {ok} THEN {xs}[CAST({k.sql} AS BIGINT)] END)", "num",
                    errs(k.err, _errs(c, [rng], [], []), bad))  # real Excel: k's error comes first
    register(name)(fn)


_kth("LARGE", "DESC")
_kth("SMALL", "ASC")


@register("RANK")
@register("RANK.EQ")
def rank(c, args, sheet):
    """1 + how many numbers in the range beat this one (descending unless order is nonzero); #N/A when the number
    is not in the range."""
    arity("RANK", args, 2, 3)
    x = num(c.expr(args[0], sheet))
    rng = _range_arg(c, args[1], sheet, "RANK")
    order = num(c.expr(args[2], sheet)) if len(args) == 3 and args[2] is not None else Expr("CAST(0 AS DOUBLE)", "num")
    beats = f"(CASE WHEN {order.sql} <> 0 THEN num < {x.sql} ELSE num > {x.sql} END)"
    where = f"FROM {CELLS} WHERE {rng.where()} AND kind = 'n'"
    found = f"EXISTS (SELECT 1 {where} AND num = {x.sql})"
    range_err = rng.err(c)
    unchecked = f"(CASE WHEN {range_err} IS NOT NULL THEN '{UNCHECKED}' END)" if range_err != "NULL" else "NULL"
    return Expr(f"(SELECT CAST(1 + COUNT(*) FILTER (WHERE {beats}) AS DOUBLE) {where})", "num",
                errs(x.err, order.err, unchecked, f"(CASE WHEN NOT {found} THEN '#N/A' END)"))


def _spread(name: str, sample: bool, root: bool):
    def fn(c, args, sheet):
        if not args:
            raise Unsupported(f"{name} needs arguments")
        ranges, arrays, values = _split(c, args, sheet, name)
        rows = " UNION ALL ".join([*(f"SELECT num FROM {CELLS} WHERE {r.where()} AND kind = 'n'" for r in ranges),
                                   *(f"SELECT num FROM {a.rel()} WHERE kind IN ('n')" for a in arrays),
                                   *(f"SELECT {v.sql} AS num" for v in values)])
        least = 2 if sample else 1
        # two passes, as Excel: the mean first, then the squared deviations from it
        var = (f"(SELECT SUM((num - m) * (num - m)) / NULLIF(COUNT(*) - {int(sample)}, 0) "
               f"FROM ({rows}) spread_rows, (SELECT avg(num) AS m FROM ({rows}) spread_mean) spread_m)")
        count = f"(SELECT COUNT(*) FROM ({rows}) spread_count)"
        sql = f"(CASE WHEN {count} >= {least} THEN {f'sqrt({var})' if root else var} END)"
        return Expr(sql, "num", errs(_errs(c, ranges, arrays, values),
                                     f"(CASE WHEN {count} < {least} THEN '#DIV/0!' END)"))
    register(name)(fn)


for _name, _sample, _root in [("STDEV", True, True), ("STDEV.S", True, True), ("STDEV.P", False, True),
                              ("STDEVP", False, True), ("VAR", True, False), ("VAR.S", True, False),
                              ("VAR.P", False, False), ("VARP", False, False)]:
    _spread(_name, _sample, _root)


def _percentile(c, rng, k: Expr, k_bad: str) -> Expr:
    """Excel's inclusive percentile: h = (n - 1) * k, x[floor h] + (h - floor h) * (x[floor h + 1] - x[floor h])."""
    xs = _sorted(rng)
    n = f"COALESCE(len({xs}), 0)"
    ok = f"(NOT ({k_bad}) AND {n} > 0)"
    h = f"(({n} - 1) * {k.sql})"
    i = f"CAST(floor({h}) AS BIGINT) + 1"  # 1-based
    lo, hi = f"{xs}[{i}]", f"{xs}[LEAST({i} + 1, {n})]"
    sql = f"(CASE WHEN {ok} THEN {lo} + ({h} - floor({h})) * ({hi} - {lo}) END)"
    return Expr(sql, "num", errs(k.err, _errs(c, [rng], [], []), f"(CASE WHEN NOT {ok} THEN '#NUM!' END)"))


@register("PERCENTILE")
@register("PERCENTILE.INC")
def percentile(c, args, sheet):
    arity("PERCENTILE", args, 2, 2)
    rng = _range_arg(c, args[0], sheet, "PERCENTILE")
    k = num(c.expr(args[1], sheet))
    return _percentile(c, rng, k, f"{k.sql} < 0 OR {k.sql} > 1")


@register("QUARTILE")
@register("QUARTILE.INC")
def quartile(c, args, sheet):
    arity("QUARTILE", args, 2, 2)
    rng = _range_arg(c, args[0], sheet, "QUARTILE")
    q = num(c.expr(args[1], sheet))  # truncated, but checked before: real Excel says QUARTILE(x, -0.1) is #NUM!
    k = Expr(f"(trunc({q.sql}) / 4)", "num", q.err)
    return _percentile(c, rng, k, f"{q.sql} < 0 OR {q.sql} > 4")
