"""PMT, PV, FV, NPER, NPV: Microsoft's documented closed forms.

With growth g = (1 + rate)^nper and type t (0: payments at the end of a period, 1: at the start):
    pv * g + pmt * (1 + rate * t) * (g - 1) / rate + fv = 0      (rate <> 0)
    pv + pmt * nper + fv = 0                                       (rate = 0)
each function solves this for one unknown. Edge cases whose Excel answer has not been checked against real Excel
(zero periods, a rate of -100% or less, a type other than 0 or 1, a growth that overflows) are UNCHECKED, which the
verifier reports as UNRESOLVED.
"""
from __future__ import annotations

from cellproof.expr import CELLS, UNCHECKED, ZERO, Expr, errs, num
from cellproof.fn.aggregate import _rng
from cellproof.functions import arity, register
from cellproof.parser import RefNode


def _args(c, args, sheet, name: str, lo: int, hi: int) -> list[Expr]:
    arity(name, args, lo, hi)
    given = [num(c.expr(a, sheet)) if a is not None else ZERO for a in args]
    return given + [ZERO] * (hi - len(args))


def _unchecked(cond: str) -> str:
    return f"(CASE WHEN {cond} THEN '{UNCHECKED}' END)"


def _growth(r: str, n: str) -> tuple[str, str]:
    """(1 + r)^n and its guard: 1 + r <= 0 (a fractional power of a negative number) or an overflow is unchecked."""
    g = f"(CASE WHEN 1 + {r} > 0 THEN POWER(1 + {r}, {n}) END)"
    return g, _unchecked(f"1 + {r} <= 0 OR NOT isfinite({g})")


def _type(t: Expr) -> str:
    return _unchecked(f"{t.sql} NOT IN (0, 1)")


def _solve(name: str, lo: int, hi: int, value, guards):
    """value(r, n, x, y, t, g) -> SQL; guards(...) -> extra error parts after the arguments' own errors."""
    def fn(c, args, sheet):
        a = _args(c, args, sheet, name, lo, hi)
        r, n = a[0].sql, a[1].sql
        g, g_err = _growth(r, n)
        parts = [x.sql for x in a] + [g]
        return Expr(value(*parts), "num", errs(*(x.err for x in a), _type(a[4]), g_err, *guards(*parts)))
    register(name)(fn)


# PMT(rate, nper, pv, [fv], [type])
_solve("PMT", 3, 5,
       lambda r, n, pv, fv, t, g: (f"(CASE WHEN {r} = 0 THEN -({pv} + {fv}) / NULLIF({n}, 0) "
                                   f"ELSE -({r} * ({fv} + {pv} * {g})) / NULLIF((1 + {r} * {t}) * ({g} - 1), 0) END)"),
       lambda r, n, pv, fv, t, g: [_unchecked(f"{n} = 0 OR ({r} <> 0 AND {g} = 1)")])

# FV(rate, nper, pmt, [pv], [type])
_solve("FV", 3, 5,
       lambda r, n, pmt, pv, t, g: (f"(CASE WHEN {r} = 0 THEN -({pv} + {pmt} * {n}) "
                                    f"ELSE -({pv} * {g} + {pmt} * (1 + {r} * {t}) * ({g} - 1) / {r}) END)"),
       lambda *_: [])

# PV(rate, nper, pmt, [fv], [type])
_solve("PV", 3, 5,
       lambda r, n, pmt, fv, t, g: (f"(CASE WHEN {r} = 0 THEN -({fv} + {pmt} * {n}) "
                                    f"ELSE -({fv} + {pmt} * (1 + {r} * {t}) * ({g} - 1) / {r}) / NULLIF({g}, 0) END)"),
       lambda r, n, pmt, fv, t, g: [_unchecked(f"{r} <> 0 AND {g} = 0")])


@register("NPER")
def nper(c, args, sheet):
    """NPER(rate, pmt, pv, [fv], [type]): ln((pmt*(1+r*t) - fv*r) / (pmt*(1+r*t) + pv*r)) / ln(1 + r); a ratio that
    is not positive (the payments never reach fv) is #NUM!."""
    r, pmt, pv, fv, t = _args(c, args, sheet, "NPER", 3, 5)
    p = f"({pmt.sql} * (1 + {r.sql} * {t.sql}))"
    top, bottom = f"({p} - {fv.sql} * {r.sql})", f"({p} + {pv.sql} * {r.sql})"
    ratio = f"({top} / NULLIF({bottom}, 0))"
    sql = (f"(CASE WHEN {r.sql} = 0 THEN -({pv.sql} + {fv.sql}) / NULLIF({pmt.sql}, 0) "
           f"WHEN 1 + {r.sql} > 0 AND {ratio} > 0 THEN ln({ratio}) / ln(1 + {r.sql}) END)")
    guard = (f"(CASE WHEN {r.sql} = 0 AND {pmt.sql} = 0 THEN '{UNCHECKED}' WHEN {r.sql} = 0 THEN NULL "
             f"WHEN 1 + {r.sql} <= 0 OR {bottom} = 0 THEN '{UNCHECKED}' WHEN {ratio} <= 0 THEN '#NUM!' END)")
    return Expr(sql, "num", errs(r.err, pmt.err, pv.err, fv.err, t.err, _type(t), guard))


@register("NPV")
def npv(c, args, sheet):
    """NPV(rate, value1, ...): sum of value_i / (1 + rate)^i. In a reference only numbers count (text, TRUE/FALSE
    and empty cells are skipped and take no period); values typed into the formula are converted."""
    arity("NPV", args, 2, 255)
    r = num(c.expr(args[0], sheet))
    rows, typed, ranges = [], [], []
    for i, a in enumerate(args[1:]):
        if isinstance(a, RefNode):
            rng = _rng(c, a, sheet)
            ranges.append(rng)
            rows.append(f"SELECT {i} AS a, row_num AS rr, col_num AS cc, num FROM {CELLS} "
                        f"WHERE {rng.where()} AND kind = 'n'")
        else:
            v = num(c.expr(a, sheet) if a is not None else ZERO)
            typed.append(v)
            rows.append(f"SELECT {i} AS a, 0 AS rr, 0 AS cc, {v.sql} AS num")
    periods = f"SELECT num, row_number() OVER (ORDER BY a, rr, cc) AS p FROM ({' UNION ALL '.join(rows)}) npv_rows"
    sql = (f"(SELECT COALESCE(SUM(num / POWER(1 + {r.sql}, p)), CAST(0 AS DOUBLE)) FROM ({periods}) npv_periods)")
    return Expr(sql, "num", errs(r.err, *(v.err for v in typed), *(g.err(c) for g in ranges),
                                 _unchecked(f"{r.sql} = -1")))
