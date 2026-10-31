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


NEWTON_STEPS = 100


def _newton(ctes: str, f: str, fp: str, guess: str) -> tuple[str, str]:
    """(root SQL, converged SQL) of Newton's method from guess; f and fp are SQL of the column r (and the CTEs).
    Excel stops at 0.00001 percent; this runs to the last digits, which differ from Excel's last iterate by far less
    than the verifier's 1e-9. Not converging, or leaving r > -1, is #NUM!, as in Excel."""
    it = (f"WITH RECURSIVE {ctes}, it(n, r, d) AS (SELECT 0, CAST({guess} AS DOUBLE), CAST(NULL AS DOUBLE) "
          f"UNION ALL SELECT n + 1, r - fv / fpv, fv / fpv FROM (SELECT it.n, it.r, {f} AS fv, {fp} AS fpv FROM it, "
          f"args WHERE it.n < {NEWTON_STEPS} AND it.r > -1 AND (it.d IS NULL OR abs(it.d) > 1e-13 * greatest(1, "
          f"abs(it.r)))) s) SELECT {{}} FROM it ORDER BY n DESC LIMIT 1")
    ok = "(d IS NOT NULL AND isfinite(r) AND r > -1 AND abs(d) <= 1e-10 * greatest(1, abs(r)))"
    return f"({it.format('CASE WHEN ' + ok + ' THEN r END')})", f"({it.format(ok)})"


@register("IRR")
def irr(c, args, sheet):
    """IRR(values, [guess]): the rate where the NPV of the values (the first one now) is 0. Text, TRUE/FALSE and empty
    cells are skipped. Several sign changes can have several roots, and which one Excel's iteration lands on was
    not checked: UNCHECKED."""
    arity("IRR", args, 1, 2)
    if not isinstance(args[0], RefNode):
        raise Unsupported("IRR of values typed into the formula not supported yet")
    rng = _rng(c, args[0], sheet)
    guess = num(c.expr(args[1], sheet)) if len(args) == 2 and args[1] is not None else Expr("CAST(0.1 AS DOUBLE)", "num")
    where = f"FROM {CELLS} WHERE {rng.where()} AND kind = 'n'"
    ctes = f"args AS (SELECT list(num ORDER BY row_num, col_num) AS v {where})"
    f = "list_sum(list_transform(v, lambda x, i: x / pow(1 + it.r, i - 1)))"
    fp = "list_sum(list_transform(v, lambda x, i: -(i - 1) * x / pow(1 + it.r, i)))"
    root, ok = _newton(ctes, f, fp, guess.sql)
    signs = f"(SELECT COUNT(*) FILTER (WHERE num > 0) > 0 AND COUNT(*) FILTER (WHERE num < 0) > 0 {where})"
    flips = (f"(SELECT COUNT(*) FROM (SELECT sign(num) AS s, lag(sign(num)) OVER (ORDER BY row_num, col_num) AS p "
             f"{where} AND num <> 0) irr_signs WHERE s <> p)")
    # real Excel (formulas' test.xlsx): an error in an array constant is #VALUE!; in cells, not checked
    error_in = rng.err(c)
    if error_in != "NULL":
        error_in = (f"(CASE WHEN {error_in} IS NOT NULL THEN "
                    f"'{'#VALUE!' if rng.area.sheet in c.book.arrays else UNCHECKED}' END)")
    return Expr(root, "num", errs(error_in, guess.err, f"(CASE WHEN NOT {signs} THEN '#NUM!' END)",
                                  _unchecked(f"{flips} > 1"), f"(CASE WHEN NOT {ok} THEN '#NUM!' END)"))


@register("RATE")
def rate(c, args, sheet):
    """RATE(nper, pmt, pv, [fv], [type], [guess]): the rate solving PMT's equation, by Newton's method."""
    arity("RATE", args, 3, 6)
    a = _args(c, args[:5], sheet, "RATE", 3, 5)
    n, pmt, pv, fv, t = (x.sql for x in a)
    guess = num(c.expr(args[5], sheet)) if len(args) == 6 and args[5] is not None else Expr("CAST(0.1 AS DOUBLE)", "num")
    ctes = f"args AS (SELECT {n} AS n_, {pmt} AS pmt, {pv} AS pv, {fv} AS fv, {t} AS t)"
    g, gp = "pow(1 + it.r, n_)", "(n_ * pow(1 + it.r, n_ - 1))"
    f = f"(pv * {g} + pmt * (1 + it.r * t) * ({g} - 1) / it.r + fv)"
    fp = (f"(pv * {gp} + pmt * (t * ({g} - 1) / it.r + (1 + it.r * t) * ({gp} * it.r - ({g} - 1)) / "
          f"(it.r * it.r)))")
    root, ok = _newton(ctes, f, fp, guess.sql)
    return Expr(root, "num", errs(*(x.err for x in a), guess.err, _type(a[4]),
                                  # a rate of 0 solves it (the payments only repay): Excel's answer there unchecked
                                  _unchecked(f"{n} <= 0 OR {n} <> floor({n}) OR abs({pv} + {pmt} * {n} + {fv}) "
                                             f"<= 1e-9 * greatest(abs({pv}), abs({pmt} * {n}), abs({fv}), 1) "
                                             f"OR abs({root}) < 1e-9"),
                                  f"(CASE WHEN NOT {ok} THEN '#NUM!' END)"))
