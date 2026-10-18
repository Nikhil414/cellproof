"""ABS, INT, SIGN, PI, SQRT, EXP, LN, LOG10, LOG, POWER, MOD, TRUNC, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, MROUND.

DuckDB throws on ln(0) or sqrt(-1), and one throw would stop the whole batch, so every function with a
domain is written as CASE WHEN <ok> THEN f(x) END plus the Excel error for the rest (a CASE never runs the
branch it skips, verified on DuckDB 1.5.6). Rounding to a digit first rounds to 15 significant digits, as
Excel does, so ROUNDUP(0.1*3, 1) is 0.3, not 0.4.
"""
from __future__ import annotations

from unexcel.expr import ZERO, Expr, errs, num
from unexcel.functions import arity, register
from unexcel.parser import Binary


def _guarded(value: str, bad: str, code: str) -> tuple[str, str]:
    return f"(CASE WHEN NOT ({bad}) THEN {value} END)", f"(CASE WHEN {bad} THEN '{code}' END)"


def _one(name: str, value, bad=None, code: str = "#NUM!"):
    def fn(c, args, sheet):
        arity(name, args, 1, 1)
        x = num(c.expr(args[0], sheet))
        if bad is None:
            return Expr(value(x.sql), "num", x.err)
        sql, err = _guarded(value(x.sql), bad(x.sql), code)
        return Expr(sql, "num", errs(x.err, err))
    register(name)(fn)


_one("ABS", lambda x: f"abs({x})")
_one("INT", lambda x: f"floor({x})")
_one("SIGN", lambda x: f"CAST(sign({x}) AS DOUBLE)")
_one("SQRT", lambda x: f"sqrt({x})", lambda x: f"{x} < 0")
_one("EXP", lambda x: f"exp({x})", lambda x: f"{x} > 709.782712893384")
_one("LN", lambda x: f"ln({x})", lambda x: f"{x} <= 0")
_one("LOG10", lambda x: f"log10({x})", lambda x: f"{x} <= 0")


@register("PI")
def pi(c, args, sheet):
    arity("PI", args, 0, 0)
    return Expr("pi()", "num")


@register("LOG")
def log(c, args, sheet):
    arity("LOG", args, 1, 2)
    x = num(c.expr(args[0], sheet))
    base = num(c.expr(args[1], sheet)) if len(args) == 2 and args[1] is not None else Expr("CAST(10 AS DOUBLE)", "num")
    bad = f"(CASE WHEN {x.sql} <= 0 OR {base.sql} <= 0 THEN '#NUM!' WHEN {base.sql} = 1 THEN '#DIV/0!' END)"
    return Expr(f"(CASE WHEN {x.sql} > 0 AND {base.sql} > 0 AND {base.sql} <> 1 THEN ln({x.sql}) / ln({base.sql}) END)",
                "num", errs(x.err, base.err, bad))


@register("POWER")
def power(c, args, sheet):
    arity("POWER", args, 2, 2)
    return c.binary(Binary("^", args[0], args[1]), sheet)


def _two(c, args, sheet, name: str, default=None) -> tuple[Expr, Expr]:
    arity(name, args, 1 if default is not None else 2, 2)
    x = num(c.expr(args[0], sheet))
    y = num(c.expr(args[1], sheet)) if len(args) == 2 and args[1] is not None else (default or ZERO)
    return x, y


def xl15(v: str) -> str:
    """v rounded to 15 significant digits, the precision Excel works at."""
    return f"(CASE WHEN {v} = 0 THEN 0 WHEN NOT isfinite({v}) THEN {v} ELSE round({v}, 14 - CAST(floor(log10(abs({v}))) AS INTEGER)) END)"


@register("MOD")
def mod(c, args, sheet):
    x, y = _two(c, args, sheet, "MOD")
    # Real Excel: MOD(1, 0.1) = 0.09999999999999995, the exact floating-point remainder, not x - y*floor(x/y).
    # DuckDB's % is the C remainder (sign of the dividend); its fmod() function is not, so it is not used.
    rem = f"({x.sql} % {y.sql})"
    sql, err = _guarded(f"(CASE WHEN {rem} <> 0 AND ({rem} < 0) <> ({y.sql} < 0) THEN {rem} + {y.sql} ELSE {rem} END)",
                        f"{y.sql} = 0", "#DIV/0!")
    return Expr(sql, "num", errs(x.err, y.err, err))  # Excel: the result takes the divisor's sign


def _digits(name: str, rnd: str):
    def fn(c, args, sheet):
        x, d = _two(c, args, sheet, name, default=ZERO if name == "TRUNC" else None)
        p = f"power(10, trunc({d.sql}))"
        # shortcut: scaling outside DOUBLE gives #NUM!, revisit with Excel boundary fixtures.
        scaled = f"abs({x.sql}) * {p}"
        ok = f"(isfinite({p}) AND {p} > 0 AND isfinite({scaled}))"
        sql = f"(CASE WHEN {d.sql} > 308 THEN {x.sql} WHEN {ok} THEN sign({x.sql}) * {rnd}({xl15(scaled)}) / {p} END)"
        bad = f"(CASE WHEN {d.sql} <= 308 AND NOT {ok} THEN '#NUM!' END)"
        return Expr(sql, "num", errs(x.err, d.err, bad))
    register(name)(fn)


_digits("ROUNDUP", "ceil")
_digits("ROUNDDOWN", "floor")
_digits("TRUNC", "floor")


def _multiple(name: str, rnd: str, zero_code: str | None):
    def fn(c, args, sheet):
        x, m = _two(c, args, sheet, name)
        bad = f"({x.sql} > 0 AND {m.sql} < 0)" if name != "MROUND" else f"({x.sql} * {m.sql} < 0)"
        if zero_code:
            bad = f"({bad} OR {m.sql} = 0)"
        value = f"(CASE WHEN {m.sql} = 0 THEN 0 ELSE {rnd}({xl15(f'{x.sql} / {m.sql}')}) * {m.sql} END)"
        sql, _ = _guarded(value, bad, "#NUM!")
        code = (f"(CASE WHEN {m.sql} = 0 THEN '{zero_code}' WHEN {bad} THEN '#NUM!' END)" if zero_code
                else f"(CASE WHEN {bad} THEN '#NUM!' END)")
        return Expr(sql, "num", errs(x.err, m.err, code))
    register(name)(fn)


_multiple("CEILING", "ceil", None)
_multiple("FLOOR", "floor", "#DIV/0!")
_multiple("MROUND", "round", None)
