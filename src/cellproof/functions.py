"""Excel function -> SQL. Each entry: fn(compiler, args, sheet) -> Expr.

Ranges are read from the cells table (inputs plus the formula cells of earlier batches).
Compiler.ready() checks every formula cell in a range was computed in an earlier batch.
Every function returns its Excel error in Expr.err; see expr.errs for the order rule.
"""
from __future__ import annotations

from cellproof.expr import (CELLS, NO_ERR, ZERO, Expr, _text_bool, any_of, as_bool, errs, keep_locale, lit_str, num,
                            pick, unify, where_area)
from cellproof.model import Unsupported, a1
from cellproof.parser import Bool, Num, RefNode, Str
from cellproof.refs import Area, resolve

FUNCS = {}


def register(name):
    def deco(fn):
        FUNCS[name] = fn
        return fn

    return deco


def arity(name, args, lo, hi):
    if not lo <= len(args) <= hi:
        raise Unsupported(f"{name} needs {lo}-{hi} arguments, got {len(args)}")


def _looks_numeric(text: str) -> bool:
    try:
        float(text.replace(",", "").strip())
        return True
    except ValueError:
        return False


def note_text_numbers(c, keys):
    """Excel skips text in SUM-like ranges, even text that looks like a number. Tell the user."""
    for key in keys.text_cells():
        cell = c.book.cells[key]
        if _looks_numeric(cell.value):
            c.findings.append(f"{a1(key)}: number stored as text ({cell.value!r}) is skipped by Excel")


def first_error(c, area: Area, keys) -> str:
    """SQL for the first error in a range, row by row, or NO_ERR when no cell in it can hold one."""
    if not keys.may_err:
        return NO_ERR
    return f"(SELECT txt FROM {CELLS} WHERE {where_area(area)} AND kind = 'e' ORDER BY row_num, col_num LIMIT 1)"


def range_sum(c, area: Area) -> Expr:
    """SUM over a range: text, TRUE/FALSE and blanks are skipped; the first error wins."""
    keys = c.ready(area)
    note_text_numbers(c, keys)
    return Expr(f"(SELECT COALESCE(SUM(num), CAST(0 AS DOUBLE)) FROM {CELLS} WHERE {where_area(area)} AND kind = 'n')",
                "num", first_error(c, area, keys))


@register("SUM")
def sum_(c, args, sheet):
    parts = [range_sum(c, resolve(arg.text, sheet, c.book)) if isinstance(arg, RefNode) else num(c.expr(arg, sheet))
             for arg in args]
    if not parts:
        return ZERO
    # Real Excel: SUM(H3,1/0) with H3 = #N/A is #DIV/0!. Typed-in expressions are evaluated first.
    typed = [p for a, p in zip(args, parts) if not isinstance(a, RefNode)]
    cells = [p for a, p in zip(args, parts) if isinstance(a, RefNode)]
    return Expr("(" + " + ".join(p.sql for p in parts) + ")", "num", errs(*(p.err for p in typed + cells)))


@register("IF")
def if_(c, args, sheet):
    arity("IF", args, 2, 3)
    cond = as_bool(c.expr(args[0], sheet))
    yes = c.expr(args[1], sheet)
    no = c.expr(args[2], sheet) if len(args) == 3 else Expr("FALSE", "bool")
    yes, no = unify(yes, no)
    taken = NO_ERR if yes.err == no.err == NO_ERR else f"(CASE WHEN {cond.sql} THEN {yes.err} ELSE {no.err} END)"
    return pick(cond.sql, yes, no, errs(cond.err, taken))


@register("ROUND")
def round_(c, args, sheet):
    arity("ROUND", args, 2, 2)
    x, digits = num(c.expr(args[0], sheet)), num(c.expr(args[1], sheet))
    # DuckDB ROUND on DOUBLE rounds half away from zero like Excel (2.675 -> 2.68, verified on 1.5.6)
    return Expr(f"ROUND({x.sql}, CAST(TRUNC({digits.sql}) AS INTEGER))", "num", errs(x.err, digits.err))


def logical_args(c, name: str, args, sheet: str) -> tuple[list[str], str]:
    """AND/OR/XOR arguments: (TRUE/FALSE SQL per argument, NULL when it is skipped; their error).

    Real Excel: text and empty cells in a reference are skipped, and so is typed text other than TRUE/FALSE
    (AND(TRUE,"0") is TRUE); an error in an argument comes first (AND(TRUE,"0",#REF!) is #REF!); nothing left to judge
    is #VALUE!."""
    if not args:
        raise Unsupported(f"{name} needs arguments")
    values, arg_errs, conv_errs = [], [], []
    for a in args:
        if isinstance(a, RefNode) and not resolve(a.text, sheet, c.book).is_single:
            raise Unsupported(f"{name} over ranges not supported yet")
        x = c.expr(a, sheet)
        arg_errs.append(x.err)
        if isinstance(a, RefNode) and x.type in ("txt", "blank"):
            values.append("CAST(NULL AS BOOLEAN)")
        elif isinstance(a, RefNode) and x.type == "any":
            values.append(f"(CASE WHEN {x.kind} = 's' THEN NULL ELSE {x.sql} <> 0 END)")
        elif x.type == "txt":
            values.append(_text_bool(x.sql))
        elif x.type == "any":
            values.append(f"(CASE WHEN {x.kind} = 's' THEN {_text_bool(x.txt)} ELSE {x.sql} <> 0 END)")
        else:
            b = as_bool(Expr(x.sql, x.type, NO_ERR, x.txt, x.kind))  # without x.err: only the conversion's error
            values.append(b.sql)
            conv_errs.append(b.err)
    none = "(CASE WHEN " + " AND ".join(f"{v} IS NULL" for v in values) + " THEN '#VALUE!' END)"
    return values, errs(*arg_errs, *conv_errs, none)


def _logic(name, joiner, skip):
    def fn(c, args, sheet):
        values, err = logical_args(c, name, args, sheet)
        return Expr("(" + f" {joiner} ".join(f"COALESCE({v}, {skip})" for v in values) + ")", "bool", err)

    return fn


FUNCS["AND"] = _logic("AND", "AND", "TRUE")
FUNCS["OR"] = _logic("OR", "OR", "FALSE")


@register("IFERROR")
def iferror(c, args, sheet):
    arity("IFERROR", args, 2, 2)
    value, fallback = unify(c.expr(args[0], sheet), c.expr(args[1], sheet))
    if value.err == NO_ERR:
        return value
    ok = f"(({value.err}) IS NULL)"
    err = NO_ERR if fallback.err == NO_ERR else f"(CASE WHEN {ok} THEN NULL ELSE {fallback.err} END)"
    return pick(ok, value, fallback, errs(keep_locale(value.err), err))
