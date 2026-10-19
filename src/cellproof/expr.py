"""Typed SQL expressions, Excel-style type coercion and Excel error codes."""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Literal

from cellproof.model import Key, Unsupported

CELLS = "__cellproof_cells__"  # every cell known before the current batch; replaced when SQL is rendered
NO_ERR = "NULL"  # the err of an expression that cannot fail
Type = Literal["num", "txt", "bool", "blank", "err", "any"]


@dataclass(frozen=True)
class Expr:
    sql: str  # the value; for "any" its number part (numbers, TRUE/FALSE as 1/0); meaningless when err is set
    type: Type
    err: str = NO_ERR  # SQL giving the Excel error code ('#DIV/0!') or NULL
    txt: str = "NULL"  # "any" only: the text part
    kind: str = "NULL"  # "any" only: SQL giving 'n', 's' or 'b' - the value's type, known only when the SQL runs


ZERO = Expr("CAST(0 AS DOUBLE)", "num")
BLANK = Expr("NULL", "blank")  # an empty cell; always coerced before use, never rendered raw


def lit_num(v: float) -> Expr:
    if not math.isfinite(v):  # 1E+400 in a crafted file: Excel's limit is about 1.8E+308
        raise Unsupported(f"number {v!r} is beyond Excel's range")
    return Expr(f"CAST({v!r} AS DOUBLE)", "num")  # bare literals would be DECIMAL in DuckDB


def lit_str(v: str) -> Expr:
    return Expr("'" + v.replace("'", "''") + "'", "txt")


def error(code: str) -> Expr:
    """A value that is always this Excel error."""
    return Expr("NULL", "err", lit_str(code).sql)


def errs(*parts: str) -> str:
    """Excel reports the first error it meets, left to right."""
    live = [p for p in parts if p != NO_ERR]
    if not live:
        return NO_ERR
    return live[0] if len(live) == 1 else f"COALESCE({', '.join(live)})"


def any_of(kind: str, num_sql: str, txt_sql: str, err: str = NO_ERR) -> Expr:
    return Expr(num_sql, "any", err, txt_sql, kind)


def parse_num(t: str) -> str:
    """Text -> number the way Excel's arithmetic reads it (' 3 ', '-1.5e3'), else NULL. 'inf', '1,000' -> NULL."""
    return (f"(CASE WHEN regexp_matches(trim({t}), '^[+-]?([0-9]+[.]?[0-9]*|[.][0-9]+)([eE][+-]?[0-9]+)?$') "
            f"THEN CAST(trim({t}) AS DOUBLE) END)")


# Real Excel writes 1E-11 ... 1E+19 as plain digits (20 digits for 1E+19) and 1E+20 and up as 1E+20.
# ponytail: PLAIN_FROM is only known to be <= -11; nothing smaller has been tried in real Excel yet.
PLAIN_FROM, PLAIN_TO = -11, 19


def num_text(x: str) -> str:
    """Number -> text the way Excel writes it in a formula: up to 15 significant digits (0.1+0.2 -> '0.3')."""
    digits = "rtrim(replace(split_part(s, 'e', 1), '.', ''), '0')"  # the 15 significant digits, no trailing zeros
    plain = ("CASE WHEN e >= 0 THEN (CASE WHEN length(d) <= e + 1 THEN d || repeat('0', e + 1 - length(d)) "
             "ELSE substr(d, 1, e + 1) || '.' || substr(d, e + 2) END) "
             "ELSE '0.' || repeat('0', -e - 1) || d END")
    return (f"(SELECT CASE WHEN v = 0 THEN '0' "
            f"WHEN e IS NULL OR e < {PLAIN_FROM} OR e > {PLAIN_TO} THEN upper(printf('%.15g', v)) "
            f"ELSE (CASE WHEN v < 0 THEN '-' ELSE '' END) || ({plain}) END "
            f"FROM (SELECT v, {digits} AS d, TRY_CAST(split_part(s, 'e', 2) AS INTEGER) AS e "
            f"FROM (SELECT v, printf('%.14e', abs(v)) AS s FROM (SELECT {x} AS v))))")


def _value_error(bad: str) -> str:
    return f"(CASE WHEN {bad} THEN '#VALUE!' END)"


def num(e: Expr) -> Expr:
    if e.type == "num":
        return e
    if e.type == "blank":
        return ZERO
    if e.type == "bool":
        return Expr(f"CAST({e.sql} AS DOUBLE)", "num", e.err)
    if e.type == "err":
        return Expr("CAST(NULL AS DOUBLE)", "num", e.err)
    if e.type == "txt":
        n = parse_num(e.sql)
        return Expr(n, "num", errs(e.err, _value_error(f"{n} IS NULL")))
    n = parse_num(e.txt)
    return Expr(f"(CASE WHEN {e.kind} = 's' THEN {n} ELSE {e.sql} END)", "num",
                errs(e.err, _value_error(f"{e.kind} = 's' AND {n} IS NULL")))


def _bool_text(b: str) -> str:
    return f"(CASE WHEN {b} THEN 'TRUE' ELSE 'FALSE' END)"


def txt(e: Expr) -> Expr:
    if e.type == "txt":
        return e
    if e.type == "blank":
        return Expr("''", "txt")
    if e.type == "err":
        return Expr("CAST(NULL AS VARCHAR)", "txt", e.err)
    if e.type == "num":
        return Expr(num_text(e.sql), "txt", e.err)
    if e.type == "bool":
        return Expr(_bool_text(e.sql), "txt", e.err)
    return Expr(f"(CASE {e.kind} WHEN 's' THEN {e.txt} WHEN 'b' THEN {_bool_text(f'{e.sql} <> 0')} "
                f"ELSE {num_text(e.sql)} END)", "txt", e.err)


def _text_bool(t: str) -> str:
    return f"(CASE upper({t}) WHEN 'TRUE' THEN TRUE WHEN 'FALSE' THEN FALSE END)"


def as_bool(e: Expr) -> Expr:
    if e.type == "bool":
        return e
    if e.type == "blank":
        return Expr("FALSE", "bool")
    if e.type == "num":
        return Expr(f"({e.sql} <> 0)", "bool", e.err)
    if e.type == "err":
        return Expr("CAST(NULL AS BOOLEAN)", "bool", e.err)
    if e.type == "txt":
        b = _text_bool(e.sql)
        return Expr(b, "bool", errs(e.err, _value_error(f"{b} IS NULL")))
    b = _text_bool(e.txt)
    return Expr(f"(CASE WHEN {e.kind} = 's' THEN {b} ELSE {e.sql} <> 0 END)", "bool",
                errs(e.err, _value_error(f"{e.kind} = 's' AND {b} IS NULL")))


def as_any(e: Expr) -> Expr:
    """Any value as the run-time-typed form."""
    if e.type == "any":
        return e
    if e.type in ("blank", "num"):
        e = num(e)
        return any_of("'n'", e.sql, "CAST(NULL AS VARCHAR)", e.err)
    if e.type == "txt":
        return any_of("'s'", "CAST(NULL AS DOUBLE)", e.sql, e.err)
    if e.type == "bool":
        return any_of("'b'", f"CAST({e.sql} AS DOUBLE)", "CAST(NULL AS VARCHAR)", e.err)
    return any_of("'n'", "CAST(NULL AS DOUBLE)", "CAST(NULL AS VARCHAR)", e.err)  # err


COERCE = {"num": num, "txt": txt, "bool": as_bool, "any": as_any}
FLEX = ("blank", "err")  # types that take on the other side's type


def unify(a: Expr, b: Expr) -> tuple[Expr, Expr]:
    """Give two results one type (IF branches, IFERROR). Different types become 'any', as in Excel."""
    return tuple(unify_all([a, b]))


def unify_all(es: list[Expr]) -> list[Expr]:
    fixed = {e.type for e in es} - set(FLEX)
    if not fixed:
        return [num(e) for e in es]
    target = fixed.pop() if len(fixed) == 1 else "any"
    return [COERCE[target](e) for e in es]


def pick(cond: str, yes: Expr, no: Expr, err: str = NO_ERR) -> Expr:
    """CASE WHEN cond THEN yes ELSE no, for two results of one type (see unify), keeping every part."""
    def case(a: str, b: str) -> str:
        return a if a == b else f"(CASE WHEN {cond} THEN {a} ELSE {b} END)"
    return Expr(case(yes.sql, no.sql), yes.type, err, case(yes.txt, no.txt), case(yes.kind, no.kind))


RANK = {"num": 0, "txt": 1, "bool": 2}  # Excel orders numbers < text < TRUE/FALSE across types


def _rank(kind: str) -> str:
    return f"(CASE {kind} WHEN 'n' THEN 0 WHEN 's' THEN 1 ELSE 2 END)"


def _blank_like(other: Expr) -> Expr:
    """An empty cell compared with something acts as 0, "" or FALSE - whatever the other side is."""
    if other.type == "txt":
        return Expr("''", "txt")
    if other.type == "bool":
        return Expr("FALSE", "bool")
    if other.type == "any":
        return any_of(other.kind, "CAST(0 AS DOUBLE)", "''")
    return ZERO


def compare(a: Expr, b: Expr, op: str) -> Expr:
    """Excel comparison: text ignores case; different types compare by type (number < text < logical)."""
    if a.type == "blank" and b.type == "blank":
        a, b = ZERO, ZERO
    elif a.type == "blank":
        a = _blank_like(b)
    elif b.type == "blank":
        b = _blank_like(a)
    if a.type == "err":
        a = num(a)
    if b.type == "err":
        b = num(b)
    err = errs(a.err, b.err)
    if a.type == b.type and a.type != "any":
        if a.type == "txt":
            return Expr(f"(LOWER({a.sql}) {op} LOWER({b.sql}))", "bool", err)
        return Expr(f"({a.sql} {op} {b.sql})", "bool", err)
    if a.type != "any" and b.type != "any":
        return Expr(f"({RANK[a.type]} {op} {RANK[b.type]})", "bool", err)
    a, b = as_any(a), as_any(b)
    ra, rb = _rank(a.kind), _rank(b.kind)
    return Expr(f"(CASE WHEN {ra} <> {rb} THEN {ra} {op} {rb} WHEN {ra} = 1 THEN LOWER({a.txt}) {op} LOWER({b.txt}) "
                f"ELSE {a.sql} {op} {b.sql} END)", "bool", err)


def finalize(e: Expr) -> Expr:
    return ZERO if e.type == "blank" else e  # =A1 on an empty cell shows 0 in Excel


def where_cell(key: Key) -> str:
    sheet, row, col = key
    return f"sheet = {lit_str(sheet).sql} AND row_num = {row} AND col_num = {col}"


def where_area(area, alias: str = "") -> str:
    p = f"{alias}." if alias else ""
    return (f"{p}sheet = {lit_str(area.sheet).sql} AND {p}row_num BETWEEN {area.r1} AND {area.r2} "
            f"AND {p}col_num BETWEEN {area.c1} AND {area.c2}")
