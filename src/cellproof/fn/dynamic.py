"""FILTER, SORT, UNIQUE: functions whose result is an array as big as the data says.

An array is a relation: one row per element with its 1-based position (i, j) and the cells table's value columns
(kind, num, txt, err; kind NULL for an empty cell). Ranges, array constants and these functions nest in any order
(SORT(UNIQUE(FILTER(...)))). Each cell of the block Excel spilled the result into picks its own element
(compiler.compile_cell); the anchor of a dynamic spill also checks the result is exactly as big as that block.

What Excel's answer was not checked for is UNCHECKED (UNRESOLVED), never guessed: SORT keys that are empty, errors
or text with characters other than letters, digits and spaces (Excel's text order is its own collation), and
UNIQUE over empty cells or errors.
"""
from __future__ import annotations

from dataclasses import dataclass

from cellproof.expr import CELLS, SETTINGS, UNCHECKED, Expr, as_any, lit_str
from cellproof.fn.aggregate import _array, _base, _truth
from cellproof.fn.lookup import int_literal
from cellproof.functions import register
from cellproof.model import Unsupported
from cellproof.parser import Bool, Call, RefNode
from cellproof.refs import MAX_ROW, resolve

DYNAMIC = {"FILTER", "SORT", "UNIQUE"}


@dataclass(frozen=True)
class Rel:
    sql: str  # a subquery with columns i, j, kind, num, txt, err
    rows: int | None  # its size when known before it runs
    cols: int | None


def _uid(c) -> str:
    return "d" + c.walias()[1:]


def _range(c, node, sheet: str) -> Rel:
    area = resolve(node.text, sheet, c.book)
    r2 = area.r2
    if r2 == MAX_ROW:  # ponytail: a whole column stops at the last row any sheet holds (its cells are all empty)
        r2 = max(area.r1, max(c.book.max_row.values(), default=1))
    c.ready(area)
    sql = (f"(SELECT CAST(gr.r - {area.r1} + 1 AS BIGINT) AS i, CAST(gc.c - {area.c1} + 1 AS BIGINT) AS j, "
           f"x.kind AS kind, x.num AS num, x.txt AS txt, CASE WHEN x.kind = 'e' THEN x.txt END AS err "
           f"FROM range({area.r1}, {r2 + 1}) gr(r) CROSS JOIN range({area.c1}, {area.c2 + 1}) gc(c) "
           f"LEFT JOIN {CELLS} x ON x.sheet = {lit_str(area.sheet).sql} AND x.row_num = gr.r AND x.col_num = gc.c)")
    return Rel(sql, r2 - area.r1 + 1, area.c2 - area.c1 + 1)


def relation(c, node, sheet: str) -> Rel:
    """An array argument as a relation: a range, an array constant, or FILTER/SORT/UNIQUE of one."""
    if isinstance(node, RefNode):
        return _range(c, node, sheet)
    if isinstance(node, Call) and node.name in BUILDERS:
        return BUILDERS[node.name](c, node.args, sheet)
    raise Unsupported("an array argument other than a range, an array constant, FILTER, SORT or UNIQUE "
                      "not supported yet")


def _whole(u: str, ctes: list[str], body: str, bad: str, empty: str | None = None) -> str:
    """WITH ctes, {u}_w AS body: that relation, or one element holding the error `bad` when it is set, or `empty`
    (a one-row SELECT) when it has no elements. Every input is a MATERIALIZED CTE: written and run once, however
    deep FILTER/SORT/UNIQUE nest."""
    ctes = ctes + [f"{u}_w AS MATERIALIZED {body}", f"{u}_b AS MATERIALIZED (SELECT ({bad}) AS e)"]
    parts = [f"SELECT * FROM {u}_w WHERE (SELECT e FROM {u}_b) IS NULL",
             f"SELECT 1, 1, 'e', NULL, NULL, e FROM {u}_b WHERE e IS NOT NULL"]
    if empty is not None:
        parts.append(f"{empty} WHERE (SELECT e FROM {u}_b) IS NULL AND NOT EXISTS (SELECT 1 FROM {u}_w)")
    return f"(WITH {', '.join(ctes)} " + " UNION ALL ".join(parts) + ")"


def _flag(node, name: str, default: bool) -> bool:
    if node is None:
        return default
    if isinstance(node, Bool):
        return node.value
    n = int_literal(node)
    if n is None:
        raise Unsupported(f"{name} with a by_col/exactly_once argument computed at run time not supported yet")
    return n != 0


def filter_(c, args, sheet: str) -> Rel:
    """FILTER(array, include, [if_empty]): the rows (include one column high as the array) or the columns (include
    one row as wide) whose include element is true; none -> if_empty (without it, UNCHECKED: see below). A text include element is #VALUE!,
    an error one is that error, for the whole result."""
    if not 2 <= len(args) <= 3:
        raise Unsupported("FILTER needs two or three arguments")
    a = relation(c, args[0], sheet)
    inc = _base(c, args[1], sheet, "FILTER") if isinstance(args[1], RefNode) else _array(c, args[1], sheet, "FILTER", "raw")
    if not inc.srcs:
        raise Unsupported("FILTER with an include that is one value not supported yet")
    if a.rows is None or a.cols is None:
        raise Unsupported("FILTER of an array whose size is only known at run time not supported yet")
    ih, iw = inc.shape[0] + 1, inc.shape[1] + 1
    first = inc.srcs[0].area
    if iw == 1 and ih == a.rows:
        pos, other = f"(row_num - {first.r1} + 1)", "i"
    elif ih == 1 and iw == a.cols:
        pos, other = f"(col_num - {first.c1} + 1)", "j"
    else:
        return Rel("(SELECT 1 AS i, 1 AS j, 'e' AS kind, NULL AS num, NULL AS txt, '#VALUE!' AS err)", 1, 1)
    keep, kerr = _truth("kind", "num", "txt", "err")
    u = _uid(c)
    ctes = [f"{u}_a AS MATERIALIZED {a.sql}",
            f"{u}_f AS MATERIALIZED (SELECT {pos} AS k, {keep} AS keep, {kerr} AS kerr FROM {inc.rel()} {u}_src)"]
    bad = f"(SELECT kerr FROM {u}_f WHERE kerr IS NOT NULL ORDER BY k LIMIT 1)"
    kept = (f"(SELECT {'dense_rank() OVER (ORDER BY a.i)' if other == 'i' else 'a.i'} AS i, "
            f"{'dense_rank() OVER (ORDER BY a.j)' if other == 'j' else 'a.j'} AS j, a.kind, a.num, a.txt, a.err "
            f"FROM {u}_a a JOIN {u}_f f ON f.k = a.{other} WHERE f.keep)")
    empty = None
    if len(args) == 3 and args[2] is not None:
        e = as_any(c.expr(args[2], sheet))
        empty = f"SELECT 1, 1, {e.kind}, {e.sql}, {e.txt}, {e.err}"
    else:  # real Excel disagrees with itself: #CALC! in Microsoft's docs, #VALUE! in formulas' test.xlsx
        empty = f"SELECT 1, 1, 'e', NULL, NULL, '{UNCHECKED}'"
    return Rel(_whole(u, ctes, kept, bad, empty), None, None)


SAFE_TEXT = "'^[A-Za-z0-9 ]*$'"


def sort_(c, args, sheet: str) -> Rel:
    """SORT(array, [sort_index], [sort_order], [by_col]): rows ordered by one column, numbers before text before
    TRUE/FALSE, text ignoring case; rows with equal keys keep their order."""
    if not 1 <= len(args) <= 4:
        raise Unsupported("SORT needs one to four arguments")
    a = relation(c, args[0], sheet)
    idx = 1 if len(args) < 2 or args[1] is None else int_literal(args[1])
    order = 1 if len(args) < 3 or args[2] is None else int_literal(args[2])
    if idx is None or order is None:
        raise Unsupported("SORT with a sort_index or sort_order that is not one whole number not supported yet")
    if _flag(args[3] if len(args) == 4 else None, "SORT", False):
        raise Unsupported("SORT by columns not supported yet")
    if order not in (1, -1) or idx < 1 or (a.cols is not None and idx > a.cols):
        return Rel("(SELECT 1 AS i, 1 AS j, 'e' AS kind, NULL AS num, NULL AS txt, '#VALUE!' AS err)", 1, 1)
    d = "ASC" if order == 1 else "DESC"
    u = _uid(c)
    ctes = [f"{u}_a AS MATERIALIZED {a.sql}", f"{u}_k AS MATERIALIZED (SELECT * FROM {u}_a WHERE j = {idx})"]
    odd = (f"(SELECT CASE WHEN EXISTS (SELECT 1 FROM {u}_k WHERE kind IS NULL OR kind = 'e' OR "
           f"(kind = 's' AND NOT regexp_matches(COALESCE(txt, ''), {SAFE_TEXT}))) THEN '{UNCHECKED}' END)")
    rank = "CASE kind WHEN 'n' THEN 0 WHEN 's' THEN 1 ELSE 2 END"
    new = (f"(SELECT i AS oi, row_number() OVER (ORDER BY {rank} {d}, num {d}, lower(COALESCE(txt, '')) {d}, i) "
           f"AS ni FROM {u}_k)")
    sql = f"(SELECT n.ni AS i, a.j, a.kind, a.num, a.txt, a.err FROM {u}_a a JOIN {new} n ON n.oi = a.i)"
    return Rel(_whole(u, ctes, sql, odd), a.rows, a.cols)


def unique_(c, args, sheet: str) -> Rel:
    """UNIQUE(array, [by_col], [exactly_once]): the distinct rows in the order first seen (text ignoring case);
    exactly_once keeps only the rows that occur once."""
    if not 1 <= len(args) <= 3:
        raise Unsupported("UNIQUE needs one to three arguments")
    a = relation(c, args[0], sheet)
    if _flag(args[1] if len(args) > 1 else None, "UNIQUE", False):
        raise Unsupported("UNIQUE by columns not supported yet")
    once = _flag(args[2] if len(args) > 2 else None, "UNIQUE", False)
    u = _uid(c)
    key = ("CASE kind WHEN 's' THEN 's' || lower(COALESCE(txt, '')) WHEN 'b' THEN 'b' || CAST(num AS VARCHAR) "
           "ELSE 'n' || CAST(num AS VARCHAR) END")
    ctes = [f"{u}_a AS MATERIALIZED {a.sql}"]
    odd = f"(SELECT CASE WHEN EXISTS (SELECT 1 FROM {u}_a WHERE kind IS NULL OR kind = 'e') THEN '{UNCHECKED}' END)"
    rows = f"(SELECT i, string_agg({key}, chr(31) ORDER BY j) AS rk FROM {u}_a GROUP BY i)"
    groups = f"(SELECT rk, min(i) AS first, COUNT(*) AS n FROM {rows} g GROUP BY rk)"
    keep = f"(SELECT first, row_number() OVER (ORDER BY first) AS ni FROM {groups} h" + (
        " WHERE n = 1)" if once else ")")
    sql = f"(SELECT k.ni AS i, a.j, a.kind, a.num, a.txt, a.err FROM {u}_a a JOIN {keep} k ON k.first = a.i)"
    empty = "SELECT 1, 1, 'e', NULL, NULL, '#CALC!'"
    return Rel(_whole(u, ctes, sql, odd, empty), None, a.cols)


BUILDERS = {"FILTER": filter_, "SORT": sort_, "UNIQUE": unique_}


def element(c, node, sheet: str, i: int, j: int, block: tuple[int, int], dynamic: bool) -> Expr:
    """This cell's element (i, j) of the array, 0-based. A result one row (column) high repeats down (across), as in
    a legacy block; past the result a cell is #N/A. The anchor of a dynamic spill is #SPILL! when the result is not
    exactly as big as the block Excel spilled it into (the translation disagrees with Excel's own result)."""
    rel = relation(c, node, sheet).sql
    u = _uid(c)
    # ponytail: the relation runs once per value column (4x); one struct per cell if big spills get slow
    head = (f"WITH {u}_r AS MATERIALIZED {rel}, {u}_d AS (SELECT max(i) AS h, max(j) AS w FROM {u}_r), "
            f"{u}_e AS (SELECT r.* FROM {u}_r r, {u}_d d WHERE r.i = (CASE WHEN d.h = 1 THEN 1 ELSE {i + 1} END) "
            f"AND r.j = (CASE WHEN d.w = 1 THEN 1 ELSE {j + 1} END))")

    def get(col: str) -> str:
        return f"({head} SELECT {col} FROM {u}_e)"
    kind = f"COALESCE(NULLIF({get('kind')}, 'e'), 'n')"  # an empty element spills as 0
    num = f"COALESCE({get('num')}, CAST(0 AS DOUBLE))"
    txt = f"COALESCE({get('txt')}, '')"
    size = f"(SELECT h <> {block[0]} OR w <> {block[1]} FROM {u}_d)" if dynamic and (i, j) == (0, 0) else "FALSE"
    markers = ", ".join(f"'{m}'" for m in SETTINGS)
    first = f"(SELECT err FROM {u}_r WHERE i = 1 AND j = 1)"  # an UNRESOLVED marker wins over the size check
    err = (f"({head} SELECT CASE WHEN {first} IN ({markers}) THEN {first} WHEN {size} THEN '#SPILL!' WHEN NOT EXISTS (SELECT 1 FROM {u}_e) THEN '#N/A' "
           f"ELSE (SELECT err FROM {u}_e) END)")  # no FROM: a missing element still gives one row
    return Expr(num, "any", err, txt, kind)


def _inside(name):
    def fn(c, args, sheet):
        raise Unsupported(f"{name} inside another calculation not supported yet")
    register(name)(fn)


for _n in DYNAMIC:
    _inside(_n)
