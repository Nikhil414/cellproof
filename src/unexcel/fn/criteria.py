"""COUNTIF, COUNTIFS, SUMIF, SUMIFS, AVERAGEIF, AVERAGEIFS, MINIFS, MAXIFS.

A criterion is a number (matches equal numbers, and text that reads as that number), TRUE/FALSE, or text
that may start with =, <>, <, >, <= or >= and may use the wildcards * and ? (~ escapes them). Text criteria
can be built while the sheet runs (">"&B1), so they are parsed in SQL, not in Python.
Empty cells have no row in the cells table. They match "" and "=" and "<>x", so they are counted from
the range's size.
"""
from __future__ import annotations

from unexcel.expr import CELLS, NO_ERR, Expr, errs, lit_str, parse_num, txt, where_area
from unexcel.functions import arity, note_text_numbers, register
from unexcel.model import Unsupported
from unexcel.parser import RefNode
from unexcel.refs import Area, resolve

OPS = (">=", "<=", "<>", ">", "<", "=")


def like_pattern(rhs: str) -> str:
    """Excel wildcards -> a LIKE pattern (ESCAPE '\\'): * -> %, ? -> _, ~* and ~? are literal."""
    p = f"lower({rhs})"
    for a, b in (("'\\'", "'\\\\'"), ("'%'", "'\\%'"), ("'_'", "'\\_'"), ("'~*'", "chr(1)"), ("'~?'", "chr(2)"),
                 ("'*'", "'%'"), ("'?'", "'_'"), ("chr(1)", "'*'"), ("chr(2)", "'?'")):
        p = f"replace({p}, {a}, {b})"
    return p


def match_sql(alias: str, crit: Expr, q: str) -> tuple[str, str, str]:
    """(SQL true when cell `alias` meets the criterion, SQL true when an empty cell does, a one-row table `q`).

    A text criterion is parsed once, in the one-row table `q` the caller puts in its FROM clause:
    repeating the criterion SQL inside every comparison overflowed DuckDB's expression depth.
    """
    if crit.err != NO_ERR:
        raise Unsupported("criteria that can be an error not supported yet")
    if crit.type == "blank":
        raise Unsupported("criteria from an empty cell not supported yet")
    k, n, t = f"{alias}.kind", f"{alias}.num", f"COALESCE({alias}.txt, '')"  # the seed CSV reads '' back as NULL
    if crit.type == "num":
        return f"COALESCE(({k} = 'n' AND {n} = {crit.sql}) OR ({k} = 's' AND {parse_num(t)} = {crit.sql}), FALSE)", \
            "FALSE", f"(SELECT 1 AS one) {q}"
    if crit.type == "bool":
        return f"COALESCE({k} = 'b' AND {n} = CAST({crit.sql} AS DOUBLE), FALSE)", "FALSE", f"(SELECT 1 AS one) {q}"
    ops = "(CASE " + " ".join(f"WHEN c LIKE '{o}%' THEN '{o}'" for o in OPS) + " ELSE '' END)"
    parsed = (f"(SELECT op, rhs, {parse_num('rhs')} AS rn, "
              f"(CASE upper(rhs) WHEN 'TRUE' THEN 1.0 WHEN 'FALSE' THEN 0.0 END) AS rb, {like_pattern('rhs')} AS pat "
              f"FROM (SELECT op, substr(c, length(op) + 1) AS rhs FROM (SELECT c, {ops} AS op "
              f"FROM (SELECT {txt(crit).sql} AS c)))) {q}")
    op, rhs, rn, rb = f"{q}.op", f"{q}.rhs", f"{q}.rn", f"{q}.rb"

    def order(a: str, b: str) -> str:
        return " ".join(["(CASE", *(f"WHEN {op} = '{o}' THEN {a} {o} {b}" for o in (">", "<", ">=", "<=")), "END)"])

    eq = (f"COALESCE(CASE WHEN {rn} IS NOT NULL THEN ({k} = 'n' AND {n} = {rn}) OR ({k} = 's' AND {parse_num(t)} = {rn}) "
          f"WHEN {rb} IS NOT NULL THEN {k} = 'b' AND {n} = {rb} "
          f"ELSE {k} = 's' AND lower({t}) LIKE {q}.pat ESCAPE '\\' END, FALSE)")
    ordered = (f"COALESCE(CASE WHEN {rn} IS NOT NULL THEN {k} = 'n' AND {order(n, rn)} "
               f"WHEN {rb} IS NOT NULL THEN {k} = 'b' AND {order(n, rb)} "
               f"ELSE {k} = 's' AND {order(f'lower({t})', f'lower({rhs})')} END, FALSE)")
    # Real Excel: "<>5" skips only the number 5. The text "5" matches =5 but still counts as "<>5".
    same = (f"COALESCE(CASE WHEN {rn} IS NOT NULL THEN {k} = 'n' AND {n} = {rn} "
            f"WHEN {rb} IS NOT NULL THEN {k} = 'b' AND {n} = {rb} "
            f"ELSE {k} = 's' AND lower({t}) LIKE {q}.pat ESCAPE '\\' END, FALSE)")
    # Real Excel: "<>" alone counts every cell that exists, including a formula that returns "".
    match = (f"(CASE WHEN {op} = '<>' AND {rhs} = '' THEN TRUE WHEN {op} = '<>' THEN NOT {same} "
             f"WHEN {op} IN ('', '=') THEN {eq} ELSE {ordered} END)")
    blank = f"(({op} IN ('', '=') AND {rhs} = '') OR ({op} = '<>' AND {rhs} <> ''))"
    return match, blank, parsed


def _range(c, node, sheet: str, fname: str) -> tuple[Area, object]:
    if not isinstance(node, RefNode):
        raise Unsupported(f"{fname} needs a cell range")
    area = resolve(node.text, sheet, c.book)
    return area, c.ready(area)


def _column(c, node, sheet: str, fname: str, height: int | None = None) -> tuple[Area, object]:
    area, keys = _range(c, node, sheet, fname)
    if area.c1 != area.c2 or (height is not None and area.r2 - area.r1 != height):
        raise Unsupported(f"{fname} ranges must be single columns of the same height")
    return area, keys


def _pairs(c, args, sheet: str, fname: str, height: int | None) -> list[tuple[Area, Expr]]:
    pairs = []
    for i in range(0, len(args), 2):
        area, _ = _column(c, args[i], sheet, fname, height)
        height = area.r2 - area.r1
        pairs.append((area, c.expr(args[i + 1], sheet)))
    return pairs


def _joined(pairs, base: str, base_r1) -> tuple[str, str, str]:
    """(one-row criteria tables, LEFT JOINs of each criteria range to the driving rows, AND of conditions, AND of blanks)."""
    parsed, joins, conds, blanks = [], [], [], []
    for i, (area, crit) in enumerate(pairs, 1):
        k = f"k{i}"
        joins.append(f"LEFT JOIN {CELLS} {k} ON {k}.sheet = {lit_str(area.sheet).sql} "
                     f"AND {k}.col_num = {area.c1} AND {k}.row_num = {base} + {area.r1 - base_r1}")
        match, blank, source = match_sql(k, crit, f"q{i}")
        parsed.append(source)
        conds.append(f"(CASE WHEN {k}.kind IS NULL THEN {blank} ELSE {match} END)")
        blanks.append(blank)
    return parsed, " ".join(joins), " AND ".join(conds), " AND ".join(blanks)


@register("COUNTIF")
def countif(c, args, sheet):
    arity("COUNTIF", args, 2, 2)
    area, _ = _range(c, args[0], sheet, "COUNTIF")
    match, blank, source = match_sql("x", c.expr(args[1], sheet), "q1")
    size = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1)
    inside = f"FROM {CELLS} x WHERE {where_area(area, 'x')}"
    return Expr(f"(SELECT (SELECT COUNT(*) FILTER (WHERE {match}) {inside}) "
                f"+ CASE WHEN {blank} THEN {size} - (SELECT COUNT(*) {inside}) ELSE 0 END FROM {source})", "num")


@register("COUNTIFS")
def countifs(c, args, sheet):
    if not args or len(args) % 2:
        raise Unsupported("COUNTIFS needs criteria pairs")
    pairs = _pairs(c, args, sheet, "COUNTIFS", None)
    height = pairs[0][0].r2 - pairs[0][0].r1 + 1
    rows = " UNION ".join(f"SELECT row_num - {a.r1} AS off FROM {CELLS} WHERE {where_area(a)}" for a, _ in pairs)
    tables, joins, cond, blanks = _joined(pairs, "o.off", 0)
    return Expr(f"(WITH o AS ({rows}) SELECT (SELECT COUNT(*) FILTER (WHERE {cond}) FROM o {joins}) "
                f"+ CASE WHEN {blanks} THEN {height} - (SELECT COUNT(*) FROM o) ELSE 0 END FROM {', '.join(tables)})", "num")


def _aggregate(c, target: Area, pairs, how: str) -> Expr:
    """SUM / AVERAGE / MIN / MAX of the target cells whose rows meet every criterion."""
    keys = c.ready(target)
    note_text_numbers(c, keys)
    tables, joins, cond, _ = _joined(pairs, "s.row_num", target.r1)
    crossed = " ".join(f"CROSS JOIN {t}" for t in tables)
    src = f"FROM {CELLS} s {crossed} {joins} WHERE {where_area(target, 's')} AND {cond}"
    numbers = "FILTER (WHERE s.kind = 'n')"
    err = NO_ERR
    if keys.may_err:  # a matched error cell makes the result that error
        err = f"(SELECT s.txt {src} AND s.kind = 'e' ORDER BY s.row_num LIMIT 1)"
    if how == "SUM":
        return Expr(f"(SELECT COALESCE(SUM(s.num) {numbers}, CAST(0 AS DOUBLE)) {src})", "num", err)
    if how in ("MIN", "MAX"):
        return Expr(f"(SELECT COALESCE({how}(s.num) {numbers}, CAST(0 AS DOUBLE)) {src})", "num", err)
    count = f"(SELECT COUNT(*) {numbers} {src})"
    return Expr(f"(SELECT SUM(s.num) {numbers} / NULLIF(COUNT(*) {numbers}, 0) {src})", "num",
                errs(err, f"(CASE WHEN {count} = 0 THEN '#DIV/0!' END)"))


def _resized(area: Area, like: Area) -> Area:
    """Excel's SUMIF uses the sum range's top-left cell and the criteria range's size."""
    return Area(area.sheet, area.r1, area.c1, area.r1 + (like.r2 - like.r1), area.c1)


def _if(how: str):
    def fn(c, args, sheet):
        fname = f"{how}IF" if how in ("SUM", "AVERAGE") else how
        arity(fname, args, 2, 3)
        area, _ = _column(c, args[0], sheet, fname)
        target = area
        if len(args) == 3:
            if not isinstance(args[2], RefNode):
                raise Unsupported(f"{fname} needs a cell range")
            target = _resized(resolve(args[2].text, sheet, c.book), area)
        return _aggregate(c, target, [(area, c.expr(args[1], sheet))], how)
    return fn


def _ifs(how: str):
    def fn(c, args, sheet):
        fname = f"{how}IFS"
        if len(args) < 3 or len(args) % 2 == 0:
            raise Unsupported(f"{fname} needs a sum range and criteria pairs")
        target, _ = _column(c, args[0], sheet, fname)
        pairs = _pairs(c, args[1:], sheet, fname, target.r2 - target.r1)
        return _aggregate(c, target, pairs, how)
    return fn


register("SUMIF")(_if("SUM"))
register("AVERAGEIF")(_if("AVERAGE"))
register("SUMIFS")(_ifs("SUM"))
register("AVERAGEIFS")(_ifs("AVERAGE"))
register("MINIFS")(_ifs("MIN"))
register("MAXIFS")(_ifs("MAX"))
