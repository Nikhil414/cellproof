"""CONCATENATE, CONCAT, TEXTJOIN, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, PROPER, SUBSTITUTE, REPLACE,
FIND, SEARCH, REPT, VALUE, TEXT, EXACT, CHAR, CLEAN.

Numbers become text the way Excel writes them (expr.num_text: 15 significant digits). Positions are 1-based
and counted in characters, as in Excel. TEXT supports the number formats in NUMBER_FORMATS; dates.py adds
date formats to FORMATS.
"""
from __future__ import annotations

from unexcel.expr import BLANK, CELLS, NO_ERR, Expr, errs, num, num_text, txt, where_area
from unexcel.functions import arity, register
from unexcel.model import Unsupported
from unexcel.parser import Bool, RefNode, Str
from unexcel.refs import resolve


def _t(c, node, sheet: str) -> Expr:
    return txt(c.expr(node, sheet) if node is not None else BLANK)


def _n(c, node, sheet: str) -> Expr:
    return num(c.expr(node, sheet) if node is not None else BLANK)


def _bad(cond: str) -> str:
    return f"(CASE WHEN {cond} THEN '#VALUE!' END)"


CELL_TEXT = ("(CASE kind WHEN 's' THEN COALESCE(txt, '') WHEN 'b' THEN (CASE WHEN num <> 0 THEN 'TRUE' ELSE 'FALSE' END) "
             f"ELSE {num_text('num')} END)")


def _range_texts(c, node, sheet: str) -> tuple[str, str, int]:
    """(SQL list of the range's cell texts in row order, first error, number of empty cells in the range)."""
    area = resolve(node.text, sheet, c.book)
    keys = c.ready(area)
    where = f"FROM {CELLS} WHERE {where_area(area)}"
    err = NO_ERR
    if keys.may_err:
        err = f"(SELECT txt {where} AND kind = 'e' ORDER BY row_num, col_num LIMIT 1)"
    blanks = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1) - len(keys)
    return (f"(SELECT COALESCE(list({CELL_TEXT} ORDER BY row_num, col_num), []) {where} AND kind <> 'e')", err,
            blanks)


@register("CONCATENATE")
def concatenate(c, args, sheet):
    if not args:
        raise Unsupported("CONCATENATE needs arguments")
    parts = [_t(c, a, sheet) for a in args]
    return Expr("(" + " || ".join(p.sql for p in parts) + ")", "txt", errs(*(p.err for p in parts)))


def _pieces(c, args, sheet: str, keep_blanks: bool) -> tuple[list[str], list[str]]:
    lists, err = [], []
    for a in args:
        if isinstance(a, RefNode) and not resolve(a.text, sheet, c.book).is_single:
            items, e, blanks = _range_texts(c, a, sheet)
            if keep_blanks and blanks:
                raise Unsupported("TEXTJOIN keeping the empty cells of a range not supported yet")
            lists.append(items)
            err.append(e)
        else:
            p = _t(c, a, sheet)
            lists.append(f"[{p.sql}]")
            err.append(p.err)
    return lists, err


@register("CONCAT")
def concat(c, args, sheet):
    if not args:
        raise Unsupported("CONCAT needs arguments")
    lists, err = _pieces(c, args, sheet, keep_blanks=False)
    return Expr(f"array_to_string(list_concat({', '.join(lists)}), '')" if len(lists) > 1
                else f"array_to_string({lists[0]}, '')", "txt", errs(*err))


@register("TEXTJOIN")
def textjoin(c, args, sheet):
    if len(args) < 3 or not isinstance(args[1], Bool):
        raise Unsupported("TEXTJOIN needs a delimiter, TRUE/FALSE written in the formula, and values")
    delim = _t(c, args[0], sheet)
    lists, err = _pieces(c, args[2:], sheet, keep_blanks=not args[1].value)
    items = f"list_concat({', '.join(lists)})" if len(lists) > 1 else lists[0]
    if args[1].value:
        items = f"list_filter({items}, x -> x <> '')"
    return Expr(f"array_to_string({items}, {delim.sql})", "txt", errs(delim.err, *err))


def _count_arg(c, args, sheet: str, i: int) -> Expr:
    if len(args) > i and args[i] is not None:
        return _n(c, args[i], sheet)
    return Expr("CAST(1 AS DOUBLE)", "num")


def _side(name: str, fn: str):
    def impl(c, args, sheet):
        arity(name, args, 1, 2)
        t, n = _t(c, args[0], sheet), _count_arg(c, args, sheet, 1)
        return Expr(f"{fn}({t.sql}, CAST(trunc(least(greatest({n.sql}, 0), length({t.sql}))) AS BIGINT))", "txt",
                    errs(t.err, n.err, _bad(f"{n.sql} < 0")))
    register(name)(impl)


_side("LEFT", "left")
_side("RIGHT", "right")


@register("MID")
def mid(c, args, sheet):
    arity("MID", args, 3, 3)
    t, start, n = _t(c, args[0], sheet), _n(c, args[1], sheet), _n(c, args[2], sheet)
    return Expr(f"substr({t.sql}, CAST(trunc(least(greatest({start.sql}, 1), length({t.sql}) + 1)) AS BIGINT), "
                f"CAST(trunc(least(greatest({n.sql}, 0), length({t.sql}))) AS BIGINT))", "txt",
                errs(t.err, start.err, n.err, _bad(f"{start.sql} < 1 OR {n.sql} < 0")))


def _text1(name: str, sql, typ: str = "txt"):
    def impl(c, args, sheet):
        arity(name, args, 1, 1)
        t = _t(c, args[0], sheet)
        return Expr(sql(t.sql), typ, t.err)
    register(name)(impl)


_text1("LEN", lambda t: f"CAST(length({t}) AS DOUBLE)", "num")
_text1("TRIM", lambda t: f"regexp_replace(trim({t}, ' '), ' +', ' ', 'g')")  # Excel also squeezes inner spaces
_text1("UPPER", lambda t: f"upper({t})")
_text1("LOWER", lambda t: f"lower({t})")
_text1("PROPER", lambda t: (f"array_to_string(list_transform(range(length({t})), i -> CASE WHEN i = 0 OR NOT "
                            f"regexp_matches({t}[i], '[[:alpha:]]') THEN upper({t}[i + 1]) ELSE lower({t}[i + 1]) END), '')"))
_text1("CLEAN", lambda t: f"regexp_replace({t}, '[\\x00-\\x1f]', '', 'g')")


@register("SUBSTITUTE")
def substitute(c, args, sheet):
    arity("SUBSTITUTE", args, 3, 4)
    t, old, new = (_t(c, a, sheet) for a in args[:3])
    if len(args) == 3:
        return Expr(f"(CASE WHEN {old.sql} = '' THEN {t.sql} ELSE replace({t.sql}, {old.sql}, {new.sql}) END)", "txt",
                    errs(t.err, old.err, new.err))
    n = _n(c, args[3], sheet)
    parts = f"string_split({t.sql}, {old.sql})"
    k = f"CAST(trunc(least(greatest({n.sql}, 0), len({parts}))) AS BIGINT)"
    return Expr(f"(CASE WHEN {old.sql} = '' OR len({parts}) <= {k} THEN {t.sql} ELSE "
                f"array_to_string({parts}[1:{k}], {old.sql}) || {new.sql} || array_to_string({parts}[{k} + 1:], {old.sql}) END)",
                "txt", errs(t.err, old.err, new.err, n.err, _bad(f"{n.sql} < 1")))


@register("REPLACE")
def replace_(c, args, sheet):
    arity("REPLACE", args, 4, 4)
    t, start, n, new = _t(c, args[0], sheet), _n(c, args[1], sheet), _n(c, args[2], sheet), _t(c, args[3], sheet)
    s = f"CAST(trunc(least(greatest({start.sql}, 1), length({t.sql}) + 1)) AS BIGINT)"
    k = f"CAST(trunc(least(greatest({n.sql}, 0), length({t.sql}))) AS BIGINT)"
    return Expr(f"(substr({t.sql}, 1, {s} - 1) || {new.sql} || substr({t.sql}, {s} + {k}))", "txt",
                errs(t.err, start.err, n.err, new.err, _bad(f"{start.sql} < 1 OR {n.sql} < 0")))


def _find(name: str, sensitive: bool):
    def impl(c, args, sheet):
        arity(name, args, 2, 3)
        what, within = _t(c, args[0], sheet), _t(c, args[1], sheet)
        start = _count_arg(c, args, sheet, 2)
        s = f"TRY_CAST(trunc({start.sql}) AS BIGINT)"
        rest = f"substr({within.sql}, {s})"
        if sensitive:
            pos = f"strpos({rest}, {what.sql})"
            found = f"{pos} > 0"
        else:  # SEARCH: ignores case, * and ? are wildcards, ~ escapes them
            pat = f"regexp_escape(lower({what.sql}))"  # escaped: * -> \\*, ? -> \\?, ~ -> \\~
            for a, b in (("'\\~\\*'", "chr(1)"), ("'\\~\\?'", "chr(2)"), ("'\\*'", "'.*'"), ("'\\?'", "'.'"),
                         ("chr(1)", "'\\*'"), ("chr(2)", "'\\?'")):
                pat = f"replace({pat}, {a}, {b})"
            found = f"regexp_matches(lower({rest}), {pat})"
            pos = f"(length(regexp_extract(lower({rest}), '(?s)^(.*?)(' || {pat} || ')', 1)) + 1)"
        bad = f"{s} IS NULL OR {s} < 1 OR {s} > length({within.sql}) + 1 OR NOT ({found})"
        return Expr(f"(CASE WHEN NOT ({bad}) THEN CAST({pos} + {s} - 1 AS DOUBLE) END)", "num",
                    errs(what.err, within.err, start.err, _bad(bad)))
    register(name)(impl)


_find("FIND", sensitive=True)
_find("SEARCH", sensitive=False)


@register("REPT")
def rept(c, args, sheet):
    arity("REPT", args, 2, 2)
    t, n = _t(c, args[0], sheet), _n(c, args[1], sheet)
    bad = f"({n.sql} < 0 OR length({t.sql}) * trunc({n.sql}) > 32767)"
    sql = (f"(CASE WHEN NOT {bad} THEN CASE WHEN {t.sql} = '' THEN '' "
           f"ELSE repeat({t.sql}, CAST(trunc({n.sql}) AS BIGINT)) END END)")
    return Expr(sql, "txt", errs(t.err, n.err, _bad(bad)))


@register("VALUE")
def value(c, args, sheet):
    arity("VALUE", args, 1, 1)
    return num(c.expr(args[0], sheet))


@register("EXACT")
def exact(c, args, sheet):
    arity("EXACT", args, 2, 2)
    a, b = _t(c, args[0], sheet), _t(c, args[1], sheet)
    return Expr(f"({a.sql} = {b.sql})", "bool", errs(a.err, b.err))


CP1252 = bytes(range(128, 160)).decode("cp1252", errors="replace")


@register("CHAR")
def char(c, args, sheet):
    arity("CHAR", args, 1, 1)
    n = _n(c, args[0], sheet)
    k = f"TRY_CAST(trunc({n.sql}) AS INTEGER)"
    table = "'" + CP1252.replace("'", "''") + "'"
    return Expr(f"(CASE WHEN {k} BETWEEN 128 AND 159 THEN substr({table}, {k} - 127, 1) "
                f"WHEN {k} BETWEEN 1 AND 255 THEN chr({k}) END)", "txt",
                errs(n.err, _bad(f"COALESCE({k} < 1 OR {k} > 255, TRUE)")))  # Windows-1252, the character set of Excel for Windows


def _fixed(decimals: int, grouped: bool, percent: bool):
    def sql(x: str) -> str:
        v = f"round({x} * 100, {decimals})" if percent else f"round({x}, {decimals})"
        v = f"(CASE WHEN {v} = 0 THEN CAST(0 AS DOUBLE) ELSE {v} END)"  # real Excel: TEXT(-0.001,"0.00") is "0.00", not "-0.00"
        body = f"format('{{:,.{decimals}f}}', {v})" if grouped else f"printf('%.{decimals}f', {v})"
        return f"({body} || '%')" if percent else body
    return sql


NUMBER_FORMATS = {"General": num_text}
for d in range(5):
    tail = "." + "0" * d if d else ""
    NUMBER_FORMATS[f"0{tail}"] = _fixed(d, False, False)
    NUMBER_FORMATS[f"#,##0{tail}"] = _fixed(d, True, False)
    NUMBER_FORMATS[f"0{tail}%"] = _fixed(d, False, True)
FORMATS = dict(NUMBER_FORMATS)  # dates.py adds the date formats


@register("TEXT")
def text(c, args, sheet):
    arity("TEXT", args, 2, 2)
    if not isinstance(args[1], Str):
        raise Unsupported("TEXT format must be text written in the formula")
    fmt = FORMATS.get(args[1].value) or FORMATS.get(args[1].value.lower())
    if fmt is None:
        raise Unsupported(f"TEXT format {args[1].value!r} not supported yet")
    if getattr(fmt, "is_date", False) and c.book.date1904:
        raise Unsupported("1904 date system not supported yet")
    x = c.expr(args[0], sheet)
    if x.type == "txt":  # TEXT("abc", "0.00") passes text through
        return x
    x = num(x)
    return Expr(fmt(x.sql), "txt", x.err)
