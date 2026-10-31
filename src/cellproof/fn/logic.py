"""IFS, SWITCH, NOT, XOR, IFNA, ISNA, ISERROR, ISERR, ISNUMBER, ISTEXT, ISLOGICAL, ISBLANK, NA, TRUE, FALSE.

The IS functions look at a value without failing: an error in their argument is their answer, not their error.
"""
from __future__ import annotations

from cellproof.expr import NO_ERR, ZERO, Expr, as_bool, compare, error, errs, keep_locale, num, num_strict, pick, unify, unify_all
from cellproof.functions import arity, logical_args, register
from cellproof.model import Unsupported, a1
from cellproof.parser import Binary, Call, Postfix, RefNode, Unary
from cellproof.refs import resolve


def _chain(branches: list[tuple[Expr, Expr]], default: Expr | None) -> Expr:
    """CASE WHEN cond1 THEN value1 ... ELSE default (or #N/A): IFS and SWITCH."""
    values = unify_all([v for _, v in branches] + ([default] if default is not None else []))
    fallback = values[-1] if default is not None else None

    def case(attr: str) -> str:
        parts = [getattr(v, attr) for v in values]
        if len(set(parts)) == 1:
            return parts[0]
        whens = " ".join(f"WHEN {cnd.sql} THEN {getattr(v, attr)}" for (cnd, _), v in zip(branches, values))
        return f"(CASE {whens} ELSE {getattr(fallback, attr) if fallback else 'NULL'} END)"
    err_whens = []
    for (cnd, _), v in zip(branches, values):
        if cnd.err != NO_ERR:  # a condition that fails stops the chain, as in Excel
            err_whens.append(f"WHEN ({cnd.err}) IS NOT NULL THEN {cnd.err}")
        err_whens.append(f"WHEN {cnd.sql} THEN {v.err}")
    last = fallback.err if fallback else "'#N/A'"
    return Expr(case("sql"), values[0].type, f"(CASE {' '.join(err_whens)} ELSE {last} END)",
                case("txt"), case("kind"))


@register("IFS")
def ifs(c, args, sheet):
    if not args or len(args) % 2:
        raise Unsupported("IFS needs condition/value pairs")
    pairs = [(as_bool(c.expr(args[i], sheet)), c.expr(args[i + 1], sheet)) for i in range(0, len(args), 2)]
    return _chain(pairs, None)


@register("SWITCH")
def switch(c, args, sheet):
    if len(args) < 3:
        raise Unsupported("SWITCH needs a value and value/result pairs")
    def zero_if_blank(e):  # real Excel: SWITCH reads an empty cell as 0 (it matches 0, not FALSE or "")
        return ZERO if e.type == "blank" else e
    x = zero_if_blank(c.expr(args[0], sheet))
    rest = args[1:]
    default = c.expr(rest[-1], sheet) if len(rest) % 2 else None
    pairs = [(compare(x, zero_if_blank(c.expr(rest[i], sheet)), "="), c.expr(rest[i + 1], sheet)) for i in range(0, len(rest) - 1, 2)]
    return _chain(pairs, default)


@register("NOT")
def not_(c, args, sheet):
    arity("NOT", args, 1, 1)
    x = as_bool(c.expr(args[0], sheet))
    return Expr(f"(NOT {x.sql})", "bool", x.err)


@register("XOR")
def xor(c, args, sheet):
    values, err = logical_args(c, "XOR", args, sheet)
    total = " + ".join(f"CAST(COALESCE({v}, FALSE) AS INTEGER)" for v in values)
    return Expr(f"(({total}) % 2 = 1)", "bool", err)


@register("IFNA")
def ifna(c, args, sheet):
    arity("IFNA", args, 2, 2)
    value, fallback = unify(c.expr(args[0], sheet), c.expr(args[1], sheet))
    if value.err == NO_ERR:
        return value
    na = f"COALESCE(({value.err}) = '#N/A', FALSE)"
    return pick(f"(NOT {na})", value, fallback, f"(CASE WHEN {na} THEN {fallback.err} ELSE {value.err} END)")


def is_arg(c, node, sheet: str) -> Expr:
    """An IS function's argument: in an array formula a range is padded per element, not sized (see _element)."""
    if c.spill is not None and isinstance(node, RefNode):
        from cellproof.refs import resolve
        area = resolve(node.text, sheet, c.book)
        if not area.is_single:
            return c._element(area, sized=False)
    return c.expr(node, sheet)


def _is(name: str, test):
    def fn(c, args, sheet):
        arity(name, args, 1, 1)
        x = is_arg(c, args[0], sheet)
        return Expr(test(x), "bool", keep_locale(x.err))
    register(name)(fn)


_is("ISNA", lambda x: f"COALESCE(({x.err}) = '#N/A', FALSE)")
_is("ISERROR", lambda x: f"(({x.err}) IS NOT NULL)")
_is("ISERR", lambda x: f"COALESCE(({x.err}) <> '#N/A', FALSE)")


def _kind_is(kind: str, typ: str):
    def test(x: Expr) -> str:
        if x.type == "any":
            return f"(({x.err}) IS NULL AND {x.kind} = '{kind}')"
        return f"(({x.err}) IS NULL)" if x.type == typ else "FALSE"
    return test


_is("ISNUMBER", _kind_is("n", "num"))
_is("ISTEXT", _kind_is("s", "txt"))
_is("ISLOGICAL", _kind_is("b", "bool"))


@register("ISBLANK")
def isblank(c, args, sheet):
    arity("ISBLANK", args, 1, 1)
    if not isinstance(args[0], RefNode):
        return Expr("FALSE", "bool")  # ISBLANK("") is FALSE: only an empty cell is blank
    area = resolve(args[0].text, sheet, c.book)
    key = area.key if area.is_single else None
    if not area.is_single:
        if c.spill is None:
            raise Unsupported("ISBLANK of a range where one value is expected not supported yet")
        key = c.element_key(area)
        if key is None:  # beyond the range: Excel's #N/A padding, which is not blank
            return Expr("FALSE", "bool")
    cell = c.book.cells.get(key)
    empty = cell is None or (cell.formula is None and cell.kind == "blank")
    return Expr("TRUE" if empty else "FALSE", "bool")


@register("NA")
def na(c, args, sheet):
    arity("NA", args, 0, 0)
    return error("#N/A")


@register("TRUE")
def true(c, args, sheet):
    arity("TRUE", args, 0, 0)
    return Expr("TRUE", "bool")


@register("FALSE")
def false(c, args, sheet):
    arity("FALSE", args, 0, 0)
    return Expr("FALSE", "bool")


LET_PREFIX = "_xlpm."  # a file stores LET's names as _xlpm.name; a cell reference never starts so


def _subst(node, env: dict):
    """node with every bound name replaced by its value's tree; an inner LET's own names shadow env's."""
    if isinstance(node, RefNode):
        return env.get(node.text.lower(), node)
    if isinstance(node, (Unary, Postfix)):
        return type(node)(node.op, _subst(node.operand, env))
    if isinstance(node, Binary):
        return Binary(node.op, _subst(node.left, env), _subst(node.right, env))
    if isinstance(node, Call):
        if node.name != "LET":
            return Call(node.name, tuple(None if a is None else _subst(a, env) for a in node.args))
        local, out, args = dict(env), [], node.args
        for i in range(0, len(args) - 1, 2):
            out += [args[i], None if args[i + 1] is None else _subst(args[i + 1], local)]  # the value sees env
            if isinstance(args[i], RefNode):
                local.pop(args[i].text.lower(), None)  # from here on the name is the inner LET's
        if len(args) % 2:
            out.append(None if args[-1] is None else _subst(args[-1], local))
        return Call("LET", tuple(out))
    return node


@register("LET")
def let(c, args, sheet):
    """LET(name1, value1, ..., calculation): each name stands for its value's formula. Substituting the trees
    (rather than computing each value once) keeps a range a range (LET(r, A1:A9, SUM(r))); formulas have no side
    effects, so the result is the same, and a value that is never used never shows its error, as in Excel."""
    if len(args) < 3 or len(args) % 2 == 0:
        raise Unsupported("LET needs name/value pairs and a calculation")
    env = {}
    for i in range(0, len(args) - 1, 2):
        name = args[i]
        if not (isinstance(name, RefNode) and name.text.lower().startswith(LET_PREFIX)):
            raise Unsupported(f"LET name {getattr(name, 'text', '?')!r} is not a name")
        env[name.text.lower()] = _subst(args[i + 1], env)
    return c.expr(_subst(args[-1], env), sheet)


@register("SINGLE")
def single(c, args, sheet):
    """@x (stored _xlfn.SINGLE(x)): implicit intersection, even inside an array formula."""
    arity("SINGLE", args, 1, 1)
    a = args[0]
    if isinstance(a, RefNode):
        from cellproof.refs import resolve
        area = resolve(a.text, sheet, c.book)
        if not area.is_single:
            spill, c.spill = c.spill, None
            try:
                return c._intersect(area, a.text)
            finally:
                c.spill = spill
    return c.expr(a, sheet)


def _position(name: str, which: int):
    """ROW/COLUMN: of the formula's own cell, of one cell, or - in an array formula - of this cell's element."""
    def fn(c, args, sheet):
        arity(name, args, 0, 1)
        if not args or args[0] is None:
            if c.at is None:
                raise Unsupported(f"{name}() in a block of rows not supported yet (stays row by row)")
            return Expr(f"CAST({c.at[which]} AS DOUBLE)", "num")
        a = args[0]
        if not isinstance(a, RefNode):
            raise Unsupported(f"{name} of something other than a reference")
        if c.rowrefs is not None and a.text in c.rowrefs:
            raise Unsupported(f"{name} of a reference that moves with the row not supported yet")
        from cellproof.refs import resolve
        area = resolve(a.text, sheet, c.book)
        if area.is_single:
            return Expr(f"CAST({area.key[which]} AS DOUBLE)", "num")
        if c.spill is None:  # legacy Excel takes the first row/column, modern Excel spills: not checked
            raise Unsupported(f"{name} of a range outside an array formula not supported yet")
        # real Excel (formulas' test.xlsx): ROW of a range is one column of its row numbers, COLUMN one row of its
        # column numbers, so they repeat across (down) the block
        first, count = (area.r1, area.r2 - area.r1 + 1) if which == 1 else (area.c1, area.c2 - area.c1 + 1)
        at = c.spill[0] if which == 1 else c.spill[1]
        c.spill_dims.append((count, 1) if which == 1 else (1, count))
        at = 0 if count == 1 else at
        return error("#N/A") if at >= count else Expr(f"CAST({first + at} AS DOUBLE)", "num")
    register(name)(fn)


_position("ROW", 1)
_position("COLUMN", 2)


@register("N")
def n_(c, args, sheet):
    """A number (dates are numbers) is itself, TRUE 1, FALSE 0, text 0, an error that error."""
    arity("N", args, 1, 1)
    a = args[0]
    if isinstance(a, RefNode):
        from cellproof.refs import resolve
        area = resolve(a.text, sheet, c.book)
        if not area.is_single:
            if c.spill is None:
                raise Unsupported("N of a range outside an array formula not supported yet")
            # real Excel (formulas' test.xlsx): N of a range is N of its first cell, in every cell of the block
            a = RefNode(f"'{area.sheet.replace(chr(39), chr(39) * 2)}'!{a1((area.sheet, area.r1, area.c1)).split('!')[-1]}")
    x = c.expr(a, sheet)
    if x.type in ("num", "blank"):
        return num(x)
    if x.type == "bool":
        return Expr(f"CAST({x.sql} AS DOUBLE)", "num", x.err)
    if x.type in ("txt", "err"):
        return Expr("CAST(0 AS DOUBLE)", "num", x.err)
    return Expr(f"(CASE WHEN {x.kind} = 's' THEN CAST(0 AS DOUBLE) ELSE {x.sql} END)", "num", x.err)


def _parity(name: str, odd: bool):
    def fn(c, args, sheet):
        arity(name, args, 1, 1)
        x = num_strict(c.expr(args[0], sheet))  # an Analysis ToolPak function: TRUE/FALSE is #VALUE!
        return Expr(f"(abs(trunc({x.sql})) % 2 {'=' if odd else '<>'} 1)", "bool", x.err)
    register(name)(fn)


_parity("ISODD", True)
_parity("ISEVEN", False)
