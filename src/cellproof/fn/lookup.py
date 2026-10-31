"""VLOOKUP, HLOOKUP, MATCH, INDEX, XLOOKUP, LOOKUP, CHOOSE.

The value found is read from the cells table. Its type comes from the cells the range holds: one type ->
that type; mixed, or text with empty cells -> 'any' (known when the SQL runs). An empty cell found reads
as 0, like Excel. Exact matching of text ignores case and (except XLOOKUP's default) honours wildcards.

Every search is an aggregate (arg_min/arg_max by position) over a table that does not depend on the formula's row;
only the WHERE does. DuckDB plans such a subquery as one join for a whole region, where ORDER BY ... LIMIT 1 or a
LEFT JOIN on the row's values cannot be planned at all. Approximate matches on data in order are such an aggregate;
on data out of order, Excel's answer is its binary-search path itself, compiled as such (Walk).
"""
from __future__ import annotations

import math

from cellproof.expr import CELLS, NO_ERR, STALE, Expr, any_of, errs, lit_str, num, unify_all, where_area
from cellproof.fn.criteria import like_pattern
from cellproof.functions import arity, register
from cellproof.model import Unsupported
from cellproof.parser import Bool, Num, RefNode, Str, Unary
from cellproof.refs import Area, resolve


def _area(c, node, sheet: str, fname: str) -> tuple[Area, object]:
    if not isinstance(node, RefNode):
        raise Unsupported(f"{fname} needs a cell range")
    area = resolve(node.text, sheet, c.book)
    return area, c.ready(area)


def _type_of(c, area: Area, keys) -> str:
    kinds = set(keys.kinds) - {"e"}  # an error found is returned as that error, whatever the type
    full = len(keys) == (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1)
    if kinds <= {"n"}:
        return "num"  # an empty cell found reads as 0
    if full and kinds == {"s"}:
        return "txt"
    if full and kinds == {"b"}:
        return "bool"
    return "any"


def _read(pick, typ: str, err: str, r: str = "k.r_") -> Expr:
    """The cell found. pick(x, default) is x of that cell, default when the cell is empty, NULL when nothing is
    found; r prefixes the cell's columns."""
    if typ == "num":
        return Expr(pick(f"{r}num", "0"), "num", err)
    if typ == "txt":
        return Expr(pick(f"{r}txt", "''"), "txt", err)
    if typ == "bool":
        return Expr(pick(f"{r}num <> 0", "FALSE"), "bool", err)
    return any_of(pick(f"{r}kind", "'n'"), pick(f"{r}num", "0"), pick(f"{r}txt", "''"), err)


def _found_error(pick, missing: str) -> str:
    """'missing' when nothing is found, the cell's own error when it holds one, else NULL."""
    found = pick("CASE WHEN k.r_kind = 'e' THEN k.r_txt END", "''")
    return f"(CASE WHEN {found} IS NULL THEN '{missing}' WHEN {found} = '' THEN NULL ELSE {found} END)"


def _static_text(node) -> bool:
    return isinstance(node, Str) and not any(ch in node.value for ch in "*?~")


def _exact(k: str, x: Expr, wild: bool) -> str:
    """SQL true when cell `k` equals the lookup value x (text ignores case; wildcards when `wild`)."""
    text = f"lower(COALESCE({k}.txt, ''))"
    t = (f"{text} LIKE {like_pattern(x.txt if x.type == 'any' else x.sql)} ESCAPE '\\'" if wild
         else f"{text} = lower({x.txt if x.type == 'any' else x.sql})")
    if x.type == "num":
        return f"({k}.kind = 'n' AND {k}.num = {x.sql})"
    if x.type == "txt":
        return f"({k}.kind = 's' AND {t})"
    if x.type == "bool":
        return f"({k}.kind = 'b' AND {k}.num = CAST({x.sql} AS DOUBLE))"
    if x.type == "any":
        return (f"(CASE {x.kind} WHEN 'n' THEN {k}.kind = 'n' AND {k}.num = {x.sql} "
                f"WHEN 's' THEN {k}.kind = 's' AND {t} ELSE {k}.kind = 'b' AND {k}.num = {x.sql} END)")
    raise Unsupported(f"lookup of a value of type {x.type}")


def _near(k: str, x: Expr, cmp: str) -> tuple[str, str]:
    """(condition, sort key) for approximate match: values of x's type that are <= x (cmp '<=') or >= x."""
    if x.type == "num":
        return f"({k}.kind = 'n' AND {k}.num {cmp} {x.sql})", f"{k}.num"
    if x.type == "txt":
        return f"({k}.kind = 's' AND lower({k}.txt) {cmp} lower({x.sql}))", f"lower({k}.txt)"
    raise Unsupported("approximate match on this kind of value not supported yet")


def _line_values(c, line: Area, fname: str) -> list:
    """The (kind, value) of the line's non-empty cells; Unsupported on an error or mixed kinds. Empty cells are
    skipped, as Excel's aggregate skips them on data in order."""
    seen = []
    for key in line.cells():
        cell = c.book.cells.get(key)
        if cell is None or cell.kind == "blank":
            continue
        if cell.kind == "e":
            raise Unsupported(f"{fname} approximate match over a range holding an error")
        seen.append((cell.kind, cell.value.lower() if isinstance(cell.value, str) else cell.value))
    if len({kind for kind, _ in seen}) > 1:
        raise Unsupported(f"{fname} approximate match over a range of mixed types")
    return seen


def _require_sorted(c, line: Area, descending: bool, fname: str) -> None:
    """Excel's approximate match is a binary search. On data that is not in order its answer depends on the search
    path (real Excel: MATCH(150, {0,100,500}, -1) is #N/A although 500 >= 150), so cellproof refuses to guess.
    The order is judged from the saved values, which the verifier then holds the computed values to."""
    values = [v for _, v in _line_values(c, line, fname)]
    if not all(a >= b if descending else a <= b for a, b in zip(values, values[1:])):
        raise Unsupported(f"{fname} approximate match needs {'descending' if descending else 'ascending'} sorted data: "
                          "on unsorted data Excel's answer depends on its binary search")


def _stale(found, err: str) -> str:
    """A compiled search is valid only while the table holds the values it was compiled from (Walk.guard)."""
    guard = getattr(found, "guard", None)
    return f"(CASE WHEN {guard} THEN '{STALE}' ELSE {err} END)" if guard else err


MAX_WALK_STATES = 1000  # the tree's SQL is inlined per pick: a few copies of ~60 characters per probe


def _walk_states(numbers: list[float]) -> list[tuple]:
    """The decision tree of Excel's bisection over the numbers, left to right: (depth, low, high, probe position
    1-based, probed value). The interval [low, high) is the x-range whose search path reaches the probe; None is
    unbounded. Every child's interval is strictly smaller, so the states form a tree, not a cycle."""
    states, todo = {}, [(0, len(numbers) - 1, 1, None, None)]
    while todo:
        lo, hi, depth, ilo, ihi = todo.pop()
        if lo > hi or (lo, hi) in states:
            continue
        mid = (lo + hi) // 2
        states[(lo, hi)] = (depth, ilo, ihi, mid + 1, numbers[mid])
        # a child's x-interval is the parent's narrowed by this branch: right means x >= the probed value,
        # left means x < it; the other side's bound carries down
        todo.append((mid + 1, hi, depth + 1,
                     numbers[mid] if ilo is None else max(ilo, numbers[mid]), ihi))
        todo.append((lo, mid - 1, depth + 1,
                     ilo, numbers[mid] if ihi is None else min(ihi, numbers[mid])))
    return list(states.values())


class Walk:
    """Excel's type-1 approximate match on data out of order: its binary search itself, compiled.

    The search path is fixed by the table's input values (formulas or empty cells would change it, so they are
    refused), so the bisection's decision tree is built here and the SQL answers with the deepest reached probe
    whose value <= x - the last probe the search moves right from, which is Excel's answer. The SQL re-checks
    every probed value against the saved one: a table that changed after translation is #STALE (UNRESOLVED by the
    verifier), never a silently wrong number. Type -1 out of order is not emulated: Excel's answer there (real
    Excel: MATCH(999, {0,1000,5000}, -1) is #N/A) is not derivable from the file, so it stays UNRESOLVED."""

    def __init__(self, x: Expr, line: Area, states: list[tuple], back=None):
        self._first = line.r1 if line.c1 == line.c2 else line.c1
        self._col = "row_num" if line.c1 == line.c2 else "col_num"
        self._other = f"col_num = {line.c1}" if line.c1 == line.c2 else f"row_num = {line.r1}"
        self._sheet = lit_str(line.sheet).sql
        rows = ", ".join(
            f"({d}, {'CAST(NULL AS DOUBLE)' if ilo is None else f'CAST({ilo!r} AS DOUBLE)'}, "
            f"{'CAST(NULL AS DOUBLE)' if ihi is None else f'CAST({ihi!r} AS DOUBLE)'}, {p}, CAST({v!r} AS DOUBLE))"
            for d, ilo, ihi, p, v in states)
        tree = f"(VALUES {rows}) t(d, ilo, ihi, p, v)"
        reach = f"(t.ilo IS NULL OR {x.sql} >= t.ilo) AND (t.ihi IS NULL OR {x.sql} < t.ihi)"
        join = (f"JOIN {CELLS} k ON k.sheet = {self._sheet} AND k.{self._col} = {self._first - 1} + t.p "
                f"AND k.{self._other} AND k.kind = 'n'")
        self.pos = f"(SELECT arg_max(t.p, t.d) FROM {tree} {join} WHERE {reach} AND k.num <= {x.sql})"
        self.guard = (f"(SELECT count(*) FROM {tree} LEFT JOIN {CELLS} k ON k.sheet = {self._sheet} "
                      f"AND k.{self._col} = {self._first - 1} + t.p AND k.{self._other} "
                      f"WHERE k.kind IS DISTINCT FROM 'n' OR k.num IS DISTINCT FROM t.v) > 0")
        self._hit = f"(EXISTS (SELECT 1 FROM {tree} {join} WHERE {reach} AND k.num <= {x.sql}))"
        self.back = back

    def pick(self, value: str, default: str | None = None) -> str:
        if self.back is None:
            return f"(SELECT {value})"  # MATCH reads the position itself
        b0, b1 = self.back
        src = (f"FROM (SELECT k.*, r.kind AS r_kind, r.num AS r_num, r.txt AS r_txt FROM {CELLS} k "
               f"LEFT JOIN {CELLS} r ON r.sheet = {b0} AND {b1} "
               f"WHERE k.sheet = {self._sheet} AND k.{self._col} = {self._first - 1} + {self.pos} "
               f"AND k.{self._other}) k")
        return f"(SELECT COALESCE({value}, {default}) {src})" if default is not None else f"(SELECT {value} {src})"

    def hit(self) -> str:
        return self._hit


def _le_search(c, x: Expr, line: Area, fname: str, back=None):
    """The 'largest value <= x' search (MATCH 1, VLOOKUP/HLOOKUP approximate, LOOKUP). Excel answers from its
    binary search: on data in order that is the largest value <= x, one aggregate; out of order it is the search
    path itself, compiled as such (Walk) when the table is plain input numbers. Text or MATCH -1 out of order
    stays UNRESOLVED: Excel's answer is not derivable from the file."""
    seen = _line_values(c, line, fname)
    values = [v for _, v in seen]
    if all(a <= b for a, b in zip(values, values[1:])):
        return Search(x, line, "le", wild=False, back=back)
    unsorted = (f"{fname} approximate match needs ascending sorted data: on unsorted data Excel's answer depends "
                "on its binary search")
    if x.type != "num" or any(kind != "n" for kind, _ in seen):
        raise Unsupported(unsorted)
    vertical = line.c1 == line.c2
    numbers = []
    for i in range(line.r2 - line.r1 + 1 if vertical else line.c2 - line.c1 + 1):
        row = line.r1 + i if vertical else line.r1
        col = line.c1 if vertical else line.c1 + i
        cell = c.book.cells.get((line.sheet, row, col))
        if cell is None or (cell.kind == "blank" and cell.formula is None):
            raise Unsupported(f"{fname} approximate match on data out of order compiles Excel's search, which "
                              "needs the table without empty cells")
        if cell.formula is not None:
            raise Unsupported(f"{fname} approximate match on data out of order compiles Excel's search, which "
                              "needs the table's values written as inputs")
        if cell.kind != "n" or not math.isfinite(cell.value):
            raise Unsupported(unsorted)
        numbers.append(cell.value)
    states = _walk_states(numbers)
    if len(states) > MAX_WALK_STATES:
        raise Unsupported(f"{fname} approximate match on data out of order compiles Excel's search, and the "
                          f"search is too large to compile ({len(states)} probes)")
    return Walk(x, line, states, back)


class Search:
    """The cell `k` of the 1-D range `line` that a lookup finds. With `back` = (sheet SQL, join condition on k and r)
    also the cell r it returns, read as k.r_kind, k.r_num, k.r_txt (NULL when r is empty)."""

    def __init__(self, x: Expr, line: Area, mode: str, wild: bool, last: bool | None = None, back=None):
        self.pos = "k.row_num" if line.c1 == line.c2 else "k.col_num"
        table = f"{CELLS} k"
        if back:
            table = (f"(SELECT k.*, r.kind AS r_kind, r.num AS r_num, r.txt AS r_txt FROM {CELLS} k "
                     f"LEFT JOIN {CELLS} r ON r.sheet = {back[0]} AND {back[1]} WHERE {where_area(line, 'k')}) k")
        if mode == "exact":  # the first match, the last when `last`
            cond, self.agg, self.key = _exact("k", x, wild), "arg_max" if last else "arg_min", self.pos
        elif mode == "le":  # largest value <= x; the last of equal ones, as Excel's search on sorted data
            cond, key = _near("k", x, "<=")
            self.agg, self.key = "arg_max", f"({key}, {'-' if last is False else ''}{self.pos})"
        else:  # "ge": smallest value >= x, the first of equal ones
            cond, key = _near("k", x, ">=")
            self.agg, self.key = "arg_min", f"({key}, {'-' if last else ''}{self.pos})"
        self.src = f"FROM {table} WHERE {where_area(line, 'k')} AND {cond}"

    def pick(self, value: str, default: str | None = None) -> str:
        if default is not None:  # arg_min skips a NULL value: the cell would look not found
            value = f"COALESCE({value}, {default})"
        return f"(SELECT {self.agg}({value}, {self.key}) {self.src})"

    def hit(self) -> str:
        return f"(EXISTS (SELECT 1 {self.src}))"


def int_literal(node) -> int | None:
    """A whole number written in the formula (-1 parses as minus applied to 1), else None."""
    if isinstance(node, Num):
        return int(node.value)
    if isinstance(node, Unary) and isinstance(node.operand, Num):
        return -int(node.operand.value)
    return None


def _exact_flag(node) -> bool:
    """VLOOKUP/HLOOKUP 4th argument: FALSE or 0 -> exact; omitted, TRUE or 1 -> approximate."""
    if node is None:
        return True
    if isinstance(node, Bool):
        return not node.value
    if isinstance(node, Num):
        return node.value == 0
    raise Unsupported("lookup match type must be TRUE/FALSE written in the formula")


def _xlookup_x(c, args, sheet) -> Expr:
    x = c.expr(args[0], sheet)
    if x.type == "blank":
        raise Unsupported("lookup of an empty cell not supported yet")
    return x


def _table_lookup(fname: str, by_row: bool):
    def fn(c, args, sheet):
        arity(fname, args, 3, 4)
        exact = len(args) == 4 and _exact_flag(args[3])
        if not isinstance(args[2], Num):
            raise Unsupported(f"{fname} index must be a number written in the formula")
        table, _ = _area(c, args[1], sheet, fname)
        offset = int(args[2].value) - 1
        if by_row:
            line = Area(table.sheet, table.r1, table.c1, table.r2, table.c1)
            target = Area(table.sheet, table.r1, table.c1 + offset, table.r2, table.c1 + offset)
            ok = table.c1 <= target.c1 <= table.c2
            hit = f"r.row_num = k.row_num AND r.col_num = {target.c1}"
        else:
            line = Area(table.sheet, table.r1, table.c1, table.r1, table.c2)
            target = Area(table.sheet, table.r1 + offset, table.c1, table.r1 + offset, table.c2)
            ok = table.r1 <= target.r1 <= table.r2
            hit = f"r.col_num = k.col_num AND r.row_num = {target.r1}"
        if not ok:
            raise Unsupported(f"{fname} {'column' if by_row else 'row'} index outside the table")
        x = _xlookup_x(c, args, sheet)
        if exact:
            found = Search(x, line, "exact", wild=True, back=("k.sheet", hit))
        else:
            found = _le_search(c, x, line, fname, back=("k.sheet", hit))
        keys = c.ready(target)
        return _read(found.pick, _type_of(c, target, keys),
                     _stale(found, errs(x.err, _found_error(found.pick, "#N/A"))))
    return fn


register("VLOOKUP")(_table_lookup("VLOOKUP", by_row=True))
register("HLOOKUP")(_table_lookup("HLOOKUP", by_row=False))


def _line(c, node, sheet: str, fname: str) -> Area:
    area, _ = _area(c, node, sheet, fname)
    if area.r1 != area.r2 and area.c1 != area.c2:
        raise Unsupported(f"{fname} needs a single row or column")
    return area


@register("MATCH")
def match(c, args, sheet):
    arity("MATCH", args, 2, 3)
    kind = 1 if len(args) == 2 or args[2] is None else int_literal(args[2])
    if kind not in (-1, 0, 1):
        raise Unsupported("MATCH type must be -1, 0 or 1 written in the formula")
    line = _line(c, args[1], sheet, "MATCH")
    x = _xlookup_x(c, args, sheet)
    if kind == -1:
        _require_sorted(c, line, True, "MATCH")
        found = Search(x, line, "ge", wild=True)
    else:
        found = _le_search(c, x, line, "MATCH") if kind == 1 else Search(x, line, "exact", wild=True)
    first = line.r1 if line.c1 == line.c2 else line.c1
    return Expr(found.pick(f"CAST({found.pos} - {first} + 1 AS DOUBLE)"), "num",
                _stale(found, errs(x.err, f"(CASE WHEN NOT {found.hit()} THEN '#N/A' END)")))


@register("XLOOKUP")
def xlookup(c, args, sheet):
    arity("XLOOKUP", args, 3, 6)
    modes = []
    for i, default in ((4, 0), (5, 1)):
        node = args[i] if len(args) > i else None
        mode = default if node is None else int_literal(node)
        if mode is None:
            raise Unsupported("XLOOKUP modes must be numbers written in the formula")
        modes.append(mode)
    match_mode, search_mode = modes
    if match_mode not in (0, -1, 1, 2) or search_mode not in (1, -1):
        raise Unsupported("XLOOKUP binary search modes not supported yet")
    look = _line(c, args[1], sheet, "XLOOKUP")
    back = _line(c, args[2], sheet, "XLOOKUP")
    by_row = look.c1 == look.c2
    if (by_row and (back.c1 != back.c2 or back.r2 - back.r1 != look.r2 - look.r1)) or \
            (not by_row and (back.r1 != back.r2 or back.c2 - back.c1 != look.c2 - look.c1)):
        raise Unsupported("XLOOKUP return range must be one row or column the size of the lookup range")
    x = _xlookup_x(c, args, sheet)
    mode = {0: "exact", 2: "exact", -1: "le", 1: "ge"}[match_mode]
    hit = (f"r.row_num = k.row_num + {back.r1 - look.r1} AND r.col_num = {back.c1}" if by_row
           else f"r.col_num = k.col_num + {back.c1 - look.c1} AND r.row_num = {back.r1}")
    search = Search(x, look, mode, wild=match_mode == 2, last=search_mode == -1, back=(lit_str(back.sheet).sql, hit))
    keys = c.ready(back)
    found = _read(search.pick, _type_of(c, back, keys), errs(x.err, _found_error(search.pick, "#N/A")))
    if len(args) < 4 or args[3] is None:
        return found
    missing = c.expr(args[3], sheet)
    found, missing = unify_all([found, missing])
    hit_sql = search.hit()
    pick_err = f"(CASE WHEN {hit_sql} THEN {found.err} ELSE {missing.err} END)"

    def case(a: str, b: str) -> str:
        return f"(CASE WHEN {hit_sql} THEN {a} ELSE {b} END)"
    return Expr(case(found.sql, missing.sql), found.type, errs(x.err, pick_err), case(found.txt, missing.txt),
                case(found.kind, missing.kind))


@register("LOOKUP")
def lookup(c, args, sheet):
    """Always approximate: the largest value <= x. Vector form searches args[1] and returns the cell at the same
    position of args[2] (either direction); array form searches the first row of a block wider than tall, else its
    first column, and returns the last row or column."""
    arity("LOOKUP", args, 2, 3)
    if len(args) == 3 and args[2] is not None:
        look, back = _line(c, args[1], sheet, "LOOKUP"), _line(c, args[2], sheet, "LOOKUP")
        if len(list(look.cells())) != len(list(back.cells())):
            raise Unsupported("LOOKUP result range must be the same size as the lookup range")
    else:
        area, _ = _area(c, args[1], sheet, "LOOKUP")
        if area.c2 - area.c1 > area.r2 - area.r1:
            look = Area(area.sheet, area.r1, area.c1, area.r1, area.c2)
            back = Area(area.sheet, area.r2, area.c1, area.r2, area.c2)
        else:
            look = Area(area.sheet, area.r1, area.c1, area.r2, area.c1)
            back = Area(area.sheet, area.r1, area.c2, area.r2, area.c2)
    x = _xlookup_x(c, args, sheet)
    step = f"k.row_num - {look.r1}" if look.c1 == look.c2 else f"k.col_num - {look.c1}"
    hit = (f"r.col_num = {back.c1} AND r.row_num = {back.r1} + {step}" if back.c1 == back.c2
           else f"r.row_num = {back.r1} AND r.col_num = {back.c1} + {step}")
    found = _le_search(c, x, look, "LOOKUP", back=(lit_str(back.sheet).sql, hit))
    keys = c.ready(back)
    return _read(found.pick, _type_of(c, back, keys),
                 _stale(found, errs(x.err, _found_error(found.pick, "#N/A"))))


@register("INDEX")
def index(c, args, sheet):
    arity("INDEX", args, 2, 3)
    area, keys = _area(c, args[0], sheet, "INDEX")
    height, width = area.r2 - area.r1 + 1, area.c2 - area.c1 + 1
    first = num(c.expr(args[1], sheet))
    second = num(c.expr(args[2], sheet)) if len(args) == 3 and args[2] is not None else None
    if second is None:
        if height == 1:
            row, col = Expr("CAST(1 AS DOUBLE)", "num"), first  # one row: the number picks the column
        elif width == 1:
            row, col = first, Expr("CAST(1 AS DOUBLE)", "num")
        else:
            raise Unsupported("INDEX of a block needs a row and a column")
    else:
        row, col = first, second
    r = f"(CASE WHEN trunc({row.sql}) BETWEEN 1 AND {height} THEN CAST(trunc({row.sql}) AS BIGINT) END)"
    k = f"(CASE WHEN trunc({col.sql}) BETWEEN 1 AND {width} THEN CAST(trunc({col.sql}) AS BIGINT) END)"
    src = (f"FROM {CELLS} r WHERE r.sheet = {lit_str(area.sheet).sql} "
           f"AND r.row_num = {area.r1} - 1 + {r} AND r.col_num = {area.c1} - 1 + {k}")

    def pick(value: str, default: str) -> str:  # at most one cell matches; none when it is empty
        return f"COALESCE((SELECT max({value}) {src}), {default})"
    bounds = (f"(CASE WHEN trunc({row.sql}) < 0 OR trunc({col.sql}) < 0 THEN '#VALUE!' "  # real Excel: negative
              f"WHEN {r} IS NULL OR {k} IS NULL OR {r} < 1 OR {r} > {height} OR {k} < 1 OR {k} > {width} THEN '#REF!' END)")
    found = f"(SELECT max(CASE WHEN r.kind = 'e' THEN r.txt END) {src})"
    return _read(pick, _type_of(c, area, keys), errs(row.err, col.err, bounds, found), "r.")


@register("CHOOSE")
def choose(c, args, sheet):
    if len(args) < 2:
        raise Unsupported("CHOOSE needs an index and values")
    i = num(c.expr(args[0], sheet))
    values = unify_all([c.expr(a, sheet) for a in args[1:]])
    n = f"TRY_CAST(trunc({i.sql}) AS BIGINT)"

    def case(parts: list[str]) -> str:
        if len(set(parts)) == 1:
            return parts[0]
        return "(CASE " + " ".join(f"WHEN {n} = {j} THEN {p}" for j, p in enumerate(parts, 1)) + " END)"
    bad = f"(CASE WHEN {n} IS NULL OR {n} < 1 OR {n} > {len(values)} THEN '#VALUE!' END)"
    err = errs(i.err, bad, case([v.err for v in values]) if any(v.err != NO_ERR for v in values) else NO_ERR)
    return Expr(case([v.sql for v in values]), values[0].type, err, case([v.txt for v in values]),
                case([v.kind for v in values]))
