"""Order formula cells so every cell comes after the cells it reads.

Ranges are not expanded into cell sets (that is quadratic for a running SUM($B$1:Bn)); they are O(n log n) when the
cells they read are ordered in ascending order (see the ponytail comment in the main loop for the bottom-up case). Per
column, a range is an interval of that column's formula cells, and a Fenwick tree of "not yet ordered" counts says
whether the interval is empty and which cell to wait for next. Cells are released level by level: a node's level is one
more than the last node it waited for, so every cell that needs only inputs comes first, then the cells that need those.

A group of cells (a region) is one node: it waits for everything its members read from outside the group, and when it is
released all its members are ordered at once, consecutively, row-major.
"""
from __future__ import annotations

import heapq
from bisect import bisect_left, bisect_right
from collections import defaultdict

from cellproof.fenwick import Fenwick
from cellproof.model import Book, Key, Unsupported
from cellproof.r1c1 import lexed
from cellproof.refs import resolve

CYCLE = "circular reference (or depends on one)"

Column = tuple[str, int]  # (sheet, col)


def formula_order(book: Book, groups=()) -> tuple[list[Key], dict[Key, str]]:
    formulas = {k for k, c in book.cells.items() if c.formula is not None}
    grouped = {k for g in groups for k in g}
    members = [sorted(g) for g in groups] + [[k] for k in sorted(formulas - grouped)]
    node_of = {k: i for i, m in enumerate(members) for k in m}
    frows: dict[Column, list[int]] = defaultdict(list)  # rows holding a formula, per column, sorted
    for sheet, row, col in sorted(formulas):
        frows[(sheet, col)].append(row)
    fcols: dict[str, list[int]] = defaultdict(list)  # columns holding a formula, per sheet, sorted
    for sheet, col in sorted(frows):
        fcols[sheet].append(col)
    unordered = {column: Fenwick([1] * len(rows)) for column, rows in frows.items()}
    span: dict[tuple[int, Column], tuple[int, int]] = {}  # (group node, column) -> its members' index interval
    for n, m in enumerate(members[:len(groups)]):
        per: dict[Column, list[int]] = defaultdict(list)
        for sheet, row, col in m:
            per[(sheet, col)].append(bisect_left(frows[(sheet, col)], row))
        for column, idx in per.items():
            if max(idx) - min(idx) + 1 == len(idx):  # contiguous: a range over it can be clipped
                span[(n, column)] = (min(idx), max(idx) + 1)

    singles: dict[Key, set[Key]] = {}  # cell -> formula cells it reads one by one
    slices: dict[Key, list[tuple[Column, int, int]]] = {}  # cell -> (column, first, past-last) index into frows
    lex = lexed(book)
    for key in formulas:
        cell = book.cells[key]
        one: set[Key] = set()
        many: list[tuple[Column, int, int]] = []
        shape, texts = lex[key]
        try:
            for text in (texts if shape is not None and shape.parses() else ()):  # = refs(parse(cell.formula))
                area = resolve(text, cell.sheet, book)
                if area.is_single:
                    # a read of another member of the same group is not a wait; a cell reading itself still is
                    if area.key in formulas and (node_of[area.key] != node_of[key] or area.key == key):
                        one.add(area.key)
                    continue
                cols = fcols.get(area.sheet, [])
                for col in cols[bisect_left(cols, area.c1):bisect_right(cols, area.c2)]:
                    rows = frows[(area.sheet, col)]
                    i, j = bisect_left(rows, area.r1), bisect_right(rows, area.r2)
                    own = span.get((node_of[key], (area.sheet, col)))
                    # a range over the group's own members is not a wait (regions only admit such ranges when they
                    # read earlier rows of an earlier column: a growing SUM), only the cells outside the group are
                    parts = [(i, min(j, own[0])), (max(i, own[1]), j)] if own else [(i, j)]
                    for a, b in parts:
                        if b > a:
                            many.append(((area.sheet, col), a, b))
        except (Unsupported, RecursionError):
            pass  # the compiler reports the real reason for this cell
        singles[key], slices[key] = one, many

    users = defaultdict(list)  # cell -> nodes waiting for it, once per edge
    for key, reads in singles.items():
        for dep in reads:
            users[dep].append(node_of[key])
    blocked_on: dict[tuple[Column, int], list] = defaultdict(list)  # (column, index) -> slices waiting for that cell
    waiting = [0] * len(members)
    for key in formulas:
        node = node_of[key]
        waiting[node] += len(singles[key]) + len(slices[key])
        for column, i, j in slices[key]:
            blocked_on[(column, j - 1)].append((node, column, i, j))
    level = [0] * len(members)
    ready = [(0, members[n][0], n) for n in range(len(members)) if waiting[n] == 0]
    heapq.heapify(ready)
    order = []

    def release(user: int, after: int) -> None:
        level[user] = max(level[user], level[after] + 1)
        waiting[user] -= 1
        if waiting[user] == 0:
            heapq.heappush(ready, (level[user], members[user][0], user))

    while ready:
        _, _, node = heapq.heappop(ready)
        for key in members[node]:
            order.append(key)
            column = (key[0], key[2])
            pos = bisect_left(frows[column], key[1])
            unordered[column].add(pos, -1)
            for user in users[key]:
                release(user, node)
            for user, ucol, i, j in blocked_on.pop((column, pos), ()):
                fw = unordered[ucol]
                if fw.range(i, j) == 0:
                    release(user, node)
                else:  # wait for the last cell of the range that is still unordered
                    # ponytail: bottom-up chain (B_r = B_{r+1}+1 with C_r = SUM($B$1:B_r)) moves every waiting slice to
                    # the previous cell one step at a time, so ordering is O(n^2 log n) there (4000 rows: 28.8 s vs 7.9 s
                    # for the old graph). Upgrade: register each slice on canonical Fenwick/segment nodes with counters,
                    # or move whole waiter lists in one step.
                    blocked_on[(ucol, fw.kth(fw.prefix(j)))].append((user, ucol, i, j))
    return order, {k: CYCLE for k in sorted(formulas - set(order))}


def order_with_groups(book: Book, groups: list[list[Key]], internally_acyclic: bool = False):
    """Order with groups as nodes. A group that creates a cycle the cells do not have is dissolved and ordering
    repeated; the set of groups only shrinks, so this ends. Returns (order, unresolved, kept groups).

    A group can hide a true cycle only through reads among its own members. Callers whose groups cannot have such a
    cycle pass internally_acyclic=True, and the extra ungrouped pass then runs only when the grouped order leaves cells
    unresolved. Regions qualify because find_regions enforces two things: their same-row reads are acyclic
    (column_order), and no absolute cell, range or name may read the region's own formula cells. The only reads of
    other rows it admits are a running total's own previous row and a growing SUM($B$k:B{r}) of an earlier column,
    both of earlier rows only (formula_order clips such ranges to the cells outside the group). Relaxing any of this
    breaks the invariant."""
    groups = [list(g) for g in groups]
    base_bad = None
    while True:
        order, bad = formula_order(book, groups)
        if not groups or (internally_acyclic and not bad):
            return order, bad, groups
        if base_bad is None:
            base_bad = set(formula_order(book)[1])
        # cycles that grouping created, and true cycles that grouping hid (a node hides a cycle inside itself)
        extra = (set(bad) - base_bad) | (base_bad - set(bad))
        if not extra:
            return order, bad, groups
        kept = [g for g in groups if not (set(g) & extra)]
        if len(kept) == len(groups):
            return order, bad, groups
        groups = kept
