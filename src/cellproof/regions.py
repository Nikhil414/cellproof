"""Find row-wise regions: blocks of columns that repeat one formula down the same rows.

A region column is a run of at least MIN_ROWS consecutive formula cells with the same R1C1 key. Columns whose runs
cover the same rows form one region. A column joins only if every reference in its formula is a same-row cell, or
reads the same cells on every row; anything that moves with the row (a previous row, a growing range) stays in cell
mode. The region compiler decides later whether the SQL can really be built.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.utils.cell import get_column_letter

from cellproof.model import Book, Key, Unsupported
from cellproof.parser import Binary, Call, RefNode, Unary, Postfix, parse, refs
from cellproof.r1c1 import ROWREF, RUNREF, classify, growing, lexed, map_refs, prev_self
from cellproof.refs import resolve

MIN_ROWS = 3


@dataclass
class Region:
    sheet: str
    r1: int
    r2: int
    cols: list[int]
    template: dict[int, object] = field(default_factory=dict)  # col -> AST, same-row refs as __ROWREF_n__
    reads: dict[int, set] = field(default_factory=dict)  # col -> source columns read in the same row
    # col -> (side, op) of a running total X{r-1} op E ("left") or E + X{r-1} ("right"); its template is E
    recur: dict[int, tuple] = field(default_factory=dict)
    runs: dict[int, list] = field(default_factory=dict)  # col -> [(source col, start row)] of its growing SUMs

    def keys(self) -> list[Key]:
        return [(self.sheet, r, c) for c in self.cols for r in range(self.r1, self.r2 + 1)]

    @property
    def first_key(self) -> Key:
        return (self.sheet, self.r1, min(self.cols))


@dataclass
class RegionScan:
    regions: list[Region]
    rejected: list[str]  # why pasted-down columns did not become regions: the map for the next slice


def column_order(cols, reads) -> tuple[list[int], list[int]]:
    """Columns in dependency order (smallest ready column first, so the SQL follows the sheet); and the stuck ones."""
    left, order = set(cols), []
    while left:
        ready = sorted(c for c in left if not (reads[c] & left))
        if not ready:
            break
        order.append(ready[0])
        left.discard(ready[0])
    return order, sorted(left)


def _runs(book: Book):
    """(sheet, col, r1, r2) for every maximal run of one pasted-down formula, at least MIN_ROWS rows."""
    lex = lexed(book)
    for (sheet, col), rows in sorted(book._rows_by_col.items()):
        run, key0 = [], None
        for row in rows + [None]:
            cell = book.cells[(sheet, row, col)] if row is not None else None
            ok = (cell is not None and cell.formula is not None and not cell.array and cell.kind != "blank"
                  and cell.key not in book.unread)  # not read, or DuckDB rejected it once: cell mode says why
            shape = lex[(sheet, row, col)][0] if ok else None
            ok, key = shape is not None, shape.key if shape is not None else None
            if ok and run and row == run[-1] + 1 and key == key0:
                run.append(row)
                continue
            if len(run) >= MIN_ROWS:
                yield sheet, col, run[0], run[-1]
            run, key0 = ([row], key) if ok else ([], None)


def _place(sheet: str, col: int, r1: int, r2: int) -> str:
    letter = get_column_letter(col)
    return f"{sheet} column {letter} rows {r1}-{r2}"


def find_regions(book: Book) -> RegionScan:
    groups: dict[tuple, list[int]] = {}
    for sheet, col, r1, r2 in _runs(book):
        groups.setdefault((sheet, r1, r2), []).append(col)
    regions, rejected = [], []
    for (sheet, r1, r2), cols in groups.items():
        template, reads, keep, recur, runs = {}, {}, [], {}, {}
        for col in cols:
            cell = book.cells[(sheet, r1, col)]
            try:
                parsed = parse(cell.formula)
                if any(t.upper().startswith(("__ROWREF_", "__RUN_")) for t in refs(parsed)):
                    raise Unsupported("uses a name spelled like cellproof's own placeholders")
                ast, rec = _running(parsed, r1, col)
                sums: list = []
                ast = _map_sums(ast, r1, col, sums)
                found, why = {s for s, _ in sums}, None
                for text in refs(ast):
                    if text.startswith("__RUN_"):
                        continue
                    kind, source = classify(text, r1)
                    if kind == "bad":
                        why = f"reads {text}, which moves with the row (previous row, growing or relative range, other sheet)"
                    elif kind == "same_row":
                        found.add(source)
                    else:  # the same cells on every row: must not be the region's own formula cells
                        area = resolve(text, sheet, book)
                        if area.sheet == sheet and area.r1 <= r2 and area.r2 >= r1:
                            lo, hi = max(area.r1, r1), min(area.r2, r2)
                            hit = [c for c in cols if area.c1 <= c <= area.c2]  # columns first, then rows
                            if any(book.cells.get((sheet, r, c)) is not None
                                   and book.cells[(sheet, r, c)].formula is not None
                                   for c in hit for r in range(lo, hi + 1)):
                                why = f"reads {text}, which includes formula cells of its own rows"
                if not why:
                    mapped = map_refs(ast, lambda n: RefNode(ROWREF.format(classify(n.text, r1)[1]))
                                      if classify(n.text, r1)[0] == "same_row" else n)
            except (Unsupported, RecursionError) as err:
                why = "formula is nested too deeply" if isinstance(err, RecursionError) else str(err)
            if why:
                rejected.append(f"{_place(sheet, col, r1, r2)} stays in cell mode: {why}")
                continue
            template[col] = mapped
            reads[col] = found
            if rec:
                recur[col] = rec
            if sums:
                runs[col] = sums
            keep.append(col)
        _, stuck = column_order(keep, reads)
        for col in stuck:
            rejected.append(f"{_place(sheet, col, r1, r2)} stays in cell mode: its row-wise reads form a cycle")
        keep = [c for c in keep if c not in stuck]
        if keep:
            regions.append(Region(sheet, r1, r2, sorted(keep), {c: template[c] for c in keep},
                                  {c: reads[c] for c in keep}, {c: recur[c] for c in keep if c in recur},
                                  {c: runs[c] for c in keep if c in runs}))
    return RegionScan(regions, rejected)


def _running(ast, row: int, col: int):
    """(E, (side, op)) when the formula is X{r-1} op t1 op t2 ... (+/- only, X{r-1} first) or E + X{r-1}; else
    (ast, None). X - a + b is X + (-a + b): same value, and Excel reports the errors in the same order (X, a, b)."""
    if not (isinstance(ast, Binary) and ast.op in ("+", "-")):
        return ast, None
    if ast.op == "+" and isinstance(ast.right, RefNode) and prev_self(ast.right.text, row, col):
        return ast.left, ("right", "+")
    terms, node = [], ast
    while isinstance(node, Binary) and node.op in ("+", "-"):
        terms.append((node.op, node.right))
        node = node.left
    if not (isinstance(node, RefNode) and prev_self(node.text, row, col)):
        return ast, None
    terms.reverse()
    if len(terms) == 1:
        return terms[0][1], ("left", terms[0][0])
    e = terms[0][1] if terms[0][0] == "+" else Unary("-", terms[0][1])
    for op, t in terms[1:]:
        e = Binary(op, e, t)
    return e, ("left", "+")


def _map_sums(node, row: int, col: int, found: list):
    """The tree with every SUM(<growing range>) replaced by a RUNREF placeholder; found gets (source, start)."""
    if isinstance(node, Call):
        if node.name == "SUM" and len(node.args) == 1 and isinstance(node.args[0], RefNode):
            g = growing(node.args[0].text, row, col)
            if g:
                found.append(g)
                return RefNode(RUNREF.format(*g))
        return Call(node.name, tuple(None if a is None else _map_sums(a, row, col, found) for a in node.args))
    if isinstance(node, (Unary, Postfix)):
        return type(node)(node.op, _map_sums(node.operand, row, col, found))
    if isinstance(node, Binary):
        return Binary(node.op, _map_sums(node.left, row, col, found), _map_sums(node.right, row, col, found))
    return node


def _reads_outside_formula(book: Book, region: Region, col: int) -> bool:
    """Whether a column reads a formula cell outside the region (same-row, or through an absolute cell/range/name)."""
    for s in region.reads[col] - set(region.cols):
        for r in range(region.r1, region.r2 + 1):
            cell = book.cells.get((region.sheet, r, s))
            if cell is not None and cell.formula is not None:
                return True
    seed = book.cells.get((region.sheet, region.r1 - 1, col)) if col in region.recur else None
    if seed is not None and seed.formula is not None:
        return True
    for src, start in region.runs.get(col, []):
        for r in range(start, region.r1):
            cell = book.cells.get((region.sheet, r, src))
            if cell is not None and cell.formula is not None:
                return True
    for text in refs(region.template[col]):
        if text.startswith(("__ROWREF_", "__RUN_")):
            continue
        try:
            area = resolve(text, region.sheet, book)
        except Unsupported:
            return True
        if any(book.cells[k].formula is not None for k in area.present(book)):
            return True
    return False


def _pure_columns(book: Book, region: Region) -> list[int]:
    """Columns that read only plain inputs or other pure columns. They wait for no formula cell outside the region,
    so no cycle can pass through them: they can always stay one block."""
    pure: set[int] = set()
    for c in column_order(region.cols, region.reads)[0]:
        if (region.reads[c] & set(region.cols)) <= pure and not _reads_outside_formula(book, region, c):
            pure.add(c)
    return sorted(pure)


def _part(region: Region, cols: list[int]) -> Region:
    return Region(region.sheet, region.r1, region.r2, cols, {c: region.template[c] for c in cols},
                  {c: region.reads[c] for c in cols}, {c: region.recur[c] for c in cols if c in region.recur},
                  {c: region.runs[c] for c in cols if c in region.runs})


def order_with_regions(book: Book, regions: list[Region]):
    """Dependency order with each region as one node: (order, unresolved, regions that survived).

    A region that cannot be ordered as one block (a cell between its columns: a running total, a total that a share
    column divides by) is split once into its pure columns and the rest; the rest stays a block if it can, else it
    goes back to cell mode. The pure part never cycles, so splitting only ever adds coverage."""
    from cellproof.graph import order_with_groups

    pending, split = list(regions), set()  # split: ids of pieces made by a split, never split again
    while True:
        order, bad, kept = order_with_groups(book, [r.keys() for r in pending], internally_acyclic=True)
        kept_sets = [set(g) for g in kept]
        survivors = [r for r in pending if set(r.keys()) in kept_sets]
        pieces = []
        for r in pending:
            if set(r.keys()) in kept_sets or id(r) in split:
                continue
            pure = _pure_columns(book, r)
            if pure and len(pure) < len(r.cols):
                parts = [_part(r, pure), _part(r, [c for c in r.cols if c not in pure])]
                split.update(id(p) for p in parts)
                pieces += parts
        if not pieces:
            return order, bad, survivors
        pending = survivors + pieces
