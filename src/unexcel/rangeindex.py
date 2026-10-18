"""What the compiler needs to know about a rectangle of cells, in O(columns x log rows) instead of O(cells).

Per (sheet, column) one Fenwick tree holds a packed counter for every present cell: how many cells of each kind
(n s b e a blank), how many can hold an error, and how many formula cells are finished. Input cells are counted
when the index is built; a formula cell is counted by complete() when its batch closes.
"""
from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections import defaultdict
from dataclasses import dataclass

from unexcel.fenwick import Fenwick
from unexcel.model import Book, Key

FIELDS = ("n", "s", "b", "e", "a", "blank", "merr", "fdone")
BITS = 40  # summary() adds counters across columns: 1,048,576 rows x 16,384 columns = 2**34 < 2**40, no carry
SHIFT = {name: BITS * i for i, name in enumerate(FIELDS)}
MASK = (1 << BITS) - 1
Column = tuple[str, int]


def field(total: int, name: str) -> int:
    return (total >> SHIFT[name]) & MASK


@dataclass(frozen=True)
class Summary:
    present: int  # non-empty cells in the rectangle
    formulas: int  # of those, formula cells
    total: int  # packed counters, read with field()


class RangeIndex:
    def __init__(self, book: Book):
        self.book = book
        self.finished: set[Key] = set()
        self.rows = book._rows_by_col  # (sheet, col) -> sorted rows that hold a cell
        self.cols: dict[str, list[int]] = defaultdict(list)  # sheet -> sorted columns that hold a cell
        for sheet, col in sorted(self.rows):
            self.cols[sheet].append(col)
        self.fen: dict[Column, Fenwick] = {}
        self.formula_prefix: dict[Column, list[int]] = {}  # formula cells among the first i cells of the column
        self.text_rows: dict[Column, list[int]] = {}  # rows of input text cells
        for (sheet, col), rows in self.rows.items():
            values, prefix, text = [], [0], []
            for row in rows:
                cell = book.cells[(sheet, row, col)]
                if cell.formula is not None:
                    values.append(0)
                    prefix.append(prefix[-1] + 1)
                    continue
                values.append((1 << SHIFT[cell.kind]) + ((1 << SHIFT["merr"]) if cell.kind == "e" else 0))
                prefix.append(prefix[-1])
                if cell.kind == "s":
                    text.append(row)
            self.fen[(sheet, col)] = Fenwick(values)
            self.formula_prefix[(sheet, col)] = prefix
            self.text_rows[(sheet, col)] = text

    def complete(self, key: Key, kind: str, may_err: bool) -> None:
        """A formula cell is finished: kind is its value kind (n s b e a), may_err whether it can hold an error."""
        sheet, row, col = key
        rows = self.rows.get((sheet, col), [])
        pos = bisect_left(rows, row)
        cell = self.book.cells.get(key)
        if pos >= len(rows) or rows[pos] != row or cell is None or cell.formula is None:
            raise ValueError(f"{key} is not a formula cell in the index")
        if key in self.finished:
            raise ValueError(f"{key} was already completed")
        self.finished.add(key)
        add = (1 << SHIFT[kind]) + (1 << SHIFT["fdone"]) + ((1 << SHIFT["merr"]) if may_err else 0)
        self.fen[(sheet, col)].add(pos, add)

    def _slices(self, sheet: str, r1: int, c1: int, r2: int, c2: int):
        cols = self.cols.get(sheet, [])
        for col in cols[bisect_left(cols, c1):bisect_right(cols, c2)]:
            rows = self.rows[(sheet, col)]
            yield (sheet, col), bisect_left(rows, r1), bisect_right(rows, r2)

    def summary(self, sheet: str, r1: int, c1: int, r2: int, c2: int) -> Summary:
        present = formulas = total = 0
        for column, i, j in self._slices(sheet, r1, c1, r2, c2):
            if j > i:
                present += j - i
                prefix = self.formula_prefix[column]
                formulas += prefix[j] - prefix[i]
                total += self.fen[column].range(i, j)
        return Summary(present, formulas, total)

    def text_cells(self, sheet: str, r1: int, c1: int, r2: int, c2: int) -> list[Key]:
        """Input text cells in the rectangle, row-major."""
        found = []
        for (s, col), _, _ in self._slices(sheet, r1, c1, r2, c2):
            rows = self.text_rows[(s, col)]
            found.extend((s, r, col) for r in rows[bisect_left(rows, r1):bisect_right(rows, r2)])
        return sorted(found)
