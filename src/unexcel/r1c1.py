"""Reference classification and R1C1 keys: which formulas are 'the same formula pasted down a column'."""
from __future__ import annotations

import re

from openpyxl.formula.tokenizer import Token, Tokenizer, TokenizerError
from openpyxl.utils.cell import column_index_from_string

from unexcel.model import Unsupported
from unexcel.parser import Binary, Call, Postfix, RefNode, Unary, parse

CELL = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$")
ROW_ONLY = re.compile(r"^\$?[0-9]+$")
COLUMN_ONLY = re.compile(r"^\$?[A-Za-z]{1,3}$")
NAME = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.]*$")
ROWREF = "__ROWREF_{}__"


def _split(text: str) -> tuple[str, str]:
    if "!" in text:
        sheet, ref = text.rsplit("!", 1)
        return sheet + "!", ref
    return "", text


def classify(text: str, row: int) -> tuple[str, int]:
    """How a reference behaves when its formula is copied down.
    ('same_row', col): one cell of the same sheet and the same row, row part not anchored.
    ('abs', 0): reads the same cells on every row (all row parts anchored, whole columns, names).
    ('bad', 0): anything else that moves with the row (previous/next row, relative ranges, other sheets, whole rows)."""
    if "[" in text or "]" in text:
        return "bad", 0  # structured (Table1[@Col]) and external ([Book]Sheet!A1) references: fail closed
    prefix, ref = _split(text)
    parts = ref.split(":")
    relative = []
    for part in parts:
        m = CELL.match(part)
        if m:
            relative.append((not m.group(3), column_index_from_string(m.group(2).upper()), int(m.group(4))))
        elif ROW_ONLY.match(part):
            return "bad", 0  # whole-row reference
        elif COLUMN_ONLY.match(part) or NAME.match(part):
            continue  # whole columns (A:A) and defined names read the same cells on every row
        else:
            return "bad", 0  # anything else is not a reference we can prove is row-independent
    if not any(is_relative for is_relative, _, _ in relative):
        return "abs", 0
    if len(parts) == 1 and not prefix and relative[0][2] == row:
        return "same_row", relative[0][1]
    return "bad", 0


RUNREF = "__RUN_{}_{}__"  # a growing SUM over column {} from row {}: a window column of the region


def prev_self(text: str, row: int, col: int) -> bool:
    """The formula's own cell one row up (relative row, same sheet): the previous value of a running total."""
    m = CELL.match(text)
    return (bool(m) and not m.group(3) and int(m.group(4)) == row - 1 >= 1
            and column_index_from_string(m.group(2).upper()) == col)


def growing(text: str, row: int, col: int):
    """(source column, start row) of $B$k:B{row} (start row anchored, end row relative on this row, one column other
    than the formula's own, same sheet), else None."""
    parts = text.split(":")
    if len(parts) != 2:
        return None
    a, b = CELL.match(parts[0]), CELL.match(parts[1])
    if not a or not b or not a.group(3) or b.group(3):
        return None
    src = column_index_from_string(a.group(2).upper())
    start, end = int(a.group(4)), int(b.group(4))
    if column_index_from_string(b.group(2).upper()) != src or src == col or end != row or start > row:
        return None
    return src, start


def _norm(part: str, row: int, col: int) -> str:
    m = CELL.match(part)
    if not m:
        return part
    c = column_index_from_string(m.group(2).upper())
    cs = f"C{c}" if m.group(1) else f"C[{c - col}]"
    rs = f"R{m.group(4)}" if m.group(3) else f"R[{int(m.group(4)) - row}]"
    return rs + cs


def r1c1_key(formula: str, row: int, col: int) -> str:
    """The formula with every reference written relative to its own cell: equal keys = same formula pasted down."""
    out = []
    for tok in Tokenizer("=" + formula).items:
        if tok.type == Token.OPERAND and tok.subtype == Token.RANGE:
            prefix, ref = _split(tok.value)
            out.append(prefix + ":".join(_norm(p, row, col) for p in ref.split(":")))
        else:
            out.append(tok.value)
    return "".join(out)


MAX_ROW = 1048576


class Shape:
    """A formula's tokens with the row numbers of its relative cell references taken out, so the same formula pasted
    to another row of the column can be predicted as text instead of tokenized again (the tokenizer is most of the
    compile time). Equal text means equal tokens: only the digits inside reference operands change, and the tokenizer
    never splits an operand on a digit."""

    __slots__ = ("key", "formula", "segments", "refs", "_parses")

    def __init__(self, formula: str, row: int, col: int):
        self.formula, self._parses = formula, None
        key, self.segments, refs = [], [], []  # segments: fixed text, or a reference as [text or row offset, ...]
        for tok in Tokenizer("=" + formula).items:
            if tok.type == Token.OPERAND and tok.subtype == Token.RANGE:
                prefix, ref = _split(tok.value)
                key.append(prefix + ":".join(_norm(p, row, col) for p in ref.split(":")))
                refs.append(tok.value)
                pieces = [prefix]
                for i, part in enumerate(ref.split(":")):
                    m = CELL.match(part)
                    pieces.append(":" if i else "")
                    if m and not m.group(3):
                        pieces += [m.group(1) + m.group(2), int(m.group(4)) - row]
                    else:
                        pieces.append(part)
                self.segments.append(pieces)
            else:
                key.append(tok.value)
                self.segments.append(tok.value)
        self.key, self.refs = "".join(key), tuple(refs)  # refs: this cell's reference texts, as written

    def at(self, row: int) -> tuple[str, tuple[str, ...]] | None:
        """(formula text, reference texts) of this formula pasted to `row`; None when a reference leaves the sheet."""
        text, refs = [], []
        for seg in self.segments:
            if isinstance(seg, str):
                text.append(seg)
                continue
            out = []
            for p in seg:
                if isinstance(p, int):
                    r = row + p
                    if not 1 <= r <= MAX_ROW:
                        return None
                    out.append(str(r))
                else:
                    out.append(p)
            ref = "".join(out)
            refs.append(ref)
            text.append(ref)
        return "".join(text), tuple(refs)

    def parses(self) -> bool:
        """Whether the parser accepts it; the same for every cell of the shape (the parser never reads reference text)."""
        if self._parses is None:
            try:
                parse(self.formula)
                self._parses = True
            except (Unsupported, RecursionError):
                self._parses = False
        return self._parses


def lexed(book) -> dict:
    """Formula cell -> (Shape or None when it does not tokenize, its reference texts left to right). Built once per
    book; a run of one formula pasted down a column is tokenized once."""
    got = book.__dict__.get("_lexed")
    if got is not None:
        return got
    got = {}
    for (sheet, col), rows in book._rows_by_col.items():
        shape, last = None, None
        for row in rows:
            cell = book.cells[(sheet, row, col)]
            if cell.formula is None:
                continue
            seen = shape.at(row) if shape is not None and last == row - 1 else None
            if seen is None or seen[0] != cell.formula:
                try:
                    shape = Shape(cell.formula, row, col)
                    seen = (cell.formula, shape.refs)
                except TokenizerError:
                    shape, seen = None, None
            got[(sheet, row, col)] = (shape, seen[1] if seen else ())
            last = row
    book.__dict__["_lexed"] = got
    return got


def map_refs(node, f):
    """The tree with every RefNode replaced by f(node)."""
    if isinstance(node, RefNode):
        return f(node)
    if isinstance(node, (Unary, Postfix)):
        return type(node)(node.op, map_refs(node.operand, f))
    if isinstance(node, Binary):
        return Binary(node.op, map_refs(node.left, f), map_refs(node.right, f))
    if isinstance(node, Call):
        return Call(node.name, tuple(None if a is None else map_refs(a, f) for a in node.args))
    return node
