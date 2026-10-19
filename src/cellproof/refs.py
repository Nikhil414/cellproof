"""Reference text ('Sales'!C:C, $A$1, Rate) -> a concrete rectangle of cells."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator

from openpyxl.utils.cell import range_boundaries

from cellproof.model import Book, Key, Unsupported

MAX_ROW = 1_048_576


@dataclass(frozen=True)
class Area:
    sheet: str
    r1: int
    c1: int
    r2: int
    c2: int

    @property
    def key(self) -> Key:
        return (self.sheet, self.r1, self.c1)

    @property
    def is_single(self) -> bool:
        return self.r1 == self.r2 and self.c1 == self.c2

    def cells(self) -> Iterator[Key]:
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                yield (self.sheet, r, c)

    def present(self, book: Book) -> list[Key]:
        """Non-empty cells in the area, row-major. Use this, not cells(), to read a range."""
        return book.present(self.sheet, self.r1, self.c1, self.r2, self.c2)


def resolve(text: str, sheet: str, book: Book) -> Area:
    ref = text
    if "!" in text:
        sheet, ref = text.rsplit("!", 1)
        if sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        if "[" in sheet:
            raise Unsupported(f"external link {text}")
    elif ref.upper() in book.names:
        sheet, ref = book.names[ref.upper()]
    if sheet not in book.sheets:
        raise Unsupported(f"unknown sheet {sheet!r}")
    try:
        if ":" not in ref and not any(ch.isdigit() for ch in ref):
            raise ValueError  # "I" alone is a name; openpyxl would read it as column I, which Excel writes as I:I
        c1, r1, c2, r2 = range_boundaries(ref.replace("$", ""))
    except ValueError:
        raise Unsupported(f"unknown name {text!r}") from None
    if c1 is None:
        raise Unsupported(f"whole-row reference {text}")
    return Area(sheet, r1 or 1, c1, r2 or MAX_ROW, c2)  # C:C is every row, as in Excel; Area.present skips empties
