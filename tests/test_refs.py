import pytest

from unexcel.model import Book, Unsupported
from unexcel.refs import Area, resolve

BOOK = Book(cells={}, sheets=["S", "My Sheet", "Bob's"], names={"RATE": ("My Sheet", "B2")},
            max_row={"S": 10, "My Sheet": 5, "Bob's": 1})


def test_resolve_forms():
    assert resolve("A1", "S", BOOK) == Area("S", 1, 1, 1, 1)
    assert resolve("$B$2:C3", "S", BOOK) == Area("S", 2, 2, 3, 3)
    assert resolve("'My Sheet'!A1", "S", BOOK) == Area("My Sheet", 1, 1, 1, 1)
    assert resolve("'Bob''s'!A1", "S", BOOK) == Area("Bob's", 1, 1, 1, 1)
    assert resolve("C:C", "S", BOOK) == Area("S", 1, 3, 1_048_576, 3)  # every row, as in Excel
    assert resolve("rate", "S", BOOK) == Area("My Sheet", 2, 2, 2, 2)


@pytest.mark.parametrize("text, msg", [
    ("[Book2.xlsx]S!A1", "external link"),
    ("Nope!A1", "unknown sheet"),
    ("1:1", "whole-row"),
    ("Missing", "unknown name"),
])
def test_resolve_rejects(text, msg):
    with pytest.raises(Unsupported, match=msg):
        resolve(text, "S", BOOK)


def test_area_helpers():
    area = Area("S", 1, 1, 2, 2)
    assert list(area.cells()) == [("S", 1, 1), ("S", 1, 2), ("S", 2, 1), ("S", 2, 2)]
    assert area.key == ("S", 1, 1) and not area.is_single
    assert Area("S", 3, 3, 3, 3).is_single
