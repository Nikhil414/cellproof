from unexcel.parser import RefNode, parse, refs
from unexcel.r1c1 import classify, map_refs, r1c1_key


def test_same_row_references():
    assert classify("A2", 2) == ("same_row", 1)
    assert classify("$A2", 2) == ("same_row", 1)  # anchored column, relative row: still row-wise
    assert classify("C2", 2) == ("same_row", 3)


def test_references_that_read_the_same_cells_on_every_row():  # Review Focus 2
    assert classify("A$2", 2) == ("abs", 0)
    assert classify("$A$2", 2) == ("abs", 0)
    assert classify("$A$2", 5) == ("abs", 0)  # anchored, even when it points into the region's own rows
    assert classify("$A$1:$A$9", 3) == ("abs", 0)
    assert classify("A:A", 2) == ("abs", 0)
    assert classify("Rate", 2) == ("abs", 0)  # a name
    assert classify("Other!$A$1", 2) == ("abs", 0)  # Review Focus 4


def test_references_that_move_with_the_row_are_bad():  # Review Focus 4
    assert classify("A1", 2)[0] == "bad"  # previous row
    assert classify("A3", 2)[0] == "bad"  # next row
    assert classify("Other!A2", 2)[0] == "bad"  # relative on another sheet
    assert classify("A2:B2", 2)[0] == "bad"  # relative range
    assert classify("$A$2:A9", 2)[0] == "bad"  # growing range
    assert classify("2:2", 2)[0] == "bad"  # whole row


def test_adversarial_references_fail_closed():
    assert classify("Table1[@Col]", 2) == ("bad", 0)  # structured: this row, moves with the row
    assert classify("Table1[Col]", 2) == ("bad", 0)
    assert classify("[Book]Sheet!$A$1", 2) == ("bad", 0)  # external workbook
    assert classify("'My Sheet'!A2", 2) == ("bad", 0)  # relative on another sheet
    assert classify("'It''s'!$A$1", 2) == ("abs", 0)  # quoted sheet name with an escaped quote
    assert classify("a2", 2) == ("same_row", 1)  # lower case
    assert classify("A$2:B$9", 2) == ("abs", 0)
    assert classify("$A2:$B9", 2) == ("bad", 0)  # relative rows in a range
    assert classify("XFD1048576", 1048576) == ("same_row", 16384)  # last cell of the sheet
    assert classify("Tax_Rate", 2) == ("abs", 0)  # a defined name
    assert classify("A#B", 2) == ("bad", 0)  # not a reference or a valid name


def test_r1c1_key_is_equal_for_a_formula_pasted_down():
    assert r1c1_key("A2*1.1", 2, 2) == r1c1_key("A3*1.1", 3, 2)
    assert r1c1_key("A2*$J$1", 2, 2) == r1c1_key("A9*$J$1", 9, 2)
    assert r1c1_key('IF(A2>1,"x","y")', 2, 3) == r1c1_key('IF(A7>1,"x","y")', 7, 3)


def test_r1c1_key_differs_when_the_formula_differs():
    assert r1c1_key("A3*1.1", 3, 2) != r1c1_key("A3*1.1", 4, 2)  # a different row of the same text
    assert r1c1_key("A2*$J$1", 2, 2) != r1c1_key("A2*$J$2", 2, 2)
    assert r1c1_key("A2*1.1", 2, 2) != r1c1_key("A2*1.2", 2, 2)
    assert r1c1_key("A2*1.1", 2, 2) != r1c1_key("A2*1.1", 2, 3)  # relative column offset differs


def test_map_refs_replaces_every_reference():
    ast = parse("A2*B2+SUM($C$1:$C$9)")
    out = map_refs(ast, lambda n: RefNode("X") if n.text in ("A2", "B2") else n)
    assert list(refs(out)) == ["X", "X", "$C$1:$C$9"]
