import pytest

from cellproof.model import Unsupported
from cellproof.parser import Binary, Bool, Call, Num, Postfix, RefNode, Str, Unary, parse, refs


def test_precedence_matches_excel():
    assert parse("-2^2") == Binary("^", Unary("-", Num(2)), Num(2))  # Excel: 4
    assert parse("1+2*3") == Binary("+", Num(1), Binary("*", Num(2), Num(3)))
    assert parse("2^3^2") == Binary("^", Binary("^", Num(2), Num(3)), Num(2))  # left-assoc: 64
    assert parse('1+2&"a"="3a"') == Binary("=", Binary("&", Binary("+", Num(1), Num(2)), Str("a")), Str("3a"))


def test_percent_text_bool_and_refs():
    assert parse("5%") == Postfix("%", Num(5))
    assert parse('"say ""hi"""') == Str('say "hi"')
    assert parse("TRUE") == Bool(True)
    assert parse("'My Sheet'!$A$1") == RefNode("'My Sheet'!$A$1")


def test_functions_and_empty_args():
    assert parse("SUM(A1:A3, 2)") == Call("SUM", (RefNode("A1:A3"), Num(2)))
    assert parse("_xlfn.XLOOKUP(A1,B:B,C:C)").name == "XLOOKUP"
    assert parse("IF(A1,,1)") == Call("IF", (RefNode("A1"), None, Num(1)))
    assert parse("NOW()") == Call("NOW", ())


@pytest.mark.parametrize("bad", ["SUM(1,", "{1,A1}", "A1:B2 B1:C3", "(1+2"])
def test_unparseable_raises_unsupported(bad):
    with pytest.raises(Unsupported):
        parse(bad)


def test_refs_walks_the_tree_in_order():
    assert list(refs(parse("SUM(A1:A3)+B1*IF(C1,1,2)"))) == ["A1:A3", "B1", "C1"]
