"""lexed() predicts a pasted-down formula's text instead of tokenizing it; it must equal tokenizing every cell."""
from xlsxkit import book_of

from cellproof.model import Unsupported
from cellproof.parser import parse, refs
from cellproof.r1c1 import lexed, r1c1_key

TEMPLATES = [
    "A{r}*2", "$A{r}+A${r}+$A$1", "SUM($B$1:B{r})", "VLOOKUP(A{r},Tbl!$A$1:$B$9,2,FALSE)",
    "'My Sheet'!A{r}+'a!b'!C{r}", "IF(A{r}>1E+5,\"A{r}\",B{r} & \" x \")", "a{r}+B{r}", "SUM(A:A)+A{r}",
    "{r}:{r}", "Rates*A{r}", "A{r} + 1.5E-3", "COUNTIF($E$1:$E$4,\">\"&A{r})", "A0{r}", "INDEX(A:A,{r})",
    "Table1[@Col]+A{r}", "-A{r}%", "SUM(A{r}:C{r})/3", "A{r}:A1048576", "((A{r}))",
    'INDIRECT("A1")+A{r}', 'INDIRECT( "My Sheet!A1" , TRUE)*A{r}', 'INDIRECT("A"&{r})', 'INDIRECT("B2",FALSE)+A{r}',
]


def _book(template, rows):
    return book_of({"S": {f"D{r}": (f"={template.format(r=r)}", 0.0) for r in rows},
                    "My Sheet": {"A1": 1.0}, "a!b": {"A1": 1.0}})


def _direct(formula, row, col):
    try:
        texts = tuple(refs(parse(formula)))
    except (Unsupported, RecursionError):
        texts = ()
    return r1c1_key(formula, row, col), texts


def test_lexed_equals_tokenizing_every_cell():
    rows = list(range(1, 12)) + list(range(98, 102)) + [999, 1000, 1001, 1048575, 1048576]
    for template in TEMPLATES:
        book = _book(template, rows)
        lex = lexed(book)
        for key, cell in book.cells.items():
            if cell.formula is None:
                continue
            shape, texts = lex[key]
            assert shape is not None
            want_key, want_refs = _direct(cell.formula, key[1], key[2])
            assert shape.key == want_key, (template, key)
            assert (texts if shape.parses() else ()) == want_refs, (template, key, texts, want_refs)


def test_a_run_is_tokenized_once(monkeypatch):
    import cellproof.r1c1 as r1c1
    made = []
    real = r1c1.Shape.__init__
    monkeypatch.setattr(r1c1.Shape, "__init__", lambda self, *a: (made.append(a), real(self, *a))[1])
    lexed(_book("VLOOKUP(A{r},Tbl!$A$1:$B$9,2,FALSE)+$C{r}", range(1, 500)))
    assert len(made) == 1
