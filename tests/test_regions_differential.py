import random

from xlsxkit import book_of

from cellproof.compiler import compile_book
from cellproof.verifier import execute

TEMPLATES = [
    "={a}{r}*1.1", "=IF({x}{r}>10,{x}{r},0)", "={x}{r}-{y}{r}", "={a}{r}/{b}{r}", '={c}{r}&"!"',
    "={x}{r}>{y}{r}", '=IF({a}{r}>5,{a}{r},"n/a")', "={x}{r}*$H$1", "=SUM($A$2:$A${last})*{x}{r}",
    "=ROUND({x}{r},1)", "=IFERROR({a}{r}/{b}{r},0)", "=SUM($A$2:A{r})", "={x}{p}+1",
    '=IF({x}{r}>3,"big","small")', "={x}{r}&{c}{r}", "=ABS({x}{r})", "=MAX({a}{r},{b}{r})",
]


def make_book(rng):
    n = rng.randint(3, 14)
    last = n + 1
    spec = {"A1": "a", "B1": "b", "C1": "c", "H1": 2.0}
    for r in range(2, last + 1):
        if not (r == 3 and rng.random() < 0.15):  # sometimes a blank in A
            spec[f"A{r}"] = float(rng.randint(0, 9))
        spec[f"B{r}"] = float(rng.randint(0, 9))
        spec[f"C{r}"] = rng.choice(["x", "yy", "3", ""])
    if rng.random() < 0.15:
        spec["B3"] = "#N/A"
    letters = ["A", "B", "C"]  # columns a template may read: inputs, then the formula columns built so far
    for col in "DEFGIJK"[:rng.randint(2, 6)]:
        template = rng.choice(TEMPLATES)
        x, y = rng.choice(letters), rng.choice(letters)  # fixed for the whole column: it is one formula pasted down
        for r in range(2, last + 1):
            spec[f"{col}{r}"] = (template.format(a="A", b="B", c="C", x=x, y=y, r=r, p=r - 1, last=last), 0.0)
        letters.append(col)
    return book_of({"S": spec})


def same(a, b):
    return (a[0] == b[0] and a[2] == b[2]
            and (a[1] == b[1] or (a[1] is not None and b[1] is not None and abs(a[1] - b[1]) <= 1e-9 * max(1, abs(a[1])))))


def test_regions_never_change_a_status_or_a_value():
    rng = random.Random(2026)
    used = 0
    for seed in range(400):
        book = make_book(rng)
        with_regions = compile_book(book, regions=True)
        without = compile_book(book, regions=False)
        assert with_regions.unresolved == without.unresolved, seed
        a, b = execute(book, with_regions), execute(book, without)
        assert a.keys() == b.keys(), seed
        for key in a:
            assert same(a[key], b[key]), (seed, key, a[key], b[key])
        used += any(hasattr(batch, "wide_sql") for batch in with_regions.batches)
    assert used > 100  # regions were really exercised, not silently skipped
