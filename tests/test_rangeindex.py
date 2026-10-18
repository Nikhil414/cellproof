import random

import pytest

from xlsxkit import book_of

from unexcel.rangeindex import FIELDS, SHIFT, RangeIndex, field


def test_summary_matches_a_brute_force_count():
    rng = random.Random(3)
    for it in range(1500):
        n = rng.randint(2, 20)
        spec = {}
        for r in range(1, n + 1):
            for c in "ABCD":
                x = rng.random()
                if x < 0.2:
                    spec[f"{c}{r}"] = float(r)
                elif x < 0.35:
                    spec[f"{c}{r}"] = "txt"
                elif x < 0.4:
                    spec[f"{c}{r}"] = "#N/A"
                elif x < 0.45:
                    spec[f"{c}{r}"] = True
                elif x < 0.5:
                    spec[f"{c}{r}"] = None
                elif x < 0.85:
                    spec[f"{c}{r}"] = ("=1", 1.0)
        book = book_of({"S": spec, "T": {"A1": 1.0}})
        idx = RangeIndex(book)
        done = {}
        for key, cell in book.cells.items():
            if cell.formula is not None and rng.random() < 0.6:
                kind, err = rng.choice([("n", False), ("s", False), ("b", False), ("e", True), ("a", True), ("n", True)])
                idx.complete(key, kind, err)
                done[key] = (kind, err)
        for _ in range(25):
            r1, c1 = rng.randint(1, n), rng.randint(1, 4)
            r2, c2 = rng.randint(r1, n), rng.randint(c1, 4)
            s = idx.summary("S", r1, c1, r2, c2)
            keys = book.present("S", r1, c1, r2, c2)
            want = dict.fromkeys(FIELDS, 0)
            formulas = 0
            for k in keys:
                cell = book.cells[k]
                if cell.formula is None:
                    want[cell.kind] += 1
                    want["merr"] += cell.kind == "e"
                else:
                    formulas += 1
                    if k in done:
                        kind, err = done[k]
                        want[kind] += 1
                        want["merr"] += err
                        want["fdone"] += 1
            assert (s.present, s.formulas) == (len(keys), formulas), it
            assert {f: field(s.total, f) for f in FIELDS} == want, it
            texts = [k for k in keys if book.cells[k].formula is None and book.cells[k].kind == "s"]
            assert idx.text_cells("S", r1, c1, r2, c2) == texts


def test_a_sheet_with_no_cells_in_the_range_is_empty():
    idx = RangeIndex(book_of({"S": {"A1": 1.0}, "T": {}}))
    assert idx.summary("T", 1, 1, 10, 10).present == 0
    assert idx.summary("S", 5, 5, 9, 9).present == 0


def test_fields_do_not_carry_into_each_other():
    total = (1 << SHIFT["n"]) * (2**22) + (1 << SHIFT["s"])
    assert field(total, "n") == 2**22
    assert field(total, "s") == 1


def test_complete_rejects_bad_keys_and_double_completion():
    idx = RangeIndex(book_of({"S": {"A1": 1.0, "A2": ("=1", 1.0)}}))
    for bad in [("S", 1, 1), ("S", 5, 1), ("S", 2, 9), ("X", 1, 1)]:
        with pytest.raises(ValueError):
            idx.complete(bad, "n", False)
    idx.complete(("S", 2, 1), "n", False)
    with pytest.raises(ValueError):
        idx.complete(("S", 2, 1), "n", False)
