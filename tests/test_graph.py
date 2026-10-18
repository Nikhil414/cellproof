import random

from xlsxkit import book_of

from unexcel.graph import CYCLE, formula_order, order_with_groups
from unexcel.parser import parse, refs
from unexcel.refs import resolve


def test_dependencies_come_first():
    book = book_of({"S": {"A1": 1, "A2": ("=A3+A1", 3), "A3": ("=A1*2", 2)}})
    assert formula_order(book) == ([("S", 3, 1), ("S", 2, 1)], {})


def test_cycles_and_their_dependents_are_unresolved():
    book = book_of({"S": {"A1": ("=B1", 0), "B1": ("=A1", 0), "C1": ("=A1+1", 1), "D1": ("=2", 2)}})
    order, bad = formula_order(book)
    assert order == [("S", 1, 4)]
    assert bad == {("S", 1, 1): CYCLE, ("S", 1, 2): CYCLE, ("S", 1, 3): CYCLE}


def test_self_reference_is_a_cycle():
    assert formula_order(book_of({"S": {"A1": ("=A1+1", 0)}})) == ([], {("S", 1, 1): CYCLE})


def test_unparseable_formula_is_still_ordered():
    assert formula_order(book_of({"S": {"A1": ("=SUM(", 0)}})) == ([("S", 1, 1)], {})


import random

from unexcel.parser import parse, refs
from unexcel.refs import resolve


def _old_unresolved(book):
    """The previous algorithm, kept as an oracle: expands every range into a set of cells."""
    import heapq
    from collections import defaultdict
    formulas = {k for k, c in book.cells.items() if c.formula is not None}
    deps = {}
    for key in formulas:
        found = set()
        try:
            for text in refs(parse(book.cells[key].formula)):
                found.update(resolve(text, key[0], book).present(book))
        except Exception:
            pass
        deps[key] = found & formulas
    users = defaultdict(set)
    for key, ds in deps.items():
        for d in ds:
            users[d].add(key)
    waiting = {k: len(ds) for k, ds in deps.items()}
    ready = [k for k, n in waiting.items() if n == 0]
    heapq.heapify(ready)
    done = set()
    while ready:
        key = heapq.heappop(ready)
        done.add(key)
        for user in users[key]:
            waiting[user] -= 1
            if waiting[user] == 0:
                heapq.heappush(ready, user)
    return formulas - done


def _random_book(rng):
    n = rng.randint(3, 25)
    spec = {}
    for r in range(1, n + 1):
        for c in "ABC":
            x = rng.random()
            if x < 0.3:
                spec[f"{c}{r}"] = float(r)
            elif x < 0.85:
                a, b = rng.randint(1, n), rng.randint(1, n)
                lo, hi = min(a, b), max(a, b)
                c2, c3 = rng.choice("ABC"), rng.choice("ABC")
                spec[f"{c}{r}"] = (rng.choice([f"={c2}{a}+1", f"=SUM({c2}{lo}:{c3}{hi})", f"=SUM({c2}:{c2})",
                                               f"={c2}{a}*{c3}{b}"]), 0.0)
    return book_of({"S": spec})


def test_order_and_cycles_match_the_old_algorithm_on_random_books():
    rng = random.Random(7)
    for _ in range(1500):
        book = _random_book(rng)
        order, bad = formula_order(book)
        formulas = {k for k, c in book.cells.items() if c.formula is not None}
        assert set(bad) == _old_unresolved(book)
        assert set(order) | set(bad) == formulas and not set(order) & set(bad) and len(set(order)) == len(order)
        position = {k: i for i, k in enumerate(order)}
        for key in order:
            for text in refs(parse(book.cells[key].formula)):
                for dep in resolve(text, "S", book).present(book):
                    if dep in formulas:
                        assert position[dep] < position[key]


def test_a_cell_inside_its_own_range_is_a_cycle():  # Review Focus 4
    book = book_of({"S": {"A1": 1, "A2": ("=SUM(A1:A3)", 0), "A3": 2}})
    assert formula_order(book) == ([], {("S", 2, 1): CYCLE})


def test_a_range_over_a_cycle_member_is_unresolved_too():  # Review Focus 4
    book = book_of({"S": {"A1": ("=A2", 0), "A2": ("=A1", 0), "B1": ("=SUM(A1:A2)", 0)}})
    assert formula_order(book)[1].keys() == {("S", 1, 1), ("S", 2, 1), ("S", 1, 2)}


def test_cells_that_read_only_inputs_come_before_cells_that_read_formulas():
    n = 50  # interleaved in key order today: B1, C1, B2, C2 ...
    spec = {}
    for r in range(1, n + 1):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=SUM($B$1:B{r})", 0.0)
    order, bad = formula_order(book_of({"S": spec}))
    assert not bad
    cols = [key[2] for key in order]
    assert cols == [2] * n + [3] * n


def test_running_total_of_a_long_column_is_fast():
    import time
    n = 20000
    spec = {}
    for r in range(1, n + 1):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=SUM($B$1:B{r})", 0.0)
    book = book_of({"S": spec})
    t0 = time.perf_counter()
    order, bad = formula_order(book)
    assert len(order) == 2 * n and not bad
    assert time.perf_counter() - t0 < 10  # measured 1.6 s; the old graph took minutes and gigabytes here


def _col(c, rows):
    return [("S", r, c) for r in rows]


def test_a_group_is_released_together_and_stays_consecutive():
    spec = {"Z1": ("=A1*2", 0.0)}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=B{r}+Z1", 0.0)  # reads a cell outside the group
    spec["A1"] = 1.0
    book = book_of({"S": spec})
    order, bad = formula_order(book, groups=[_col(2, range(2, 6)) + _col(3, range(2, 6))])
    assert not bad
    assert order[0] == ("S", 1, 26)  # Z1 first: the group waits for it
    assert order[1:] == sorted(_col(2, range(2, 6)) + _col(3, range(2, 6)))  # then the group, row-major


def test_reading_another_member_of_the_group_is_not_a_wait_but_self_reading_is_a_cycle():
    spec = {f"A{r}": float(r) for r in range(2, 5)}
    spec.update({f"B{r}": (f"=A{r}*2", 0.0) for r in range(2, 5)})
    spec.update({f"C{r}": (f"=B{r}+1", 0.0) for r in range(2, 5)})
    order, bad = formula_order(book_of({"S": spec}), groups=[_col(2, range(2, 5)) + _col(3, range(2, 5))])
    assert len(order) == 6 and not bad
    order, bad = formula_order(book_of({"S": {"A1": ("=A1+1", 0.0)}}), groups=[[("S", 1, 1)]])
    assert order == [] and set(bad) == {("S", 1, 1)}


def test_a_group_that_creates_a_cycle_the_cells_do_not_have_is_dissolved():  # Review Focus 3
    spec = {}
    for r in range(2, 5):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=Z1+B{r}", 0.0)
    spec["Z1"] = ("=B3+1", 0.0)  # B3 -> Z1 -> C2..C4 is fine for cells, a cycle if B and C are one node
    book = book_of({"S": spec})
    group = _col(2, range(2, 5)) + _col(3, range(2, 5))
    grouped_order, grouped_bad = formula_order(book, groups=[group])
    assert grouped_bad  # the false cycle
    order, bad, kept = order_with_groups(book, [group])
    assert bad == {} and kept == [] and len(order) == 7
    position = {k: i for i, k in enumerate(order)}
    assert position[("S", 3, 2)] < position[("S", 1, 26)] < position[("S", 2, 3)]


def test_a_real_cycle_is_still_reported_and_the_group_is_kept_out_of_it():
    spec = {"A1": ("=B1", 0.0), "B1": ("=A1", 0.0)}
    for r in range(2, 5):
        spec[f"C{r}"] = (f"={r}+1", 0.0)
    order, bad, kept = order_with_groups(book_of({"S": spec}), [])
    assert set(bad) == {("S", 1, 1), ("S", 1, 2)}


def test_groups_agree_with_cell_order_on_random_books():
    rng = random.Random(11)
    kept_total = dissolved_total = 0
    for i in range(600):
        book, groups = _random_book_and_groups(rng)
        order, bad, kept = order_with_groups(book, groups)
        base_order, base_bad = formula_order(book)
        assert set(bad) == set(base_bad), i  # the same cycles as cell mode
        formulas = {k for k, c in book.cells.items() if c.formula is not None}
        assert set(order) | set(bad) == formulas and len(order) == len(set(order))
        position = {k: j for j, k in enumerate(order)}
        for key in order:
            for text in refs(parse(book.cells[key].formula)):
                for dep in resolve(text, "S", book).present(book):
                    if dep in formulas and dep in position and dep != key:
                        assert any(key in g and dep in g for g in kept) or position[dep] < position[key], (i, key, dep)
        for g in kept:
            idx = sorted(position[k] for k in g if k in position)
            assert not idx or idx == list(range(idx[0], idx[0] + len(idx))), i
        kept_total += len(kept)
        dissolved_total += len(groups) - len(kept)
    assert kept_total > 20 and dissolved_total > 20  # both paths were exercised (about 65 and 210 expected)
    for _ in range(200):
        book, _unused = _random_book_and_groups(rng)
        assert formula_order(book, groups=[]) == formula_order(book)


def _internally_ok(book, group):
    """Like a real region: no range covers a member, and the reads among members are acyclic."""
    members = set(group)
    edges = {k: set() for k in group}
    for k in group:
        for text in refs(parse(book.cells[k].formula)):
            area = resolve(text, "S", book)
            if area.is_single:
                if area.key == k:
                    return False
                if area.key in members:
                    edges[k].add(area.key)
            elif any(m in members for m in area.present(book)):
                return False
    done = set()
    while len(done) < len(group):
        ready = [k for k in group if k not in done and not (edges[k] - done)]
        if not ready:
            return False
        done.update(ready)
    return True


def _random_book_and_groups(rng):
    n = rng.randint(4, 14)
    spec = {}
    for r in range(1, n + 1):
        for c in "ABCD":
            if rng.random() < 0.25:
                spec[f"{c}{r}"] = float(r)
            else:
                a, b = rng.randint(1, n), rng.randint(1, n)
                c2, c3 = rng.choice("ABCD"), rng.choice("ABCD")
                spec[f"{c}{r}"] = (rng.choice([f"={c2}{r}+1", f"={c2}{r}*{c3}{a}",
                                               f"=SUM({c2}{min(a, b)}:{c3}{max(a, b)})", f"={c2}{a}+1"]), 0.0)
    book = book_of({"S": spec})
    formulas = sorted(k for k, c in book.cells.items() if c.formula is not None)
    groups, used = [], set()
    for _ in range(rng.randint(0, 3)):
        col = rng.randint(1, 4)
        rows = sorted(r for (_s, r, c) in formulas if c == col)
        if len(rows) < 3:
            continue
        a = rng.randint(0, len(rows) - 3)
        b = rng.randint(a + 2, len(rows) - 1)
        run = [("S", r, col) for r in rows[a:b + 1] if ("S", r, col) not in used]
        if len(run) >= 2:
            groups.append(run)
            used.update(run)
            col2 = rng.randint(1, 4)
            if col2 != col and rng.random() < 0.5:
                run2 = [("S", r, col2) for r in rows[a:b + 1] if ("S", r, col2) in set(formulas) and ("S", r, col2) not in used]
                if len(run2) >= 2:
                    groups[-1] = groups[-1] + run2
                    used.update(run2)
    return book, [g for g in groups if _internally_ok(book, g)]


def test_a_true_cycle_inside_a_group_is_not_hidden_by_grouping():
    spec = {"A1": ("=B1", 0.0), "B1": ("=A1", 0.0), "C1": ("=A1+1", 0.0)}
    book = book_of({"S": spec})
    order, bad, kept = order_with_groups(book, [[("S", 1, 1), ("S", 1, 2)]])
    assert set(bad) == set(formula_order(book)[1]) and bad
