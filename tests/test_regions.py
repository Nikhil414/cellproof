from xlsxkit import book_of

from unexcel.regions import MIN_ROWS, column_order, find_regions


def sheet(n=6):
    spec = {"A1": "amount", "B1": "tax"}
    for r in range(2, n + 2):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*1.1", 0.0)  # row-wise
        spec[f"C{r}"] = (f"=MAX($A$2:A{r})", 0.0)  # growing range outside SUM: stays in cell mode
        spec[f"D{r}"] = (f"=A{r - 1}+1", 0.0)  # previous row: stays in cell mode
    return book_of({"S": spec})


def spans(scan):
    return [(r.sheet, r.r1, r.r2, r.cols) for r in scan.regions]


def test_finds_only_the_row_wise_column_and_says_why_the_rest_stayed():
    scan = find_regions(sheet())
    assert spans(scan) == [("S", 2, 7, [2])]
    notes = " ".join(scan.rejected)
    assert "S column C rows 2-7" in notes and "S column D rows 2-7" in notes


def test_a_run_shorter_than_min_rows_is_not_a_region():
    assert MIN_ROWS == 3
    book = book_of({"S": {"A1": 1.0, "A2": 2.0, "B1": ("=A1*2", 0.0), "B2": ("=A2*2", 0.0)}})
    assert find_regions(book).regions == []


def test_columns_with_the_same_rows_form_one_region():
    spec = {}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=B{r}+1", 0.0)
    scan = find_regions(book_of({"S": spec}))
    assert spans(scan) == [("S", 2, 5, [2, 3])]
    region = scan.regions[0]
    assert region.reads == {2: {1}, 3: {2}}
    assert region.first_key == ("S", 2, 2)
    assert region.keys()[0] == ("S", 2, 2) and len(region.keys()) == 8


def test_different_row_spans_are_separate_regions():
    spec = {}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
    for r in range(2, 8):
        spec[f"C{r}"] = ("=$A$2+B$2", 0.0)  # the same formula on every row, rows 2..7
    scan = find_regions(book_of({"S": spec}))
    assert spans(scan) == [("S", 2, 5, [2]), ("S", 2, 7, [3])]


def test_a_range_over_the_regions_own_formula_cells_keeps_the_column_out():
    spec = {}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=SUM($B$2:$B$5)+A{r}", 0.0)  # absolute range over column B's formula cells
    scan = find_regions(book_of({"S": spec}))
    assert spans(scan) == [("S", 2, 5, [2])]


def test_a_same_row_cycle_between_columns_keeps_both_out():
    spec = {}
    for r in range(2, 6):
        spec[f"B{r}"] = (f"=C{r}+1", 0.0)
        spec[f"C{r}"] = (f"=B{r}+1", 0.0)
        spec[f"D{r}"] = (f"=A{r}*2", 0.0)
        spec[f"A{r}"] = float(r)
    assert spans(find_regions(book_of({"S": spec}))) == [("S", 2, 5, [4])]


def test_array_formulas_and_formulas_without_a_saved_value_break_runs():
    spec = {f"A{r}": float(r) for r in range(2, 8)}
    spec.update({f"B{r}": (f"=A{r}*2", 0.0) for r in range(2, 8)})
    spec["B5"] = ("=A5*2", None)  # no saved value: kind blank
    scan = find_regions(book_of({"S": spec}))
    assert spans(scan) == [("S", 2, 4, [2])]  # rows 6..7 are too few for a second run


def test_column_order_is_dependency_order_with_the_smallest_ready_column_first():
    reads = {2: {1}, 3: {2}, 4: {1}, 5: {3, 4}}
    assert column_order([2, 3, 4, 5], reads) == ([2, 3, 4, 5], [])
    assert column_order([3, 2], {3: {2}, 2: {3}}) == ([], [2, 3])


def test_a_deeply_nested_formula_stays_in_cell_mode_instead_of_crashing():
    spec = {}
    for r in range(2, 6):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = ("=" + "+".join([f"A{r}"] * 1500), 0.0)
    scan = find_regions(book_of({"S": spec}))
    assert scan.regions == []
    assert "S column B rows 2-5 stays in cell mode" in " ".join(scan.rejected)


def test_regions_are_ordered_with_a_single_graph_pass(monkeypatch):
    """Regions cannot hide a cycle inside themselves, so the extra ungrouped pass is skipped when nothing is unresolved."""
    import unexcel.graph as graph
    from unexcel.regions import order_with_regions

    calls = []
    real = graph.formula_order
    monkeypatch.setattr(graph, "formula_order", lambda book, groups=(): calls.append(1) or real(book, groups))
    spec = {}
    for r in range(2, 8):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
        spec[f"C{r}"] = (f"=B{r}+1", 0.0)
    book = book_of({"S": spec})
    order, bad, kept = order_with_regions(book, find_regions(book).regions)
    assert len(kept) == 1 and not bad and len(order) == 12
    assert len(calls) == 1


def test_a_true_cycle_elsewhere_still_runs_the_full_guard():
    from unexcel.regions import order_with_regions

    spec = {"Z1": ("=Z2", 0.0), "Z2": ("=Z1", 0.0)}  # a true cycle outside the region
    for r in range(2, 8):
        spec[f"A{r}"] = float(r)
        spec[f"B{r}"] = (f"=A{r}*2", 0.0)
    book = book_of({"S": spec})
    order, bad, kept = order_with_regions(book, find_regions(book).regions)
    assert set(bad) == {("S", 1, 26), ("S", 2, 26)} and len(kept) == 1 and len(order) == 6
