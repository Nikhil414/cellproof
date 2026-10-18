import random

from unexcel.fenwick import Fenwick


def test_prefix_range_and_add_match_a_plain_list():
    rng = random.Random(1)
    for _ in range(200):
        n = rng.randint(0, 40)
        values = [rng.randint(0, 5) for _ in range(n)]
        fw = Fenwick(values)
        for _ in range(30):
            if n and rng.random() < 0.4:
                i, d = rng.randrange(n), rng.randint(0, 3)
                values[i] += d
                fw.add(i, d)
            i = rng.randint(0, n)
            j = rng.randint(i, n)
            assert fw.prefix(i) == sum(values[:i])
            assert fw.range(i, j) == sum(values[i:j])


def test_kth_finds_the_position_of_the_kth_unit():
    rng = random.Random(2)
    for _ in range(200):
        n = rng.randint(1, 40)
        values = [rng.randint(0, 3) for _ in range(n)]
        fw = Fenwick(values)
        units = [i for i, v in enumerate(values) for _ in range(v)]
        for k, pos in enumerate(units, 1):
            assert fw.kth(k) == pos


def test_big_packed_integers_work_for_prefix_sums():
    fw = Fenwick([1 << 100, 1 << 100, 0])
    fw.add(2, 1 << 90)
    assert fw.range(0, 3) == (2 << 100) + (1 << 90)
