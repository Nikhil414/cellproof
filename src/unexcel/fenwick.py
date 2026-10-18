"""Prefix sums over positions 0..n-1 with point updates and k-th lookup, all O(log n)."""
from __future__ import annotations


class Fenwick:
    def __init__(self, values: list[int]):
        n = len(values)
        self.n = n
        t = [0] + list(values)
        for i in range(1, n + 1):  # O(n) build
            j = i + (i & -i)
            if j <= n:
                t[j] += t[i]
        self.t = t
        self.top = 1 << n.bit_length()

    def add(self, i: int, delta: int) -> None:
        """Add delta to position i."""
        i += 1
        t, n = self.t, self.n
        while i <= n:
            t[i] += delta
            i += i & -i

    def prefix(self, i: int) -> int:
        """Sum of positions 0..i-1."""
        s, t = 0, self.t
        while i > 0:
            s += t[i]
            i -= i & -i
        return s

    def range(self, i: int, j: int) -> int:
        """Sum of positions i..j-1."""
        return self.prefix(j) - self.prefix(i)

    def kth(self, k: int) -> int:
        """0-based position of the k-th unit (k >= 1) when every value is a non-negative count."""
        pos, t, n = 0, self.t, self.n
        step = self.top
        while step:
            nxt = pos + step
            if nxt <= n and t[nxt] < k:
                pos = nxt
                k -= t[nxt]
            step >>= 1
        return pos
