# Changelog

## 0.2.0 (2026-10-19)

- **Renamed** from `unexcel` to `cellproof` (the old name clashed with an R package): `pip install cellproof`, command
  `cellproof convert`, import `cellproof`, `CELLPROOF_PERF` for the benchmark, dbt project and database `cellproof`.
- **Lookups row-wise:** VLOOKUP, HLOOKUP, MATCH, INDEX and XLOOKUP pasted down a column compile into one statement.
  1,000 VLOOKUPs: verify 5.7 s -> 0.1 s.
- **Criteria row-wise:** COUNTIFS (and SUMIF/AVERAGEIF(S) over a range holding an error) with a criterion from the row
  no longer fall back to cell mode. 1,000 COUNTIFS: verify 29 s -> 0.1 s.
- **Faster compile:** a pasted-down formula is tokenized once per run. 200k-formula benchmark 31 s -> 26 s; 50k
  VLOOKUPs 8.9 s -> 4.6 s.
- A lookup outside a region is about 20% slower than in 0.1.0 (one SQL form for both modes).

## 0.1.0 (2026-10-18)

First release, as `unexcel`.
