Python 3.12.14 on Darwin arm64

## Recalculation (hash indexes + aggregate memo on)

| rows | formulas | parse s | graph s | eval s | total s | formulas/s | accuracy vs ground truth |
|---:|---:|---:|---:|---:|---:|---:|---:|
| 1,000 | 12,000 | 0.186 | 0.027 | 0.049 | 0.260 | 46,120 | 12,000/12,000 |
| 10,000 | 120,000 | 2.134 | 0.300 | 0.509 | 2.945 | 40,749 | 120,000/120,000 |
| 50,000 | 600,000 | 10.963 | 1.708 | 2.703 | 15.102 | 39,729 | 600,000/600,000 |

## Effect of criteria hash index and aggregate memo

| rows | caches off s | caches on s | speedup |
|---:|---:|---:|---:|
| 500 | 0.560 | 0.120 | 4.7x |
| 1,000 | 2.038 | 0.240 | 8.5x |
| 2,000 | 8.125 | 0.490 | 16.6x |

## Full conversion of a 10,000-row .xlsx (one hardcoded override at row 9)

| metric | value |
|---|---:|
| wall time (load + recalc + translate + write + validate + audit) | 4.98 s |
| recalc stage | 3.124 s |
| translate stage | 0.517 s |
| write stage | 0.028 s |
| validate stage | 0.223 s |
| calc columns translated | 10/12 |
| row-level checks agreeing with Excel | 99,998/100,000 |
| cells flagged as differing from their column pattern | 2 |
| .twbx size | 57 KiB |
