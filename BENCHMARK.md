# Benchmark

## Environment

- JDK 21.0.10, OpenJDK 64-Bit Server VM
- Apple Silicon (aarch64)
- JMH 1.35, Fork 1, Warmup 3 iterations, Measurement 5 iterations
- GC profiler enabled (`gc.alloc.rate.norm` = bytes allocated per operation)
- Test data: mixed types (String + int + double + boolean), `XSSFWorkbook` (shared string table), 6 columns
- All read benchmarks produce the same POJO for fair comparison

## Read — Throughput

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| FastExcel Reader | 2.5 ms | 20.9 ms | 102.7 ms | 209.1 ms |
| jackson-spreadsheet | 3.0 ms | 23.6 ms | 103.9 ms | 220.3 ms |
| EasyExcel | 4.1 ms | 32.4 ms | 139.1 ms | 296.4 ms |
| Poiji | 11.4 ms | 100.1 ms | 448.4 ms | 889.4 ms |
| Apache POI UserModel | 12.4 ms | 106.2 ms | 592.6 ms | 1274.3 ms |

- jackson-spreadsheet and FastExcel are within 5% of each other at 50K+ rows.
- EasyExcel is ~35% slower at 100K rows.
- Poiji wraps POI UserModel with annotation mapping — faster than raw POI due to internal optimizations, but still 4x slower than jackson-spreadsheet.
- POI UserModel loads the full workbook into memory — 6x slower at 100K rows.

## Read — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| jackson-spreadsheet | 38 MB | 192 MB | 395 MB |
| FastExcel Reader | 41 MB | 204 MB | 406 MB |
| EasyExcel | 43 MB | 210 MB | 418 MB |
| POI UserModel | 234 MB | 1162 MB | 2347 MB |
| Poiji | 303 MB | 1455 MB | 2909 MB |

- jackson-spreadsheet has the lowest allocation among all libraries thanks to SoA shared string layout and direct String return (no POI RichTextString wrapping).
- Poiji allocates more than raw POI due to annotation reflection and internal object creation overhead.

## Write — Throughput

Poiji is read-only and not included.

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| EasyExcel | 3.4 ms | 6.3 ms | 20.1 ms | 39.3 ms |
| FastExcel | 2.9 ms | 18.5 ms | 99.8 ms | 191.9 ms |
| Apache POI SXSSF | 5.7 ms | 29.3 ms | 134.1 ms | 265.3 ms |
| jackson-spreadsheet | 6.5 ms | 33.3 ms | 152.3 ms | 319.2 ms |

- jackson-spreadsheet uses POI `SXSSFWorkbook` internally. The ~20% overhead vs raw SXSSF comes from Jackson serialization and schema-driven cell styling.
- EasyExcel is fastest for write due to its own optimized streaming writer.

## SharedStrings — InMemory vs FileBacked

String-heavy workload (4 string columns, mostly unique values, `XSSFWorkbook`):

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | 19.0 ms | 86.4 ms | 175.7 ms |
| FileBacked (H2 MVStore) | 29.3 ms | 119.4 ms | 248.8 ms |

- FileBacked is ~40% slower due to disk I/O per cache miss and MVStore overhead.
- FileBacked's value is **peak heap reduction** for extremely large shared string tables, not throughput. When the SST exceeds available heap, InMemory causes OOM while FileBacked stays constant at ~4 MB cache.

## How to Run

```bash
./gradlew jmh
```

Filter specific benchmarks:

```bash
./gradlew jmh -Pjmh.includes="ReadBenchmark"
./gradlew jmh -Pjmh.includes="WriteBenchmark"
./gradlew jmh -Pjmh.includes="MemoryBenchmark"
./gradlew jmh -Pjmh.includes="SharedStringsBenchmark"
```

GC profiler is enabled by default. Results are saved to `build/results/jmh/results.txt`.
