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
| FastExcel Reader | 2.2 ms | 21.6 ms | 100.8 ms | 215.3 ms |
| jackson-spreadsheet | 2.9 ms | 19.5 ms | 103.7 ms | 224.3 ms |
| Fesod | 3.3 ms | 27.5 ms | 134.1 ms | 280.7 ms |
| Poiji | 9.6 ms | 87.4 ms | 420.0 ms | 837.9 ms |
| Apache POI UserModel | 11.5 ms | 96.0 ms | 525.7 ms | 1175.7 ms |

- jackson-spreadsheet and FastExcel are within 5% of each other at 50K+ rows.
- Fesod is ~25% slower at 100K rows.
- Poiji wraps POI UserModel with annotation mapping — faster than raw POI due to internal optimizations, but still 4x slower than jackson-spreadsheet.
- POI UserModel loads the full workbook into memory — 6x slower at 100K rows.

## Read — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| jackson-spreadsheet | 36 MB | 180 MB | 352 MB |
| Fesod | 39 MB | 193 MB | 384 MB |
| FastExcel Reader | 40 MB | 208 MB | 407 MB |
| POI UserModel | 223 MB | 1113 MB | 2224 MB |
| Poiji | 274 MB | 1373 MB | 2739 MB |

- jackson-spreadsheet has the lowest allocation among all libraries thanks to SoA shared string layout and direct String return (no POI RichTextString wrapping).
- Poiji allocates more than raw POI due to annotation reflection and internal object creation overhead.

## Write — Throughput

Poiji is read-only and not included.

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| FastExcel | 2.8 ms | 15.9 ms | 77.0 ms | 161.0 ms |
| Apache POI SXSSF | 4.9 ms | 28.4 ms | 135.6 ms | 262.9 ms |
| jackson-spreadsheet | 6.3 ms | 34.1 ms | 162.4 ms | 331.9 ms |
| Fesod | 6.5 ms | 35.8 ms | 170.4 ms | 333.0 ms |

- FastExcel is fastest — it generates OOXML XML directly without Apache POI.
- jackson-spreadsheet uses POI `SXSSFWorkbook` internally. The ~26% overhead vs raw SXSSF comes from Jackson serialization and schema-driven cell styling.
- Fesod wraps POI `SXSSFWorkbook` with handler chain, cglib BeanMap, and converter overhead — comparable to jackson-spreadsheet.

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
