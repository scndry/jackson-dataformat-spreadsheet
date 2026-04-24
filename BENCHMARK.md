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
| jackson-spreadsheet | 2.5 ms | 17.5 ms | 84.0 ms | 182.7 ms |
| FastExcel Reader | 2.2 ms | 21.8 ms | 101.4 ms | 211.8 ms |
| Fesod | 3.3 ms | 27.2 ms | 134.5 ms | 264.6 ms |
| Poiji | 10.5 ms | 88.3 ms | 428.2 ms | 854.4 ms |
| Apache POI UserModel | 11.3 ms | 102.7 ms | 593.9 ms | 1059.0 ms |

- jackson-spreadsheet is fastest at 10K+ rows — 14% faster than FastExcel at 100K.
- Fesod is ~45% slower at 100K rows.
- Poiji wraps POI UserModel with annotation mapping — faster than raw POI due to internal optimizations, but still 4.7x slower than jackson-spreadsheet.
- POI UserModel loads the full workbook into memory — 5.8x slower at 100K rows.

## Read — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| jackson-spreadsheet | 38 MB | 189 MB | 378 MB |
| Fesod | 41 MB | 206 MB | 400 MB |
| FastExcel Reader | 43 MB | 219 MB | 428 MB |
| POI UserModel | 233 MB | 1161 MB | 2332 MB |
| Poiji | 288 MB | 1438 MB | 2876 MB |

- jackson-spreadsheet has the lowest allocation among all libraries thanks to SoA shared string layout and direct String return (no POI RichTextString wrapping).
- Poiji allocates more than raw POI due to annotation reflection and internal object creation overhead.

## Write — Throughput

Poiji is read-only and not included.

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet (default) | 5.2 ms | 19.6 ms | 77.3 ms | 153.4 ms |
| FastExcel | 3.0 ms | 17.0 ms | 86.3 ms | 168.8 ms |
| Apache POI SXSSF | 6.1 ms | 31.6 ms | 148.1 ms | 292.5 ms |
| jackson-spreadsheet (POI User Model) | 7.6 ms | 37.3 ms | 177.1 ms | 349.4 ms |
| Fesod | 7.8 ms | 40.8 ms | 170.2 ms | 351.7 ms |

- jackson-spreadsheet is fastest at 50K+ rows — 9% faster than FastExcel at 100K.
- The default path builds a POI XSSFWorkbook skeleton for package metadata, then streams worksheet and sharedStrings via StringBuilder directly to ZipOutputStream.
- The POI User Model path uses `SXSSFWorkbook` internally. The ~19% overhead vs raw SXSSF comes from Jackson serialization and schema-driven cell styling.
- Fesod wraps POI `SXSSFWorkbook` with handler chain, cglib BeanMap, and converter overhead — comparable to jackson-spreadsheet POI User Model.

## Write — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| FastExcel | 16 MB | 78 MB | 156 MB |
| jackson-spreadsheet (default) | 21 MB | 96 MB | 191 MB |
| Apache POI SXSSF | 22 MB | 109 MB | 214 MB |
| jackson-spreadsheet (POI User Model) | 27 MB | 130 MB | 258 MB |
| Fesod | 50 MB | 241 MB | 482 MB |

- Default path allocates 22% more than FastExcel but 11% less than raw SXSSF — the POI skeleton adds a fixed cost, but StringBuilder streaming avoids per-cell POI object overhead.
- Fesod allocates 2.5x more than jackson-spreadsheet due to cglib BeanMap and converter chain.

## SharedStrings — FILE_BACKED Variants

`FILE_BACKED_SHARED_STRINGS` applies to both read and write paths. String-heavy workload (4 string columns, mostly unique values):

**Read:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | 17.0 ms | 73.2 ms | 133.6 ms |
| FileBacked (H2 MVStore) | 30.9 ms | 135.8 ms | 264.2 ms |
| FileBacked + Encrypted | 34.3 ms | 146.2 ms | 289.2 ms |

**Write:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | 19.6 ms | 77.3 ms | 153.4 ms |
| FileBacked (H2 MVStore) | 36.4 ms | 168.8 ms | 317.2 ms |
| FileBacked + Encrypted | 41.1 ms | 224.2 ms | 496.6 ms |

- Read: FileBacked is ~98% slower, mitigated by LRU cache for repeated lookups. Encryption adds ~9%.
- Write: FileBacked is ~107% slower. Encryption adds ~57%.
- FileBacked's value is **peak heap reduction**, not throughput. When the SST exceeds available heap, InMemory causes OOM while FileBacked stays constant.

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
