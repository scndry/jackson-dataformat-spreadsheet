# Benchmark

## Environment

- JDK 21.0.10, OpenJDK 64-Bit Server VM
- Apple Silicon (aarch64)
- JMH 1.35, Fork 1, Warmup 3 iterations, Measurement 5 iterations
- GC profiler enabled (`gc.alloc.rate.norm` = bytes allocated per operation)
- Test data: mixed types (String + int + double + boolean), `XSSFWorkbook` (shared string table), 6 columns
- All benchmarks use each library's default configuration for fair comparison
- Apache POI mode: `XSSFWorkbook` for read (User Model), `SXSSFWorkbook` for write (POI's streaming write API)

### Library Versions

| Library | Version |
|---------|---------|
| jackson-dataformat-spreadsheet | 1.4.0 |
| Apache POI | 5.5.1 |
| FastExcel | 0.20.0 |
| Fesod | 2.0.1-incubating |
| Poiji | 5.4.0 |

## Read — Throughput

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | 2.8 ms | 21.8 ms | 91.2 ms | 198.1 ms |
| FastExcel | 2.2 ms | 20.4 ms | 105.8 ms | 212.3 ms |
| Fesod | 3.6 ms | 29.4 ms | 138.4 ms | 279.4 ms |
| Poiji | 9.9 ms | 91.2 ms | 429.7 ms | 843.4 ms |
| Apache POI | 12.1 ms | 107.0 ms | 522.5 ms | 1197.6 ms |

## Read — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| jackson-spreadsheet | 38 MB | 189 MB | 378 MB |
| Fesod | 42 MB | 206 MB | 400 MB |
| FastExcel | 42 MB | 219 MB | 428 MB |
| Apache POI | 235 MB | 1167 MB | 2333 MB |
| Poiji | 288 MB | 1438 MB | 2876 MB |

## Write — Throughput

Poiji is read-only and not included.

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | 4.5 ms | 17.8 ms | 76.9 ms | 150.3 ms |
| FastExcel | 2.8 ms | 16.7 ms | 83.7 ms | 165.8 ms |
| Apache POI | 6.2 ms | 30.9 ms | 147.0 ms | 283.0 ms |
| Fesod | 7.1 ms | 36.6 ms | 166.9 ms | 336.7 ms |

## Write — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| FastExcel | 16 MB | 78 MB | 156 MB |
| jackson-spreadsheet | 21 MB | 96 MB | 191 MB |
| Apache POI | 23 MB | 108 MB | 207 MB |
| Fesod | 50 MB | 247 MB | 480 MB |

## jackson — Streaming vs POI User Model

Default path uses streaming. `USE_POI_USER_MODEL` switches to POI's Sheet/Row/Cell API.
Apache POI (raw, no jackson) shown as reference.

**Read:**

| Mode | 1K rows | 10K rows | 50K rows | 100K rows | Memory (100K) |
|------|---------|----------|----------|-----------|---------------|
| Streaming (default) | 2.8 ms | 21.8 ms | 91.2 ms | 198.1 ms | 378 MB |
| POI User Model | 15.8 ms | 124.1 ms | 673.6 ms | 1459.8 ms | 2943 MB |
| Apache POI (reference) | 12.1 ms | 107.0 ms | 522.5 ms | 1197.6 ms | 2333 MB |

**Write:**

| Mode | 1K rows | 10K rows | 50K rows | 100K rows | Memory (100K) |
|------|---------|----------|----------|-----------|---------------|
| Streaming (default) | 4.5 ms | 17.8 ms | 76.9 ms | 150.3 ms | 191 MB |
| POI User Model | 7.2 ms | 35.4 ms | 166.9 ms | 334.5 ms | 258 MB |
| Apache POI (reference) | 6.2 ms | 30.9 ms | 147.0 ms | 283.0 ms | 207 MB |

- Read: Streaming is 7.4x faster and uses 87% less memory than POI User Model at 100K rows.
- Write: Streaming is 2.2x faster and uses 26% less memory than POI User Model at 100K rows.
- POI User Model is slower than raw Apache POI due to Jackson serialization and schema overhead.

## SharedStrings — FILE_BACKED Variants

`FILE_BACKED_SHARED_STRINGS` applies to both read and write paths. String-heavy workload (4 string columns, mostly unique values):

**Read:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | 16.7 ms | 64.7 ms | 136.7 ms |
| FileBacked (H2 MVStore) | 30.9 ms | 130.3 ms | 252.5 ms |
| FileBacked + Encrypted | 32.9 ms | 144.8 ms | 290.4 ms |

**Write:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | 17.8 ms | 76.9 ms | 150.3 ms |
| FileBacked (H2 MVStore) | 35.5 ms | 164.8 ms | 329.0 ms |
| FileBacked + Encrypted | 39.6 ms | 228.4 ms | 480.5 ms |

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
