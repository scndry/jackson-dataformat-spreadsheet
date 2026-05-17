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
| jackson-dataformat-spreadsheet | 1.6.4 |
| Apache POI | 5.5.1 |
| FastExcel | 0.20.0 |
| Fesod | 2.0.1-incubating |
| Poiji | 5.4.0 |

## Read — Throughput

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | 2.4 ms | 17.9 ms | 83.7 ms | 190.1 ms |
| FastExcel | 2.6 ms | 21.1 ms | 104.1 ms | 207.6 ms |
| Fesod | 3.4 ms | 28.7 ms | 133.9 ms | 265.6 ms |
| Poiji | 9.0 ms | 84.8 ms | 421.3 ms | 809.4 ms |
| Apache POI | 11.0 ms | 95.8 ms | 507.4 ms | 1172.8 ms |

## Read — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| jackson-spreadsheet | 36 MB | 180 MB | 360 MB |
| Fesod | 40 MB | 197 MB | 381 MB |
| FastExcel | 40 MB | 208 MB | 407 MB |
| Apache POI | 226 MB | 1114 MB | 2227 MB |
| Poiji | 274 MB | 1373 MB | 2739 MB |

## Write — Throughput

Poiji is read-only and not included.

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | 3.9 ms | 16.1 ms | 70.6 ms | 138.3 ms |
| FastExcel | 2.6 ms | 15.8 ms | 76.5 ms | 151.7 ms |
| Apache POI | 5.2 ms | 29.1 ms | 135.2 ms | 269.0 ms |
| Fesod | 5.9 ms | 34.1 ms | 160.3 ms | 322.6 ms |

## Write — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| jackson-spreadsheet | 14 MB | 62 MB | 125 MB |
| FastExcel | 15 MB | 75 MB | 149 MB |
| Apache POI | 21 MB | 103 MB | 204 MB |
| Fesod | 48 MB | 230 MB | 458 MB |

## jackson — Streaming vs POI User Model

Default path uses streaming. `USE_POI_USER_MODEL` switches to POI's Sheet/Row/Cell API.
Apache POI (raw, no jackson) shown as reference.

**Read:**

| Mode | 1K rows | 10K rows | 50K rows | 100K rows | Memory (100K) |
|------|---------|----------|----------|-----------|---------------|
| Streaming (default) | 2.4 ms | 17.9 ms | 83.7 ms | 190.1 ms | 360 MB |
| POI User Model | 14.0 ms | 120.7 ms | 602.2 ms | 1348.2 ms | 2807 MB |
| Apache POI (reference) | 11.0 ms | 95.8 ms | 507.4 ms | 1172.8 ms | 2227 MB |

**Write:**

| Mode | 1K rows | 10K rows | 50K rows | 100K rows | Memory (100K) |
|------|---------|----------|----------|-----------|---------------|
| Streaming (default) | 3.9 ms | 16.1 ms | 70.6 ms | 138.3 ms | 125 MB |
| POI User Model | 5.7 ms | 33.6 ms | 159.5 ms | 317.7 ms | 246 MB |
| Apache POI (reference) | 5.2 ms | 29.1 ms | 135.2 ms | 269.0 ms | 204 MB |

- Read: Streaming is 7.1x faster and uses 87% less memory than POI User Model at 100K rows.
- Write: Streaming is 2.3x faster and uses 49% less memory than POI User Model at 100K rows.
- POI User Model is slower than raw Apache POI due to Jackson serialization and schema overhead.

## SharedStrings — FILE_BACKED Variants

`FILE_BACKED_SHARED_STRINGS` applies to both read and write paths. String-heavy workload (4 string columns, mostly unique values):

**Read:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | 14.9 ms | 62.7 ms | 130.9 ms |
| FileBacked (H2 MVStore) | 26.7 ms | 117.5 ms | 230.4 ms |
| FileBacked + Encrypted | 32.0 ms | 136.0 ms | 242.8 ms |

**Write:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | 17.8 ms | 74.7 ms | 161.3 ms |
| FileBacked (H2 MVStore) | 42.9 ms | 212.0 ms | 418.2 ms |
| FileBacked + Encrypted | 50.2 ms | 282.9 ms | 633.0 ms |

- FileBacked's value is **peak heap reduction**, not throughput. When the SST exceeds available heap, InMemory causes OOM while FileBacked stays constant.

## Write — AutoSize Cost (POI mode)

`autoSize` is supported only in `USE_POI_USER_MODEL` mode (the streaming default doesn't measure column widths). Cost compared against the same POI mode without `autoSize`:

| Mode | 1K rows | 10K rows | 100K rows |
|------|---------|----------|-----------|
| Without autoSize | 5.9 ms | 33.2 ms | 314.8 ms |
| With autoSize | 14.7 ms | 55.0 ms | 451.3 ms |
| Overhead | 2.51x | 1.66x | 1.43x |

- Sampling strategy (full first 100 rows + every 100th row thereafter) keeps overhead bounded — overhead amortizes from 2.92x at 1K rows to 1.40x at 100K rows.
- For exact-fit guarantees on outlier-long values, set `@DataColumn(width = N)` explicitly.

## How to Run

```bash
./gradlew jmh
```

Filter specific benchmarks (and route results to per-benchmark files for separate runs):

```bash
./gradlew jmh -Pjmh.includes="ReadBenchmark"          -Pjmh.resultsFile="build/results/jmh/read.txt"
./gradlew jmh -Pjmh.includes="WriteBenchmark"         -Pjmh.resultsFile="build/results/jmh/write.txt"
./gradlew jmh -Pjmh.includes="MemoryBenchmark"        -Pjmh.resultsFile="build/results/jmh/memory.txt"
./gradlew jmh -Pjmh.includes="SharedStringsBenchmark" -Pjmh.resultsFile="build/results/jmh/sharedstrings.txt"
./gradlew jmh -Pjmh.includes="AutoSizeBenchmark"      -Pjmh.resultsFile="build/results/jmh/autosize.txt"
```

Run benchmarks separately to avoid system noise accumulation across categories. GC profiler is enabled by default.
