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
| jackson-dataformat-spreadsheet | 1.6.0 |
| Apache POI | 5.5.1 |
| FastExcel | 0.20.0 |
| Fesod | 2.0.1-incubating |
| Poiji | 5.4.0 |

## Read — Throughput

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | 2.1 ms | 17.7 ms | 93.2 ms | 196.0 ms |
| FastExcel | 2.4 ms | 21.1 ms | 105.7 ms | 252.0 ms |
| Fesod | 3.3 ms | 27.8 ms | 136.3 ms | 266.6 ms |
| Poiji | 9.3 ms | 86.9 ms | 418.4 ms | 847.7 ms |
| Apache POI | 10.6 ms | 101.8 ms | 541.1 ms | 1094.5 ms |

## Read — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| jackson-spreadsheet | 36 MB | 180 MB | 360 MB |
| Fesod | 39 MB | 197 MB | 384 MB |
| FastExcel | 41 MB | 209 MB | 407 MB |
| Apache POI | 222 MB | 1115 MB | 2223 MB |
| Poiji | 274 MB | 1372 MB | 2743 MB |

## Write — Throughput

Poiji is read-only and not included.

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | 3.4 ms | 16.7 ms | 74.8 ms | 143.7 ms |
| FastExcel | 2.7 ms | 15.7 ms | 77.9 ms | 153.9 ms |
| Apache POI | 4.3 ms | 27.6 ms | 133.4 ms | 267.6 ms |
| Fesod | 4.8 ms | 32.9 ms | 154.5 ms | 314.3 ms |

## Write — Memory (bytes/op)

| Library | 10K rows | 50K rows | 100K rows |
|---------|----------|----------|-----------|
| FastExcel | 15 MB | 75 MB | 149 MB |
| jackson-spreadsheet | 20 MB | 92 MB | 182 MB |
| Apache POI | 21 MB | 104 MB | 207 MB |
| Fesod | 48 MB | 230 MB | 458 MB |

## jackson — Streaming vs POI User Model

Default path uses streaming. `USE_POI_USER_MODEL` switches to POI's Sheet/Row/Cell API.
Apache POI (raw, no jackson) shown as reference.

**Read:**

| Mode | 1K rows | 10K rows | 50K rows | 100K rows | Memory (100K) |
|------|---------|----------|----------|-----------|---------------|
| Streaming (default) | 2.1 ms | 17.7 ms | 93.2 ms | 196.0 ms | 360 MB |
| POI User Model | 15.2 ms | 123.3 ms | 600.4 ms | 1393.1 ms | 2807 MB |
| Apache POI (reference) | 10.6 ms | 101.8 ms | 541.1 ms | 1094.5 ms | 2225 MB |

**Write:**

| Mode | 1K rows | 10K rows | 50K rows | 100K rows | Memory (100K) |
|------|---------|----------|----------|-----------|---------------|
| Streaming (default) | 3.4 ms | 16.7 ms | 74.8 ms | 143.7 ms | 182 MB |
| POI User Model | 4.6 ms | 32.6 ms | 155.3 ms | 317.9 ms | 246 MB |
| Apache POI (reference) | 4.3 ms | 27.6 ms | 133.4 ms | 267.6 ms | 207 MB |

- Read: Streaming is 7.1x faster and uses 87% less memory than POI User Model at 100K rows.
- Write: Streaming is 2.2x faster and uses 26% less memory than POI User Model at 100K rows.
- POI User Model is slower than raw Apache POI due to Jackson serialization and schema overhead.

## SharedStrings — FILE_BACKED Variants

`FILE_BACKED_SHARED_STRINGS` applies to both read and write paths. String-heavy workload (4 string columns, mostly unique values):

**Read:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | 13.7 ms | 62.7 ms | 134.7 ms |
| FileBacked (H2 MVStore) | 28.6 ms | 120.0 ms | 222.9 ms |
| FileBacked + Encrypted | 33.5 ms | 151.2 ms | 256.7 ms |

**Write:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | 18.8 ms | 90.9 ms | 184.4 ms |
| FileBacked (H2 MVStore) | 41.0 ms | 207.5 ms | 419.8 ms |
| FileBacked + Encrypted | 49.9 ms | 274.9 ms | 626.3 ms |

- FileBacked's value is **peak heap reduction**, not throughput. When the SST exceeds available heap, InMemory causes OOM while FileBacked stays constant.

## Write — AutoSize Cost (POI mode)

`autoSize` is supported only in `USE_POI_USER_MODEL` mode (the streaming default doesn't measure column widths). Cost compared against the same POI mode without `autoSize`:

| Mode | 1K rows | 10K rows | 100K rows |
|------|---------|----------|-----------|
| Without autoSize | 4.7 ms | 31.9 ms | 336.6 ms |
| With autoSize | 13.7 ms | 55.4 ms | 469.9 ms |
| Overhead | 2.92x | 1.74x | 1.40x |

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

---

<sub>Anonymous, aggregated usage tracking via [Scarf](https://about.scarf.sh/). No personal information is collected.</sub>
<img referrerpolicy="no-referrer-when-downgrade" src="https://static.scarf.sh/a.png?x-pxid=5217c9e5-45fd-4941-a10b-354455edbd72&page=BENCHMARK.md" />
