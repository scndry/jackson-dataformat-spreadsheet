# Benchmark

## Environment

- JDK 21.0.10, OpenJDK 64-Bit Server VM on Apple Silicon (aarch64)
- JMH 1.35, Fork 1, Warmup 3 × 1 s, Measurement 5 × 1 s
- GC profiler enabled

### Library Versions

| Library | Version |
|---------|---------|
| jackson-dataformat-spreadsheet | 1.8.0 |
| FastExcel | 0.20.0 |
| Apache Fesod | 2.0.1-incubating |
| Apache POI | 5.5.1 |

## Measurement Scope

These benchmarks measure a single combination of design choices. Other production scenarios — listed below — are not represented and may shift the numbers.

- **Schema** (`BenchRow`, 10 columns): `Long id` + four `String` (`name`, `category`, `status`, `description`) + `int quantity` + `double price` + `BigDecimal amount` + `LocalDate dueDate` + `LocalDateTime createdAt`
- **Source format**: `SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)` — streaming write with shared string table enabled. Inline-string sources change `InMemorySharedStringsLookup` / `StringCache` behavior.
- **Apache POI mode**: `XSSFWorkbook` for read (User Model), `SXSSFWorkbook` for write.
- **POJO mapping** via `BenchRow` setter/getter on every library. Cell-level use cases that bypass POJO mapping are not measured.
- **Per-op workbook open** (cold). Warm-mapper / reused-workbook scenarios reduce per-op cost.
- **Default `ArrayList` capacity** (no row-count hint). Pre-sized collectors trim allocation.
- **`AverageTime` mode** at short windows (1–5 s) for per-op cost. Sustained Throughput (60 s window) covers long-run GC pressure separately.
- **Default library configuration** otherwise — no per-library tuning.

## Read

### Throughput

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | 3.26 ms | 29.10 ms | 141.43 ms | 270.39 ms |
| FastExcel | **2.59 ms** | **25.70 ms** | **129.13 ms** | **257.12 ms** |
| Apache Fesod | 5.08 ms | 46.84 ms | 225.37 ms | 537.61 ms |
| Apache POI | 18.02 ms | 169.95 ms | 886.21 ms | 1964.68 ms |
| jackson-spreadsheet (POI User Model) | 20.49 ms | 190.94 ms | 978.04 ms | 2193.62 ms |

FastExcel is faster at every scale (5–21%); the gap is narrowest at 100K (5%) and widest at 1K (21%). Apache POI's User Model loads the entire file up front and is 5.5–7.3× slower than streaming readers — first-N early exit pays the same cost as a full read. Apache POI also ships a SAX event API (`XSSFReader` + `XSSFSheetXMLHandler`) for streaming reads; that path is not measured here.

### Allocation per op

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | **6.5 MB** | **57 MB** | **275 MB** | **548 MB** |
| FastExcel | 7.1 MB | 68 MB | 337 MB | 673 MB |
| Apache Fesod | 8.6 MB | 76 MB | 380 MB | 841 MB |
| Apache POI | 42 MB | 407 MB | 2046 MB | 4053 MB |
| jackson-spreadsheet (POI User Model) | 45 MB | 435 MB | 2181 MB | 4384 MB |

jackson allocates 9–19% less than FastExcel across all scales, ~7× less than Apache POI. POI User Model retains the full workbook in memory; heap budget must scale with file size.

## Write

### Throughput

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | 5.48 ms | 33.27 ms | **153.62 ms** | **303.45 ms** |
| FastExcel | **3.92 ms** | **31.00 ms** | 166.76 ms | 349.70 ms |
| Apache Fesod | 10.15 ms | 80.78 ms | 394.27 ms | 781.53 ms |
| Apache POI (SXSSF) | 8.26 ms | 66.12 ms | 323.79 ms | 655.88 ms |
| jackson-spreadsheet (POI User Model) | 9.45 ms | 79.36 ms | 399.27 ms | 799.91 ms |

jackson overtakes FastExcel at 50K+ (8% at 50K, 13% at 100K). FastExcel leads smaller workloads by 7–28%.

### Allocation per op

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | **3.3 MB** | **23 MB** | **112 MB** | **224 MB** |
| FastExcel | 3.6 MB | 31 MB | 157 MB | 313 MB |
| Apache Fesod | 12 MB | 106 MB | 524 MB | 1049 MB |
| Apache POI (SXSSF) | 6.5 MB | 49 MB | 249 MB | 496 MB |
| jackson-spreadsheet (POI User Model) | 7.1 MB | 57 MB | 275 MB | 551 MB |

jackson writes with 26–29% less allocation than FastExcel from 10K onward; Apache Fesod allocates 3.6–4.7× as much.

## Sustained Throughput

10K-row source looped over 60 seconds. Surfaces steady-state allocation rate and GC pause accumulation.

**Read:**

| Library | ops/sec | gc.alloc.rate | gc.count | gc.time |
|---------|---------|---------------|----------|---------|
| jackson-spreadsheet | 35.39 | **1996 MB/sec** | **256** | **242 ms** |
| FastExcel | **37.10** | 2490 MB/sec | 319 | 268 ms |

**Write:**

| Library | ops/sec | gc.alloc.rate | gc.count | gc.time |
|---------|---------|---------------|----------|---------|
| jackson-spreadsheet | 34.05 | **862 MB/sec** | **199** | **99 ms** |
| FastExcel | **34.55** | 1143 MB/sec | 212 | 179 ms |

- Read: FastExcel is ~5% faster but allocates 25% more per second and triggers 25% more GC cycles.
- Write: FastExcel is ~1.5% faster but allocates 33% more per second with ~1.8× longer GC time.

## SharedStrings

`FILE_BACKED_SHARED_STRINGS` applies to both read and write paths.

**Read:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | **14.89 ms** | **65.31 ms** | **126.78 ms** |
| FileBacked (H2 MVStore) | 26.57 ms | 106.53 ms | 204.89 ms |
| FileBacked + Encrypted | 30.29 ms | 121.66 ms | 239.50 ms |

**Write:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | **18.26 ms** | **77.92 ms** | **156.40 ms** |
| FileBacked (H2 MVStore) | 41.25 ms | 202.51 ms | 388.60 ms |
| FileBacked + Encrypted | 47.55 ms | 288.00 ms | 633.41 ms |

FileBacked trades throughput for **peak heap reduction** — InMemory OOMs when the SST exceeds available heap.

## Running & Reproducing

```bash
./gradlew jmh
```

Filter specific benchmarks:

```bash
./gradlew jmh -Pjmh.includes='\.ReadBenchmark\.'                -Pjmh.resultsFile='build/results/jmh/read.txt'
./gradlew jmh -Pjmh.includes='\.WriteBenchmark\.'               -Pjmh.resultsFile='build/results/jmh/write.txt'
./gradlew jmh -Pjmh.includes='\.SustainedThroughputBenchmark\.' -Pjmh.resultsFile='build/results/jmh/sustained.txt'
./gradlew jmh -Pjmh.includes='\.SharedStringsBenchmark\.'       -Pjmh.resultsFile='build/results/jmh/sharedstrings.txt'
```

Run benchmarks separately to avoid system noise accumulation across categories.
