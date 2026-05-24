# Benchmark

## Environment

- JDK 21.0.10, OpenJDK 64-Bit Server VM on Apple Silicon (aarch64)
- JMH 1.35, Fork 1, Warmup 3 × 1 s, Measurement 5 × 1 s
- GC profiler enabled

### Library Versions

| Library | Version |
|---------|---------|
| jackson-dataformat-spreadsheet | 1.7.0 |
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
| jackson-spreadsheet | 3.30 ms | 29.41 ms | 145.03 ms | 277.67 ms |
| FastExcel | **2.90 ms** | **25.94 ms** | **137.10 ms** | **263.30 ms** |
| Apache Fesod | 5.17 ms | 46.23 ms | 232.03 ms | 554.32 ms |
| Apache POI | 18.45 ms | 170.88 ms | 979.53 ms | 2071.82 ms |
| jackson-spreadsheet (POI User Model) | 20.81 ms | 194.31 ms | 1142.17 ms | 2342.11 ms |

FastExcel is faster at every scale (5–14%); the gap is narrowest at 100K (5%) and widest at 1K (14%). Apache POI's User Model loads the entire file up front and is 6–8× slower than streaming readers — first-N early exit pays the same cost as a full read. Apache POI also ships a SAX event API (`XSSFReader` + `XSSFSheetXMLHandler`) for streaming reads; that path is not measured here.

### Allocation per op

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | **6.5 MB** | **57 MB** | **275 MB** | **548 MB** |
| FastExcel | 7.4 MB | 68 MB | 352 MB | 675 MB |
| Apache Fesod | 8.6 MB | 76 MB | 380 MB | 841 MB |
| Apache POI | 42 MB | 407 MB | 2043 MB | 4086 MB |
| jackson-spreadsheet (POI User Model) | 45 MB | 435 MB | 2193 MB | 4392 MB |

jackson allocates 12–22% less than FastExcel across all scales, 7× less than Apache POI. POI User Model retains the full workbook in memory; heap budget must scale with file size.

## Write

### Throughput

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | 5.53 ms | 37.38 ms | **167.72 ms** | **333.89 ms** |
| FastExcel | **4.10 ms** | **31.29 ms** | 171.45 ms | 352.81 ms |
| Apache Fesod | 10.51 ms | 81.57 ms | 388.84 ms | 783.00 ms |
| Apache POI (SXSSF) | 8.36 ms | 66.41 ms | 328.76 ms | 644.25 ms |
| jackson-spreadsheet (POI User Model) | 9.34 ms | 77.66 ms | 389.50 ms | 803.11 ms |

jackson overtakes FastExcel at 50K+ (2% at 50K, 5% at 100K). FastExcel leads smaller workloads by 19–35%.

### Allocation per op

| Library | 1K rows | 10K rows | 50K rows | 100K rows |
|---------|---------|----------|----------|-----------|
| jackson-spreadsheet | **3.4 MB** | **23 MB** | **119 MB** | **237 MB** |
| FastExcel | 3.5 MB | 31 MB | 157 MB | 313 MB |
| Apache Fesod | 12 MB | 106 MB | 524 MB | 1046 MB |
| Apache POI (SXSSF) | 6.5 MB | 49 MB | 249 MB | 497 MB |
| jackson-spreadsheet (POI User Model) | 7.1 MB | 55 MB | 272 MB | 543 MB |

jackson writes with 24–27% less allocation than FastExcel from 10K onward; Apache Fesod allocates 3–4× as much.

## Sustained Throughput

10K-row source looped over 60 seconds. Surfaces steady-state allocation rate and GC pause accumulation.

**Read:**

| Library | ops/sec | gc.alloc.rate | gc.count | gc.time |
|---------|---------|---------------|----------|---------|
| jackson-spreadsheet | 35.04 | **1993 MB/sec** | **256** | **241 ms** |
| FastExcel | **39.43** | 2706 MB/sec | 347 | 312 ms |

**Write:**

| Library | ops/sec | gc.alloc.rate | gc.count | gc.time |
|---------|---------|---------------|----------|---------|
| jackson-spreadsheet | 30.18 | **749 MB/sec** | **173** | **84 ms** |
| FastExcel | **35.18** | 1172 MB/sec | 269 | 223 ms |

- Read: FastExcel is ~13% faster but allocates 36% more per second and triggers 36% more GC cycles.
- Write: FastExcel is ~17% faster but allocates 56% more per second and triggers 56% more GC cycles with ~2.7× longer GC time.

## SharedStrings

`FILE_BACKED_SHARED_STRINGS` applies to both read and write paths.

**Read:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | **14.46 ms** | **67.08 ms** | **125.56 ms** |
| FileBacked (H2 MVStore) | 24.40 ms | 106.52 ms | 204.54 ms |
| FileBacked + Encrypted | 29.93 ms | 129.02 ms | 241.48 ms |

**Write:**

| Strategy | 10K rows | 50K rows | 100K rows |
|----------|----------|----------|-----------|
| InMemory (default) | **17.42 ms** | **75.31 ms** | **160.64 ms** |
| FileBacked (H2 MVStore) | 42.01 ms | 200.35 ms | 362.73 ms |
| FileBacked + Encrypted | 50.51 ms | 304.18 ms | 666.64 ms |

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
