# Benchmark

## Environment

- JDK 21.0.10, OpenJDK 64-Bit Server VM on Apple Silicon (aarch64)
- JMH 1.35, Fork 1, Warmup 3 × 1 s, Measurement 5 × 1 s
- GC profiler enabled

### Library Versions

| Library | Version |
|---------|---------|
| jackson-dataformat-spreadsheet | 1.7.0-SNAPSHOT |
| FastExcel | 0.20.0 |
| Apache POI | 5.5.1 |
| Apache Fesod | 2.0.1-incubating |

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

| Library | 1K rows | 10K rows | 50K rows | 100K rows | 500K rows |
|---------|---------|----------|----------|-----------|-----------|
| jackson-spreadsheet | 3.47 ms | 28.93 ms | 141.94 ms | 287.36 ms | 1447.12 ms |
| FastExcel | **2.83 ms** | **26.70 ms** | **134.68 ms** | **264.63 ms** | **1290.10 ms** |
| Apache Fesod | 5.17 ms | 48.61 ms | 240.29 ms | 536.97 ms | 2744.46 ms |
| jackson-spreadsheet (POI User Model) | 21.14 ms | 193.34 ms | 985.82 ms | 2255.86 ms | 11875.93 ms |
| Apache POI | 18.45 ms | 174.12 ms | 892.69 ms | 1997.64 ms | 10714.94 ms |

FastExcel is faster at every scale (5–18%); the gap is narrowest around 50K (5%) and widest at 1K (18%). Apache POI's User Model loads the entire file up front and is 6–8× slower than streaming readers — first-N early exit pays the same cost as a full read. Apache POI also ships a SAX event API (`XSSFReader` + `XSSFSheetXMLHandler`) for streaming reads; that path is not measured here.

A 10-column 100%-string probe (`internal/StringOnlyBenchmark`) reverses the order: jackson reads 12–23% faster with 27–33% less allocation.

### Allocation per op

| Library | 1K rows | 10K rows | 50K rows | 100K rows | 500K rows |
|---------|---------|----------|----------|-----------|-----------|
| jackson-spreadsheet | **6.8 MB** | **62 MB** | **290 MB** | **579 MB** | **2997 MB** |
| FastExcel | 7.8 MB | 71 MB | 370 MB | 722 MB | 3537 MB |
| Apache Fesod | 9.1 MB | 80 MB | 398 MB | 882 MB | 4390 MB |
| Apache POI | 44 MB | 426 MB | 2142 MB | 4287 MB | 21249 MB |
| jackson-spreadsheet (POI User Model) | 47 MB | 459 MB | 2285 MB | 4603 MB | 22824 MB |

jackson allocates 13–22% less than FastExcel across all scales, 7× less than Apache POI. At 500K, POI User Model retains the full workbook in memory (~22 GB allocation per op); heap budget must scale with file size.

## Write

### Throughput

| Library | 1K rows | 10K rows | 50K rows | 100K rows | 500K rows |
|---------|---------|----------|----------|-----------|-----------|
| jackson-spreadsheet | 5.42 ms | 37.94 ms | **166.64 ms** | **327.20 ms** | **1682.09 ms** |
| FastExcel | **3.95 ms** | **30.96 ms** | 172.14 ms | 353.14 ms | 1816.80 ms |
| jackson-spreadsheet (POI User Model) | 9.28 ms | 77.21 ms | 391.74 ms | 786.24 ms | 4008.07 ms |
| Apache POI (SXSSF) | 8.57 ms | 66.63 ms | 328.96 ms | 664.48 ms | 3235.75 ms |
| Apache Fesod | 10.57 ms | 80.73 ms | 395.32 ms | 790.65 ms | 3874.08 ms |

jackson overtakes FastExcel at 50K+ (3% at 50K, 7% at 100K, 7% at 500K). FastExcel leads smaller workloads by 23–37%.

### Allocation per op

| Library | 1K rows | 10K rows | 50K rows | 100K rows | 500K rows |
|---------|---------|----------|----------|-----------|-----------|
| jackson-spreadsheet | 3.8 MB | **26 MB** | **136 MB** | **271 MB** | **1333 MB** |
| FastExcel | **3.7 MB** | 33 MB | 165 MB | 329 MB | 1649 MB |
| jackson-spreadsheet (POI User Model) | 7.6 MB | 58 MB | 297 MB | 593 MB | 2934 MB |
| Apache POI (SXSSF) | 6.8 MB | 51 MB | 260 MB | 520 MB | 2563 MB |
| Apache Fesod | 13 MB | 111 MB | 550 MB | 1096 MB | 5500 MB |

jackson writes with 17–21% less allocation than FastExcel from 10K onward; Apache Fesod allocates 3–4× as much.

## Sustained Throughput

10K-row source looped over 60 seconds. Surfaces steady-state allocation rate and GC pause accumulation.

**Read:**

| Library | ops/sec | gc.alloc.rate | gc.count | gc.time |
|---------|---------|---------------|----------|---------|
| jackson-spreadsheet | 35.06 | **1975 MB/sec** | **253** | **241 ms** |
| FastExcel | **39.68** | 2663 MB/sec | 341 | 302 ms |

**Write:**

| Library | ops/sec | gc.alloc.rate | gc.count | gc.time |
|---------|---------|---------------|----------|---------|
| jackson-spreadsheet | 29.54 | **798 MB/sec** | **184** | **94 ms** |
| FastExcel | **34.96** | 1165 MB/sec | 215 | 190 ms |

- Read: FastExcel is ~13% faster but allocates 35% more per second and triggers 35% more GC cycles.
- Write: FastExcel is ~18% faster but allocates 46% more per second and triggers 17% more GC cycles with 2× longer GC time.

## SharedStrings

`FILE_BACKED_SHARED_STRINGS` applies to both read and write paths.

**Read:**

| Strategy | 10K rows | 50K rows | 100K rows | 500K rows |
|----------|----------|----------|-----------|-----------|
| InMemory (default) | **14.77 ms** | **67.51 ms** | **125.49 ms** | **647.11 ms** |
| FileBacked (H2 MVStore) | 27.93 ms | 112.36 ms | 210.68 ms | 1058.98 ms |
| FileBacked + Encrypted | 30.32 ms | 126.76 ms | 250.05 ms | 1237.06 ms |

**Write:**

| Strategy | 10K rows | 50K rows | 100K rows | 500K rows |
|----------|----------|----------|-----------|-----------|
| InMemory (default) | **17.22 ms** | **75.11 ms** | **154.55 ms** | **795.75 ms** |
| FileBacked (H2 MVStore) | 41.48 ms | 199.13 ms | 403.46 ms | 2003.43 ms |
| FileBacked + Encrypted | 48.49 ms | 322.45 ms | 656.90 ms | 3730.20 ms |

FileBacked trades throughput for **peak heap reduction** — InMemory OOMs when the SST exceeds available heap.

## Running & Reproducing

```bash
./gradlew jmh
```

Filter specific benchmarks:

```bash
./gradlew jmh -Pjmh.includes='\.ReadBenchmark\.'                -Pjmh.resultsFile='build/results/jmh/read.txt'
./gradlew jmh -Pjmh.includes='\.WriteBenchmark\.'               -Pjmh.resultsFile='build/results/jmh/write.txt'
./gradlew jmh -Pjmh.includes='\.SharedStringsBenchmark\.'       -Pjmh.resultsFile='build/results/jmh/sharedstrings.txt'
./gradlew jmh -Pjmh.includes='\.SustainedThroughputBenchmark\.' -Pjmh.resultsFile='build/results/jmh/sustained.txt'
```

Run benchmarks separately to avoid system noise accumulation across categories.
