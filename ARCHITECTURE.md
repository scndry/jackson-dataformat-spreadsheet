# Architecture

A Jackson streaming dataformat module for SpreadsheetML — same pattern as `jackson-dataformat-csv`, applied to spreadsheets.

| jackson-dataformat-csv | jackson-dataformat-spreadsheet |
|---|---|
| `CsvMapper extends ObjectMapper` | `SpreadsheetMapper extends ObjectMapper` |
| `CsvParser extends CsvParserBase extends ParserMinimalBase` | `SheetParser extends ParserMinimalBase` |
| `CsvGenerator extends GeneratorBase` | `SheetGenerator extends GeneratorBase` |
| `CsvFactory extends JsonFactory` | `SpreadsheetFactory extends JsonFactory` |
| `CsvSchema` via `FormatVisitor` | `SpreadsheetSchema` via `FormatVisitor` |
| `CsvParser.Feature implements FormatFeature` | `SheetParser.Feature`, `SpreadsheetFactory.Feature` |
| `CsvMapper.Builder extends MapperBuilder` | `SpreadsheetMapper.Builder extends MapperBuilder` |

```
              ┌─────────────────────────────────────────────────┐
              │               SpreadsheetMapper                 │
              │             (extends ObjectMapper)              │
              └─────────────────┬────────────┬──────────────────┘
                                │            │
                   ┌────────────▼────┐  ┌────▼──────────────┐
                   │   SheetParser   │  │  SheetGenerator   │
                   │(ParserMinimal-  │  │  (GeneratorBase)  │
                   │      Base)      │  │                   │
                   └────────┬────────┘  └─────────┬─────────┘
                            │                     │
                ┌───────────▼────────┐  ┌─────────▼──────────┐
                │    SheetReader     │  │     SheetWriter    │
                │    (interface)     │  │     (interface)    │
                └──┬───────────┬─────┘  └─────┬───────────┬──┘
                   │           │              │           │
         ┌─────────▼──┐ ┌──────▼──────┐┌──────▼─────┐ ┌───▼─────────┐
         │  POISheet  │ │  SSMLSheet  ││  POISheet  │ │  SSMLSheet  │
         │   Reader   │ │   Reader    ││   Writer   │ │    Writer   │
         │   (XLS)    │ │ (XLSX/StAX) ││   (POI)    │ │   (SSML)    │
         └────────────┘ └─────────────┘└────────────┘ └─────────────┘
```

## Scope

This module handles **data binding between POJOs and spreadsheet cells**. Nothing more.

| This module | POI (or user code) |
|---|---|
| POJO → flat cells (write) | Workbook lifecycle, file I/O |
| Flat cells → POJO (read) | Charts, images, pivot tables |
| Nested object flattening/reconstruction | Formulas, conditional formatting |
| Header generation, cell styling | Macros (.xlsm) |
| Schema generation from class structure | Cell-level random access |

`Sheet` is an I/O type — like `OutputStream` in jackson-core. The module reads from and writes to a POI `Sheet` but does not manage the `Workbook` lifecycle. Users open a workbook with POI, pass a `Sheet` to the mapper, and close the workbook themselves. This enables template-based writing, multi-sheet workbooks, and post-processing with POI — all outside the module's concern.

## Design Decisions

### Pull-Pull Streaming

Jackson is a pull parser — the client calls `nextToken()`.
StAX is a pull parser — the client calls `next()`.
This library bridges the two naturally:

```
Jackson (pull)          SheetParser (pull)         StAX (pull)
    │                        │                        │
    ├─ nextToken() ─────────►├─ SheetReader.next() ──►├─ XMLStreamReader.next()
    │◄─ VALUE_STRING ────────┤◄─ CELL_VALUE ──────────┤◄─ START_ELEMENT <c>
    │                        │                        │
```

No buffering, no threading, no adapter — pull aligns with pull.

The alternative is SAX (push model), which [POI's Event API](https://poi.apache.org/components/spreadsheet/how-to.html#event_api) uses.
SAX pushes events to a handler callback, but Jackson pulls tokens from a parser.
Bridging push→pull requires an intermediate buffer or a separate thread — complexity that defeats the purpose of streaming.

This library bypasses POI's Event API entirely and parses OOXML XML directly via StAX.

## Token Model

The core problem: mapping flat `(row, column)` cells to hierarchical Jackson tokens.

### Cell to Token Translation

A spreadsheet with nested POJO schema `Employee { Address address; Employment employment; }`:

```
Column index:    0     1        2             3              4
Column pointer:  /id   /name   /address/zip  /address/city  /employment/title
```

For one row of cell values `[1, "Alice", "12345", "Seoul", "SRE"]`, the parser emits:

```
START_OBJECT
  FIELD_NAME "id"
  VALUE_NUMBER_INT 1
  FIELD_NAME "name"
  VALUE_STRING "Alice"
  FIELD_NAME "address"
  START_OBJECT
    FIELD_NAME "zip"
    VALUE_STRING "12345"
    FIELD_NAME "city"
    VALUE_STRING "Seoul"
  END_OBJECT
  FIELD_NAME "employment"
  START_OBJECT
    FIELD_NAME "title"
    VALUE_STRING "SRE"
  END_OBJECT
END_OBJECT
```

One flat row of cells becomes a nested object tree.
Jackson's data-binding layer deserializes this token stream into `Employee` — no special deserializer needed.

### ColumnPointer

`ColumnPointer` is the mechanism that drives this translation.
Each schema column has a pointer representing its path in the object hierarchy:

```
ColumnPointer.empty().resolve("address").resolve("zip")              → address/zip
ColumnPointer.empty().resolve("address").resolve("zip").getParent()  → address
pointer1.relativize(pointer2)                                        → relative path
```

When the parser moves from one cell to the next, it computes the relative path between the current and previous `ColumnPointer`.
If the parent changes (e.g., `address/city` → `employment/title`), the parser emits `END_OBJECT` for the old scope and `START_OBJECT` + `FIELD_NAME` for the new one.

### SheetToken

The low-level token enum bridging `SheetReader` and `SheetParser`:

```
SHEET_DATA_START  →  START_ARRAY
ROW_START         →  START_OBJECT
CELL_VALUE        →  FIELD_NAME + value token(s)
ROW_END           →  END_OBJECT (× nesting depth)
SHEET_DATA_END    →  END_ARRAY
```

`SheetReader` produces `SheetToken`.
`SheetParser` consumes `SheetToken` and emits `JsonToken` with proper nesting.

## Read Path

```
mapper.readValue(file, Employee.class)
  │
  ├─ schemaGenerator.generate(Employee.class)
  │    └─ Jackson FormatVisitor → annotation introspection → SpreadsheetSchema
  │
  ├─ SpreadsheetFactory.createParser(file)
  │    ├─ USE_POI_USER_MODEL? → POISheetReader (POI object model)
  │    ├─ PackageUtil.isOOXML(file)?
  │    │   ├─ yes → SSMLSheetReader (StAX streaming)
  │    │   └─ no  → POISheetReader  (POI object model)
  │    └─ new SheetParser(reader, formatFeatures)
  │
  └─ SheetParser.nextToken() loop
       ├─ SheetReader.next() → SheetToken
       ├─ row/column bounds check against schema
       ├─ ColumnPointer scope tracking
       ├─ JsonToken emission (with nesting)
       └─ Jackson BeanDeserializer consumes tokens → Employee
```

### Dual Reader Strategy

Both implement `SheetReader`. Format is auto-detected via ZIP magic bytes (`FileMagic.OOXML`).

| | SSMLSheetReader | POISheetReader |
|---|---|---|
| Format | XLSX (OOXML) | XLS / XLSX (any POI-supported) |
| Parsing | StAX `XMLStreamReader` on raw XML | POI `Sheet` / `Row` / `Cell` API |
| Memory | Streaming — constant regardless of file size | Full workbook in memory |
| Cell types | OOXML `STCellType` schema enum | POI `CellType` enum via `CellFormat` |
| Shared strings | Custom `SharedStringLookup` (lazy StAX) | POI's built-in `SharedStringsTable` |
| When used | Auto-detected for XLSX files | XLS files, or direct `Sheet` input |

When the input is an `InputStream`, OOXML files are copied to a temporary file — ZIP random access requires seekable I/O.

### SharedStrings

XLSX files store all string cell values in a shared string table (`xl/sharedStrings.xml`).
POI's `SharedStringsTable` wraps each entry in `XSSFRichTextString` with per-entry object allocation.

This library uses `SharedStringLookup`, a custom interface that returns `String` directly — no `RichTextString` wrapping:

```java
// POI's interface
public interface SharedStrings {
    RichTextString getItemAt(int idx);
}

// This library's interface
public interface SharedStringLookup {
    String getItemAt(int idx);
}
```

**Read** — Two `SharedStringLookup` implementations:

- **InMemorySharedStringLookup** (default) — All character data in a single `char[]` buffer with `int[]` offset/length arrays. Eliminates per-String object overhead. Entries are parsed lazily via StAX — only when first accessed.
- **FileBackedSharedStringLookup** — H2 MVStore with 4 MB page cache and 1024-entry LRU lookup cache. Constant heap usage regardless of table size. Prevents OOM when the SST exceeds available heap. Optional AES encryption protects sensitive data at rest.

**Write** — Two `SharedStringStore` implementations for the SSML writer:

- **InMemorySharedStringStore** (default) — SoA layout: one `char[]` slab with `int[]` offset/length arrays and open-addressing hash index. Lower GC pressure than `HashMap<String, Integer>`.
- **FileBackedSharedStringStore** — H2 MVStore with single-entry cache optimized for sequential access during SST serialization. Constant heap usage for high-cardinality string columns.

Result: lowest memory allocation among all tested libraries at 100K rows. See [BENCHMARK.md](BENCHMARK.md).

## Write Path

```
mapper.writeValue(file, employee)
  │
  ├─ SpreadsheetFactory.createGenerator(file)
  │    ├─ USE_POI_USER_MODEL?
  │    │   ├─ yes → POISheetWriter → SheetGenerator
  │    │   └─ no  → XSSFSheet? → SSMLSheetWriter → SheetGenerator
  │    └─ non-XSSF fallback → POISheetWriter → SheetGenerator
  │
  └─ Jackson BeanSerializer
       ├─ writeStartObject()
       │    └─ SheetGenerator creates scope context
       ├─ writeFieldName("address")
       │    └─ stored in context for ColumnPointer resolution
       ├─ writeString("12345")
       │    └─ SheetGenerator.verifyValueWrite()
       │         └─ SheetStreamContext.currentReference() → CellAddress
       │              └─ SheetWriter.setReference() + writeString()
       ├─ writeEndObject()
       │    └─ mergeScopedColumns(pointer, row, size)
       └─ close()
            └─ SheetWriter.write() + close()
```

### Dual Writer Strategy

| | SSMLSheetWriter (default) | POISheetWriter (`USE_POI_USER_MODEL`) |
|---|---|---|
| Packaging | POI skeleton + ZipOutputStream | POI `Workbook.write()` |
| Cell writing | StringBuilder → ZipOutputStream | POI `Cell.setCellValue()` |
| Shared strings | `SharedStringStore` (in-memory or file-backed) | POI managed |
| Styles | POI `CellStyle.getIndex()` from skeleton | POI `CellStyle` API |
| Performance | ~143 ms / 100K rows | ~323 ms / 100K rows |
| Format | XLSX only (requires XSSFWorkbook) | XLSX, XLS |

SSMLSheetWriter creates a POI XSSFWorkbook with styles and an empty sheet, saves it to a temp file as a complete XLSX skeleton. It then opens the skeleton as a zip, copies all entries except `sheet1.xml` and `sharedStrings.xml`, and writes those two parts via StringBuilder streaming. POI guarantees the metadata (styles, rels, content types, core properties); StringBuilder provides the throughput.

## Schema Generation

Uses Jackson's `JsonFormatVisitor` — the same mechanism Jackson uses internally for JSON Schema generation:

```
schemaGenerator.generate(Employee.class)
  │
  ├─ serializerProvider.acceptJsonFormatVisitor(type, visitor)
  │    │
  │    └─ FormatVisitorWrapper
  │         ├─ expectObjectFormat() → ObjectFormatVisitor
  │         │    ├─ property("id", IntSerializer)
  │         │    │    └─ Column(pointer="/id", type=int)
  │         │    ├─ property("address", BeanSerializer<Address>)
  │         │    │    └─ recursive visit with pointer="/address"
  │         │    │         ├─ property("zip", StringSerializer)
  │         │    │         │    └─ Column(pointer="/address/zip", type=String)
  │         │    │         └─ property("city", StringSerializer)
  │         │    │              └─ Column(pointer="/address/city", type=String)
  │         │    └─ ...
  │         └─ expectArrayFormat() → ArrayFormatVisitor
  │              └─ recursive visit for element type
  │
  └─ SpreadsheetSchema(columns, origin, useHeader, styles)
```

## POI Boundary

The read path and the write path have different relationships with POI:

| | POI | This library |
|---|---|---|
| **XLSX read (default)** | ZIP package open (`OPCPackage`) | XML parsing (StAX), SharedStrings, cell type resolution |
| **XLSX read (`USE_POI_USER_MODEL`)** | Everything (`XSSFWorkbook`, `Sheet`, `Row`, `Cell`) | Token translation only |
| **XLS read** | Everything (`HSSFWorkbook`, `Sheet`, `Row`, `Cell`) | Token translation only |
| **XLSX write (default)** | Skeleton generation (`XSSFWorkbook.write()`) | StringBuilder streaming for worksheet + SST, zip entry copy for metadata |
| **XLSX write (`USE_POI_USER_MODEL`)** | WorkbookProvider (default: `XSSFWorkbook`), `CellStyle`, `Font` | Schema-driven cell routing, merge logic |
| **Styling** | `CellStyle` / `Font` API | `StylesBuilder` (fluent builder layer) |

By default, both XLSX read and write paths bypass POI's cell model — the read path uses direct StAX parsing, the write path uses StringBuilder streaming with a POI skeleton for package metadata. `USE_POI_USER_MODEL` reverts both paths to full POI cell-model processing.

`POICompat` provides a reflection-based shim for APIs that differ across POI versions (`Date1904Support.isDate1904()` in POI 4.1.1+, `OPCPackage.isStrictOoxmlFormat()` in POI 5.1.0+). Methods are looked up at class load time; missing methods fall back to safe defaults.

---

Public API surface: `SpreadsheetMapper`, `SheetInput`, `SheetOutput`, `@DataGrid`, `@DataColumn`, `SpreadsheetSchema`, `SheetParser.Feature`, `SpreadsheetFactory.Feature`.

Everything under `poi/` is implementation detail — swappable without affecting the streaming contract.
