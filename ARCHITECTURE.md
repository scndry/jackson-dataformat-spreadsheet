# Architecture

A Jackson streaming dataformat module for SpreadsheetML — same pattern as `jackson-dataformat-csv`, applied to spreadsheets.

| jackson-dataformat-csv | jackson-dataformat-spreadsheet |
|---|---|
| `CsvMapper extends ObjectMapper` | `SpreadsheetMapper extends ObjectMapper` |
| `CsvParser extends ParserMinimalBase` | `SheetParser extends ParserMinimalBase` |
| `CsvGenerator extends GeneratorBase` | `SheetGenerator extends GeneratorBase` |
| `CsvFactory extends JsonFactory` | `SpreadsheetFactory extends JsonFactory` |
| `CsvSchema` via `FormatVisitor` | `SpreadsheetSchema` via `FormatVisitor` |
| `CsvParser.Feature implements FormatFeature` | `SheetParser.Feature`, `SpreadsheetFactory.Feature` |
| `CsvMapper.Builder extends MapperBuilder` | `SpreadsheetMapper.Builder extends MapperBuilder` |

```
              ┌────────────────────────────────────────────────┐
              │               SpreadsheetMapper                │
              │             (extends ObjectMapper)             │
              └─────────────────┬────────────┬─────────────────┘
                                │            │
                 ┌──────────────▼────┐  ┌────▼──────────────┐
                 │    SheetParser    │  │  SheetGenerator   │
                 │  (ParserMinimal-  │  │  (GeneratorBase)  │
                 │       Base)       │  │                   │
                 └──────────┬────────┘  └─────────┬─────────┘
                            │                     │
                ┌───────────▼────────┐  ┌─────────▼──────────┐
                │    SheetReader     │  │    SheetWriter     │
                │    (interface)     │  │    (interface)     │
                └──┬───────────┬─────┘  └─────┬───────────┬──┘
                   │           │              │           │
         ┌─────────▼──┐ ┌──────▼──────┐┌──────▼─────┐ ┌───▼─────────┐
         │  POISheet  │ │  SSMLSheet  ││  POISheet  │ │  SSMLSheet  │
         │   Reader   │ │   Reader    ││   Writer   │ │   Writer    │
         │   (POI)    │ │ (Streaming) ││   (POI)    │ │ (Streaming) │
         └────────────┘ └─────────────┘└────────────┘ └─────────────┘
```

## Jackson's Three Layers

Jackson processes data in three layers. This library plugs into all three:

| Layer | Jackson | This library | Role |
|-------|---------|-------------|------|
| **Data Binding** | ObjectMapper, ObjectReader, ObjectWriter | SpreadsheetMapper, SpreadsheetReader, SpreadsheetWriter | POJO serialization — **inherited, not reimplemented** |
| **Streaming** | JsonParser, JsonGenerator, JsonFactory | SheetParser, SheetGenerator, SpreadsheetFactory | Flat cell ↔ nested token translation — **the core abstraction** |
| **Format I/O** | UTF8StreamJsonParser reads bytes | SheetReader / SheetWriter (interface) | Raw format read/write — **swappable** (Streaming or POI User Model) |

Layer 3 is fully inherited — all Jackson annotations, overloads, and features work as-is. Layer 2 is where the library's core logic lives. Layer 1 is pluggable behind an interface.

## Scope

This library handles **data binding between POJOs and spreadsheet cells**. Nothing more.

| This library | POI (or user code) |
|---|---|
| POJO → flat cells (write) | Workbook lifecycle, file I/O |
| Flat cells → POJO (read) | Charts, images, pivot tables |
| Nested object flattening/reconstruction | Formulas, conditional formatting |
| Header generation, cell styling | Macros (.xlsm) |
| Schema generation from class structure | Cell-level random access |

`Sheet` is an I/O type — like `OutputStream` in jackson-core. The library reads from and writes to a POI `Sheet` but does not manage the `Workbook` lifecycle. Users open a workbook with POI, pass a `Sheet` to the mapper, and close the workbook themselves. This enables template-based writing, multi-sheet workbooks, and post-processing with POI — all outside the library's concern.

## Token Model

The core problem: mapping flat `(row, column)` cells to hierarchical Jackson tokens.

### Cell to Token Translation

A spreadsheet with nested POJO schema `Employee { Address address; Employment employment; }`:

```
Column index:    0     1        2             3              4
Column pointer:  /id   /name   /address/zip  /address/city  /employment/title
```

The mapping runs both ways — flat cells ↔ nested Jackson tokens. For one row of cell values `[1, "Alice", "12345", "Seoul", "SRE"]`, the parser emits:

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

Jackson's data-binding layer deserializes this token stream into `Employee` (read), or serializes `Employee` into the same stream that the generator consumes (write) — no special deserializer/serializer needed.

### ColumnPointer

`ColumnPointer` is the mechanism that drives this translation — shared by both the parser and the generator.
Each schema column has a pointer representing its path in the object hierarchy:

```
ColumnPointer.empty().resolve("address").resolve("zip")              → address/zip
ColumnPointer.empty().resolve("address").resolve("zip").getParent()  → address
pointer1.relativize(pointer2)                                        → relative path
```

When the parser or generator moves between cells, it computes the relative path between the current and previous `ColumnPointer`.
If the parent changes (e.g., `address/city` → `employment/title`), the parser emits `END_OBJECT` for the old scope and `START_OBJECT` + `FIELD_NAME` for the new one; the generator does the reverse, consuming `END_OBJECT` / `START_OBJECT` boundaries to switch scope.

`ColumnPointer` is custom rather than reusing Jackson's `JsonPointer` because spreadsheet columns flatten arrays in place — `items[0]/qty` and `items[1]/qty` map to the same column, not distinct paths. `ColumnPointer.resolveArray()` collapses the index, while `JsonPointer` would preserve it.

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
  └─ SpreadsheetSchema(columns, ...)
```

## Read Path

```
mapper.readValue(file, Employee.class)
  │
  ├─ schemaGenerator.generate(Employee.class)
  │    └─ Jackson FormatVisitor → annotation introspection → SpreadsheetSchema
  │
  ├─ SpreadsheetFactory.createParser(file)
  │    ├─ USE_POI_USER_MODEL OR non-OOXML(input)?
  │    │   ├─ yes → POISheetReader (POI User Model)
  │    │   └─ no  → SSMLSheetReader (Streaming)
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
| Memory | Streaming — cell data not held in memory (SST grows with unique strings) | Full workbook in memory |
| Cell types | OOXML `STCellType` schema enum | POI `CellType` enum |
| Shared strings | Custom `SharedStringsLookup` (lazy StAX) | POI's built-in `SharedStringsTable` |
| When used | Auto-detected for XLSX files | XLS files, direct `Sheet` input, or `USE_POI_USER_MODEL` |

When the input is an `InputStream`, OOXML files are copied to a temporary file — ZIP random access requires seekable I/O. The POI User Model path bypasses this copy: the stream is handed directly to POI, which decompresses the entire ZIP into memory (OOM risk on large files). Setting `USE_POI_USER_MODEL` forces this path even for OOXML input — the escape hatch for disk-write-restricted environments.

### Pull-Pull Streaming

The streaming read path bridges two pull APIs — Jackson's `nextToken()` and StAX's `next()` — via `SheetParser`. The bridge token (`SheetToken`) carries flat sheet events (`SHEET_DATA_START`, `ROW_START`, `CELL_VALUE`, `ROW_END`, `SHEET_DATA_END`) which `SheetParser` translates into nested `JsonToken` for Jackson:

```
Jackson (pull)          SheetParser (pull)         StAX (pull)
    │                        │                        │
    ├─ nextToken() ─────────►├─ SheetReader.next() ──►├─ XMLStreamReader.next()
    │◄─ VALUE_STRING ────────┤◄─ CELL_VALUE ──────────┤◄─ START_ELEMENT <c>
    │                        │                        │
```

No buffering, no threading, no adapter — pull aligns with pull.

The alternative is SAX (push model), which [POI's Event API](https://poi.apache.org/components/spreadsheet/how-to.html#event_api) uses. SAX pushes events to a handler callback, but Jackson pulls tokens from a parser. Bridging push→pull requires an intermediate buffer or a separate thread — complexity that defeats the purpose of streaming. This library bypasses POI's Event API entirely and parses OOXML XML directly via StAX.

## Write Path

```
mapper.writeValue(file, employee)
  │
  ├─ SpreadsheetFactory.createGenerator(file)
  │    ├─ USE_POI_USER_MODEL OR non-XSSF(sheet)?
  │    │   ├─ yes → POISheetWriter (POI User Model)
  │    │   └─ no  → SSMLSheetWriter (Streaming)
  │    └─ new SheetGenerator(writer)
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
       │    └─ mergeScopedColumns(...)
       └─ close()
            └─ SheetWriter.write() + close()
```

### Dual Writer Strategy

| | SSMLSheetWriter | POISheetWriter |
|---|---|---|
| Packaging | POI scaffold + ZipOutputStream | POI `Workbook.write()` |
| Cell writing | SoA buffer → StringBuilder → ZipOutputStream | POI `Cell.setCellValue()` |
| Shared strings | `SharedStringsStore` (in-memory or file-backed) | POI managed |
| Styles | POI `CellStyle.getIndex()` from scaffold | POI `CellStyle` API |
| Format | XLSX only (requires XSSFWorkbook) | XLSX, XLS |
| When used | OOXML output + XSSFSheet workbook | Direct `Sheet` output, non-XSSF custom workbook, or `USE_POI_USER_MODEL` |

### Scaffold-Based Streaming

The split **separates correctness from throughput**. The throughput bottleneck is `sheetData` + `sharedStrings.xml` — one entry per cell, scaling with row count. Everything else (relationships, content types, core properties, theme, drawing rels, namespace declarations, even `mergeCells`/`autoFilter`/`conditionalFormatting`) is bounded in size and OOXML spec-sensitive. Hand-rolling those in StringBuilder is not a throughput question; it's a compatibility risk.

`SSMLSheetWriter` lets POI generate a complete XLSX scaffold with the target schema's styles and an empty sheet, then patches only the two streaming entries:

```
setSchema()
  ├─ XSSFWorkbook (styles + empty sheet) → temp file
  ├─ schema.configureSheet() applies sheet-level features to the scaffold
  └─ split sheet1.xml at <sheetData>:
        head, tail   (POI-generated; copied verbatim)
        data         (<row> entries — streamed at write time)

close()
  ├─ for each zip entry:
  │     sheet1.xml        → head + streamed rows + tail
  │     sharedStrings.xml → SharedStringsStore-driven streaming
  │     others            → copied as-is
  └─ delete temp file
```

POI owns OOXML correctness. A row-linked SoA cell buffer feeds the StringBuilder per fragment, owning per-cell throughput. The scaffold is the handoff between them.

## SharedStrings

XLSX SST is the memory bottleneck for string-heavy workloads. The streaming reader uses `SharedStringsLookup` (returns `String` directly, no `RichTextString` wrapping); the streaming writer uses `SharedStringsStore`. Both default to SoA buffers.

`FILE_BACKED_SHARED_STRINGS` switches both to H2 MVStore — constant heap on tables exceeding available heap. `ENCRYPT_FILE_BACKED_STORE` adds AES encryption for sensitive data at rest.

## Nested List

A schema with `List<T>` (or array) fields maps N rows to one outer record on read. `@DataColumn(anchor = true)` marks the outer-record boundary; the parser buffers cells into a record tree until an anchor change closes the record, and a depth-first walk emits `{outer, child1[], child2[], ...}`, so sibling lists, outer-after-list, and N-depth nesting fall out of the tree.

On write, the same schema emits one outer record across N rows. When the schema has an outer field declared after a nested list or multiple sibling lists at the same scope, outer fields back-write into past rows — the library streaming writer keeps cell metadata in an SoA cell buffer for O(1) back-write; the POI path uses random Cell access.

## POI Boundary

The read path and the write path have different relationships with POI:

| | POI | This library |
|---|---|---|
| **XLSX read (default)** | ZIP package open (`OPCPackage`) | XML parsing (StAX), SharedStrings, cell type resolution |
| **Non-streaming read** | Everything (`Workbook`, `Sheet`, `Row`, `Cell`) | Token translation only |
| **XLSX write (default)** | Scaffold generation (`XSSFWorkbook.write()`) | StringBuilder streaming for worksheet + SST, zip entry copy for metadata |
| **Non-streaming write** | `Workbook.write()`, `CellStyle`, `Font` | Schema-driven cell routing, merge logic |
| **Styling** | `CellStyle` / `Font` API | `StylesBuilder` (fluent builder layer) |
| **Sheet-level features** | `Sheet.createFreezePane`, `setAutoFilter`, `ConditionalFormatting` API | `GridConfigurer` (fluent builder layer) |

By default, both XLSX read and write paths bypass POI's User Model — the read path uses direct StAX parsing, the write path uses StringBuilder streaming with a POI scaffold for package metadata. Non-streaming paths (XLS, or XLSX with `USE_POI_USER_MODEL`) use POI's full User Model.

`POICompat` absorbs POI version differences internally, letting the library support POI 4.1.1+.

---

Public API surface (types users typically import): `SpreadsheetMapper`, `SpreadsheetReader`, `SpreadsheetWriter`, `@DataGrid`, `@DataColumn`, `@DataColumnGroup`, `SheetInput`, `SheetOutput`, `SpreadsheetSchema`, `StylesBuilder`, `GridConfigurer`, `SheetMappingIterator`, `SheetLocation`, `SheetParser.Feature`, `SpreadsheetFactory.Feature`.

Everything under `poi/` is implementation detail — swappable without affecting the streaming contract.
