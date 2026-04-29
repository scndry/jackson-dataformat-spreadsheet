# Usage Guide

**jackson-dataformat-spreadsheet** is a Jackson streaming dataformat module for reading and writing Excel spreadsheets (XLSX/XLS) as POJOs. It extends Jackson's `ObjectMapper` — the same API you already use for JSON, CSV, and XML.

```java
// JSON
ObjectMapper json = new ObjectMapper();
Employee e = json.readValue(jsonFile, Employee.class);

// Excel — same pattern
SpreadsheetMapper mapper = new SpreadsheetMapper();
Employee e = mapper.readValue(xlsxFile, Employee.class);
```

No cell-level iteration. No column index counting. No manual type casting.

> For runnable code examples, see [jackson-spreadsheet-examples](https://github.com/scndry/jackson-spreadsheet-examples).

## How It Differs from Apache POI

Apache POI gives you cells. This library gives you POJOs.

```java
// Apache POI — manual cell-by-cell extraction
try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
    Sheet sheet = wb.getSheetAt(0);
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
        Row row = sheet.getRow(i);
        Product p = new Product();
        p.name = row.getCell(0).getStringCellValue();
        p.category = row.getCell(1).getStringCellValue();
        p.quantity = (int) row.getCell(2).getNumericCellValue();
        p.price = row.getCell(3).getNumericCellValue();
        p.inStock = row.getCell(4).getBooleanCellValue();
        products.add(p);
    }
}

// jackson-spreadsheet — one line
List<Product> products = mapper.readValues(file, Product.class);
```

The column-to-field mapping is driven by the class structure. Add a field, the column appears. Rename a field, the column follows. Remove a field, the column disappears. No index to update.

But this is not a POI replacement. POI types (`Sheet`, `Workbook`) are first-class I/O targets. Open a workbook with POI, pass a `Sheet` to the mapper for data binding, then continue using POI for charts, formulas, conditional formatting — whatever the mapper doesn't cover. Data binding on top of POI, not instead of it.

## Requirements

- Java 8+
- Jackson 2.14.0+
- Apache POI 4.1.1+ (Strict OOXML requires 5.1.0+)

## Installation

**Maven:**
```xml
<dependency>
    <groupId>io.github.scndry</groupId>
    <artifactId>jackson-dataformat-spreadsheet</artifactId>
    <version>1.4.0</version>
</dependency>
```

**Gradle:**
```groovy
implementation "io.github.scndry:jackson-dataformat-spreadsheet:1.4.0"
```

## Quick Start

```java
import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;

import java.io.File;
import java.util.Arrays;
import java.util.List;

public class QuickStart {

    @DataGrid
    public static class Product {
        @DataColumn("Name")
        public String name;
        public int quantity;
        public double price;

        public Product() {}
        public Product(String name, int quantity, double price) {
            this.name = name;
            this.quantity = quantity;
            this.price = price;
        }
    }

    public static void main(String[] args) throws Exception {
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        // Write
        List<Product> products = Arrays.asList(
            new Product("Apple", 10, 1.50),
            new Product("Banana", 20, 0.80));
        mapper.writeValue(new File("products.xlsx"), products, Product.class);

        // Read
        List<Product> loaded = mapper.readValues(
            new File("products.xlsx"), Product.class);
    }
}
```

Output (`products.xlsx`):

| Name | quantity | price |
|------|----------|-------|
| Apple | 10 | 1.5 |
| Banana | 20 | 0.8 |

## Supported I/O Types

| Direction | Supported types |
|-----------|----------------|
| **Read** | `File`, `Path`, `InputStream`, `SheetInput<T>`, `Sheet` |
| **Write** | `File`, `Path`, `OutputStream`, `SheetOutput<T>`, `Sheet` |

`File`, `Path`, `InputStream`, `OutputStream` are for the default (first) sheet. `Path` delegates to `File` internally.

`SheetInput` selects a sheet to read by name or index. `SheetOutput` names the sheet to write.

`Sheet` is a POI type for direct workbook control — multi-sheet writes, formula post-processing, template population.

## Reading

Read an Excel file into Java POJOs:

```java
SpreadsheetMapper mapper = new SpreadsheetMapper();

// Single row (first data row)
Product p = mapper.readValue(file, Product.class);

// All rows
List<Product> list = mapper.readValues(file, Product.class);

// From InputStream
List<Product> list = mapper.readValues(stream, Product.class);
```

Use `SheetInput` when you need to select a specific sheet by name or index:

```java
// Specific sheet by name
List<Product> list = mapper.readValues(SheetInput.source(file, "Products"), Product.class);

// Specific sheet by index (0-based)
List<Product> list = mapper.readValues(SheetInput.source(file, 0), Product.class);
```

By default, columns are matched by position — the spreadsheet's column order must match the field declaration order. Enable `columnReordering(true)` to match by header name instead (see [Column Reordering](#column-reordering)).

### Streaming Read

For large files, process rows one at a time with constant memory:

```java
SpreadsheetReader reader = mapper.sheetReaderFor(Product.class);
try (SheetMappingIterator<Product> iter = reader.readValues(file)) {
    while (iter.hasNext()) {
        Product p = iter.next();
    }
}
```

`readValues(file, type)` loads all rows into a `List`. For large files or row-level processing, prefer `SheetMappingIterator`.

#### Error Handling with Location

`SheetLocation.of(exception)` extracts the row and column where the error occurred, or returns `null` if unavailable:

```java
try (SheetMappingIterator<Product> iter = reader.readValues(file)) {
    while (iter.hasNext()) {
        try {
            Product p = iter.next();
            validate(p);
        } catch (Exception e) {
            SheetLocation loc = SheetLocation.of(e);
            if (loc != null) {
                log.warn("Row {}, Col {}: {}", loc.getRow(), loc.getColumn(), e.getMessage());
            }
            // skip and continue
        }
    }
}
```

#### Batch Processing

Collect rows in batches for bulk database inserts:

```java
List<Product> batch = new ArrayList<>(1000);
try (SheetMappingIterator<Product> iter = reader.readValues(file)) {
    while (iter.hasNext()) {
        batch.add(iter.next());
        if (batch.size() >= 1000) {
            process(batch);
            batch.clear();
        }
    }
}
if (!batch.isEmpty()) {
    process(batch);
}
```

## Writing

```java
SpreadsheetMapper mapper = new SpreadsheetMapper();

// Single object
mapper.writeValue(file, product);

// Collection — element type required (Java type erasure)
mapper.writeValue(file, products, Product.class);

// To OutputStream
mapper.writeValue(outputStream, products, Product.class);

// To byte array (in-memory Excel generation)
byte[] bytes = mapper.writeValueAsBytes(products, Product.class);
```

Use `SheetOutput` when you need to specify a sheet name:

```java
mapper.writeValue(SheetOutput.target(file, "Products"), products, Product.class);
```

### Streaming (Default)

By default, XLSX read/write uses streaming — bypassing POI's User Model for direct XML generation. No configuration needed:

```java
SpreadsheetMapper mapper = new SpreadsheetMapper();
mapper.writeValue(file, products, Product.class);  // streaming writer
List<Product> list = mapper.readValues(file, Product.class);  // streaming reader
```

For large files with high-cardinality string columns, file-backed shared strings keep heap usage constant:

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
    .build();
```

Streaming write path limitations (use `USE_POI_USER_MODEL` when these are needed):

- **Auto-size columns** — not supported. Use `@DataColumn(width = N)` for fixed widths.
- **XLS format** — Streaming is XLSX only. XLS automatically uses POI regardless of this setting.
- **Direct Sheet/Workbook access** — `createGenerator(Sheet)` always uses POI. Streaming applies only to File/OutputStream targets.

To fall back to POI's User Model:

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
    .build();
```

### Streaming Write

Write rows one at a time using Jackson's `SequenceWriter`:

```java
SpreadsheetWriter writer = mapper.sheetWriterFor(Product.class);
try (SequenceWriter seq = writer.writeValues(file)) {
    for (Product p : products) {
        seq.write(p);
    }
}
```

### POI Integration

POI types (`Sheet`, `Workbook`) are first-class I/O targets. The mapper reads from and writes to a POI `Sheet` directly — you control the workbook lifecycle.

**Multiple sheets in one workbook:**

```java
try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
    mapper.writeValue(wb.createSheet("Products"), products, Product.class);
    mapper.writeValue(wb.createSheet("Orders"), orders, Order.class);

    try (OutputStream out = new FileOutputStream(file)) {
        wb.write(out);
    }
}
```

**Template-based writing** — open an existing Excel template with pre-formatted headers, charts, formulas, and conditional formatting. Write only the data area; everything else is preserved:

```java
// template.xlsx: styled headers, charts, validation rules already in place
try (XSSFWorkbook wb = new XSSFWorkbook(new File("template.xlsx"))) {
    Sheet sheet = wb.getSheet("Report");

    // Mapper writes data starting at the configured origin
    // Template formatting outside the data area is untouched
    SpreadsheetMapper mapper = SpreadsheetMapper.builder()
        .origin("A2")      // header row is already in the template
        .useHeader(false)   // don't overwrite the template's header
        .build();
    mapper.writeValue(sheet, records, Record.class);

    try (OutputStream out = new FileOutputStream(new File("output.xlsx"))) {
        wb.write(out);
    }
}
```

**Reading from an open workbook:**

```java
try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
    List<Product> products = mapper.readValues(wb.getSheet("Products"), Product.class);
    List<Order> orders = mapper.readValues(wb.getSheet("Orders"), Order.class);
}
```

**Post-processing with POI** — data binding first, then POI for anything the mapper doesn't cover:

```java
try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
    Sheet sheet = wb.createSheet("Sales");
    mapper.writeValue(sheet, salesData, SalesRow.class);

    // Add a SUM formula below the data area with POI
    int lastRow = sheet.getLastRowNum();
    Row totalRow = sheet.createRow(lastRow + 1);
    totalRow.createCell(2).setCellFormula(
        "SUM(C2:C" + (lastRow + 1) + ")");

    try (OutputStream out = new FileOutputStream(file)) {
        wb.write(out);
    }
}
```

## Nested Objects

Flat spreadsheets map to nested POJOs automatically. This is a unique feature — no other Excel library supports bidirectional nested object mapping.

| id | name | zipcode | city | title | salary |
|----|------|---------|------|-------|--------|
| 1 | Alice | 12345 | Seoul | SRE | 80000 |

```java
@DataGrid
class Employee {
    int id;
    String name;
    Address address;       // flattened: zipcode, city
    Employment employment; // flattened: title, salary
}

class Address {
    String zipcode;
    String city;
}

class Employment {
    String title;
    long salary;
}
```

```java
// Flat spreadsheet -> nested POJO
Employee e = mapper.readValue(file, Employee.class);
// Employee { id=1, name="Alice",
//   address=Address { zipcode="12345", city="Seoul" },
//   employment=Employment { title="SRE", salary=80000 } }

// Nested POJO -> flat spreadsheet
mapper.writeValue(file, employees, Employee.class);
```

The schema is generated from the class hierarchy. Each nested object's fields become contiguous columns. No manual column index mapping needed.

By default, nested field headers use the path from the parent (e.g., `address/zipcode`). Use `@DataColumn("Zipcode")` to set a custom header name.

### Nested Lists

Nested lists are serialized into multi-row output. *Deserialization is not currently supported.*

```java
@DataGrid
class Order {
    @DataColumn(value = "Order ID", merge = OptBoolean.TRUE)
    int orderId;
    List<Item> items;
    @DataColumn(value = "Total", merge = OptBoolean.TRUE)
    int total;
}

class Item {
    @DataColumn("Product") String product;
    @DataColumn("Qty") int qty;
}
```

```java
Order order = new Order(1, Arrays.asList(
    new Item("Apple", 3),
    new Item("Banana", 5)), 8);
mapper.writeValue(output, order);
```

```
+----------+---------+-----+-------+
| Order ID | Product | Qty | Total |
+----------+---------+-----+-------+
|          | Apple   |   3 |       |
|    1     +---------+-----+   8   |
|          | Banana  |   5 |       |
+----------+---------+-----+-------+
```

The [SXSSFWorkbook row access window](https://poi.apache.org/apidocs/4.1/org/apache/poi/xssf/streaming/SXSSFWorkbook.html#SXSSFWorkbook-int-) is adjusted automatically to fit the list size. If you see `"Cannot write to row N (already flushed)"`, the nested list expanded beyond the window. Increase it:

```java
SpreadsheetMapper mapper = new SpreadsheetMapper(
    new SpreadsheetFactory(() -> new SXSSFWorkbook(500),
        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS));
```

## Annotations

### @DataGrid

Marks the root type for schema generation. Attributes set class-level defaults for all columns.

| Attribute | Default | Description |
|-----------|---------|-------------|
| `columnStyle` | `""` | Default cell style for data cells |
| `columnHeaderStyle` | `""` | Default cell style for header cells |
| `columnWidth` | `-1` (auto) | Default column width in character units |
| `autoSizeColumn` | `DEFAULT` | Auto-size columns to fit content |
| `minColumnWidth` | `-1` (none) | Minimum column width |
| `maxColumnWidth` | `255` | Maximum column width |
| `mergeColumn` | `DEFAULT` | Merge cells vertically for repeated values |

### @DataColumn

Customizes individual column properties. Unset attributes inherit from the enclosing `@DataGrid`.

| Attribute | Default | Description |
|-----------|---------|-------------|
| `value` | field name | Column header name |
| `comment` | `""` | Header cell comment text |
| `style` | `""` | Cell style for data cells |
| `headerStyle` | `""` | Cell style for the header cell |
| `width` | `-1` (auto) | Column width in character units |
| `autoSize` | `DEFAULT` | Auto-size to fit content |
| `minWidth` | `-1` (none) | Minimum column width |
| `maxWidth` | `255` | Maximum column width |
| `merge` | `DEFAULT` | Merge cells vertically |

`autoSize` may not be accurate for [fullwidth forms](https://en.wikipedia.org/wiki/Halfwidth_and_fullwidth_forms) like [CJK characters](https://en.wikipedia.org/wiki/CJK_characters).

### Attribute Resolution Order

Column attributes resolve in priority order:

1. `@DataColumn` on the property
2. `@DataGrid` on the declaring class
3. `@DataGrid` on the enclosing (parent) class

```java
@DataGrid(autoSizeColumn = OptBoolean.FALSE)
class Foo {
    @DataColumn(style = "highlight") int a;
    int b;
}

@DataGrid(columnStyle = "base")
class Bar {
    Foo foo;
}
```

| Column | style | autoSize |
|--------|-------|----------|
| bar.foo.a | `"highlight"` (from @DataColumn) | `FALSE` (from Foo's @DataGrid) |
| bar.foo.b | `"base"` (from Bar's @DataGrid) | `FALSE` (from Foo's @DataGrid) |

### Jackson Annotations

`SpreadsheetMapper` extends `ObjectMapper`. Standard Jackson annotations work on `@DataGrid` classes.

**Supported:**

| Annotation | Read | Write | Effect |
|---|:---:|:---:|---|
| `@JsonProperty` | Yes | Yes | Column name |
| `@JsonIgnore` | Yes | Yes | Exclude column |
| `@JsonIgnoreProperties` | Yes | Yes | Class-level multi-field exclusion |
| `@JsonCreator` | Yes | — | Constructor-based deserialization |
| `@JsonPropertyOrder` | Yes | Yes | Column order |
| `@JsonInclude(NON_NULL)` | — | Yes | Skip null cells |
| `@JsonNaming` | Yes | Yes | Naming strategy (SNAKE_CASE, etc.) |
| `@JsonGetter` / `@JsonSetter` | Yes | Yes | Custom accessor names |
| `@JsonAutoDetect` | Yes | Yes | Private field access |
| `@JsonValue` / `@JsonCreator` on enum | Yes | Yes | Custom enum cell values |
| `@JsonEnumDefaultValue` | Yes | — | Unknown enum fallback |
| `@JsonSerialize` / `@JsonDeserialize` | Yes | Yes | Custom type conversion |
| `@JsonFormat(shape = STRING)` | — | Yes | Force string cell for numeric types |
| `@JsonUnwrapped` | Yes | Yes | Flatten nested object — headers use leaf name (`x`) instead of path (`inner/x`); supports `prefix`/`suffix` for collision avoidance |
| `@JsonIncludeProperties` | Yes | Yes | Whitelist fields |
| `@JsonFilter` | — | Yes | Programmatic column filtering |
| `@JsonAlias` | Yes | — | Alternative header names for reading (requires `columnReordering(true)`) |
| `@JsonView` | — | Yes | View-based column filtering (via `sheetWriterForWithView(type, view)`) |
| `@JsonTypeInfo` + `@JsonSubTypes` | — | Yes | Polymorphic types (`As.PROPERTY`, union schema) |
| Mix-in | Yes | Yes | Apply `@DataGrid` + annotations to third-party classes |

**Limited:**

| Annotation | Reason |
|---|---|
| `@JsonAnySetter` / `@JsonAnyGetter` | Dynamic properties cannot map to a fixed schema |

**`@JsonView` usage:**

```java
// sheetWriterForWithView(type, view) generates view-filtered schema
mapper.sheetWriterFor(Report.class, Views.Summary.class)
    .writeValue(file, reports);
```

Note: `mapper.writerWithView()` does not work — it bypasses schema generation. Use `sheetWriterForWithView(type, view)` instead.

## Styling

Register named cell styles with `StylesBuilder` and reference them from annotations:

```java
StylesBuilder styles = new StylesBuilder()
    .cellStyle("currency")
        .dataFormat("#,##0.00")
        .font().bold().end()
        .end()
    .cellStyle("header")
        .fillForegroundColor(IndexedColors.GREY_25_PERCENT)
        .fillPattern().solidForeground()
        .font().bold().end()
        .border().thin()
        .end()
    .cellStyle("customHeader", /*cloneStyleFrom*/ "header")
        .fillForegroundColor(198, 239, 206)
        .end();

SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .stylesBuilder(styles)
    .build();
```

```java
@DataGrid(columnHeaderStyle = "header")
class Invoice {
    @DataColumn(value = "Amount", style = "currency")
    double amount;
}
```

### Built-in Data Format Shortcuts

| Method | Pattern |
|--------|---------|
| `general()` | `"General"` |
| `text()` | `"@"` |
| `numberInt()` | `"0"` |
| `numberFloat()` | `"0.00"` |
| `numberIntWithComma()` | `"#,##0"` |
| `numberFloatWithComma()` | `"#,##0.00"` |
| `date()` | `"yyyy-mm-dd"` |
| `dateTime()` | `"yyyy-mm-dd hh:mm:ss"` |

Custom patterns can be passed directly via `dataFormat(String)`. See [Number format codes](https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68) for the full syntax.

### Excel Dates

Excel stores dates as numeric serial values. `ExcelDateModule` is registered by default, automatically converting between Java date types and Excel date numbers.

Supported: `Date`, `Calendar`, `LocalDate`, `LocalDateTime`.

No configuration needed. Read an Excel date cell and get a `LocalDate`. Write a `LocalDate` and get an Excel-formatted date.

## Sheet-Level Features

`GridConfigurer` controls sheet-level features anchored on the data grid:

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .stylesBuilder(new StylesBuilder()
        .cellStyle("highlight")
            .fillForegroundColor(IndexedColors.RED)
            .fillPattern().solidForeground()
            .end())
    .gridConfigurer(new GridConfigurer()
        .freezePane(0, 1)
        .autoFilter()
        .conditionalFormatting()
            .column("score").greaterThanOrEqual("80").style("highlight").end())
    .build();
```

`freezePane(colSplit, rowSplit)` delegates to POI `Sheet#createFreezePane` — the leftmost `colSplit` columns and topmost `rowSplit` rows stay fixed while scrolling. `autoFilter()` enables the filter dropdown across all schema columns; the range is computed from the schema and row count. Both the streaming default and `USE_POI_USER_MODEL` write paths apply these identically.

### Conditional Formatting

`column(name)` matches `@DataColumn(value)`, the field name, or `@JsonAlias`. `style(name)` matches a `cellStyle(name)` in `StylesBuilder`. Both resolve at write time — typos throw `IllegalArgumentException` listing the available names.

```java
.conditionalFormatting().column("score").greaterThan("90").style("good").end()
.conditionalFormatting().column("status").equalTo("\"URGENT\"").style("warn").end()  // text: quote the value
.conditionalFormatting().column("price").between("100", "500").style("warn").end()
```

Operators: `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`, `equalTo`, `notEqualTo`, `between`, `notBetween`.

Style → DXF: fill, font, and border only. Alignment and wrap-text are silently skipped.

## Configuration

### Origin

```java
// Data grid starts at B2 instead of default A1
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .origin("B2")   // or .origin(1, 1) or .origin(new CellAddress("B2"))
    .build();
```

### Header

```java
// Default: first row is header, data starts at row 2
SpreadsheetMapper mapper = new SpreadsheetMapper();

// No header: data starts at origin row
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .useHeader(false)
    .build();

// Per-schema toggle
SpreadsheetSchema schema = mapper.sheetSchemaFor(Entry.class)
    .withUseHeader(false);
```

### Column Reordering

By default, columns are matched by position. Enable `columnReordering` to match by header name instead:

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .columnReordering(true)
    .build();
```

The header row is read and schema columns are reordered to match the file's column order. Extra columns in the file are ignored. Missing columns get default values. `@JsonAlias` names are also checked during header matching.

### Column Names

Override how column names are resolved:

```java
// Lambda — transform property names
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .columnNameResolver(prop -> prop.getName().toUpperCase())
    .build();

// Custom annotation — resolve from your own annotation
ColumnNameResolver resolver = ColumnNameResolver.annotated(
    MyColumn.class, MyColumn::displayName);
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .columnNameResolver(resolver)
    .build();
```

### Parser Features

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SheetParser.Feature.BREAK_ON_BLANK_ROW)
    .build();
```

| Feature | Default | Description |
|---------|---------|-------------|
| `BLANK_ROW_AS_NULL` | enabled | Blank rows are deserialized as `null` |
| `BREAK_ON_BLANK_ROW` | disabled | Stop reading at the first blank row |

### Factory Features

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
    .build();
```

| Feature | Default | Description |
|---------|---------|-------------|
| `USE_POI_USER_MODEL` | disabled | Use POI's User Model (Sheet/Row/Cell) for all read/write, bypassing streaming |
| `FILE_BACKED_SHARED_STRINGS` | disabled | Store shared strings on disk for read and write (requires `com.h2database:h2`) |
| `ENCRYPT_FILE_BACKED_STORE` | disabled | Encrypt the file-backed store with AES (requires `FILE_BACKED_SHARED_STRINGS`) |

## Format Support

### XLSX and XLS

Both formats are supported through the same API. The format is auto-detected via ZIP magic bytes — no configuration needed.

```java
// Same code works for both
mapper.readValues(new File("data.xlsx"), Product.class); // XLSX (OOXML)
mapper.readValues(new File("data.xls"), Product.class);  // XLS (legacy)
```

The two formats use different read paths internally:

| | XLSX (default) | XLSX (`USE_POI_USER_MODEL`) | XLS |
|---|---|---|---|
| Read path | StAX streaming | Apache POI (UserModel) | Apache POI (UserModel) |
| Write path | StringBuilder + POI skeleton | POI WorkbookProvider (default: `XSSFWorkbook`) | POI `HSSFWorkbook` |

For large files, XLSX is strongly recommended.

### InputStream Handling

When reading from an `InputStream`, OOXML (XLSX) files are copied to a temporary file first because ZIP format requires random access. The temp file is cleaned up when the parser is closed.

For best performance with large XLSX files, prefer `File` input over `InputStream`.

## Performance

At 100K rows (mixed types, shared string table):

**Read:**

| Library | Time | Memory |
|---------|------|--------|
| jackson-spreadsheet | 198 ms | 378 MB |
| FastExcel | 212 ms | 428 MB |
| Fesod | 279 ms | 400 MB |
| Poiji | 843 ms | 2876 MB |
| Apache POI | 1198 ms | 2333 MB |

**Write:**

| Library | Time | Memory |
|---------|------|--------|
| jackson-spreadsheet | 150 ms | 191 MB |
| FastExcel | 166 ms | 156 MB |
| Apache POI | 283 ms | 207 MB |
| Fesod | 337 ms | 480 MB |

Fastest read and write throughput among all libraries. See [BENCHMARK.md](BENCHMARK.md) for full results.

### Low-Memory Mode for Large Files

For extremely large XLSX files that cause `OutOfMemoryError`:

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
    .build();
```

If the spreadsheet contains sensitive data, enable encryption to protect the temp file at rest:

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
    .enable(SpreadsheetFactory.Feature.ENCRYPT_FILE_BACKED_STORE)
    .build();
```

Requires `com.h2database:h2` on the classpath:

```xml
<dependency>
    <groupId>com.h2database</groupId>
    <artifactId>h2</artifactId>
    <version>2.2.224</version>
</dependency>
```

Trades ~40% throughput for constant heap usage regardless of string table size. The temporary file is automatically deleted when the reader is closed.

## FAQ

**Q: Does `@DataGrid` have to be on the root class?**
Yes. `@DataGrid` marks the root type for schema generation. Nested classes (like `Address` in an `Employee`) do not need it.

**Q: What happens if a column in the spreadsheet doesn't match any field?**
It is ignored. Only columns matching the schema (derived from the class structure) are read.

**Q: Can I read multiple sheets from one file?**
Yes. Use `SheetInput.source(file, "SheetName")` or `SheetInput.source(file, sheetIndex)` to target a specific sheet. For multiple sheets, use direct `Sheet` access via POI `Workbook`.

**Q: Why does InputStream reading create a temp file?**
XLSX is a ZIP archive. ZIP requires random access (seek), which `InputStream` does not support. The library copies the stream to a temp file, reads it, and deletes the file when done.

**Q: How is the column order determined?**
By the field declaration order in the Java class. Nested object fields are flattened in-place.

**Q: Can I use this with Spring Boot?**
Yes. Create a `SpreadsheetMapper` bean and inject it. It is thread-safe like `ObjectMapper`.

## References

- [ARCHITECTURE.md](ARCHITECTURE.md) — Design decisions, data flow
- [BENCHMARK.md](BENCHMARK.md) — JMH benchmark results
- [Jackson documentation](https://github.com/FasterXML/jackson-docs)
- [Apache POI](https://poi.apache.org/components/spreadsheet/index.html)
