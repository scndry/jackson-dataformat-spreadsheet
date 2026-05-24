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

POI types (`Sheet`, `Workbook`) are first-class I/O targets — see [POI Integration](#poi-integration).

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
    <version>1.7.0</version>
</dependency>
```

**Gradle:**
```groovy
implementation "io.github.scndry:jackson-dataformat-spreadsheet:1.7.0"
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

### Formula Cells

Formula cells return the **cached computed value**, not the formula text. The reader binds the cached value (emitted when the formula was last evaluated by Excel/POI) to the typed field — no manual `FormulaEvaluator` invocation needed.

```java
@DataGrid
class Invoice {
    String item;
    int quantity;
    double unitPrice;
    // formula reads as the cached value
    double subtotal;   // `=B2*C2` (quantity*unitPrice) — e.g., 5000.0
    double total;      // `=D2*1.1` (subtotal*1.1) — e.g., 5500.0 (10% VAT)
}
```

If the cached value is missing (rare — produced by a writer that doesn't recompute), the cell reads as blank. To force re-evaluation, open the workbook with POI directly and invoke `FormulaEvaluator.evaluateAll()` before passing the `Sheet` to the mapper.

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

Invalid sheet names throw `IllegalArgumentException` at `SheetOutput.target(..., name)` per Excel constraints.

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

Flat spreadsheets map to nested POJOs automatically.

| id | name | address/zipcode | address/city | employment/title | employment/salary |
|----|------|-----------------|--------------|------------------|-------------------|
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

Nested lists are serialized into multi-row output. The reverse — reading multi-row blocks back into nested lists — is supported when the outer record carries an anchor column.

```java
@DataGrid(mergeColumn = OptBoolean.TRUE)
class Order {
    @DataColumn("Order ID") int orderId;
    List<Item> items;
    @DataColumn("Total") int total;
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

Reading the same multi-row layout back into a `List<Order>` requires marking one column per record level with `@DataColumn(anchor = true)`. Rows sharing the anchor value collapse into one outer record.

```java
@DataGrid
class Order {
    @DataColumn(value = "Order ID", anchor = true) int orderId;
    List<Item> items;
    @DataColumn("Total") int total;
}

List<Order> orders = mapper.readValues(file, Order.class);
```

## Annotations

### @DataGrid

Marks the root type for schema generation. Attributes set class-level defaults for all columns.

| Attribute | Default | Description |
|-----------|---------|-------------|
| `columnStyle` | `""` | Default cell style for data cells |
| `columnHeaderStyle` | `""` | Default cell style for leaf header cells |
| `groupHeaderStyle` | `""` | Default cell style for `@DataColumnGroup` header cells |
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

`autoSize` is not supported in the default streaming write path — enable `USE_POI_USER_MODEL` or set `width` explicitly.

When `USE_POI_USER_MODEL` is enabled, `autoSize` samples rows for bounded overhead (~1.5× write time at 100K rows) and may miss outliers between sample rows — pin a known column with `width` if exact fit matters.

### @DataColumnGroup

Renders flattened columns from a nested object field under a shared group header. The annotation stacks across nesting depth to form multi-row headers — parent group on top, leaf names at the bottom.

| Attribute | Default | Description |
|-----------|---------|-------------|
| `value` | field name | Column group header name |
| `comment` | `""` | Group header cell comment text |
| `headerStyle` | `""` | Cell style for the group header cell |
| `columnStyle` | `""` | Default cell style for child data cells (cascade) |
| `columnHeaderStyle` | `""` | Default cell style for child leaf header cells (cascade) |
| `columnWidth` | `-1` (auto) | Default width for child columns (cascade) |
| `autoSizeColumn` | `DEFAULT` | Default auto-size for child columns (cascade) |
| `minColumnWidth` | `-1` (none) | Default minimum width for child columns (cascade) |
| `maxColumnWidth` | `255` | Default maximum width for child columns (cascade) |
| `mergeColumn` | `DEFAULT` | Default merge for child columns (cascade) |

The seven cascade attributes mirror `@DataGrid`'s corresponding defaults and act as an intermediate layer between the leaf `@DataColumn` and the enclosing `@DataGrid` — see *Attribute Resolution Order* below.

```java
@DataGrid
class Employee {
    @DataColumn("ID") int id;
    @DataColumn("Name") String name;
    @DataColumnGroup(value = "Address", comment = "Customer billing address")
    Address address;
}
class Address {
    @DataColumn("City") String city;
    @DataColumn("Zip") String zip;
}
```

renders as:

```
+----+-------+---------------+
| ID | Name  |    Address    |
|    |       +-------+-------+
|    |       | City  | Zip   |
+----+-------+-------+-------+
|  1 | Alice | Seoul | 12345 |
+----+-------+-------+-------+
```

The header row count is `max(group depth) + 1`, derived automatically. Flat or shallow columns vertically merge from their hierarchy depth down to the leaf header row. Adjacent columns sharing parent path and group name merge horizontally; a different parent yields separate cells (e.g. `Q1` under `2024` vs `Q1` under `2025`). On read the parser auto-skips header rows and column reordering only matches the leaf row.

For deeper hierarchies the annotation stacks naturally:

```java
class Company {
    @DataColumn("Name") String name;
    @DataColumnGroup("2024") YearMetrics year2024;
    @DataColumnGroup("2025") YearMetrics year2025;
}
class YearMetrics {
    @DataColumnGroup("Q1") QuarterMetrics q1;
    @DataColumnGroup("Q2") QuarterMetrics q2;
}
class QuarterMetrics {
    @DataColumn("Sales") int sales;
    @DataColumn("Profit") int profit;
}
```

```
+------+---------------------------------+---------------------------------+
| Name |              2024               |              2025               |
|      +----------------+----------------+----------------+----------------+
|      |       Q1       |       Q2       |       Q1       |       Q2       |
|      +-------+--------+-------+--------+-------+--------+-------+--------+
|      | Sales | Profit | Sales | Profit | Sales | Profit | Sales | Profit |
+------+-------+--------+-------+--------+-------+--------+-------+--------+
```

Depth 3 is the practical limit; readability degrades beyond that. `setColumnReordering(true)` combined with `@DataColumnGroup` throws `IllegalStateException` at schema load.

### Attribute Resolution Order

Column attributes resolve in priority order (highest first, first-non-empty wins):

1. `@DataColumn` on the leaf property
2. Innermost enclosing `@DataColumnGroup` (cascade slots)
3. Outer enclosing `@DataColumnGroup` (recurse outward)
4. `@DataGrid` on the declaring class
5. `@DataGrid` on the enclosing (parent) class

`@DataColumnGroup.headerStyle` is level-specific — outer groups do not cascade their `headerStyle` into inner groups. Each group level falls back to `@DataGrid.groupHeaderStyle` from the declaring class, then from the enclosing class.

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
| foo/a | `"highlight"` (from @DataColumn) | `FALSE` (from Foo's @DataGrid) |
| foo/b | `"base"` (from Bar's @DataGrid) | `FALSE` (from Foo's @DataGrid) |

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
mapper.sheetWriterForWithView(Report.class, Views.Summary.class)
    .writeValue(file, reports);
```

Note: `mapper.writerWithView()` does not work — it bypasses schema generation. Use `sheetWriterForWithView(type, view)` instead.

## Styling

Register named cell styles with `StylesBuilder` and reference them from annotations. Each named style maps to exactly one POI `CellStyle` (`workbook.createCellStyle()` invoked once per name at build time), so the [64,000 cell-style per-workbook limit](https://learn.microsoft.com/en-us/office/troubleshoot/excel/excel-specifications-and-limits) is bound by style declarations rather than row count.

`StylesBuilder.simple()` returns a starter builder with per-type defaults registered (comma-grouped number formats, date/datetime formats). Drop it in early to get sensible defaults, then replace with a fully custom builder as needs grow.

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

`ExcelDateModule` is registered by default — `Date`, `Calendar`, `LocalDate`, `LocalDateTime` convert to/from Excel serial values automatically. Without a date format on the cell, Excel shows the raw serial (e.g. `46157`); `StylesBuilder.simple()` registers per-type defaults as a starter (see [Styling](#styling)).

On read, the workbook's date system (1900 / 1904) is detected; write defaults to 1900.

## Sheet-Level Features

`GridConfigurer` controls sheet-level features anchored on the data grid. The conditional formatting call below uses factory methods from `ConditionalFormats`, typically static-imported (see [Conditional Formatting](#conditional-formatting) for details):

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
        .conditionalFormatting("score",
            greaterThanOrEqual(80).style("highlight")))
    .build();
```

`freezePane(colSplit, rowSplit)` freezes the leftmost columns / topmost rows. `autoFilter()` enables the filter dropdown across all schema columns. Both work identically on streaming and `USE_POI_USER_MODEL` write paths.

### Conditional Formatting

`conditionalFormatting(String column, ConditionalFormatRule rule, ConditionalFormatRule... rules)` accepts one or more rules for a single column. Static-import the factory methods from `ConditionalFormats` for fluent chaining:

```java
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.ConditionalFormats.*;
```

Column names resolve against `@DataColumn(value)`, the field name, or `@JsonAlias`. Style names resolve against `cellStyle(name)` in `StylesBuilder`. Both resolve at write time. Alignment and wrap-text on a referenced style are silently dropped — DXF limitation.

Comparison factories (`greaterThan`, `between`, `equalTo`, ...) accept typed values (numeric, boolean, string, date) or `Formula` for cell references and Excel expressions. String operands are auto-quoted; dates emit as `DATE(y,m,d)+TIME(h,m,s)`. The factory returns a `FormatCondition`; `.style(name)` finishes it as a `ConditionalFormatRule`.

```java
// Typed primitives — single-rule call
.conditionalFormatting("score", greaterThan(90).style("good"))
.conditionalFormatting("price", between(100, 500).style("warn"))
.conditionalFormatting("status", equalTo("URGENT").style("warn"))
.conditionalFormatting("active", equalTo(true).style("info"))

// Date types — LocalDate, LocalDateTime, Date, Calendar all supported
.conditionalFormatting("createdAt", greaterThan(LocalDate.of(2026, 1, 1)).style("recent"))

// Multi-rule for one column — varargs
.conditionalFormatting("score",
    greaterThanOrEqual(80).style("good"),
    lessThan(60).style("bad"))

// Cell reference / function — formula() for raw passthrough
.conditionalFormatting("price", greaterThan(formula("$D$1")).style("warn"))
.conditionalFormatting("price", greaterThan(formula("AVERAGE($B$2:$B$100)")).style("aboveAvg"))

// Schema-aware cross-column reference — row-relative, schema-safe
.conditionalFormatting("price", greaterThan(columnRef("minPrice")).style("warn"))

// Arbitrary boolean expression rule (type="expression")
.conditionalFormatting("score", expression("AND($A1>0, $B1<100)").style("warn"))

// Color scale — 3-color visualization, no styling required
.conditionalFormatting("revenue", colorScale())                    // Excel defaults: MIN / PERCENTILE 50 / MAX
.conditionalFormatting("revenue", colorScale(0, 50_000, 100_000))   // explicit NUMBER thresholds
```

Comparison factories: `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`, `equalTo`, `notEqualTo`, `between`, `notBetween`. Equality also accepts `boolean` and `String`.

Operand types: any `Number` (`int`, `long`, `double`, `float`, `BigDecimal`, etc. — primitives are autoboxed), `boolean`, `String` (equality only — auto-escaped to Excel string literal), `LocalDate`, `LocalDateTime`, `Date`, `Calendar`, `Formula`.

#### Date and time precision

Dates emit as Excel formula expressions:

| Type | Formula |
|------|---------|
| `LocalDate` | `DATE(y,m,d)` |
| `LocalDateTime` | `DATE(y,m,d)+TIME(h,m,s)` — sub-second truncated |
| `Date` | `DATE(y,m,d)+TIME(h,m,s)` — converted via system default timezone |
| `Calendar` | `DATE(y,m,d)+TIME(h,m,s)` — converted via Calendar's timezone |

Prefer `LocalDate` / `LocalDateTime` for deterministic CF rules. `Date` carries a system-timezone dependence; the resulting formula varies with the JVM `ZoneId.systemDefault()`.

#### `cellIs` vs `expression`

| | `cellIs` operators | `expression` |
|---|---|---|
| Compares | The cell against operand(s) | Arbitrary boolean formula |
| Example | `greaterThan(100).style("warn")` | `expression("$E1<$F1").style("warn")` |
| Use when | Direct comparison fits | Need cross-cell logic, `AND`/`OR`, `ISBLANK`, etc. |

`expression(formula)` is passed verbatim to POI; do not include a leading `=`.

#### Formula escape

`formula(text)` is a power-user escape — the text is emitted verbatim into the OOXML `<formula>` element. The library does not validate Excel syntax. `columnRef(name)` resolves the schema column name to a row-relative reference (`$<col><dataStartRow>`) at write time, so Excel auto-shifts per cell in the formatting range.

#### Color scale (3-color)

A 3-color gradient based on cell values. `colorScale` returns `ConditionalFormatRule` directly — no `.style()` required, since the visualization carries its own colors.

| Form | Thresholds | Colors |
|------|-----------|--------|
| `colorScale()` | Excel defaults — MIN / PERCENTILE 50 / MAX | red → yellow → green |
| `colorScale(min, mid, max)` | Explicit NUMBER values | red → yellow → green |

Color customization, threshold types other than `NUMBER` (PERCENT / PERCENTILE / FORMULA), the 2-color variant, and other visualization types (`dataBar`, `iconSet`) are deferred to a later release.

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

// Per-schema toggle (using SchemaGenerator)
SpreadsheetSchema schema = new SchemaGenerator()
    .withUseHeader(false)
    .generate(Entry.class);
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
| `USE_POI_USER_MODEL` | disabled | Use POI's User Model (Sheet/Row/Cell) for all read/write, bypassing streaming. Also skips InputStream temp-file copy — see [InputStream Handling](#inputstream-handling). |
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
| Write path | StringBuilder + POI scaffold | POI WorkbookProvider (default: `XSSFWorkbook`) | POI `HSSFWorkbook` |

For large files, XLSX is strongly recommended.

### InputStream Handling

When reading from an `InputStream`, OOXML (XLSX) files are copied to a temporary file first because ZIP format requires random access. The temp file is cleaned up when the parser is closed.

For best performance with large XLSX files, prefer `File` input over `InputStream`.

**Disk-write-restricted environments** (AWS Lambda read-only filesystem, Kubernetes pods with `readOnlyRootFilesystem: true`, Docker `--read-only`, sandbox runners) cannot create the temp file. Two options:

1. **Use `File` input** — POI reads directly from disk, no temp file.
2. **Enable `USE_POI_USER_MODEL`** — InputStream is passed directly to POI, which holds the entire ZIP in memory (no temp file). Trade-off: higher heap usage. POI's own javadoc recommends `File` over `InputStream` when possible because of this.

```java
// In-memory read from InputStream (no temp file)
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
    .build();
List<Row> rows = mapper.readValues(inputStream, Row.class);
```

## Low-Memory Mode for Large Files

For extremely large XLSX files that cause `OutOfMemoryError`:

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
    .build();
```

If the spreadsheet contains sensitive data, enable encryption to protect the **shared-strings temp file** (H2 MVStore-backed) at rest:

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
    .enable(SpreadsheetFactory.Feature.ENCRYPT_FILE_BACKED_STORE)
    .build();
```

This protects the temp file only — the output XLSX is not encrypted. Workbook-level password protection (encrypted XLSX) is out of scope; use POI directly for that.

Requires `com.h2database:h2` on the classpath:

```xml
<dependency>
    <groupId>com.h2database</groupId>
    <artifactId>h2</artifactId>
    <version>2.2.224</version>
</dependency>
```

Trades throughput for constant heap usage regardless of string table size — see [BENCHMARK.md](BENCHMARK.md) for measured overhead. The temporary file is automatically deleted when the reader is closed.

## Logging

SLF4J `TRACE` level emits per-cell values and addresses — disable in environments handling sensitive data.

## FAQ

**Q: Does `@DataGrid` have to be on the root class?**
Yes. `@DataGrid` marks the root type for schema generation. Nested classes (like `Address` in an `Employee`) do not need it.

**Q: What happens if a column in the spreadsheet doesn't match any field?**
It is ignored. Only columns matching the schema (derived from the class structure) are read.

**Q: Can I read multiple sheets from one file?**
Yes. Use `SheetInput.source(file, "SheetName")` or `SheetInput.source(file, sheetIndex)` to target a specific sheet. For multiple sheets, use direct `Sheet` access via POI `Workbook`.

**Q: Why does InputStream reading create a temp file?**
XLSX is a ZIP archive. ZIP requires random access (seek), which `InputStream` does not support. The library copies the stream to a temp file, reads it, and deletes the file when done. In disk-write-restricted environments (Lambda read-only fs, Kubernetes `readOnlyRootFilesystem`), enable `USE_POI_USER_MODEL` to skip the temp file at the cost of higher heap usage — see [InputStream Handling](#inputstream-handling).

**Q: How is the column order determined?**
By the field declaration order in the Java class. Nested object fields are flattened in-place.

**Q: Can I use this with Spring Boot?**
Yes. Create a `SpreadsheetMapper` bean and inject it. It is thread-safe like `ObjectMapper`.

## References

- [ARCHITECTURE.md](ARCHITECTURE.md) — Design decisions, data flow
- [BENCHMARK.md](BENCHMARK.md) — JMH benchmark results
- [Jackson documentation](https://github.com/FasterXML/jackson-docs)
- [Apache POI](https://poi.apache.org/components/spreadsheet/index.html)
