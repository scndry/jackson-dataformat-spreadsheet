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
    <version>1.0.1</version>
</dependency>
```

**Gradle:**
```groovy
implementation "io.github.scndry:jackson-dataformat-spreadsheet:1.0.1"
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

## Reading

```java
SpreadsheetMapper mapper = new SpreadsheetMapper();

// Single row (first data row)
Product p = mapper.readValue(file, Product.class);

// All rows
List<Product> list = mapper.readValues(file, Product.class);

// Specific sheet by name
SheetInput<File> input = SheetInput.source(file, "Products");
List<Product> list = mapper.readValues(input, Product.class);

// Specific sheet by index
SheetInput<File> input = SheetInput.source(file, 2);
List<Product> list = mapper.readValues(input, Product.class);

// From InputStream
SheetInput<InputStream> input = SheetInput.source(stream, "Products");
List<Product> list = mapper.readValues(input, Product.class);
```

### Streaming Read

For large files, process rows one at a time with constant memory:

```java
SpreadsheetReader reader = mapper.sheetReaderFor(Product.class);
try (SheetMappingIterator<Product> iter = reader.readValues(input)) {
    while (iter.hasNext()) {
        Product p = iter.next();

        // Current cell location (row, column)
        SheetLocation location = iter.getCurrentLocation();
    }
}
```

`readValues(file, type)` loads all rows into a `List`. For 100K+ row files, prefer `SheetMappingIterator`.

## Writing

```java
SpreadsheetMapper mapper = new SpreadsheetMapper();

// Single object
mapper.writeValue(file, product);

// Collection — element type required (Java type erasure)
mapper.writeValue(file, products, Product.class);

// Specific sheet name
SheetOutput<File> output = SheetOutput.target(file, "Products");
mapper.writeValue(output, products, Product.class);

// To OutputStream
SheetOutput<OutputStream> output = SheetOutput.target(stream, "Products");
mapper.writeValue(output, products, Product.class);

// To byte array (in-memory Excel generation)
byte[] bytes = mapper.writeValueAsBytes(products, Product.class);
```

### Streaming Write

Write rows one at a time using Jackson's `SequenceWriter`:

```java
SpreadsheetWriter writer = mapper.sheetWriterFor(Product.class);
try (SequenceWriter seq = writer.writeValues(SheetOutput.target(file))) {
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

The schema is generated from the class hierarchy via Jackson's `FormatVisitor`. Each nested object's fields become contiguous columns. No manual column index mapping needed.

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

| Order ID | Product | Qty | Total |
|----------|---------|-----|-------|
| 1 | Apple | 3 | 8 |
| | Banana | 5 | |

`Order ID` and `Total` cells are merged vertically.

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
| `value` | `""` (field name) | Column header name |
| `style` | `""` | Cell style for data cells |
| `headerStyle` | `""` | Cell style for the header cell |
| `width` | `-1` (auto) | Column width in character units |
| `autoSize` | `DEFAULT` | Auto-size to fit content |
| `minWidth` | `-1` (none) | Minimum column width |
| `maxWidth` | `255` | Maximum column width |
| `merge` | `DEFAULT` | Merge cells vertically |

`autoSize` may not be accurate for fullwidth forms like CJK characters.

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

### Excel Dates

Excel stores dates as numeric serial values. `ExcelDateModule` is registered by default, automatically converting between Java date types and Excel date numbers.

Supported: `Date`, `Calendar`, `LocalDate`, `LocalDateTime`.

No configuration needed. Read an Excel date cell and get a `LocalDate`. Write a `LocalDate` and get an Excel-formatted date.

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
| `FILE_BACKED_SHARED_STRINGS` | disabled | Store shared strings on disk (requires `com.h2database:h2`) |

## Format Support

### XLSX and XLS

Both formats are supported through the same API. The format is auto-detected via ZIP magic bytes — no configuration needed.

```java
// Same code works for both
mapper.readValues(new File("data.xlsx"), Product.class); // XLSX (OOXML)
mapper.readValues(new File("data.xls"), Product.class);  // XLS (legacy)
```

The two formats use different read paths internally:

| | XLSX | XLS |
|---|---|---|
| Read path | StAX streaming (constant memory) | POI object model (full workbook in memory) |
| Write path | POI `SXSSFWorkbook` (streaming) | POI `HSSFWorkbook` |

For large files, XLSX is strongly recommended.

### InputStream Handling

When reading from an `InputStream`, OOXML (XLSX) files are copied to a temporary file first because ZIP format requires random access. The temp file is cleaned up when the parser is closed.

For best performance with large XLSX files, prefer `File` input over `InputStream`.

## Performance

At 100K rows (mixed types, shared string table):

| Library | Read | Memory |
|---------|------|--------|
| FastExcel | 209 ms | 406 MB |
| jackson-spreadsheet | 220 ms | 395 MB |
| EasyExcel | 296 ms | 418 MB |
| Poiji | 889 ms | 2909 MB |
| Apache POI UserModel | 1274 ms | 2347 MB |

Lowest memory allocation among all libraries. See [BENCHMARK.md](BENCHMARK.md) for full results.

### File-Backed SharedStrings

For extremely large XLSX files where the shared string table exceeds available heap:

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SheetParser.Feature.FILE_BACKED_SHARED_STRINGS)
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
