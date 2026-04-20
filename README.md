# jackson-dataformat-spreadsheet

[![Build](https://github.com/scndry/jackson-dataformat-spreadsheet/actions/workflows/build.yml/badge.svg)](https://github.com/scndry/jackson-dataformat-spreadsheet/actions/workflows/build.yml)
[![Maven Central](https://img.shields.io/maven-central/v/io.github.scndry/jackson-dataformat-spreadsheet)](https://central.sonatype.com/artifact/io.github.scndry/jackson-dataformat-spreadsheet)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://www.apache.org/licenses/LICENSE-2.0)

> _Let's get rid of the Spread Sh!t._

**A Java library (Jackson extension) for reading and writing Excel spreadsheets (XLSX/XLS) as POJOs.**

Map spreadsheet rows to Java objects with `ObjectMapper` — the same API you use for JSON, CSV, and XML. No cell-level iteration, no column index counting, no manual type casting.

## What It Does

```java
// Read Excel to POJOs
SpreadsheetMapper mapper = new SpreadsheetMapper();
List<Employee> employees = mapper.readValues(file, Employee.class);

// Write POJOs to Excel
mapper.writeValue(file, employees, Employee.class);
```

**That's it.** If you know Jackson, you know this library.

## When To Use This

- You need to **read XLSX/XLS files into Java objects** without writing cell-by-cell parsing code
- You need to **write Java objects to Excel** with headers, types, and styling
- You want **Jackson ecosystem integration** (custom deserializers, mix-ins, modules)
- You have **nested object structures** that need to flatten into spreadsheet columns
- You need **streaming performance** for large files (100K+ rows)

## Installation

Available on Maven Central:

**Maven:**
```xml
<dependency>
    <groupId>io.github.scndry</groupId>
    <artifactId>jackson-dataformat-spreadsheet</artifactId>
    <version>1.1.0</version>
</dependency>
```

**Gradle:**
```groovy
implementation "io.github.scndry:jackson-dataformat-spreadsheet:1.1.0"
```

### Requirements

- Java 8+
- Jackson 2.14.0+
- Apache POI 4.1.1+ (Strict OOXML requires 5.1.0+)

## Quick Start

### Define your model

```java
@DataGrid
public class Product {
    private String name;
    private int quantity;
    private double price;
    // constructors, getters, setters
}
```

### Read from Excel

```java
SpreadsheetMapper mapper = new SpreadsheetMapper();

// Single object (first row)
Product product = mapper.readValue(file, Product.class);

// All rows
List<Product> products = mapper.readValues(file, Product.class);

// Specific sheet
SheetInput<File> input = SheetInput.source(file, "Products");
List<Product> products = mapper.readValues(input, Product.class);
```

### Write to Excel

```java
List<Product> products = Arrays.asList(
    new Product("Apple", 10, 1.50),
    new Product("Banana", 20, 0.80));

mapper.writeValue(new File("output.xlsx"), products, Product.class);
```

### Nested Objects

Flat spreadsheets map to nested POJOs automatically:

| id | name | zipcode | city | title | salary |
|----|------|---------|------|-------|--------|
| 1 | Alice | 12345 | Seoul | SRE | 80000 |

```java
@DataGrid
class Employee {
    int id;
    String name;
    Address address;    // flattened: zipcode, city
    Employment employment; // flattened: title, salary
}
```

No configuration needed. Read and write — both directions work.

## How It Compares

### Performance (100K rows, mixed types, shared string table)

| Library | Read Time | Read Memory |
|---------|-----------|-------------|
| FastExcel | 209 ms | 406 MB |
| jackson-spreadsheet | 220 ms | 395 MB |
| EasyExcel | 296 ms | 418 MB |
| Poiji | 889 ms | 2909 MB |
| Apache POI UserModel | 1274 ms | 2347 MB |

### Feature Comparison

| Feature | jackson-spreadsheet | Apache POI | EasyExcel | FastExcel |
|---------|:---:|:---:|:---:|:---:|
| POJO data binding | Yes | No | Yes | No |
| Nested object support | Yes | No | No | No |
| Jackson ecosystem | Yes | No | No | No |
| Streaming read | Yes | Event API | Yes | Yes |
| Streaming write | Yes | SXSSF | Yes | Yes |
| Cell styling | Yes | Yes | Yes | No |
| XLSX read/write | Yes | Yes | Yes | Yes |
| XLS read/write | Yes | Yes | No | No |
| Annotation mapping | Yes | No | Yes | No |

## Key Features

### Annotations

```java
@DataGrid
class Product {
    @DataColumn("Product Name")
    String name;

    @DataColumn(value = "Price", style = "currency")
    double price;

    @DataColumn(merge = OptBoolean.TRUE)
    String category;
}
```

### Streaming for Large Files

```java
SpreadsheetReader reader = mapper.sheetReaderFor(Product.class);
try (SheetMappingIterator<Product> iter = reader.readValues(input)) {
    while (iter.hasNext()) {
        Product p = iter.next();
        // process one at a time
    }
}
```

### Cell Styling

```java
StylesBuilder styles = new StylesBuilder()
    .cellStyle("currency")
        .dataFormat("#,##0.00")
        .font().bold().end()
        .end();

SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .stylesBuilder(styles)
    .build();
```

### Configuration

```java
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .origin("B2")           // start at B2 instead of A1
    .useHeader(false)        // no header row
    .enable(SheetParser.Feature.BREAK_ON_BLANK_ROW)
    .build();
```

### Excel Date Handling

Built-in conversion between Java date types and Excel serial numbers. Registered by default — no setup needed.

Supported: `Date`, `Calendar`, `LocalDate`, `LocalDateTime`

## Architecture

Not a POI wrapper. Extends Jackson's streaming layer directly:

- `SheetParser extends ParserMinimalBase` — StAX pull parser
- `SheetGenerator extends GeneratorBase` — POI cell writer
- `SpreadsheetFactory extends JsonFactory` — format detection

The XLSX read path parses OOXML XML directly via StAX — no XMLBeans, no SAX, no DOM. Lightweight ECMA-376 schema types provide type safety without runtime overhead.

See [ARCHITECTURE.md](ARCHITECTURE.md) for design decisions and data flow diagrams.

## Documentation

- [GUIDE.md](GUIDE.md) — Full usage guide
- [ARCHITECTURE.md](ARCHITECTURE.md) — Design decisions, data flow
- [BENCHMARK.md](BENCHMARK.md) — JMH benchmark results

## FAQ

**Q: How is this different from Apache POI?**
POI gives you cells. This gives you POJOs. You define a class with `@DataGrid`, and `mapper.readValues()` returns typed objects. No `row.getCell(0).getStringCellValue()`.

**Q: How is this different from EasyExcel?**
EasyExcel has its own API. This extends Jackson's `ObjectMapper`, so you get the full Jackson ecosystem — custom deserializers, mix-ins, modules, polymorphic types.

**Q: Does it support nested objects?**
Yes. Nested POJOs automatically flatten to columns on write and reconstruct on read. No configuration needed.

**Q: How does performance compare?**
Comparable to FastExcel on read throughput at 100K rows, with the lowest memory allocation. 6x faster than POI UserModel. See [BENCHMARK.md](BENCHMARK.md).

**Q: What Excel formats are supported?**
XLSX (OOXML) and XLS (legacy). XLSX uses StAX streaming; XLS uses POI object model.

**Q: Is it production-ready?**
Yes. Version 1.1.0 on Maven Central. Java 8+, Jackson 2.14+, POI 4.1.1+. Listed as a [community data format module](https://github.com/FasterXML/jackson#data-format-modules) in the FasterXML jackson repository.

## License

[Apache License 2.0](LICENSE)
