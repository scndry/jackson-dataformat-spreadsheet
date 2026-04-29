# Design

Why these design choices were made — and where the library deliberately stops.

This library started from production incidents with Apache POI — OutOfMemoryError on large exports, silent type mismatches, cell index bugs after column changes. The common thread: spreadsheet I/O forced developers to think in cells instead of objects.

ORM solved the impedance mismatch between SQL tables and Java objects. Jackson solved it between JSON text and Java objects. The same mismatch existed between spreadsheet grids and Java objects — but there was no Jackson dataformat module for spreadsheets.

## Why Build on Jackson

Excel libraries typically invent their own API. POI gives you `Sheet`/`Row`/`Cell`. EasyExcel gives you `ReadListener`. Fesod gives you its own annotation set. Each library is a new world to learn.

This library takes a different path: **it extends Jackson's ObjectMapper**. Not wraps. Not adapts. Extends.

Why? Because most Java backend developers already know Jackson. They use `@JsonProperty`, `@JsonIgnore`, `@JsonView`, custom serializers — every day, for JSON. If those same annotations and the same `mapper.readValue()` / `mapper.writeValue()` API just work for Excel, there is nothing new to learn.

The question was never "how to read Excel in Java." POI already answers that. The question is: **why should reading Excel feel different from reading JSON?** It shouldn't. A spreadsheet row is a serialized object — same as a JSON line, a CSV row, or an XML element. Jackson already has the abstraction for this. The missing piece was a dataformat module for spreadsheets.

## Why Not Wrap POI

A POI wrapper would still expose the grid mental model — rows, columns, cell types. You'd get convenience methods, but you'd still think in cells.

This library doesn't wrap POI. It sits **above** POI at the streaming layer:

```
Your code → ObjectMapper (Jackson) → SheetParser/SheetGenerator (this library) → SheetReader/SheetWriter → POI or StAX
```

POI is one of two I/O backends — the other being direct StAX. The default XLSX path bypasses POI entirely (StAX + StringBuilder). POI is used when the user needs it — direct `Sheet` access, XLS format, or `USE_POI_USER_MODEL`. It's opt-in, not the foundation.

This is why the library can outperform POI for reading and writing. POI is still a dependency (skeleton generation, XLS support, style API), but the default XLSX path avoids POI's Sheet/Row/Cell overhead at runtime.

## The Mental Model Shift

Without this library, a spreadsheet is a **grid of cells**. You think in rows, columns, cell types, and indices.

```java
Row row = sheet.getRow(1);
String name = row.getCell(0).getStringCellValue();
int salary = (int) row.getCell(2).getNumericCellValue();
```

With this library, a spreadsheet is a **collection of typed objects**. You think in classes and fields.

```java
List<Employee> employees = mapper.readValues(file, Employee.class);
```

The grid doesn't disappear — it becomes an implementation detail. Like how JSON is "just text" but you never parse it character by character when you have Jackson.

## What the Library Handles

### Cell addressing → Field declaration

You don't specify which column maps to which field. The class structure *is* the schema.

```java
@DataGrid
class Employee {
    String name;      // column 0
    Address address;  // columns 1-2 (flattened)
    int salary;       // column 3
}
```

Add a field, a column appears. Remove a field, a column disappears. Reorder fields, columns reorder. The mapping is always in sync because there is no mapping — there is only the class.

Why this way? Because explicit column-index mapping is a maintenance liability. Every time the spreadsheet structure changes, someone has to update the mapping. With class-driven schema, the code and the schema are the same thing.

### Type coercion → Type declaration

You don't call `getStringCellValue()` or `getNumericCellValue()`. You declare a field type, and the value arrives as that type.

```java
String name;           // string cell -> String
int salary;            // numeric cell -> int
LocalDate startDate;   // Excel serial number -> LocalDate
boolean active;        // boolean cell -> boolean
```

Wrong type in the spreadsheet? It fails at deserialization — the same way a malformed JSON value fails. No silent `ClassCastException` from `getNumericCellValue()` on a string cell.

Why fail-fast? Because silent type coercion hides data quality problems. A string in a numeric column should break loudly at import time, not propagate through the system as zero.

### Nesting → Flattening/Reconstruction

You don't flatten objects into columns manually. You don't reconstruct them from columns manually.

```
Spreadsheet:  | name  | address/city | address/zip | salary |
              | Alice | Seoul        | 12345       | 80000  |

POJO:         Employee { name="Alice", address=Address { city="Seoul", zip="12345" }, salary=80000 }
```

Why automate this? Because nested objects are the norm in Java, but spreadsheets are flat. Every other Excel library forces you to manually bridge this gap — either flattening on write or reconstructing on read. This library does both directions automatically via Jackson's token model.

### Format I/O → Format detection

You don't choose between XLSX and XLS parsers. You pass a file.

```java
mapper.readValues(file, Employee.class);  // XLSX or XLS — auto-detected
```

Why auto-detect? Because the user shouldn't care about the container format. The data is the same. XLSX uses StAX streaming internally; XLS uses POI object model. The API is identical.

### Shared strings, styles, packaging → Configuration

You don't manage shared string tables, zip entries, or XML namespaces. You optionally configure styles and features.

```java
// Default: streaming, in-memory shared strings
mapper.writeValue(file, employees, Employee.class);

// Optional: file-backed shared strings for constant heap
SpreadsheetMapper mapper = SpreadsheetMapper.builder()
    .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
    .build();
```

Why hide this? Because OOXML packaging is an implementation detail of XLSX, not a user concern. Shared string tables exist for file size optimization. ZIP entry layout exists for OOXML compliance. Neither should leak into the data binding API.

## Where the Library Stops

The library draws a deliberate line. These are outside its scope entirely:

| Not handled | Why not |
|-------------|---------|
| **Workbook lifecycle** | Multi-sheet, template, formula — too many variations to generalize without constraining the user |
| **Charts, pivot tables, macros** | Outside data binding scope — these are document features, not data features |
| **Cell-level random access** | Contradicts the streaming model that enables performance |
| **Validation rules** | Business logic belongs in the application, not the format layer |

Sheet selection is a different case — the library provides `SheetInput` and `SheetOutput` for it, but the decision is explicit, not defaulted. The default (first sheet) works without any wrapper; naming or indexing a sheet is an opt-in choice.

These are not limitations. They are boundaries. POI handles everything on the other side. The library provides `Sheet` as an I/O type so the two coexist without friction:

```java
try (XSSFWorkbook wb = new XSSFWorkbook(template)) {
    mapper.writeValue(wb.getSheetAt(0), data, Type.class);  // library: data binding
    wb.getSheetAt(1).getRow(0).createCell(0).setCellFormula("SUM(A2:A100)");  // POI: formula
    wb.write(out);  // POI: workbook lifecycle
}
```

Why not abstract these too? Because over-abstraction locks users out. A library that tries to handle formulas, charts, and macros through annotations would either be incomplete or become a second POI. By stopping at data binding and providing `Sheet` as an escape hatch, the library stays small and POI stays accessible.

## The Litmus Test

If you can do it with Jackson for JSON, you can do it with this library for Excel:

| JSON (Jackson) | Excel (this library) |
|----------------|---------------------|
| `mapper.readValue(file, T.class)` | `mapper.readValue(file, T.class)` |
| `mapper.writeValue(file, obj)` | `mapper.writeValue(file, obj)` |
| `@JsonProperty("name")` | `@JsonProperty("name")` |
| `@JsonIgnore` | `@JsonIgnore` |
| `@JsonView(Summary.class)` | `@JsonView(Summary.class)` |
| Custom `JsonSerializer<T>` | Custom `JsonSerializer<T>` |
| MixIn for third-party class | MixIn for third-party class |

Same annotations. Same API. Different format. That's the abstraction.

If it doesn't pass this test — if the user has to learn a new API or write spreadsheet-specific code for something Jackson already handles — it's a design failure.
