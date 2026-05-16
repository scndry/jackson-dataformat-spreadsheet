package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.OptBoolean;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonPropertyOrder;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import com.fasterxml.jackson.annotation.JsonView;
import com.fasterxml.jackson.databind.MapperFeature;
import com.fasterxml.jackson.databind.PropertyNamingStrategies;
import com.fasterxml.jackson.databind.annotation.JsonNaming;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;

/**
 * PoC smoke tests for {@link DataColumnGroup} multi-row header support.
 * Covers 1-depth (Employee + Address) and 2-depth (Company + Year + Quarter)
 * via the POI User Model write path.
 */
class DataColumnGroupTest {

    SpreadsheetMapper mapper;
    @TempDir Path tempDir;

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
    }

    // -- 1-depth ---------------------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Employee {
        @DataColumn("ID") int id;
        @DataColumn("Name") String name;
        @DataColumnGroup("Address") Address address;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Address {
        @DataColumn("City") String city;
        @DataColumn("Zip") String zip;
    }

    @Test
    void singleDepthGroup_headerLayoutAndMerge() throws Exception {
        File file = tempFile("group-1depth.xlsx");

        Employee value = new Employee(1, "Alice", new Address("Seoul", "12345"));
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);

            // Two header rows. Each column writes its text at row = origin + hierarchy depth.
            //   ID/Name (h=0):     row 0 (top of vertical merge → actually rendered)
            //   Address (group):    row 0 (horizontal merge across cols 2-3)
            //   City/Zip (h=1):    row 1 (leaf)
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("ID");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Name");
            assertThat(sheet.getRow(0).getCell(2).getStringCellValue()).isEqualTo("Address");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("City");
            assertThat(sheet.getRow(1).getCell(3).getStringCellValue()).isEqualTo("Zip");

            // Data row at row 2
            assertThat((int) sheet.getRow(2).getCell(0).getNumericCellValue()).isEqualTo(1);
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Alice");
            assertThat(sheet.getRow(2).getCell(2).getStringCellValue()).isEqualTo("Seoul");
            assertThat(sheet.getRow(2).getCell(3).getStringCellValue()).isEqualTo("12345");

            // Merges:
            //   - "Address" horizontal merge (row 0, cols 2-3)
            //   - "ID" vertical merge (rows 0-1, col 0)
            //   - "Name" vertical merge (rows 0-1, col 1)
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(2);
                assertThat(r.getLastColumn()).isEqualTo(3);
            });
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastColumn()).isEqualTo(0);
            });
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(1);
            });
        }
    }

    // -- 2-depth ---------------------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Company {
        @DataColumn("Company") String company;
        @DataColumnGroup("2024") YearMetrics year2024;
        @DataColumnGroup("2025") YearMetrics year2025;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class YearMetrics {
        @DataColumnGroup("Q1") QuarterMetrics q1;
        @DataColumnGroup("Q2") QuarterMetrics q2;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class QuarterMetrics {
        @DataColumn("Sales") int sales;
        @DataColumn("Profit") int profit;
    }

    @Test
    void twoDepthGroup_headerLayoutAndMerge() throws Exception {
        File file = tempFile("group-2depth.xlsx");

        Company value = new Company("Acme",
                new YearMetrics(new QuarterMetrics(100, 20), new QuarterMetrics(110, 22)),
                new YearMetrics(new QuarterMetrics(120, 25), new QuarterMetrics(130, 28)));
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);

            // Three header rows. Each column writes its text at row = origin + hierarchy depth.
            //   Company (h=0):        row 0 (top of A1:A3 vertical merge)
            //   2024/2025 (groups):   row 0 (horizontal merge across child cols)
            //   Q1/Q2 (groups):       row 1
            //   Sales/Profit (h=2):   row 2 (leaf)
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Company");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("2024");
            assertThat(sheet.getRow(0).getCell(5).getStringCellValue()).isEqualTo("2025");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("Q1");
            assertThat(sheet.getRow(1).getCell(3).getStringCellValue()).isEqualTo("Q2");
            assertThat(sheet.getRow(1).getCell(5).getStringCellValue()).isEqualTo("Q1");
            assertThat(sheet.getRow(1).getCell(7).getStringCellValue()).isEqualTo("Q2");
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Sales");
            assertThat(sheet.getRow(2).getCell(2).getStringCellValue()).isEqualTo("Profit");
            assertThat(sheet.getRow(2).getCell(8).getStringCellValue()).isEqualTo("Profit");

            // Data row 3
            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("Acme");
            assertThat((int) sheet.getRow(3).getCell(1).getNumericCellValue()).isEqualTo(100);
            assertThat((int) sheet.getRow(3).getCell(8).getNumericCellValue()).isEqualTo(28);

            List<CellRangeAddress> merged = sheet.getMergedRegions();

            // "2024" group: row 0, cols 1-4
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(4);
            });
            // "2025" group: row 0, cols 5-8
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(5);
                assertThat(r.getLastColumn()).isEqualTo(8);
            });
            // "Q1" under 2024: row 1, cols 1-2
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(1);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(2);
            });
            // "Q1" under 2025: row 1, cols 5-6 (separate cell, despite same name — different parent)
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(1);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(5);
                assertThat(r.getLastColumn()).isEqualTo(6);
            });
            // "Company" flat column: rows 0-2, col 0 (vertical merge spanning all header rows)
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(2);
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastColumn()).isEqualTo(0);
            });
        }
    }

    // -- SSML path (default streaming) -----------------------------------

    @Test
    void singleDepthGroup_ssml() throws Exception {
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();
        File file = tempFile("group-1depth-ssml.xlsx");

        Employee value = new Employee(1, "Alice", new Address("Seoul", "12345"));
        ssmlMapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);

            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("ID");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Name");
            assertThat(sheet.getRow(0).getCell(2).getStringCellValue()).isEqualTo("Address");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("City");
            assertThat(sheet.getRow(1).getCell(3).getStringCellValue()).isEqualTo("Zip");

            assertThat((int) sheet.getRow(2).getCell(0).getNumericCellValue()).isEqualTo(1);
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Alice");

            List<CellRangeAddress> merged = sheet.getMergedRegions();
            // Address horizontal merge (row 0, cols 2-3)
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(2);
                assertThat(r.getLastColumn()).isEqualTo(3);
            });
            // ID vertical merge (rows 0-1, col 0)
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastColumn()).isEqualTo(0);
            });
            // Name vertical merge (rows 0-1, col 1)
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(1);
            });
        }
    }

    @Test
    void twoDepthGroup_ssml() throws Exception {
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();
        File file = tempFile("group-2depth-ssml.xlsx");

        Company value = new Company("Acme",
                new YearMetrics(new QuarterMetrics(100, 20), new QuarterMetrics(110, 22)),
                new YearMetrics(new QuarterMetrics(120, 25), new QuarterMetrics(130, 28)));
        ssmlMapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);

            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("Company");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("2024");
            assertThat(sheet.getRow(0).getCell(5).getStringCellValue()).isEqualTo("2025");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("Q1");
            assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Sales");

            assertThat(sheet.getRow(3).getCell(0).getStringCellValue()).isEqualTo("Acme");
            assertThat((int) sheet.getRow(3).getCell(1).getNumericCellValue()).isEqualTo(100);

            List<CellRangeAddress> merged = sheet.getMergedRegions();
            // 2024 horizontal merge (row 0, cols 1-4)
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(4);
            });
            // Company vertical merge (rows 0-2, col 0)
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(2);
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastColumn()).isEqualTo(0);
            });
        }
    }

    // -- Round-trip (write → read) ---------------------------------------

    @Test
    void singleDepthGroup_roundTrip_poi() throws Exception {
        File file = tempFile("group-1depth-rt.xlsx");
        Employee in = new Employee(1, "Alice", new Address("Seoul", "12345"));
        mapper.writeValue(file, in);
        Employee out = mapper.readValue(file, Employee.class);
        assertThat(out).usingRecursiveComparison().isEqualTo(in);
    }

    @Test
    void twoDepthGroup_roundTrip_poi() throws Exception {
        File file = tempFile("group-2depth-rt.xlsx");
        Company in = new Company("Acme",
                new YearMetrics(new QuarterMetrics(100, 20), new QuarterMetrics(110, 22)),
                new YearMetrics(new QuarterMetrics(120, 25), new QuarterMetrics(130, 28)));
        mapper.writeValue(file, in);
        Company out = mapper.readValue(file, Company.class);
        assertThat(out).usingRecursiveComparison().isEqualTo(in);
    }

    @Test
    void singleDepthGroup_roundTrip_ssml() throws Exception {
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();
        File file = tempFile("group-1depth-rt-ssml.xlsx");
        Employee in = new Employee(1, "Alice", new Address("Seoul", "12345"));
        ssmlMapper.writeValue(file, in);
        Employee out = ssmlMapper.readValue(file, Employee.class);
        assertThat(out).usingRecursiveComparison().isEqualTo(in);
    }

    @Test
    void twoDepthGroup_roundTrip_ssml() throws Exception {
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();
        File file = tempFile("group-2depth-rt-ssml.xlsx");
        Company in = new Company("Acme",
                new YearMetrics(new QuarterMetrics(100, 20), new QuarterMetrics(110, 22)),
                new YearMetrics(new QuarterMetrics(120, 25), new QuarterMetrics(130, 28)));
        ssmlMapper.writeValue(file, in);
        Company out = ssmlMapper.readValue(file, Company.class);
        assertThat(out).usingRecursiveComparison().isEqualTo(in);
    }

    // -- Reorder + multi-row header --------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class EmployeeWriteOrder {
        @DataColumn("Name") String name;
        @DataColumn("ID") int id;
        @DataColumnGroup("Address") Address address;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class EmployeeReadOrder {
        @DataColumn("ID") int id;
        @DataColumn("Name") String name;
        @DataColumnGroup("Address") Address address;
    }

    @Test
    void multiRowHeader_reorder_failsFast_poi() throws Exception {
        File file = tempFile("group-reorder.xlsx");
        EmployeeWriteOrder writeValue = new EmployeeWriteOrder(
                "Alice", 1, new Address("Seoul", "12345"));
        mapper.writeValue(file, writeValue);

        SpreadsheetMapper reorderMapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL))
                .setColumnReordering(true);

        assertThatThrownBy(() -> reorderMapper.readValue(file, EmployeeReadOrder.class))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("Column reordering is not supported with multi-row headers");
    }

    @Test
    void multiRowHeader_reorder_failsFast_ssml() throws Exception {
        SpreadsheetMapper writeMapper = new SpreadsheetMapper();
        File file = tempFile("group-reorder-ssml.xlsx");
        EmployeeWriteOrder writeValue = new EmployeeWriteOrder(
                "Alice", 1, new Address("Seoul", "12345"));
        writeMapper.writeValue(file, writeValue);

        SpreadsheetMapper reorderMapper = new SpreadsheetMapper().setColumnReordering(true);

        assertThatThrownBy(() -> reorderMapper.readValue(file, EmployeeReadOrder.class))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("Column reordering is not supported with multi-row headers");
    }

    // -- Group cell comment ----------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class EmployeeWithGroupComment {
        @DataColumn("ID") int id;
        @DataColumn("Name") String name;
        @DataColumnGroup(value = "Address", comment = "Customer billing address") Address address;
    }

    @Test
    void groupCellComment_poi() throws Exception {
        File file = tempFile("group-comment.xlsx");

        EmployeeWithGroupComment value = new EmployeeWithGroupComment(
                1, "Alice", new Address("Seoul", "12345"));
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Group cell "Address" at (row 0, col 2)
            Cell cell = sheet.getRow(0).getCell(2);
            assertThat(cell.getStringCellValue()).isEqualTo("Address");
            Comment comment = cell.getCellComment();
            assertThat(comment).isNotNull();
            assertThat(comment.getString().getString())
                    .isEqualTo("Customer billing address");
        }
    }

    @Test
    void groupCellComment_ssml() throws Exception {
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();
        File file = tempFile("group-comment-ssml.xlsx");

        EmployeeWithGroupComment value = new EmployeeWithGroupComment(
                1, "Alice", new Address("Seoul", "12345"));
        ssmlMapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(0).getCell(2);
            assertThat(cell.getStringCellValue()).isEqualTo("Address");
            Comment comment = cell.getCellComment();
            assertThat(comment).isNotNull();
            assertThat(comment.getString().getString())
                    .isEqualTo("Customer billing address");
        }
    }

    // -- value default fallback (field name) -----------------------------

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class EmployeeImplicitGroup {
        @DataColumn("ID") int id;
        @DataColumn("Name") String name;
        @DataColumnGroup Address address;   // no value — falls back to field name
    }

    @Test
    void groupValueFallsBackToFieldName() throws Exception {
        File file = tempFile("group-implicit.xlsx");
        EmployeeImplicitGroup value = new EmployeeImplicitGroup(
                1, "Alice", new Address("Seoul", "12345"));
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Group cell at (row 0, col 2) should carry the field name "address"
            assertThat(sheet.getRow(0).getCell(2).getStringCellValue()).isEqualTo("address");
        }
    }

    // -- Jackson annotation interop --------------------------------------

    @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
    @JsonSubTypes({
            @JsonSubTypes.Type(value = Dog.class, name = "dog"),
            @JsonSubTypes.Type(value = Cat.class, name = "cat")
    })
    static class Animal {
        public String name;
        public Animal() {}
    }

    static class Dog extends Animal {
        public String breed;
        public Dog() {}
        public Dog(String name, String breed) { this.name = name; this.breed = breed; }
    }

    static class Cat extends Animal {
        public boolean indoor;
        public Cat() {}
        public Cat(String name, boolean indoor) { this.name = name; this.indoor = indoor; }
    }

    @DataGrid
    static class OwnerWithPet {
        public String owner;
        @DataColumnGroup("Pet") public Animal pet;
        public OwnerWithPet() {}
        public OwnerWithPet(String owner, Animal pet) { this.owner = owner; this.pet = pet; }
    }

    @Test
    void polymorphicWithGroup() throws Exception {
        File file = tempFile("group-polymorphic.xlsx");
        mapper.writeValue(file, Arrays.asList(
                new OwnerWithPet("Alice", new Dog("Rex", "Labrador")),
                new OwnerWithPet("Bob", new Cat("Whiskers", true))),
                OwnerWithPet.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Two header rows: row 0 (group + flat top), row 1 (leaf).
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("owner");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Pet");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("type");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("pet/name");
            assertThat(sheet.getRow(1).getCell(3).getStringCellValue()).isEqualTo("pet/breed");
            assertThat(sheet.getRow(1).getCell(4).getStringCellValue()).isEqualTo("pet/indoor");

            // "Pet" group spans cols 1..4
            List<CellRangeAddress> merged = sheet.getMergedRegions();
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(4);
            });

            Row dogRow = sheet.getRow(2);
            assertThat(dogRow.getCell(0).getStringCellValue()).isEqualTo("Alice");
            assertThat(dogRow.getCell(1).getStringCellValue()).isEqualTo("dog");
            assertThat(dogRow.getCell(2).getStringCellValue()).isEqualTo("Rex");
            assertThat(dogRow.getCell(3).getStringCellValue()).isEqualTo("Labrador");
        }
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class PlainAddress {
        public String city;
        public String zip;
    }

    // -- @JsonView + @DataColumnGroup -----------------------------------

    static class GroupViews {
        static class Summary {}
        static class Detail extends Summary {}
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ViewedDetails {
        @JsonView(GroupViews.Summary.class) public String code;
        @JsonView(GroupViews.Detail.class) public String description;
    }

    @DataGrid
    static class WithViewedGroup {
        @JsonView(GroupViews.Summary.class) public String name;
        @DataColumnGroup("Details")
        @JsonView(GroupViews.Summary.class)
        public ViewedDetails details;
        public WithViewedGroup() {}
        public WithViewedGroup(String name, ViewedDetails details) {
            this.name = name;
            this.details = details;
        }
    }

    /**
     * Verifies {@code @JsonView} filtering inside a group: under the
     * Summary view only Summary-tagged inner fields appear, and the group
     * header still spans the visible columns.
     */
    @Test
    void jsonViewWithGroup_summaryOnly() throws Exception {
        SpreadsheetMapper viewMapper = SpreadsheetMapper.builder()
                .configure(MapperFeature.DEFAULT_VIEW_INCLUSION, false)
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();

        File file = tempFile("group-view.xlsx");
        viewMapper.sheetWriterForWithView(WithViewedGroup.class, GroupViews.Summary.class)
                .writeValue(file, Arrays.asList(
                        new WithViewedGroup("A", new ViewedDetails("X", "long"))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Two header rows: outer "name" + group "Details", with single inner leaf.
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Details");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("details/code");
            // "description" is filtered out.
            assertThat(sheet.getRow(0).getCell(2)).isNull();

            // Single-column group: no horizontal merge is created (library skips
            // 1-cell merges in visitGroupCell). Vertical merge for the flat
            // "name" column at col 0 still spans both header rows.
            assertThat(sheet.getMergedRegions()).noneSatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(1);
            });
            assertThat(sheet.getMergedRegions()).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastColumn()).isEqualTo(0);
            });

            Row dataRow = sheet.getRow(2);
            assertThat(dataRow.getCell(0).getStringCellValue()).isEqualTo("A");
            assertThat(dataRow.getCell(1).getStringCellValue()).isEqualTo("X");
        }
    }

    @Test
    void jsonViewWithGroup_detailIncludesAll() throws Exception {
        SpreadsheetMapper viewMapper = SpreadsheetMapper.builder()
                .configure(MapperFeature.DEFAULT_VIEW_INCLUSION, false)
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();

        File file = tempFile("group-view-detail.xlsx");
        viewMapper.sheetWriterForWithView(WithViewedGroup.class, GroupViews.Detail.class)
                .writeValue(file, Arrays.asList(
                        new WithViewedGroup("A", new ViewedDetails("X", "long"))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Detail extends Summary — both inner fields visible, group spans cols 1..2.
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Details");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("details/code");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("details/description");

            assertThat(sheet.getMergedRegions()).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(2);
            });

            Row dataRow = sheet.getRow(2);
            assertThat(dataRow.getCell(0).getStringCellValue()).isEqualTo("A");
            assertThat(dataRow.getCell(1).getStringCellValue()).isEqualTo("X");
            assertThat(dataRow.getCell(2).getStringCellValue()).isEqualTo("long");
        }
    }

    // -- @JsonNaming + @DataColumnGroup ---------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class CamelInner {
        public String streetName;
        public String postalCode;
    }

    @DataGrid
    @JsonNaming(PropertyNamingStrategies.SnakeCaseStrategy.class)
    static class WithSnakeCaseGroup {
        public String userName;
        @DataColumnGroup("Address")
        public CamelInner homeAddress;
        public WithSnakeCaseGroup() {}
        public WithSnakeCaseGroup(String userName, CamelInner homeAddress) {
            this.userName = userName;
            this.homeAddress = homeAddress;
        }
    }

    /**
     * {@code @JsonNaming} is per-class — outer property names get the
     * strategy applied (snake_case), but the inner class needs its own
     * {@code @JsonNaming} to convert its fields. The group header value
     * comes from the {@code @DataColumnGroup} literal and is untouched.
     */
    @Test
    void jsonNamingWithGroup_outerStrategyOnly() throws Exception {
        File file = tempFile("group-naming.xlsx");
        mapper.writeValue(file, new WithSnakeCaseGroup(
                "Alice", new CamelInner("Main", "12345")));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Outer: snake_case applied. Group label: literal from annotation.
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("user_name");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Address");
            // Inner leaf path: outer's @JsonNaming reaches the outer property
            // ("home_address"); inner field names stay camelCase because
            // CamelInner has no @JsonNaming of its own.
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("home_address/streetName");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("home_address/postalCode");

            Row dataRow = sheet.getRow(2);
            assertThat(dataRow.getCell(0).getStringCellValue()).isEqualTo("Alice");
            assertThat(dataRow.getCell(1).getStringCellValue()).isEqualTo("Main");
            assertThat(dataRow.getCell(2).getStringCellValue()).isEqualTo("12345");
        }
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    @JsonNaming(PropertyNamingStrategies.SnakeCaseStrategy.class)
    static class SnakeInner {
        public String streetName;
        public String postalCode;
    }

    @DataGrid
    @JsonNaming(PropertyNamingStrategies.SnakeCaseStrategy.class)
    static class WithSnakeCaseGroupBoth {
        public String userName;
        @DataColumnGroup("Address")
        public SnakeInner homeAddress;
        public WithSnakeCaseGroupBoth() {}
        public WithSnakeCaseGroupBoth(String userName, SnakeInner homeAddress) {
            this.userName = userName;
            this.homeAddress = homeAddress;
        }
    }

    @Test
    void jsonNamingWithGroup_appliedToInnerWhenAnnotated() throws Exception {
        File file = tempFile("group-naming-both.xlsx");
        mapper.writeValue(file, new WithSnakeCaseGroupBoth(
                "Alice", new SnakeInner("Main", "12345")));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("user_name");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Address");
            // Inner with its own @JsonNaming → snake_case applied here too.
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("home_address/street_name");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("home_address/postal_code");

            Row dataRow = sheet.getRow(2);
            assertThat(dataRow.getCell(0).getStringCellValue()).isEqualTo("Alice");
            assertThat(dataRow.getCell(1).getStringCellValue()).isEqualTo("Main");
            assertThat(dataRow.getCell(2).getStringCellValue()).isEqualTo("12345");
        }
    }

    // -- @JsonPropertyOrder + @DataColumnGroup --------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    @JsonPropertyOrder({"surname", "givenName"})
    static class OrderedInner {
        public String givenName;
        public String surname;
    }

    @DataGrid
    @JsonPropertyOrder({"score", "details", "id"})
    static class WithOrderAndGroup {
        public int id;
        @DataColumnGroup("Details")
        public OrderedInner details;
        public int score;
        public WithOrderAndGroup() {}
        public WithOrderAndGroup(int id, OrderedInner details, int score) {
            this.id = id;
            this.details = details;
            this.score = score;
        }
    }

    /**
     * Outer {@code @JsonPropertyOrder} controls the position of the group
     * field; inner {@code @JsonPropertyOrder} controls the order of leaf
     * columns inside the group.
     */
    @Test
    void jsonPropertyOrderWithGroup_innerAndOuter() throws Exception {
        File file = tempFile("group-order.xlsx");
        mapper.writeValue(file, new WithOrderAndGroup(
                7, new OrderedInner("Alice", "Kim"), 99));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Outer order: score, details (group spans cols 1..2), id.
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("score");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Details");
            assertThat(sheet.getRow(0).getCell(3).getStringCellValue()).isEqualTo("id");
            // Inner order: surname before givenName.
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("details/surname");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("details/givenName");

            assertThat(sheet.getMergedRegions()).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(2);
            });

            Row dataRow = sheet.getRow(2);
            assertThat((int) dataRow.getCell(0).getNumericCellValue()).isEqualTo(99);
            assertThat(dataRow.getCell(1).getStringCellValue()).isEqualTo("Kim");
            assertThat(dataRow.getCell(2).getStringCellValue()).isEqualTo("Alice");
            assertThat((int) dataRow.getCell(3).getNumericCellValue()).isEqualTo(7);
        }
    }

    // -- Mix-in + @DataColumnGroup --------------------------------------

    static class ThirdPartyOrder {
        public int id;
        public PlainAddress address;
        public ThirdPartyOrder() {}
        public ThirdPartyOrder(int id, PlainAddress address) {
            this.id = id;
            this.address = address;
        }
    }

    @DataGrid
    abstract static class ThirdPartyOrderMixin {
        @JsonProperty int id;
        @DataColumnGroup("Address") PlainAddress address;
    }

    /**
     * Verifies {@code @DataColumnGroup} can be applied via Jackson Mix-in
     * to a third-party class that the user can't modify directly. The
     * group header reaches the visitor through Jackson's annotation
     * introspection.
     */
    @Test
    void mixinWithGroup_appliesGroupHeader() throws Exception {
        mapper.addMixIn(ThirdPartyOrder.class, ThirdPartyOrderMixin.class);

        File file = tempFile("group-mixin.xlsx");
        mapper.writeValue(file, new ThirdPartyOrder(
                1, new PlainAddress("Seoul", "12345")));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("id");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Address");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("address/city");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("address/zip");

            assertThat(sheet.getMergedRegions()).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(2);
            });

            Row dataRow = sheet.getRow(2);
            assertThat((int) dataRow.getCell(0).getNumericCellValue()).isEqualTo(1);
            assertThat(dataRow.getCell(1).getStringCellValue()).isEqualTo("Seoul");
            assertThat(dataRow.getCell(2).getStringCellValue()).isEqualTo("12345");
        }
    }

    // -- Collection<NestedType> + @DataColumnGroup ---------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class LineItem {
        public String sku;
        public int qty;
    }

    @DataGrid
    static class OrderWithItemList {
        public int id;
        @DataColumnGroup("Items")
        public List<LineItem> items;
        public OrderWithItemList() {}
        public OrderWithItemList(int id, List<LineItem> items) {
            this.id = id;
            this.items = items;
        }
    }

    /**
     * Verifies the {@code @DataColumnGroup} javadoc claim that
     * {@code List<NestedType>} fields are supported — the group header
     * spans every flattened column produced by the element type.
     */
    @Test
    void collectionOfNestedWithGroup() throws Exception {
        File file = tempFile("group-collection.xlsx");
        mapper.writeValue(file, new OrderWithItemList(
                1, Arrays.asList(new LineItem("A", 2), new LineItem("B", 5))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("id");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Items");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("items/[]/sku");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("items/[]/qty");

            assertThat(sheet.getMergedRegions()).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(2);
            });
        }
    }

    // -- 3-depth nesting -----------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class TeamSection {
        public String teamLead;
        public int memberCount;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class DepartmentSection {
        public String deptName;
        @DataColumnGroup("Team") public TeamSection team;
    }

    @DataGrid
    static class CountryReport {
        public String country;
        @DataColumnGroup("Department") public DepartmentSection dept;
        public CountryReport() {}
        public CountryReport(String country, DepartmentSection dept) {
            this.country = country;
            this.dept = dept;
        }
    }

    /**
     * Verifies the multi-row header algorithm scales to 3-depth nesting:
     * outer flat column, mid-level group, inner group with leaves.
     * Header row count = max group depth + 1 = 3.
     */
    @Test
    void threeDepthNesting_layoutAndMerges() throws Exception {
        File file = tempFile("group-3depth.xlsx");
        mapper.writeValue(file, new CountryReport(
                "KR", new DepartmentSection("Eng",
                        new TeamSection("Alice", 7))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);

            // Row 0: top-level — flat column "country" + outermost group "Department".
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("country");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Department");

            // Row 1: mid-level — leaf inside Department, plus inner group "Team".
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("dept/deptName");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("Team");

            // Row 2: leaves inside Team.
            assertThat(sheet.getRow(2).getCell(2).getStringCellValue()).isEqualTo("dept/team/teamLead");
            assertThat(sheet.getRow(2).getCell(3).getStringCellValue()).isEqualTo("dept/team/memberCount");

            List<CellRangeAddress> merged = sheet.getMergedRegions();
            // "country" vertical merge across all 3 header rows (col 0).
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(2);
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastColumn()).isEqualTo(0);
            });
            // "Department" horizontal merge cols 1..3 at row 0.
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(3);
            });
            // "dept/deptName" vertical merge rows 1..2 (col 1).
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(1);
                assertThat(r.getLastRow()).isEqualTo(2);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(1);
            });
            // "Team" horizontal merge cols 2..3 at row 1.
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(1);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(2);
                assertThat(r.getLastColumn()).isEqualTo(3);
            });

            // Data row at row 3.
            Row dataRow = sheet.getRow(3);
            assertThat(dataRow.getCell(0).getStringCellValue()).isEqualTo("KR");
            assertThat(dataRow.getCell(1).getStringCellValue()).isEqualTo("Eng");
            assertThat(dataRow.getCell(2).getStringCellValue()).isEqualTo("Alice");
            assertThat((int) dataRow.getCell(3).getNumericCellValue()).isEqualTo(7);
        }
    }

    // -- @JsonIgnore inner field + @DataColumnGroup --------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ProfileWithSecret {
        public String name;
        @JsonIgnore public String passwordHash;
        public int age;
    }

    @DataGrid
    static class UserWithProfile {
        public int id;
        @DataColumnGroup("Profile")
        public ProfileWithSecret profile;
        public UserWithProfile() {}
        public UserWithProfile(int id, ProfileWithSecret profile) {
            this.id = id;
            this.profile = profile;
        }
    }

    /**
     * Verifies {@code @JsonIgnore} on an inner field excludes that field
     * from the group: the group header still spans only the visible
     * leaf columns, with the ignored field absent.
     */
    @Test
    void jsonIgnoreInnerField_groupSpansVisibleOnly() throws Exception {
        File file = tempFile("group-ignore.xlsx");
        mapper.writeValue(file, new UserWithProfile(
                1, new ProfileWithSecret("Alice", "secret", 30)));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("id");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Profile");
            // Only "name" and "age" appear under Profile.
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("profile/name");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("profile/age");
            assertThat(sheet.getRow(1).getCell(3)).isNull();

            // Group "Profile" spans cols 1..2 (passwordHash filtered out).
            assertThat(sheet.getMergedRegions()).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(2);
            });

            Row dataRow = sheet.getRow(2);
            assertThat((int) dataRow.getCell(0).getNumericCellValue()).isEqualTo(1);
            assertThat(dataRow.getCell(1).getStringCellValue()).isEqualTo("Alice");
            assertThat((int) dataRow.getCell(2).getNumericCellValue()).isEqualTo(30);
        }
    }

    // -- @JsonProperty rename inner field + @DataColumnGroup -----------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class RenamedDetails {
        @JsonProperty("Display Code") public String code;
        public int qty;
    }

    @DataGrid
    static class WithRenamedInner {
        public int id;
        @DataColumnGroup("Item")
        public RenamedDetails item;
        public WithRenamedInner() {}
        public WithRenamedInner(int id, RenamedDetails item) {
            this.id = id;
            this.item = item;
        }
    }

    /**
     * Verifies {@code @JsonProperty} on an inner field changes the leaf
     * column path inside the group. Jackson's renamed property name flows
     * through to the pointer used as the column header fallback.
     */
    @Test
    void jsonPropertyRenameInner_appliesToLeafPath() throws Exception {
        File file = tempFile("group-rename.xlsx");
        mapper.writeValue(file, new WithRenamedInner(
                1, new RenamedDetails("ABC", 3)));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("id");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Item");

            // Both leaf columns appear under the group; rename applied to "code".
            // Order is decided by Jackson's bean introspection and is not asserted here.
            Set<String> leafNames = new HashSet<>();
            leafNames.add(sheet.getRow(1).getCell(1).getStringCellValue());
            leafNames.add(sheet.getRow(1).getCell(2).getStringCellValue());
            assertThat(leafNames).containsExactlyInAnyOrder(
                    "item/Display Code", "item/qty");

            assertThat(sheet.getMergedRegions()).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(2);
            });

            // Data row at row 2: id at col 0; inner cells at cols 1/2 in
            // whichever order Jackson emitted the leaves. Pair leaves to
            // their data values to verify rename round-trip.
            Row dataRow = sheet.getRow(2);
            assertThat((int) dataRow.getCell(0).getNumericCellValue()).isEqualTo(1);
            Map<String, String> leafToData = new HashMap<>();
            for (int c = 1; c <= 2; c++) {
                String header = sheet.getRow(1).getCell(c).getStringCellValue();
                Cell cell = dataRow.getCell(c);
                String value = cell.getCellType() == CellType.STRING
                        ? cell.getStringCellValue()
                        : String.valueOf((int) cell.getNumericCellValue());
                leafToData.put(header, value);
            }
            assertThat(leafToData).containsEntry("item/Display Code", "ABC");
            assertThat(leafToData).containsEntry("item/qty", "3");
        }
    }

    // -- @JsonInclude(NON_NULL) inner + @DataColumnGroup ---------------

    @Data @NoArgsConstructor @AllArgsConstructor
    @JsonInclude(JsonInclude.Include.NON_NULL)
    static class NonNullInner {
        public String present;
        public String optional;
    }

    @DataGrid
    static class WithNonNullGroup {
        public int id;
        @DataColumnGroup("Data")
        public NonNullInner data;
        public WithNonNullGroup() {}
        public WithNonNullGroup(int id, NonNullInner data) {
            this.id = id;
            this.data = data;
        }
    }

    /**
     * {@code @JsonInclude(NON_NULL)} on the inner type skips null fields
     * during serialization — the corresponding cell is empty. The schema
     * (header layout, group span) is unchanged because it is built from
     * the type, not from instance values.
     */
    @Test
    void jsonIncludeNonNullInner_skipsNullCellOnly() throws Exception {
        File file = tempFile("group-nonnull.xlsx");
        mapper.writeValue(file, new WithNonNullGroup(
                1, new NonNullInner("X", null)));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Headers unchanged: both inner properties always reserved.
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("id");
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Data");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("data/present");
            assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("data/optional");

            assertThat(sheet.getMergedRegions()).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(2);
            });

            // Data row: present has value, optional is skipped.
            Row dataRow = sheet.getRow(2);
            assertThat((int) dataRow.getCell(0).getNumericCellValue()).isEqualTo(1);
            assertThat(dataRow.getCell(1).getStringCellValue()).isEqualTo("X");
            assertThat(dataRow.getCell(2)).isNull();
        }
    }

    // -- group header style (headerStyle + DataGrid.groupHeaderStyle cascade) ----

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class CompanyExplicitGroupStyle {
        @DataColumn("Name") String name;
        @DataColumnGroup(value = "2024", headerStyle = "groupBg") YearMetricsSmall year2024;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class YearMetricsSmall {
        @DataColumn("Sales") int sales;
        @DataColumn("Profit") int profit;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    @DataGrid(groupHeaderStyle = "groupBg")
    static class CompanyGridDefaultGroupStyle {
        @DataColumn("Name") String name;
        @DataColumnGroup("2024") YearMetricsSmall year2024;
    }

    @Test
    void groupHeaderStyle_explicit_appliedToGroupCellOnly_poi() throws Exception {
        File file = tempFile("group-headerstyle-explicit.xlsx");
        SpreadsheetMapper styled = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("groupBg")
                            .fillForegroundColor(IndexedColors.GREY_25_PERCENT)
                            .fillPattern().solidForeground()
                            .end())
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        styled.writeValue(file, new CompanyExplicitGroupStyle("Acme",
                new YearMetricsSmall(100, 20)));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Group cell at (0,1) carries the registered fill; leaf header
            // cells below it (row 1) carry the default style without fill.
            CellStyle groupStyle = sheet.getRow(0).getCell(1).getCellStyle();
            CellStyle leafStyle = sheet.getRow(1).getCell(1).getCellStyle();
            assertThat(groupStyle.getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
            assertThat(leafStyle.getFillPattern())
                    .isNotEqualTo(FillPatternType.SOLID_FOREGROUND);
        }
    }

    @Test
    void groupHeaderStyle_explicit_appliedToGroupCellOnly_ssml() throws Exception {
        SpreadsheetMapper styled = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("groupBg")
                            .fillForegroundColor(IndexedColors.GREY_25_PERCENT)
                            .fillPattern().solidForeground()
                            .end())
                .build();
        File file = tempFile("group-headerstyle-explicit-ssml.xlsx");
        styled.writeValue(file, new CompanyExplicitGroupStyle("Acme",
                new YearMetricsSmall(100, 20)));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            CellStyle groupStyle = sheet.getRow(0).getCell(1).getCellStyle();
            CellStyle leafStyle = sheet.getRow(1).getCell(1).getCellStyle();
            assertThat(groupStyle.getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
            assertThat(leafStyle.getFillPattern())
                    .isNotEqualTo(FillPatternType.SOLID_FOREGROUND);
        }
    }

    @Test
    void groupHeaderStyle_cascadesFromDataGrid() throws Exception {
        File file = tempFile("group-headerstyle-cascade.xlsx");
        SpreadsheetMapper styled = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("groupBg")
                            .fillForegroundColor(IndexedColors.GREY_25_PERCENT)
                            .fillPattern().solidForeground()
                            .end())
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        styled.writeValue(file, new CompanyGridDefaultGroupStyle("Acme",
                new YearMetricsSmall(100, 20)));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            CellStyle groupStyle = sheet.getRow(0).getCell(1).getCellStyle();
            assertThat(groupStyle.getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
        }
    }

    // -- group-level child defaults (columnStyle / columnHeaderStyle / width / merge) --

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class OrderWithGroupChildDefaults {
        @DataColumn("id") int id;
        @DataColumnGroup(value = "items", columnHeaderStyle = "itemHeaderBg")
        List<ItemSmall> items;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ItemSmall {
        @DataColumn("sku") String sku;
        @DataColumn("qty") int qty;
    }

    @Test
    void groupColumnHeaderStyle_appliesToChildLeafHeaders() throws Exception {
        File file = _debugFile("group-child-headerstyle-poi.xlsx");
        SpreadsheetMapper styled = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("itemHeaderBg")
                            .fillForegroundColor(IndexedColors.GREY_25_PERCENT)
                            .fillPattern().solidForeground()
                            .end())
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        styled.writeValue(file, new OrderWithGroupChildDefaults(1,
                Arrays.asList(new ItemSmall("A", 1), new ItemSmall("B", 2))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // id leaf header (row 0, col 0) is outside the group — no group fill.
            // id vertically merges across rows 0..1, so the text + style live at row 0.
            assertThat(sheet.getRow(0).getCell(0).getCellStyle().getFillPattern())
                    .isNotEqualTo(FillPatternType.SOLID_FOREGROUND);
            // sku / qty leaf headers (row 1, cols 1, 2) inherit the group's
            // columnHeaderStyle via the cascade.
            assertThat(sheet.getRow(1).getCell(1).getCellStyle().getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
            assertThat(sheet.getRow(1).getCell(2).getCellStyle().getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
        }
    }

    @Test
    void groupColumnHeaderStyle_appliesToChildLeafHeaders_ssml() throws Exception {
        File file = _debugFile("group-child-headerstyle-ssml.xlsx");
        SpreadsheetMapper styled = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("itemHeaderBg")
                            .fillForegroundColor(IndexedColors.GREY_25_PERCENT)
                            .fillPattern().solidForeground()
                            .end())
                .build();
        styled.writeValue(file, new OrderWithGroupChildDefaults(1,
                Arrays.asList(new ItemSmall("A", 1), new ItemSmall("B", 2))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(0).getCellStyle().getFillPattern())
                    .isNotEqualTo(FillPatternType.SOLID_FOREGROUND);
            assertThat(sheet.getRow(1).getCell(1).getCellStyle().getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
            assertThat(sheet.getRow(1).getCell(2).getCellStyle().getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
        }
    }

    // -- group columnStyle / columnWidth / mergeColumn render (POI + SSML) --

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class OrderWithGroupColumnStyle {
        @DataColumn("id") int id;
        @DataColumnGroup(value = "items", columnStyle = "itemDataBg")
        List<ItemSmall> items;
    }

    @Test
    void groupColumnStyle_appliesToChildDataCells_poi() throws Exception {
        File file = _debugFile("group-child-columnstyle-poi.xlsx");
        SpreadsheetMapper styled = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("itemDataBg")
                            .fillForegroundColor(IndexedColors.GREY_25_PERCENT)
                            .fillPattern().solidForeground()
                            .end())
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        styled.writeValue(file, new OrderWithGroupColumnStyle(1,
                Arrays.asList(new ItemSmall("A", 1))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Data row at row 2: id (col 0) has no fill; sku/qty (cols 1, 2) inherit columnStyle.
            assertThat(sheet.getRow(2).getCell(0).getCellStyle().getFillPattern())
                    .isNotEqualTo(FillPatternType.SOLID_FOREGROUND);
            assertThat(sheet.getRow(2).getCell(1).getCellStyle().getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
            assertThat(sheet.getRow(2).getCell(2).getCellStyle().getFillPattern())
                    .isEqualTo(FillPatternType.SOLID_FOREGROUND);
        }
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class OrderWithGroupWidth {
        @DataColumn("id") int id;
        @DataColumnGroup(value = "items", columnWidth = 17)
        List<ItemSmall> items;
    }

    @Test
    void groupColumnWidth_appliesToChildColumns() throws Exception {
        File file = _debugFile("group-child-width-ssml.xlsx");
        new SpreadsheetMapper().writeValue(file, new OrderWithGroupWidth(1,
                Arrays.asList(new ItemSmall("A", 1))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // POI stores width in 1/256 of a character — multiply request by 256.
            assertThat(sheet.getColumnWidth(1)).isEqualTo(17 * 256);
            assertThat(sheet.getColumnWidth(2)).isEqualTo(17 * 256);
        }
    }


    // -- leaf @DataColumn.style beats group columnStyle in real render --

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class OrderWithLeafStyleOverride {
        @DataColumn("id") int id;
        @DataColumnGroup(value = "items", columnStyle = "groupBg")
        List<ItemWithExplicit> items;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ItemWithExplicit {
        @DataColumn(value = "x", style = "leafBg") String x;
        @DataColumn("y") String y;
    }

    @Test
    void leafColumnStyle_winsOverGroupColumnStyle_inRender() throws Exception {
        File file = _debugFile("group-leaf-override-poi.xlsx");
        SpreadsheetMapper styled = SpreadsheetMapper.builder()
                .stylesBuilder(new StylesBuilder()
                        .cellStyle("groupBg")
                            .fillForegroundColor(IndexedColors.YELLOW)
                            .fillPattern().solidForeground()
                            .end()
                        .cellStyle("leafBg")
                            .fillForegroundColor(IndexedColors.GREY_25_PERCENT)
                            .fillPattern().solidForeground()
                            .end())
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        styled.writeValue(file, new OrderWithLeafStyleOverride(1,
                Arrays.asList(new ItemWithExplicit("a", "b"))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Data row 2: x (col 1) carries the leaf's style (GREY); y (col 2) inherits group (YELLOW).
            String xFill = ((XSSFCellStyle) sheet.getRow(2).getCell(1).getCellStyle())
                    .getFillForegroundXSSFColor().getARGBHex();
            String yFill = ((XSSFCellStyle) sheet.getRow(2).getCell(2).getCellStyle())
                    .getFillForegroundXSSFColor().getARGBHex();
            assertThat(xFill).endsWithIgnoringCase("C0C0C0");
            assertThat(yFill).endsWithIgnoringCase("FFFF00");
        }
    }

    private File tempFile(String name) {
        return tempDir.resolve(name).toFile();
    }

    // -- @DataColumnGroup empty value() falls back to column pointer path ----

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class FallbackRoot {
        @DataColumnGroup FallbackInner address;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class FallbackInner {
        @DataColumn("City") String city;
        @DataColumn("Zip") String zip;
    }

    @Test
    void groupValueEmpty_atRoot_usesFieldName() throws Exception {
        File file = tempFile("group-fallback-root.xlsx");

        mapper.writeValue(file, new FallbackRoot(new FallbackInner("Seoul", "12345")));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Root field's pointer is the bare field name.
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("address");
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("City");
            assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("Zip");
        }
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class FallbackOuter {
        @DataColumnGroup FallbackMid mid;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class FallbackMid {
        @DataColumnGroup FallbackLeaf leaf;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class FallbackLeaf {
        @DataColumn("v") int v;
    }

    @Test
    void groupValueEmpty_atNested_usesColumnPointerPath() throws Exception {
        File file = tempFile("group-fallback-nested.xlsx");

        mapper.writeValue(file, new FallbackOuter(new FallbackMid(new FallbackLeaf(42))));

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // Outer @DataColumnGroup with empty value → root pointer "mid".
            assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("mid");
            // Nested @DataColumnGroup with empty value → path "mid/leaf",
            // not the bare field name "leaf".
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("mid/leaf");
            assertThat(sheet.getRow(2).getCell(0).getStringCellValue()).isEqualTo("v");
            assertThat((int) sheet.getRow(3).getCell(0).getNumericCellValue()).isEqualTo(42);
        }
    }

    private static final Path DEBUG_OUTPUT_DIR = Paths.get("build/debug-output");

    private static File _debugFile(final String name) throws IOException {
        Files.createDirectories(DEBUG_OUTPUT_DIR);
        return DEBUG_OUTPUT_DIR.resolve(name).toFile();
    }
}
