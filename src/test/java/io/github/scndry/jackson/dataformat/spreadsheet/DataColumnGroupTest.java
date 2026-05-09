package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

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

    private File tempFile(String name) {
        return tempDir.resolve(name).toFile();
    }
}
