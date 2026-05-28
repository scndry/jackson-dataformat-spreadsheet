package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import com.fasterxml.jackson.annotation.JsonUnwrapped;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetFactory;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatCode;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Integration tests for {@code @DataColumn(shift)} and
 * {@code @DataColumnGroup(shift)} — verifies blank columns on write,
 * column skip on read, and round-trip equivalence under
 * {@code useHeader=false}.
 */
class DataColumnShiftTest {

    SpreadsheetMapper mapper;

    @TempDir
    Path tempDir;

    @BeforeEach
    void setUp() {
        mapper = SpreadsheetMapper.builder().useHeader(false).build();
    }

    // -- Fixture types --

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class FlatShift {
        @DataColumn String name;
        @DataColumn int qty;
        @DataColumn(shift = 1) double total;  // col 3 (col 2 = blank)
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Item {
        @DataColumn String product;
        @DataColumn int qty;
        @DataColumn(shift = 1) double amount;  // group col 3 (group col 2 = blank)
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class GroupShift {
        @DataColumn(anchor = true) String orderId;
        @DataColumnGroup(value = "Items", shift = 1) List<Item> items;  // col 2 = blank, items col 3-
        @DataColumn double total;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ItemShiftFirst {
        @DataColumn(shift = 2) String product;
        @DataColumn int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class GroupShiftPlusInnerShift {
        @DataColumn(anchor = true) String orderId;
        @DataColumnGroup(value = "Items", shift = 1) List<ItemShiftFirst> items;
        @DataColumn double total;
    }

    @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, include = JsonTypeInfo.As.PROPERTY, property = "kind")
    @JsonSubTypes({@JsonSubTypes.Type(value = Cash.class, name = "cash")})
    interface Payment {}

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Cash implements Payment {
        @DataColumn String name;
        @DataColumn double amount;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class PolymorphicShiftOnField {
        @DataColumn(shift = 1) Payment payment;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class CashWithSubtypeShift implements Payment {
        @DataColumn String name;
        @DataColumn(shift = 1) double amount;
    }

    @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, include = JsonTypeInfo.As.PROPERTY, property = "kind")
    @JsonSubTypes({@JsonSubTypes.Type(value = CashWithSubtypeShift.class, name = "cash")})
    interface PaymentWithSubtypeShift {}

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class PolymorphicSubtypeShift {
        PaymentWithSubtypeShift payment;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Inner {
        @DataColumn String city;
        @DataColumn String zip;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class WithUnwrappedShift {
        @DataColumn String name;
        @JsonUnwrapped @DataColumn(shift = 1) Inner inner;
    }

    // -- Round-trip tests --

    @Test
    void flatShift_roundTrip() throws Exception {
        File file = tempFile("flat-shift.xlsx");
        List<FlatShift> input = Arrays.asList(
                new FlatShift("Apple", 10, 15.50),
                new FlatShift("Banana", 20, 8.00));

        mapper.writeValue(file, input, FlatShift.class);
        List<FlatShift> output = mapper.readValues(file, FlatShift.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void flatShift_writeBlanksColumn() throws Exception {
        File file = tempFile("flat-shift-write.xlsx");
        mapper.writeValue(file,
                Arrays.asList(new FlatShift("Apple", 10, 15.50)),
                FlatShift.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(0);  // useHeader=false → data starts at row 0
            assertThat(row.getCell(0).getStringCellValue()).isEqualTo("Apple");
            assertThat((int) row.getCell(1).getNumericCellValue()).isEqualTo(10);
            // col 2 = shifted blank
            Cell blank = row.getCell(2);
            assertThat(blank == null || blank.toString().isEmpty()).isTrue();
            assertThat(row.getCell(3).getNumericCellValue()).isEqualTo(15.50);
        }
    }

    @Test
    void flatShift_withHeader_keepsLayoutAndLeavesGapHeaderBlank() throws Exception {
        final SpreadsheetMapper headerMapper = SpreadsheetMapper.builder().useHeader(true).build();
        File file = tempFile("flat-shift-header.xlsx");

        headerMapper.writeValue(file,
                Arrays.asList(new FlatShift("Apple", 10, 15.50)),
                FlatShift.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Row header = sheet.getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("name");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("qty");
            Cell gapHeader = header.getCell(2);
            assertThat(gapHeader == null || gapHeader.toString().isEmpty()).isTrue();
            assertThat(header.getCell(3).getStringCellValue()).isEqualTo("total");
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("Apple");
        }
    }

    @Test
    void flatShift_poiUserModel_roundTrip() throws Exception {
        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .useHeader(false)
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        File file = tempFile("flat-shift-poi.xlsx");
        List<FlatShift> input = Arrays.asList(
                new FlatShift("Apple", 10, 15.50),
                new FlatShift("Banana", 20, 8.00));

        poiMapper.writeValue(file, input, FlatShift.class);
        List<FlatShift> output = poiMapper.readValues(file, FlatShift.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void nestedGroupShift_roundTrip() throws Exception {
        File file = tempFile("group-shift.xlsx");
        List<GroupShift> input = Arrays.asList(
                new GroupShift("ORD-1", Arrays.asList(
                        new Item("Apple", 3, 9.0),
                        new Item("Banana", 5, 5.0)), 14.0));

        mapper.writeValue(file, input, GroupShift.class);
        List<GroupShift> output = mapper.readValues(file, GroupShift.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void groupShift_plus_innerFirstShift_accumulates() throws Exception {
        // Group shift (outer space) + inner first field shift (inner space) accumulate.
        // outer: col 0 orderId, col 1 blank (group shift 1),
        //        col 2-3 blank (inner shift 2), col 4 product, col 5 qty,
        //        col 6 total
        File file = tempFile("group-plus-inner-shift.xlsx");
        List<GroupShiftPlusInnerShift> input = Arrays.asList(
                new GroupShiftPlusInnerShift("ORD-1", Arrays.asList(
                        new ItemShiftFirst("Apple", 3),
                        new ItemShiftFirst("Banana", 5)), 14.0));

        mapper.writeValue(file, input, GroupShiftPlusInnerShift.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Row row0 = sheet.getRow(0);
            assertThat(row0.getCell(0).getStringCellValue()).isEqualTo("ORD-1");
            // col 1, 2, 3 = blank (group shift 1 + inner shift 2)
            for (int c = 1; c <= 3; c++) {
                Cell blank = row0.getCell(c);
                assertThat(blank == null || blank.toString().isEmpty()).isTrue();
            }
            assertThat(row0.getCell(4).getStringCellValue()).isEqualTo("Apple");
            assertThat((int) row0.getCell(5).getNumericCellValue()).isEqualTo(3);
            assertThat(row0.getCell(6).getNumericCellValue()).isEqualTo(14.0);
        }

        List<GroupShiftPlusInnerShift> output = mapper.readValues(file, GroupShiftPlusInnerShift.class);
        assertThat(output).isEqualTo(input);
    }

    @Test
    void groupShift_withHeader_collapsesToSingleLeafRow() throws Exception {
        final SpreadsheetMapper headerMapper = SpreadsheetMapper.builder().useHeader(true).build();
        File file = tempFile("group-shift-header.xlsx");

        headerMapper.writeValue(file,
                Arrays.asList(new GroupShift("ORD-1",
                        Arrays.asList(new Item("Apple", 3, 9.0)), 9.0)),
                GroupShift.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Row header = sheet.getRow(0);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("orderId");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo("items/[]/product");
            assertThat(header.getCell(3).getStringCellValue()).isEqualTo("items/[]/qty");
            assertThat(header.getCell(5).getStringCellValue()).isEqualTo("items/[]/amount");
            assertThat(header.getCell(6).getStringCellValue()).isEqualTo("total");
            for (int r = 0; r <= 1; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                for (Cell cell : row) {
                    assertThat(cell.toString()).isNotEqualTo("Items");
                }
            }
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("ORD-1");
        }
    }

    @Test
    void flatShift_withColumnReordering_roundTrips() throws Exception {
        final SpreadsheetMapper reorderMapper = SpreadsheetMapper.builder()
                .useHeader(true)
                .columnReordering(true)
                .build();
        File file = tempFile("flat-shift-reorder.xlsx");

        List<FlatShift> input = Arrays.asList(
                new FlatShift("Apple", 10, 15.50),
                new FlatShift("Banana", 20, 8.00));
        reorderMapper.writeValue(file, input, FlatShift.class);
        List<FlatShift> output = reorderMapper.readValues(file, FlatShift.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void polymorphicShift_onPolymorphicField_isAllowed() throws Exception {
        SpreadsheetSchema schema = mapper.sheetSchemaFor(PolymorphicShiftOnField.class);
        List<Column> columns = new ArrayList<>();
        for (Column c : schema) columns.add(c);
        assertThat(columns.get(0)).isNull();
        assertThat(columns.get(1)).isNotNull();
        assertThat(columns.get(1).getName()).isEqualTo("kind");
    }

    @Test
    void unwrapped_shiftOnOuterField_isSilentlyIgnored() throws Exception {
        SpreadsheetSchema schema = mapper.sheetSchemaFor(WithUnwrappedShift.class);
        int total = 0;
        int gaps = 0;
        for (Column c : schema) {
            total++;
            if (c == null) gaps++;
        }
        assertThat(total).isEqualTo(3);
        assertThat(gaps).isZero();
    }

    @Test
    void polymorphicShift_insideSubtype_isRejected() {
        assertThatThrownBy(() -> mapper.sheetSchemaFor(PolymorphicSubtypeShift.class))
                .isInstanceOf(com.fasterxml.jackson.databind.exc.InvalidDefinitionException.class)
                .hasMessageContaining("polymorphic subtype")
                .hasMessageContaining("CashWithSubtypeShift")
                .hasMessageContaining("Place shift outside the polymorphic field");
    }

    // -- Helpers --

    private File tempFile(String name) {
        return tempDir.resolve(name).toFile();
    }
}
