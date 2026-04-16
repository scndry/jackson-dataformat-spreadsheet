package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Layer 2: Cell type read tests via POI Sheet direct access.
 * Row 0 = header (schema convention), Row 1+ = data.
 */
class CellTypeReadTest {

    SpreadsheetMapper mapper;

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class StringRow { String value; }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class IntRow { int value; }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class DoubleRow { double value; }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class BoolRow { boolean value; }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NullableRow { String value; }


    @Test
    void stringCell() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0); // header
            addRow(sheet, 1, "hello");

            StringRow result = mapper.readValue(sheet, StringRow.class);
            assertThat(result.getValue()).isEqualTo("hello");
        }
    }

    @Test
    void intCell() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0); // header
            addNumericRow(sheet, 1, 42);

            IntRow result = mapper.readValue(sheet, IntRow.class);
            assertThat(result.getValue()).isEqualTo(42);
        }
    }

    @Test
    void doubleCell() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0); // header
            addNumericRow(sheet, 1, 3.14);

            DoubleRow result = mapper.readValue(sheet, DoubleRow.class);
            assertThat(result.getValue()).isEqualTo(3.14);
        }
    }

    @Test
    void booleanCell() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0); // header
            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue(true);

            BoolRow result = mapper.readValue(sheet, BoolRow.class);
            assertThat(result.isValue()).isTrue();
        }
    }

    @Test
    void blankCellIsSkipped() throws Exception {
        // Blank cells are skipped (consistent with SSML path where blank cells are absent).
        // A row with only blank cells becomes an empty object → field keeps default value.
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0); // header
            Row row = sheet.createRow(1);
            row.createCell(0); // blank

            NullableRow result = mapper.readValue(sheet, NullableRow.class);
            // Row with only blank cells → treated as blank row → null by BLANK_ROW_AS_NULL (default)
            assertThat(result).isNull();
        }
    }

    @Test
    void blankRowAsNull() throws Exception {
        SpreadsheetMapper m = SpreadsheetMapper.builder()
                .enable(SheetParser.Feature.BLANK_ROW_AS_NULL)
                .build();

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0); // header
            addRow(sheet, 1, "first");
            sheet.createRow(2); // blank data row
            addRow(sheet, 3, "third");

            List<StringRow> results = m.readValues(sheet, StringRow.class);
            assertThat(results).hasSize(3);
            assertThat(results.get(0).getValue()).isEqualTo("first");
            assertThat(results.get(1)).isNull();
            assertThat(results.get(2).getValue()).isEqualTo("third");
        }
    }

    @Test
    void breakOnBlankRow() throws Exception {
        SpreadsheetMapper m = SpreadsheetMapper.builder()
                .disable(SheetParser.Feature.BLANK_ROW_AS_NULL)
                .enable(SheetParser.Feature.BREAK_ON_BLANK_ROW)
                .build();

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0); // header
            addRow(sheet, 1, "first");
            addRow(sheet, 2, "second");
            sheet.createRow(3); // blank — should stop
            addRow(sheet, 4, "should not reach");

            List<StringRow> results = m.readValues(sheet, StringRow.class);
            assertThat(results).hasSize(2)
                    .extracting(StringRow::getValue)
                    .containsExactly("first", "second");
        }
    }

    @Test
    void errorCellThrows() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            sheet.createRow(0); // header
            Row row = sheet.createRow(1);
            row.createCell(0).setCellErrorValue(org.apache.poi.ss.usermodel.FormulaError.DIV0.getCode());

            // Error cells are not mappable — should throw, not NPE or hang
            assertThatThrownBy(() -> mapper.readValue(sheet, IntRow.class))
                    .hasMessageContaining("ERROR");
        }
    }

    // -- helpers --

    private void addRow(Sheet sheet, int rowIdx, String value) {
        sheet.createRow(rowIdx).createCell(0).setCellValue(value);
    }

    private void addNumericRow(Sheet sheet, int rowIdx, double value) {
        sheet.createRow(rowIdx).createCell(0).setCellValue(value);
    }
}
