package io.github.scndry.jackson.dataformat.spreadsheet;

import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * ERROR (and other unsupported) cells in nested reads raise
 * {@link SheetStreamReadException} — same as the flat path.
 */
class NestedErrorCellTest {

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ErrItem {
        @DataColumn("sku") String sku;
        @DataColumn("qty") Integer qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class ErrOrder {
        @DataColumn(value = "id", anchor = true, merge = OptBoolean.TRUE) Integer id;
        @DataColumnGroup("Items") List<ErrItem> items;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class FlatRow {
        @DataColumn("sku") String sku;
        @DataColumn("qty") Integer qty;
    }

    /** Inner cell holds #N/A — nested rejects at cell-arrival time. */
    @Test
    void nested_innerErrorCell_throwsUnexpected() throws Exception {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row h0 = sheet.createRow(0);
            h0.createCell(0).setCellValue("id");
            h0.createCell(1).setCellValue("Items");
            Row h1 = sheet.createRow(1);
            h1.createCell(1).setCellValue("sku");
            h1.createCell(2).setCellValue("qty");

            Row r2 = sheet.createRow(2);
            r2.createCell(0).setCellValue(1);
            r2.createCell(1).setCellValue("A");
            r2.createCell(2).setCellErrorValue(FormulaError.NA.getCode());

            assertThatThrownBy(() -> mapper.readValues(sheet, ErrOrder.class))
                    .isInstanceOf(SheetStreamReadException.class)
                    .hasMessageContaining("ERROR");
        }
    }

    /** Flat path raises SheetStreamReadException — semantic baseline for
     *  comparison with the nested path's silent behaviour. */
    @Test
    void flat_errorCell_throwsUnexpected() throws Exception {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row h = sheet.createRow(0);
            h.createCell(0).setCellValue("sku");
            h.createCell(1).setCellValue("qty");

            Row r = sheet.createRow(1);
            r.createCell(0).setCellValue("A");
            r.createCell(1).setCellErrorValue(FormulaError.NA.getCode());

            assertThatThrownBy(() -> mapper.readValues(sheet, FlatRow.class))
                    .isInstanceOf(SheetStreamReadException.class);
        }
    }
}
