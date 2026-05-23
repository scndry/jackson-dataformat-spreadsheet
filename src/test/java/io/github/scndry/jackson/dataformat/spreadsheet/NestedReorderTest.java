package io.github.scndry.jackson.dataformat.spreadsheet;

import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Pins down the FEATURE_COLUMN_REORDERING boundary for nested schemas.
 * SheetParser.setSchema rejects the combination of {@code
 * columnReordering=true} and any schema carrying
 * {@code @DataColumnGroup} (multi-row headers) — which all nested-list
 * schemas necessarily do. The check fires before RecordTreeBuffer is
 * created, so the algorithm never sees a reordered schema and the
 * snapshot it holds at construction stays authoritative.
 */
class NestedReorderTest {

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ROItem {
        @DataColumn("sku") String sku;
        @DataColumn("qty") int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class ROOrder {
        @DataColumn(value = "id", anchor = OptBoolean.TRUE, merge = OptBoolean.TRUE) Integer id;
        @DataColumn("name") String name;
        @DataColumnGroup("Items") List<ROItem> items;
    }

    @Test
    void nested_withColumnReordering_isRejectedAtSetSchema() throws Exception {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .columnReordering(true).build();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row h0 = sheet.createRow(0);
            h0.createCell(0).setCellValue("id");
            h0.createCell(1).setCellValue("name");
            h0.createCell(2).setCellValue("Items");
            Row h1 = sheet.createRow(1);
            h1.createCell(2).setCellValue("sku");
            h1.createCell(3).setCellValue("qty");

            assertThatThrownBy(() -> mapper.readValues(sheet, ROOrder.class))
                    .isInstanceOf(IllegalStateException.class)
                    .hasMessageContaining("Column reordering is not supported with multi-row headers");
        }
    }
}
