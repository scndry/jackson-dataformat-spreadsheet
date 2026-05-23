package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Direct assertions on the written header layout for nested-list
 * schemas — group header text, leaf header text, vertical merge on
 * outer columns, and horizontal merge on the group cell.
 * NestedListReadTest's round-trip cases verify these indirectly via
 * deserialisation; this test pins them as visible Excel content.
 */
class NestedListHeaderTest {

    @TempDir Path tempDir;

    @Data @NoArgsConstructor @AllArgsConstructor
    static class HeaderItem {
        @DataColumn("sku") String sku;
        @DataColumn("qty") Integer qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class HeaderOrder {
        @DataColumn(value = "id", anchor = true, merge = OptBoolean.TRUE) Integer id;
        @DataColumn("name") String name;
        @DataColumnGroup("Items") List<HeaderItem> items;
    }

    @Test
    void headerLayout_groupCellAndLeafRow() throws Exception {
        File file = tempDir.resolve("nested-header.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file, Arrays.asList(
                new HeaderOrder(1, "Alice",
                        Arrays.asList(new HeaderItem("A", 10)))),
                HeaderOrder.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Row h0 = sheet.getRow(0);
            Row h1 = sheet.getRow(1);

            // depth 0: outer column headers + group header cell
            assertThat(h0.getCell(0).getStringCellValue()).isEqualTo("id");
            assertThat(h0.getCell(1).getStringCellValue()).isEqualTo("name");
            assertThat(h0.getCell(2).getStringCellValue()).isEqualTo("Items");

            // depth 1 (leaf): inner column headers
            assertThat(h1.getCell(2).getStringCellValue()).isEqualTo("sku");
            assertThat(h1.getCell(3).getStringCellValue()).isEqualTo("qty");

            // vertical merge of outer columns (id, name) across rows 0-1
            List<CellRangeAddress> regions = sheet.getMergedRegions();
            assertThat(regions).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastColumn()).isEqualTo(0);
            });
            assertThat(regions).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(1);
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastColumn()).isEqualTo(1);
            });

            // horizontal merge of group cell ("Items") across columns 2-3
            assertThat(regions).anySatisfy(r -> {
                assertThat(r.getFirstRow()).isEqualTo(0);
                assertThat(r.getLastRow()).isEqualTo(0);
                assertThat(r.getFirstColumn()).isEqualTo(2);
                assertThat(r.getLastColumn()).isEqualTo(3);
            });
        }
    }
}
