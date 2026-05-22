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
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * BLANK_ROW_AS_NULL and BREAK_ON_BLANK_ROW carry the same semantics in
 * nested mode as in the flat path: a completely blank row closes the
 * currently-open outer record, then BLANK_ROW_AS_NULL inserts a null
 * entry into the result list and BREAK_ON_BLANK_ROW terminates the
 * iteration. With both flags off the blank row is silently dropped and
 * reading continues — matching the flat path's _handleEmptyObject
 * clear-and-continue behaviour.
 */
class NestedBlankRowFeatureTest {

    @Data @NoArgsConstructor @AllArgsConstructor
    static class BlankItem {
        @DataColumn("sku") String sku;
        @DataColumn("qty") int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class BlankOrder {
        @DataColumn(value = "id", anchor = OptBoolean.TRUE, merge = OptBoolean.TRUE) Integer id;
        @DataColumnGroup("Items") List<BlankItem> items;
    }

    /**
     * Schema headerRowCount is 2 because {@code @DataColumnGroup("Items")}
     * gives the leaf inner columns a depth-1 hierarchy. Layout (0-indexed):
     *   row 0 (header depth 0): id, "Items", ""
     *   row 1 (header depth 1, leaf): "", sku, qty
     *   row 2: 1, A, 10        ← record 1 first row
     *   row 3:  , B, 20        ← record 1 continuation
     *   row 4: (blank)         ← blank separator
     *   row 5: 2, C, 30        ← record 2 first row
     *   row 6:  , D, 40        ← record 2 continuation
     */
    private static Sheet _fixtureWithBlankSeparator(final XSSFWorkbook wb) {
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
        r2.createCell(2).setCellValue(10);

        Row r3 = sheet.createRow(3);
        r3.createCell(1).setCellValue("B");
        r3.createCell(2).setCellValue(20);

        sheet.createRow(4); // blank

        Row r5 = sheet.createRow(5);
        r5.createCell(0).setCellValue(2);
        r5.createCell(1).setCellValue("C");
        r5.createCell(2).setCellValue(30);

        Row r6 = sheet.createRow(6);
        r6.createCell(1).setCellValue("D");
        r6.createCell(2).setCellValue(40);
        return sheet;
    }

    @Test
    void defaultFlags_blankRowAsNullEntry() throws Exception {
        // Default: BLANK_ROW_AS_NULL=true, BREAK_ON_BLANK_ROW=false.
        // Blank row → close record 1 (emit) + null entry + continue.
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = _fixtureWithBlankSeparator(wb);
            List<BlankOrder> results = mapper.readValues(sheet, BlankOrder.class);

            assertThat(results).hasSize(3);
            assertThat(results.get(0).getId()).isEqualTo(1);
            assertThat(results.get(0).getItems()).extracting(BlankItem::getSku)
                    .containsExactly("A", "B");
            assertThat(results.get(1)).isNull();
            assertThat(results.get(2).getId()).isEqualTo(2);
            assertThat(results.get(2).getItems()).extracting(BlankItem::getSku)
                    .containsExactly("C", "D");
        }
    }

    @Test
    void breakOnBlankRow_stopsIteration() throws Exception {
        // Blank row → close record 1 + iteration ends; record 2 is never read.
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .disable(SheetParser.Feature.BLANK_ROW_AS_NULL)
                .enable(SheetParser.Feature.BREAK_ON_BLANK_ROW)
                .build();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = _fixtureWithBlankSeparator(wb);
            List<BlankOrder> results = mapper.readValues(sheet, BlankOrder.class);

            assertThat(results).hasSize(1);
            assertThat(results.get(0).getId()).isEqualTo(1);
            assertThat(results.get(0).getItems()).extracting(BlankItem::getSku)
                    .containsExactly("A", "B");
        }
    }

    @Test
    void bothFlagsOff_blankRowSilentlyDropped() throws Exception {
        // Blank row → close record 1 + drop (no null entry) + continue.
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .disable(SheetParser.Feature.BLANK_ROW_AS_NULL)
                .disable(SheetParser.Feature.BREAK_ON_BLANK_ROW)
                .build();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = _fixtureWithBlankSeparator(wb);
            List<BlankOrder> results = mapper.readValues(sheet, BlankOrder.class);

            assertThat(results).hasSize(2);
            assertThat(results.get(0).getId()).isEqualTo(1);
            assertThat(results.get(0).getItems()).extracting(BlankItem::getSku)
                    .containsExactly("A", "B");
            assertThat(results.get(1).getId()).isEqualTo(2);
            assertThat(results.get(1).getItems()).extracting(BlankItem::getSku)
                    .containsExactly("C", "D");
        }
    }
}
