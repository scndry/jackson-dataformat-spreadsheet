package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Edge cases for the nested-list read/write paths — input shapes that
 * common-case fixtures in {@link NestedListReadTest} and
 * {@link NestedBlankRowFeatureTest} do not exercise: trailing blank
 * cells, anchor-blank continuation rows, partial-null inner fields,
 * and cross-mode (SSML / POI) round-trips.
 */
class NestedListEdgeCaseTest {

    @Data @NoArgsConstructor @AllArgsConstructor
    static class EdgeItem {
        @DataColumn("sku") String sku;
        @DataColumn("qty") Integer qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class EdgeOrder {
        @DataColumn(value = "id", anchor = true, merge = OptBoolean.TRUE) Integer id;
        @DataColumnGroup("Items") List<EdgeItem> items;
    }

    /**
     * Outer cell missing (anchor only present in first row, then
     * continuation rows omit it entirely) plus inner cell missing on
     * the last continuation row. The reader must continue the same
     * record and emit {@code null} for the missing inner field.
     */
    @Test
    void trailingBlankCells_continuationAndPartialInner() throws Exception {
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
            r2.createCell(2).setCellValue(10);

            Row r3 = sheet.createRow(3);
            r3.createCell(1).setCellValue("B"); // qty omitted, anchor omitted

            List<EdgeOrder> result = mapper.readValues(sheet, EdgeOrder.class);

            assertThat(result).hasSize(1);
            assertThat(result.get(0).getId()).isEqualTo(1);
            assertThat(result.get(0).getItems()).hasSize(2);
            assertThat(result.get(0).getItems().get(0).getSku()).isEqualTo("A");
            assertThat(result.get(0).getItems().get(0).getQty()).isEqualTo(10);
            assertThat(result.get(0).getItems().get(1).getSku()).isEqualTo("B");
            assertThat(result.get(0).getItems().get(1).getQty()).isNull();
        }
    }

    /**
     * Anchor cell blank in continuation rows is treated as "still the
     * same outer record". Three continuation rows after the anchored
     * row collapse into one outer with four inner items.
     */
    @Test
    void anchorBlankInContinuationRow_treatedAsSameRecord() throws Exception {
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
            r2.createCell(2).setCellValue(10);

            for (int i = 0; i < 3; i++) {
                Row r = sheet.createRow(3 + i);
                r.createCell(1).setCellValue("X" + i);
                r.createCell(2).setCellValue(20 + i);
            }

            List<EdgeOrder> result = mapper.readValues(sheet, EdgeOrder.class);

            assertThat(result).hasSize(1);
            assertThat(result.get(0).getId()).isEqualTo(1);
            assertThat(result.get(0).getItems()).extracting(EdgeItem::getSku)
                    .containsExactly("A", "X0", "X1", "X2");
            assertThat(result.get(0).getItems()).extracting(EdgeItem::getQty)
                    .containsExactly(10, 20, 21, 22);
        }
    }

    /**
     * Inner fields missing in alternation — first inner row omits
     * {@code sku}, second omits {@code qty}. Each inner record
     * carries the missing field as {@code null}.
     */
    @Test
    void innerFieldNull_perRecordCorrect() throws Exception {
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
            r2.createCell(2).setCellValue(10); // sku omitted

            Row r3 = sheet.createRow(3);
            r3.createCell(1).setCellValue("B"); // qty omitted

            List<EdgeOrder> result = mapper.readValues(sheet, EdgeOrder.class);

            assertThat(result).hasSize(1);
            assertThat(result.get(0).getItems()).hasSize(2);
            assertThat(result.get(0).getItems().get(0).getSku()).isNull();
            assertThat(result.get(0).getItems().get(0).getQty()).isEqualTo(10);
            assertThat(result.get(0).getItems().get(1).getSku()).isEqualTo("B");
            assertThat(result.get(0).getItems().get(1).getQty()).isNull();
        }
    }

    /**
     * Cross-mode round-trip — SSML write read back through POI and the
     * reverse — should produce identical objects, confirming neither
     * writer leaves the other reader out of sync.
     */
    @Test
    void crossRoundtrip_ssmlAndPoiAgree() throws Exception {
        SpreadsheetMapper ssml = new SpreadsheetMapper();
        SpreadsheetMapper poiRead = new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        SpreadsheetMapper poiWrite = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));

        List<EdgeOrder> data = Arrays.asList(
                new EdgeOrder(1, Arrays.asList(
                        new EdgeItem("A", 10), new EdgeItem("B", 20))),
                new EdgeOrder(2, Arrays.asList(new EdgeItem("C", 30))));

        // SSML write -> POI read
        ByteArrayOutputStream ssmlOut = new ByteArrayOutputStream();
        ssml.writeValue(ssmlOut, data, EdgeOrder.class);
        List<EdgeOrder> readViaPoi = poiRead.readValues(
                new ByteArrayInputStream(ssmlOut.toByteArray()), EdgeOrder.class);
        assertThat(readViaPoi).isEqualTo(data);

        // POI write -> SSML read
        ByteArrayOutputStream poiOut = new ByteArrayOutputStream();
        poiWrite.writeValue(poiOut, data, EdgeOrder.class);
        List<EdgeOrder> readViaSsml = ssml.readValues(
                new ByteArrayInputStream(poiOut.toByteArray()), EdgeOrder.class);
        assertThat(readViaSsml).isEqualTo(data);
    }
}
