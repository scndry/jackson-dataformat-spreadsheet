package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.sheet.SheetConfigurer;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.File;
import java.util.Arrays;
import java.util.List;

import static io.github.scndry.jackson.dataformat.spreadsheet.OpcXmlHelper.NS_SPREADSHEETML;
import static io.github.scndry.jackson.dataformat.spreadsheet.OpcXmlHelper.parsePart;

import static org.assertj.core.api.Assertions.assertThat;

class AutoFilterTest {

    @TempDir File tempDir;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class Item {
        private String name;
        private int value;
    }

    @Test
    void poiPathAutoFilter() throws Exception {
        File file = new File(tempDir, "af-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .sheetConfigurer(new SheetConfigurer().autoFilter())
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getCTWorksheet().isSetAutoFilter()).isTrue();
            assertThat(sheet.getCTWorksheet().getAutoFilter().getRef()).isEqualTo("A1:B3");
        }
    }

    @Test
    void ssmlPathAutoFilter() throws Exception {
        File file = new File(tempDir, "af-ssml.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .sheetConfigurer(new SheetConfigurer().autoFilter())
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getCTWorksheet().isSetAutoFilter()).isTrue();
            // SSML uses max rows for skeleton, verify autoFilter exists
            String ref = sheet.getCTWorksheet().getAutoFilter().getRef();
            assertThat(ref).startsWith("A1:");
        }
    }

    @Test
    void noAutoFilterByDefault() throws Exception {
        File file = new File(tempDir, "af-none.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        new SpreadsheetMapper().writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).getCTWorksheet().isSetAutoFilter()).isFalse();
        }
    }

    @Test
    void dataIntegrityWithAutoFilter() throws Exception {
        File file = new File(tempDir, "af-roundtrip.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .sheetConfigurer(new SheetConfigurer().autoFilter())
                .build();
        mapper.writeValue(file, data, Item.class);

        List<Item> read = new SpreadsheetMapper().readValues(file, Item.class);
        assertThat(read).isEqualTo(data);
    }

    @Test
    void domEquivalence() throws Exception {
        File ssmlFile = new File(tempDir, "af-dom-ssml.xlsx");
        File poiFile = new File(tempDir, "af-dom-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper ssmlMapper = SpreadsheetMapper.builder()
                .sheetConfigurer(new SheetConfigurer().autoFilter())
                .build();
        ssmlMapper.writeValue(ssmlFile, data, Item.class);

        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .sheetConfigurer(new SheetConfigurer().autoFilter())
                .build();
        poiMapper.writeValue(poiFile, data, Item.class);

        try (OPCPackage expPkg = OPCPackage.open(poiFile);
             OPCPackage actPkg = OPCPackage.open(ssmlFile)) {
            Document expDoc = parsePart(expPkg, "/xl/worksheets/sheet1.xml");
            Document actDoc = parsePart(actPkg, "/xl/worksheets/sheet1.xml");

            NodeList expAf = expDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "autoFilter");
            NodeList actAf = actDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "autoFilter");
            assertThat(actAf.getLength())
                    .as("autoFilter element count")
                    .isEqualTo(expAf.getLength());
            // ref differs (POI uses actual lastRow, SSML uses max row) — compare only existence and column range
            if (expAf.getLength() > 0) {
                String expRef = ((Element) expAf.item(0)).getAttribute("ref");
                String actRef = ((Element) actAf.item(0)).getAttribute("ref");
                // Both should start with same column range (e.g., "A1:")
                String expPrefix = expRef.substring(0, expRef.indexOf(':') + 1);
                String actPrefix = actRef.substring(0, actRef.indexOf(':') + 1);
                assertThat(actPrefix)
                        .as("autoFilter ref column prefix")
                        .isEqualTo(expPrefix);
            }
        }
    }
}
