package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.GridConfigurer;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import java.io.File;
import java.util.Arrays;
import java.util.List;

import static io.github.scndry.jackson.dataformat.spreadsheet.OpcXmlHelper.NS_SPREADSHEETML;
import static io.github.scndry.jackson.dataformat.spreadsheet.OpcXmlHelper.parsePart;
import static org.assertj.core.api.Assertions.assertThat;

class DisplayGridlinesTest {

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
    void poiPathHideGridlines() throws Exception {
        File file = new File(tempDir, "hg-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .gridConfigurer(new GridConfigurer().displayGridlines(false))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).isDisplayGridlines()).isFalse();
        }
    }

    @Test
    void ssmlPathHideGridlines() throws Exception {
        File file = new File(tempDir, "hg-ssml.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().displayGridlines(false))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).isDisplayGridlines()).isFalse();
        }
    }

    @Test
    void gridlinesShownByDefault() throws Exception {
        File file = new File(tempDir, "hg-default.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        new SpreadsheetMapper().writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).isDisplayGridlines()).isTrue();
        }
    }

    @Test
    void dataIntegrityWithHideGridlines() throws Exception {
        File file = new File(tempDir, "hg-roundtrip.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().displayGridlines(false))
                .build();
        mapper.writeValue(file, data, Item.class);

        List<Item> read = new SpreadsheetMapper().readValues(file, Item.class);
        assertThat(read).isEqualTo(data);
    }

    @Test
    void domEquivalence() throws Exception {
        File ssmlFile = new File(tempDir, "hg-dom-ssml.xlsx");
        File poiFile = new File(tempDir, "hg-dom-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper ssmlMapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().displayGridlines(false))
                .build();
        ssmlMapper.writeValue(ssmlFile, data, Item.class);

        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .gridConfigurer(new GridConfigurer().displayGridlines(false))
                .build();
        poiMapper.writeValue(poiFile, data, Item.class);

        try (OPCPackage expPkg = OPCPackage.open(poiFile);
             OPCPackage actPkg = OPCPackage.open(ssmlFile)) {
            Document expDoc = parsePart(expPkg, "/xl/worksheets/sheet1.xml");
            Document actDoc = parsePart(actPkg, "/xl/worksheets/sheet1.xml");

            NodeList expViews = expDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "sheetView");
            NodeList actViews = actDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "sheetView");
            assertThat(actViews.getLength())
                    .as("sheetView element count")
                    .isEqualTo(expViews.getLength());
            for (int i = 0; i < expViews.getLength(); i++) {
                assertThat(actViews.item(i).isEqualNode(expViews.item(i)))
                        .as("sheetView[%d] DOM equality", i)
                        .isTrue();
            }
        }
    }
}
