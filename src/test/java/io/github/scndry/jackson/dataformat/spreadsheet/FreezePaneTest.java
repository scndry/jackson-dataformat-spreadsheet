package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.GridConfigurer;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.PaneInformation;
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

class FreezePaneTest {

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
    void poiPathFreezeHeaderRow() throws Exception {
        File file = new File(tempDir, "fp-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .gridConfigurer(new GridConfigurer().freezePane(0, 1))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            PaneInformation pane = wb.getSheetAt(0).getPaneInformation();
            assertThat(pane).isNotNull();
            assertThat(pane.isFreezePane()).isTrue();
            assertThat(pane.getHorizontalSplitPosition()).isEqualTo((short) 1);
            assertThat(pane.getVerticalSplitPosition()).isEqualTo((short) 0);
        }
    }

    @Test
    void ssmlPathFreezeHeaderRow() throws Exception {
        File file = new File(tempDir, "fp-ssml.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().freezePane(0, 1))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            PaneInformation pane = wb.getSheetAt(0).getPaneInformation();
            assertThat(pane).isNotNull();
            assertThat(pane.isFreezePane()).isTrue();
            assertThat(pane.getHorizontalSplitPosition()).isEqualTo((short) 1);
            assertThat(pane.getVerticalSplitPosition()).isEqualTo((short) 0);
        }
    }

    @Test
    void noFreezePaneByDefault() throws Exception {
        File file = new File(tempDir, "fp-none.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        new SpreadsheetMapper().writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            PaneInformation pane = wb.getSheetAt(0).getPaneInformation();
            assertThat(pane == null || !pane.isFreezePane()).isTrue();
        }
    }

    @Test
    void dataIntegrityWithFreezePane() throws Exception {
        File file = new File(tempDir, "fp-roundtrip.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().freezePane(0, 1))
                .build();
        mapper.writeValue(file, data, Item.class);

        List<Item> read = new SpreadsheetMapper().readValues(file, Item.class);
        assertThat(read).isEqualTo(data);
    }

    @Test
    void domEquivalence() throws Exception {
        File ssmlFile = new File(tempDir, "fp-dom-ssml.xlsx");
        File poiFile = new File(tempDir, "fp-dom-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper ssmlMapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().freezePane(0, 1))
                .build();
        ssmlMapper.writeValue(ssmlFile, data, Item.class);

        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .gridConfigurer(new GridConfigurer().freezePane(0, 1))
                .build();
        poiMapper.writeValue(poiFile, data, Item.class);

        try (OPCPackage expPkg = OPCPackage.open(poiFile);
             OPCPackage actPkg = OPCPackage.open(ssmlFile)) {
            Document expDoc = parsePart(expPkg, "/xl/worksheets/sheet1.xml");
            Document actDoc = parsePart(actPkg, "/xl/worksheets/sheet1.xml");

            NodeList expPanes = expDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "pane");
            NodeList actPanes = actDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "pane");
            assertThat(actPanes.getLength())
                    .as("pane element count")
                    .isEqualTo(expPanes.getLength());
            for (int i = 0; i < expPanes.getLength(); i++) {
                assertThat(actPanes.item(i).isEqualNode(expPanes.item(i)))
                        .as("pane[%d] DOM equality", i)
                        .isTrue();
            }
        }
    }
}
