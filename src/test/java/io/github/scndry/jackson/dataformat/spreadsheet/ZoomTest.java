package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.grid.GridConfigurer;
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

class ZoomTest {

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
    void poiPathZoom() throws Exception {
        File file = new File(tempDir, "zm-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .gridConfigurer(new GridConfigurer().zoom(150))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale())
                    .isEqualTo(150);
        }
    }

    @Test
    void ssmlPathZoom() throws Exception {
        File file = new File(tempDir, "zm-ssml.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().zoom(150))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).getZoomScale())
                    .isEqualTo(150);
        }
    }

    @Test
    void noZoomByDefault() throws Exception {
        File file = new File(tempDir, "zm-default.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        new SpreadsheetMapper().writeValue(file, data, Item.class);

        try (OPCPackage pkg = OPCPackage.open(file)) {
            Document doc = parsePart(pkg, "/xl/worksheets/sheet1.xml");
            NodeList views = doc.getElementsByTagNameNS(NS_SPREADSHEETML, "sheetView");
            assertThat(views.getLength()).isPositive();
            Element view = (Element) views.item(0);
            assertThat(view.getAttribute("zoomScale"))
                    .as("zoomScale attribute absent by default")
                    .isEmpty();
        }
    }

    @Test
    void dataIntegrityWithZoom() throws Exception {
        File file = new File(tempDir, "zm-roundtrip.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().zoom(125))
                .build();
        mapper.writeValue(file, data, Item.class);

        List<Item> read = new SpreadsheetMapper().readValues(file, Item.class);
        assertThat(read).isEqualTo(data);
    }

    @Test
    void domEquivalence() throws Exception {
        File ssmlFile = new File(tempDir, "zm-dom-ssml.xlsx");
        File poiFile = new File(tempDir, "zm-dom-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper ssmlMapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().zoom(150))
                .build();
        ssmlMapper.writeValue(ssmlFile, data, Item.class);

        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .gridConfigurer(new GridConfigurer().zoom(150))
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
