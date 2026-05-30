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
import org.w3c.dom.NodeList;

import java.io.File;
import java.util.Arrays;
import java.util.List;

import static io.github.scndry.jackson.dataformat.spreadsheet.OpcXmlHelper.NS_SPREADSHEETML;
import static io.github.scndry.jackson.dataformat.spreadsheet.OpcXmlHelper.parsePart;
import static org.assertj.core.api.Assertions.assertThat;

class SheetProtectionTest {

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
    void poiPathProtectSheet() throws Exception {
        File file = new File(tempDir, "sp-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .gridConfigurer(new GridConfigurer().protectSheet("secret"))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).getProtect()).isTrue();
        }
    }

    @Test
    void ssmlPathProtectSheet() throws Exception {
        File file = new File(tempDir, "sp-ssml.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().protectSheet("secret"))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).getProtect()).isTrue();
        }
    }

    @Test
    void noProtectSheetByDefault() throws Exception {
        File file = new File(tempDir, "sp-none.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        new SpreadsheetMapper().writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).getProtect()).isFalse();
        }
    }

    @Test
    void dataIntegrityWithProtectSheet() throws Exception {
        File file = new File(tempDir, "sp-roundtrip.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().protectSheet("secret"))
                .build();
        mapper.writeValue(file, data, Item.class);

        List<Item> read = new SpreadsheetMapper().readValues(file, Item.class);
        assertThat(read).isEqualTo(data);
    }

    @Test
    void nullPasswordClears() throws Exception {
        File file = new File(tempDir, "sp-null.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().protectSheet("secret").protectSheet(null))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).getProtect()).isFalse();
        }
    }

    @Test
    void wrongPasswordRejected() throws Exception {
        File file = new File(tempDir, "sp-validate.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().protectSheet("secret"))
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            assertThat(sheet.validateSheetPassword("secret")).isTrue();
            assertThat(sheet.validateSheetPassword("wrong")).isFalse();
        }
    }

    @Test
    void domEquivalence() throws Exception {
        File ssmlFile = new File(tempDir, "sp-dom-ssml.xlsx");
        File poiFile = new File(tempDir, "sp-dom-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        SpreadsheetMapper ssmlMapper = SpreadsheetMapper.builder()
                .gridConfigurer(new GridConfigurer().protectSheet("secret"))
                .build();
        ssmlMapper.writeValue(ssmlFile, data, Item.class);

        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .gridConfigurer(new GridConfigurer().protectSheet("secret"))
                .build();
        poiMapper.writeValue(poiFile, data, Item.class);

        try (OPCPackage expPkg = OPCPackage.open(poiFile);
             OPCPackage actPkg = OPCPackage.open(ssmlFile)) {
            Document expDoc = parsePart(expPkg, "/xl/worksheets/sheet1.xml");
            Document actDoc = parsePart(actPkg, "/xl/worksheets/sheet1.xml");

            NodeList expProt = expDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "sheetProtection");
            NodeList actProt = actDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "sheetProtection");
            assertThat(actProt.getLength())
                    .as("sheetProtection element count")
                    .isEqualTo(expProt.getLength());
            for (int i = 0; i < expProt.getLength(); i++) {
                assertThat(actProt.item(i).isEqualNode(expProt.item(i)))
                        .as("sheetProtection[%d] DOM equality", i)
                        .isTrue();
            }
        }
    }
}
