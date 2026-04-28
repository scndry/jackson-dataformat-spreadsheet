package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.util.CellAddress;
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

class HeaderCommentTest {

    @TempDir File tempDir;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class Item {
        @DataColumn(comment = "항목 이름")
        private String name;
        @DataColumn(comment = "수량")
        private int quantity;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class NoComment {
        private String name;
        private int value;
    }

    @Test
    void poiPathHeaderComment() throws Exception {
        File file = new File(tempDir, "comment-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            Comment c0 = sheet.getCellComment(new CellAddress(0, 0));
            Comment c1 = sheet.getCellComment(new CellAddress(0, 1));
            assertThat(c0).isNotNull();
            assertThat(c0.getString().getString()).isEqualTo("항목 이름");
            assertThat(c1).isNotNull();
            assertThat(c1.getString().getString()).isEqualTo("수량");
        }
    }

    @Test
    void ssmlPathHeaderComment() throws Exception {
        File file = new File(tempDir, "comment-ssml.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        SpreadsheetMapper mapper = SpreadsheetMapper.builder().build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            Comment c0 = sheet.getCellComment(new CellAddress(0, 0));
            Comment c1 = sheet.getCellComment(new CellAddress(0, 1));
            assertThat(c0).isNotNull();
            assertThat(c0.getString().getString()).isEqualTo("항목 이름");
            assertThat(c1).isNotNull();
            assertThat(c1.getString().getString()).isEqualTo("수량");
        }
    }

    @Test
    void noCommentByDefault() throws Exception {
        File file = new File(tempDir, "comment-none.xlsx");
        List<NoComment> data = Arrays.asList(new NoComment("A", 1));

        new SpreadsheetMapper().writeValue(file, data, NoComment.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getCellComment(new CellAddress(0, 0))).isNull();
        }
    }

    @Test
    void dataIntegrityWithComment() throws Exception {
        File file = new File(tempDir, "comment-roundtrip.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1), new Item("B", 2));

        new SpreadsheetMapper().writeValue(file, data, Item.class);

        List<Item> read = new SpreadsheetMapper().readValues(file, Item.class);
        assertThat(read).isEqualTo(data);
    }

    @Test
    void domEquivalence() throws Exception {
        File ssmlFile = new File(tempDir, "comment-dom-ssml.xlsx");
        File poiFile = new File(tempDir, "comment-dom-poi.xlsx");
        List<Item> data = Arrays.asList(new Item("A", 1));

        SpreadsheetMapper ssmlMapper = SpreadsheetMapper.builder().build();
        ssmlMapper.writeValue(ssmlFile, data, Item.class);

        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        poiMapper.writeValue(poiFile, data, Item.class);

        try (OPCPackage expPkg = OPCPackage.open(poiFile);
             OPCPackage actPkg = OPCPackage.open(ssmlFile)) {
            Document expDoc = parsePart(expPkg, "/xl/comments1.xml");
            Document actDoc = parsePart(actPkg, "/xl/comments1.xml");

            NodeList expComments = expDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "comment");
            NodeList actComments = actDoc.getElementsByTagNameNS(NS_SPREADSHEETML, "comment");
            assertThat(actComments.getLength())
                    .as("comment count")
                    .isEqualTo(expComments.getLength());
        }
    }
}
