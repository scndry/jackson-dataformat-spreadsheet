package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Accuracy checks beyond the basic single-nested-list case. Each test asserts
 * that SSML output matches POI XSSFWorkbook output (DOM equivalence) for a
 * scenario representative of real-world data shapes.
 */
class RealisticAccuracyTest {

    private static final Path DEBUG_OUTPUT_DIR = Paths.get("build/debug-output");

    // ----------------------------------------------------------------
    // 1) MULTIPLE_LISTS — record contains two nested lists, each followed
    //    by an outer field. Both outer fields anchor at the first record
    //    row (Jackson + ObjectContext.getRow() behavior).
    // ----------------------------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Item {
        @DataColumn("상품") String name;
        @DataColumn("수량") int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Payment {
        @DataColumn("결제수단") String method;
        @DataColumn("결제액") BigDecimal amount;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class OrderTwoLists {
        @DataColumn(value = "주문번호", merge = OptBoolean.TRUE) String orderNumber;
        List<Item> items;
        @DataColumn(value = "소계", merge = OptBoolean.TRUE) BigDecimal subtotal;
        List<Payment> payments;
        @DataColumn(value = "총결제", merge = OptBoolean.TRUE) BigDecimal totalPaid;
    }

    @Test
    void multipleListsPerRecord() throws Exception {
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        OrderTwoLists order = new OrderTwoLists(
                "ORD-001",
                Arrays.asList(new Item("Apple", 3), new Item("Banana", 5)),
                BigDecimal.valueOf(2000),
                Arrays.asList(new Payment("CARD", BigDecimal.valueOf(1500)),
                        new Payment("CASH", BigDecimal.valueOf(500))),
                BigDecimal.valueOf(2000));

        File ssmlFile = _debugFile("rat-multi-list-ssml.xlsx");
        File poiFile = _debugFile("rat-multi-list-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile,
                Collections.singletonList(order), OrderTwoLists.class);
        _poiMapper().writeValue(poiFile,
                Collections.singletonList(order), OrderTwoLists.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // ----------------------------------------------------------------
    // 2) DEEP_NESTED — list of list (Group → Items), with subtotal per
    //    group anchored at the group's first item row.
    // ----------------------------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class GroupItem {
        @DataColumn("품목") String label;
        @DataColumn("수량") int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Group {
        @DataColumn(value = "그룹명", merge = OptBoolean.TRUE) String groupName;
        List<GroupItem> items;
        @DataColumn(value = "그룹소계", merge = OptBoolean.TRUE) int groupSubtotal;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class GroupedOrder {
        @DataColumn(value = "주문ID", merge = OptBoolean.TRUE) int id;
        List<Group> groups;
        @DataColumn(value = "총합", merge = OptBoolean.TRUE) int total;
    }

    @Test
    void deepNestedListOfList() throws Exception {
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        GroupedOrder order = new GroupedOrder(1,
                Arrays.asList(
                        new Group("X",
                                Arrays.asList(new GroupItem("a", 1), new GroupItem("b", 2)),
                                3),
                        new Group("Y",
                                Arrays.asList(new GroupItem("c", 4)),
                                4)),
                7);

        File ssmlFile = _debugFile("rat-deep-nested-ssml.xlsx");
        File poiFile = _debugFile("rat-deep-nested-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile,
                Collections.singletonList(order), GroupedOrder.class);
        _poiMapper().writeValue(poiFile,
                Collections.singletonList(order), GroupedOrder.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // ----------------------------------------------------------------
    // 3) WIDE_TABLE — many columns (30+), with a nested list and outer
    //    fields after the list.
    // ----------------------------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class WideInner {
        @DataColumn("i1") int i1;
        @DataColumn("i2") int i2;
        @DataColumn("i3") int i3;
        @DataColumn("i4") int i4;
        @DataColumn("i5") int i5;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class WideRecord {
        @DataColumn(value = "h01", merge = OptBoolean.TRUE) String h01;
        @DataColumn(value = "h02", merge = OptBoolean.TRUE) String h02;
        @DataColumn(value = "h03", merge = OptBoolean.TRUE) String h03;
        @DataColumn(value = "h04", merge = OptBoolean.TRUE) String h04;
        @DataColumn(value = "h05", merge = OptBoolean.TRUE) String h05;
        @DataColumn(value = "h06", merge = OptBoolean.TRUE) int h06;
        @DataColumn(value = "h07", merge = OptBoolean.TRUE) int h07;
        @DataColumn(value = "h08", merge = OptBoolean.TRUE) int h08;
        @DataColumn(value = "h09", merge = OptBoolean.TRUE) double h09;
        @DataColumn(value = "h10", merge = OptBoolean.TRUE) double h10;
        List<WideInner> inners;
        @DataColumn(value = "f01", merge = OptBoolean.TRUE) String f01;
        @DataColumn(value = "f02", merge = OptBoolean.TRUE) String f02;
        @DataColumn(value = "f03", merge = OptBoolean.TRUE) String f03;
        @DataColumn(value = "f04", merge = OptBoolean.TRUE) String f04;
        @DataColumn(value = "f05", merge = OptBoolean.TRUE) int f05;
        @DataColumn(value = "f06", merge = OptBoolean.TRUE) int f06;
        @DataColumn(value = "f07", merge = OptBoolean.TRUE) int f07;
        @DataColumn(value = "f08", merge = OptBoolean.TRUE) double f08;
        @DataColumn(value = "f09", merge = OptBoolean.TRUE) double f09;
        @DataColumn(value = "f10", merge = OptBoolean.TRUE) double f10;
    }

    @Test
    void wideTable() throws Exception {
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        WideRecord record = new WideRecord(
                "h01v", "h02v", "h03v", "h04v", "h05v",
                1, 2, 3, 4.0, 5.0,
                Arrays.asList(new WideInner(11, 12, 13, 14, 15),
                        new WideInner(21, 22, 23, 24, 25)),
                "f01v", "f02v", "f03v", "f04v",
                10, 20, 30, 40.0, 50.0, 60.0);

        File ssmlFile = _debugFile("rat-wide-ssml.xlsx");
        File poiFile = _debugFile("rat-wide-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile,
                Collections.singletonList(record), WideRecord.class);
        _poiMapper().writeValue(poiFile,
                Collections.singletonList(record), WideRecord.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // ----------------------------------------------------------------
    // 4) VERY_LARGE_STRINGS — long memo (~1KB), several inners.
    // ----------------------------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class LargeStringInner {
        @DataColumn("name") String name;
        @DataColumn("desc") String description;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class LargeStringRecord {
        @DataColumn(value = "title", merge = OptBoolean.TRUE) String title;
        List<LargeStringInner> inners;
        @DataColumn(value = "memo", merge = OptBoolean.TRUE) String memo;
    }

    @Test
    void veryLargeStrings() throws Exception {
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        StringBuilder bigMemo = new StringBuilder(1024);
        for (int i = 0; i < 100; i++) bigMemo.append("긴 메모 텍스트 — Korean and English mix... ");
        StringBuilder bigDesc = new StringBuilder(2048);
        for (int i = 0; i < 200; i++) bigDesc.append("description chunk ").append(i).append(" / ");

        LargeStringRecord record = new LargeStringRecord(
                "주문 / Order — long title with mixed content abcdefghijabcdefghijabcdefghijabcdefghijabcdefghijabcdefghijabcdefghijabcdefghijabcdefghijabcdefghij",
                Arrays.asList(
                        new LargeStringInner("Apple", bigDesc.toString()),
                        new LargeStringInner("Banana", bigDesc.toString())),
                bigMemo.toString());

        File ssmlFile = _debugFile("rat-large-string-ssml.xlsx");
        File poiFile = _debugFile("rat-large-string-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile,
                Collections.singletonList(record), LargeStringRecord.class);
        _poiMapper().writeValue(poiFile,
                Collections.singletonList(record), LargeStringRecord.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // ----------------------------------------------------------------
    // 5) SPARSE_NULL — null fields (frequent in real data).
    // ----------------------------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class NullableInner {
        @DataColumn("a") String a;
        @DataColumn("b") Integer b;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NullableRecord {
        @DataColumn(value = "id", merge = OptBoolean.TRUE) Integer id;
        @DataColumn(value = "label", merge = OptBoolean.TRUE) String label;
        List<NullableInner> inners;
        @DataColumn(value = "note", merge = OptBoolean.TRUE) String note;
        @DataColumn(value = "summary", merge = OptBoolean.TRUE) Integer summary;
    }

    @Test
    void sparseNullValues() throws Exception {
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        NullableRecord record = new NullableRecord(
                1, null,
                Arrays.asList(new NullableInner("a", null),
                        new NullableInner(null, 99)),
                null, 999);

        File ssmlFile = _debugFile("rat-sparse-null-ssml.xlsx");
        File poiFile = _debugFile("rat-sparse-null-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile,
                Collections.singletonList(record), NullableRecord.class);
        _poiMapper().writeValue(poiFile,
                Collections.singletonList(record), NullableRecord.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // ----------------------------------------------------------------
    // Helpers
    // ----------------------------------------------------------------

    private static SpreadsheetMapper _poiMapper() {
        return new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new, SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
    }

    private static void _assertPartEqualIgnoringDimension(
            final File expected, final File actual, final String partName) throws Exception {
        try (OPCPackage expectedPkg = OPCPackage.open(expected);
             OPCPackage actualPkg = OPCPackage.open(actual)) {

            final Document expectedDoc = OpcXmlHelper.parsePart(expectedPkg, partName);
            final Document actualDoc = OpcXmlHelper.parsePart(actualPkg, partName);

            _removeDimensionElements(expectedDoc);
            _removeDimensionElements(actualDoc);

            assertThat(actualDoc.getDocumentElement().isEqualNode(
                    expectedDoc.getDocumentElement()))
                    .as("%s DOM equality (ignoring dimension)", partName)
                    .isTrue();
        }
    }

    private static void _removeDimensionElements(final Document doc) {
        final NodeList dimensions = doc.getElementsByTagNameNS(
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "dimension");
        for (int i = dimensions.getLength() - 1; i >= 0; i--) {
            final Node node = dimensions.item(i);
            node.getParentNode().removeChild(node);
        }
    }

    private static File _debugFile(final String name) throws IOException {
        Files.createDirectories(DEBUG_OUTPUT_DIR);
        return DEBUG_OUTPUT_DIR.resolve(name).toFile();
    }
}
