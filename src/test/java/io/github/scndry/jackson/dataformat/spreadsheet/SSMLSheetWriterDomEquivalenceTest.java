package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.fasterxml.jackson.annotation.OptBoolean;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * DOM-level equivalence tests: SSML writer output vs POI XSSFWorkbook output.
 * Verifies that sheet1.xml and sharedStrings.xml produced by SSMLSheetWriter
 * are structurally identical to POI's output.
 */
class SSMLSheetWriterDomEquivalenceTest {

    private static final Path DEBUG_OUTPUT_DIR = Paths.get("build/debug-output");

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class Entry {
        String name;
        int qty;
        boolean active;
    }

    private static final List<Entry> DATA = Arrays.asList(
            new Entry("Alice", 10, true),
            new Entry("Bob", 20, false),
            new Entry("<Special & \"Chars\">", 0, true));

    @Test
    void sheetXmlDomEquivalent() throws Exception {
        File ssmlFile = _debugFile("dom-sheet-ssml.xlsx");
        File poiFile = _debugFile("dom-sheet-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, DATA, Entry.class);
        _poiMapper().writeValue(poiFile, DATA, Entry.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // -- Nested list with outer field declared after the list ------------
    // Trigger for SSML back-reference: Jackson serializes the outer field
    // after the list, but the cell anchors at the list's first element row.
    // merge=TRUE here only affects mergeCells; the back-write is independent.

    @Data @NoArgsConstructor @AllArgsConstructor
    static class NestedItem {
        String product;
        int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NestedListEntry {
        @DataColumn(value = "id", merge = OptBoolean.TRUE) int id;
        List<NestedItem> items;
        @DataColumn(value = "total", merge = OptBoolean.TRUE) double total;
    }

    @Test
    void sheetXmlDomEquivalent_nestedList() throws Exception {
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        List<NestedListEntry> data = Arrays.asList(
                new NestedListEntry(1,
                        Arrays.asList(new NestedItem("Apple", 3), new NestedItem("Banana", 5)),
                        10.0));

        File ssmlFile = _debugFile("dom-nested-list-ssml.xlsx");
        File poiFile = _debugFile("dom-nested-list-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, data, NestedListEntry.class);
        _poiMapper().writeValue(poiFile, data, NestedListEntry.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    @Test
    void sheetXmlDomEquivalent_nestedList_largeArray() throws Exception {
        // List XML accumulation crosses the SSML buffer threshold (~512 KB).
        // Tests that flush suspension during nested array scope keeps the
        // first element row tag present for back-write of the outer field.
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        int innerCount = 10000;
        NestedItem[] items = new NestedItem[innerCount];
        for (int i = 0; i < innerCount; i++) {
            items[i] = new NestedItem("p" + i, i);
        }
        List<NestedListEntry> data = Arrays.asList(
                new NestedListEntry(1, Arrays.asList(items), 99.0));

        File ssmlFile = _debugFile("dom-nested-list-large-ssml.xlsx");
        File poiFile = _debugFile("dom-nested-list-large-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, data, NestedListEntry.class);
        _poiMapper().writeValue(poiFile, data, NestedListEntry.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    @Test
    void sheetXmlDomEquivalent_nestedList_multiRecord() throws Exception {
        // _arrayScopeDepth must cycle 0 → 1 → 0 across records when the
        // root is an array of records (mapper.writeValue(list, T.class)).
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        List<NestedListEntry> data = Arrays.asList(
                new NestedListEntry(1,
                        Arrays.asList(new NestedItem("Apple", 3), new NestedItem("Banana", 5)),
                        10.0),
                new NestedListEntry(2,
                        Arrays.asList(new NestedItem("Cherry", 2)),
                        20.0));

        File ssmlFile = _debugFile("dom-nested-list-multi-ssml.xlsx");
        File poiFile = _debugFile("dom-nested-list-multi-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, data, NestedListEntry.class);
        _poiMapper().writeValue(poiFile, data, NestedListEntry.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // -- Two nested lists in one record, outer fields between/after ------
    // Each outer-after-list field back-writes into the first row of the
    // immediately-preceding list. The two back-write targets are the same
    // first record row, but reached via different array scopes — verifies
    // _isBackReference() resolves the target row from _reference and
    // _insertCellIntoEmittedRow finds the row tag in _sb each time.

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Payment {
        @DataColumn("method") String method;
        @DataColumn("amount") int amount;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class OrderWithTwoLists {
        @DataColumn(value = "id", merge = OptBoolean.TRUE) int id;
        List<NestedItem> items;
        @DataColumn(value = "subtotal", merge = OptBoolean.TRUE) int subtotal;
        List<Payment> payments;
        @DataColumn(value = "totalPaid", merge = OptBoolean.TRUE) int totalPaid;
    }

    @Test
    void sheetXmlDomEquivalent_twoNestedListsPerRecord() throws Exception {
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        List<OrderWithTwoLists> data = Arrays.asList(
                new OrderWithTwoLists(1,
                        Arrays.asList(new NestedItem("Apple", 3), new NestedItem("Banana", 5)),
                        8,
                        Arrays.asList(new Payment("CARD", 5), new Payment("CASH", 3)),
                        8));

        File ssmlFile = _debugFile("dom-two-lists-ssml.xlsx");
        File poiFile = _debugFile("dom-two-lists-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, data, OrderWithTwoLists.class);
        _poiMapper().writeValue(poiFile, data, OrderWithTwoLists.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // -- List-of-list with an outer field after the inner list -----------
    // Inner list's outer field anchors at the inner list's first item row
    // (the enclosing Group object's parent ArrayContext + index). The
    // outer list's outer field anchors at the outermost record row.
    // Validates that nested array-scope depth cycles 0 → 1 → 2 → 1 → 2 →
    // 1 → 0 across the structure without losing back-write targets.

    @Data @NoArgsConstructor @AllArgsConstructor
    static class GroupItem {
        @DataColumn("label") String label;
        @DataColumn("qty") int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Group {
        @DataColumn(value = "groupName", merge = OptBoolean.TRUE) String groupName;
        List<GroupItem> items;
        @DataColumn(value = "groupSubtotal", merge = OptBoolean.TRUE) int groupSubtotal;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class GroupedOrder {
        @DataColumn(value = "orderId", merge = OptBoolean.TRUE) int orderId;
        List<Group> groups;
        @DataColumn(value = "total", merge = OptBoolean.TRUE) int total;
    }

    @Test
    void sheetXmlDomEquivalent_listOfList_outerAfterInnerList() throws Exception {
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        List<GroupedOrder> data = Arrays.asList(
                new GroupedOrder(1,
                        Arrays.asList(
                                new Group("X",
                                        Arrays.asList(new GroupItem("a", 1), new GroupItem("b", 2)),
                                        3),
                                new Group("Y",
                                        Arrays.asList(new GroupItem("c", 4)),
                                        4)),
                        7));

        File ssmlFile = _debugFile("dom-list-of-list-ssml.xlsx");
        File poiFile = _debugFile("dom-list-of-list-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, data, GroupedOrder.class);
        _poiMapper().writeValue(poiFile, data, GroupedOrder.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // -- Multi-row header (@DataColumnGroup) -----------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class GroupedAddress {
        String city;
        String zip;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class GroupedEntry {
        int id;
        String name;
        @DataColumnGroup("Address") GroupedAddress address;
    }

    private static final List<GroupedEntry> GROUPED_DATA = Arrays.asList(
            new GroupedEntry(1, "Alice", new GroupedAddress("Seoul", "12345")),
            new GroupedEntry(2, "Bob", new GroupedAddress("Busan", "23456")));

    @Test
    void sheetXmlDomEquivalent_multiRowHeader() throws Exception {
        File ssmlFile = _debugFile("dom-multirow-ssml.xlsx");
        File poiFile = _debugFile("dom-multirow-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, GROUPED_DATA, GroupedEntry.class);
        _poiMapper().writeValue(poiFile, GROUPED_DATA, GroupedEntry.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    @Test
    void sharedStringsXmlDomEquivalent() throws Exception {
        File ssmlFile = _debugFile("dom-sst-ssml.xlsx");
        File poiFile = _debugFile("dom-sst-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, DATA, Entry.class);
        _poiMapper().writeValue(poiFile, DATA, Entry.class);

        _assertPartEqual(poiFile, ssmlFile, "/xl/sharedStrings.xml");
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

            // POI 4.x ~ 5.2.2 omit default s="0" on cells; SSML writer always emits.
            // Strip on the SSML side so byte-equal comparison holds across versions.
            // On POI 5.2.3+ both sides emit, and the strict check is preserved.
            if (!PoiVersionProbe.isPoi523OrLater()) {
                _stripDefaultStyleAttribute(actualDoc);
            }

            assertThat(actualDoc.getDocumentElement().isEqualNode(
                    expectedDoc.getDocumentElement()))
                    .as("%s DOM equality (ignoring dimension)", partName)
                    .isTrue();
        }
    }

    private static void _assertPartEqual(
            final File expected, final File actual, final String partName) throws Exception {
        try (OPCPackage expectedPkg = OPCPackage.open(expected);
             OPCPackage actualPkg = OPCPackage.open(actual)) {

            final Document expectedDoc = OpcXmlHelper.parsePart(expectedPkg, partName);
            final Document actualDoc = OpcXmlHelper.parsePart(actualPkg, partName);

            assertThat(actualDoc.getDocumentElement().isEqualNode(
                    expectedDoc.getDocumentElement()))
                    .as("%s DOM equality", partName)
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

    private static void _stripDefaultStyleAttribute(final Document doc) {
        final NodeList cells = doc.getElementsByTagNameNS(
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "c");
        for (int i = 0; i < cells.getLength(); i++) {
            final Element cell = (Element) cells.item(i);
            if ("0".equals(cell.getAttribute("s"))) {
                cell.removeAttribute("s");
            }
        }
    }

    private static File _debugFile(final String name) throws IOException {
        Files.createDirectories(DEBUG_OUTPUT_DIR);
        return DEBUG_OUTPUT_DIR.resolve(name).toFile();
    }
}
