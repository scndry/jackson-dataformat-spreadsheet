package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.annotation.OptBoolean;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;

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

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
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
        List<NestedListEntry> data = Arrays.asList(
                new NestedListEntry(1,
                        Arrays.asList(new NestedItem("Apple", 3), new NestedItem("Banana", 5)),
                        10.0));

        File ssmlFile = _debugFile("dom-nested-list-ssml.xlsx");
        File poiFile = _debugFile("dom-nested-list-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, data, NestedListEntry.class);
        _poiMapper().writeValue(poiFile, data, NestedListEntry.class);

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    @Test
    void sheetXmlDomEquivalent_nestedList_largeArray() throws Exception {
        // List XML accumulation crosses the SSML buffer threshold (~512 KB).
        // Tests that flush suspension during nested array scope keeps the
        // first element row tag present for back-write of the outer field.
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

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    @Test
    void sheetXmlDomEquivalent_nestedList_multiRecord() throws Exception {
        // _arrayScopeDepth must cycle 0 → 1 → 0 across records when the
        // root is an array of records (mapper.writeValue(list, T.class)).
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

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
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

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
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

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
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

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // -- @DataColumnGroup styled attributes: headerStyle on the group cell,
    //    columnStyle / columnHeaderStyle cascading into the leaf children.
    //    Verifies SSML applies the resolved styles at the same cell positions
    //    that POI does (same style index for the same cell).

    @Data @NoArgsConstructor @AllArgsConstructor
    static class StyledLeaves {
        @DataColumn("city") String city;
        @DataColumn("zip") String zip;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class StyledGroupedEntry {
        @DataColumn("id") int id;
        @DataColumnGroup(value = "address",
                headerStyle = "groupHdr",
                columnStyle = "childData",
                columnHeaderStyle = "childHdr")
        StyledLeaves address;
    }

    @Test
    void sheetXmlDomEquivalent_groupStyledAttributes() throws Exception {
        File ssmlFile = _debugFile("dom-group-styled-ssml.xlsx");
        File poiFile = _debugFile("dom-group-styled-poi.xlsx");

        List<StyledGroupedEntry> data = Arrays.asList(
                new StyledGroupedEntry(1, new StyledLeaves("Seoul", "12345")),
                new StyledGroupedEntry(2, new StyledLeaves("Busan", "67890")));

        _styledMapper().writeValue(ssmlFile, data, StyledGroupedEntry.class);
        _styledPoiMapper().writeValue(poiFile, data, StyledGroupedEntry.class);

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // -- 2-level @DataColumnGroup (3 header rows) + nested List + outer
    //    fields with merge=TRUE after the list. Tests that back-write
    //    into the first record row resolves against schema.getDataRow()
    //    (origin + headerRowCount = 3, not origin + 1), so the totals
    //    block lands on the first data row — not on a header row.

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ProductDetail {
        @DataColumn("sku") String sku;
        @DataColumn("name") String name;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class LineItemGrouped {
        @DataColumnGroup("Product") ProductDetail product;
        @DataColumn("qty") int qty;
        @DataColumn("amount") BigDecimal amount;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class TotalsBlock {
        @DataColumn(value = "subtotal", merge = OptBoolean.TRUE) BigDecimal subtotal;
        @DataColumn(value = "tax",      merge = OptBoolean.TRUE) BigDecimal tax;
        @DataColumn(value = "total",    merge = OptBoolean.TRUE) BigDecimal total;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class OrderWith2LevelGroupAndList {
        @DataColumn(value = "id",       merge = OptBoolean.TRUE) int id;
        @DataColumn(value = "customer", merge = OptBoolean.TRUE) String customer;
        @DataColumnGroup("Items") List<LineItemGrouped> items;
        @DataColumnGroup("Totals") TotalsBlock totals;
        @DataColumn(value = "memo",     merge = OptBoolean.TRUE) String memo;
    }

    @Test
    void sheetXmlDomEquivalent_twoLevelGroup_nestedList_mergeTrue() throws Exception {
        List<OrderWith2LevelGroupAndList> data = Arrays.asList(
                new OrderWith2LevelGroupAndList(1, "Alice",
                        Arrays.asList(
                                new LineItemGrouped(new ProductDetail("A1", "Apple"),
                                        3, BigDecimal.valueOf(3000)),
                                new LineItemGrouped(new ProductDetail("A2", "Banana"),
                                        5, BigDecimal.valueOf(5000))),
                        new TotalsBlock(BigDecimal.valueOf(8000),
                                BigDecimal.valueOf(800),
                                BigDecimal.valueOf(8800)),
                        "Paid in full"),
                new OrderWith2LevelGroupAndList(2, "Bob",
                        Arrays.asList(
                                new LineItemGrouped(new ProductDetail("B1", "Cherry"),
                                        2, BigDecimal.valueOf(3000))),
                        new TotalsBlock(BigDecimal.valueOf(3000),
                                BigDecimal.valueOf(300),
                                BigDecimal.valueOf(3300)),
                        "Net 30"));

        File ssmlFile = _debugFile("dom-2level-group-nested-list-ssml.xlsx");
        File poiFile = _debugFile("dom-2level-group-nested-list-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, data, OrderWith2LevelGroupAndList.class);
        _poiMapper().writeValue(poiFile, data, OrderWith2LevelGroupAndList.class);

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // -- @DataColumn(shift) / @DataColumnGroup(shift) ---------------------
    // Regression guard for null-gap-column handling across writer iteration
    // paths. The nested-list case triggers mergeScopedColumns with a root
    // (empty) pointer, exercising every column-iterating path on POI write.

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class FlatShiftEntry {
        @DataColumn String name;
        @DataColumn int qty;
        @DataColumn(shift = 1) double total;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class ShiftItem {
        @DataColumn String product;
        @DataColumn int qty;
        @DataColumn(shift = 1) double amount;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class GroupShiftEntry {
        @DataColumn(anchor = true) String orderId;
        @DataColumnGroup(value = "Items", shift = 1) List<ShiftItem> items;
        @DataColumn double total;
    }

    @Test
    void sheetXmlDomEquivalent_flatShift() throws Exception {
        File ssmlFile = _debugFile("dom-shift-flat-ssml.xlsx");
        File poiFile = _debugFile("dom-shift-flat-poi.xlsx");

        List<FlatShiftEntry> data = Arrays.asList(
                new FlatShiftEntry("Apple", 10, 15.50),
                new FlatShiftEntry("Banana", 20, 8.00));

        new SpreadsheetMapper().writeValue(ssmlFile, data, FlatShiftEntry.class);
        _poiMapper().writeValue(poiFile, data, FlatShiftEntry.class);

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    @Test
    void sheetXmlDomEquivalent_groupShift_nestedList() throws Exception {
        File ssmlFile = _debugFile("dom-shift-group-ssml.xlsx");
        File poiFile = _debugFile("dom-shift-group-poi.xlsx");

        List<GroupShiftEntry> data = Arrays.asList(
                new GroupShiftEntry("ORD-1", Arrays.asList(
                        new ShiftItem("Apple", 3, 9.0),
                        new ShiftItem("Banana", 5, 5.0)), 14.0));

        new SpreadsheetMapper().writeValue(ssmlFile, data, GroupShiftEntry.class);
        _poiMapper().writeValue(poiFile, data, GroupShiftEntry.class);

        XlsxDomAssertions.assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    @Test
    void sharedStringsXmlDomEquivalent() throws Exception {
        File ssmlFile = _debugFile("dom-sst-ssml.xlsx");
        File poiFile = _debugFile("dom-sst-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, DATA, Entry.class);
        _poiMapper().writeValue(poiFile, DATA, Entry.class);

        XlsxDomAssertions.assertPartEqual(poiFile, ssmlFile, "/xl/sharedStrings.xml");
    }

    // ----------------------------------------------------------------
    // Helpers
    // ----------------------------------------------------------------

    private static SpreadsheetMapper _poiMapper() {
        return new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new, SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
    }

    private static StylesBuilder _groupStyles() {
        return new StylesBuilder()
                .cellStyle("groupHdr")
                    .fillForegroundColor(IndexedColors.YELLOW)
                    .fillPattern().solidForeground()
                    .end()
                .cellStyle("childHdr")
                    .fillForegroundColor(IndexedColors.GREY_25_PERCENT)
                    .fillPattern().solidForeground()
                    .end()
                .cellStyle("childData")
                    .font().bold().end()
                    .end();
    }

    private static SpreadsheetMapper _styledMapper() {
        return SpreadsheetMapper.builder().stylesBuilder(_groupStyles()).build();
    }

    private static SpreadsheetMapper _styledPoiMapper() {
        return SpreadsheetMapper.builder()
                .stylesBuilder(_groupStyles())
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
    }

    private static File _debugFile(final String name) throws IOException {
        Files.createDirectories(DEBUG_OUTPUT_DIR);
        return DEBUG_OUTPUT_DIR.resolve(name).toFile();
    }
}
