package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;

/**
 * Asserts SSML and POI writers produce byte-equal sheet1.xml for merge layouts
 * that exercise the inner-cell fill paths — column merges, group-header
 * merges, multi-row header merges, and nested-list outer-field merges, all
 * with a border-bearing column style so {@code _fillMergedInnerCells*} fires
 * on both writers. The existing {@code SSMLSheetWriterDomEquivalenceTest}
 * scenarios skip inner-cell fill because their columns carry no style.
 */
class MergeRegionInnerCellsTest {

    private static final Path DEBUG_OUTPUT_DIR = Paths.get("build/debug-output");

    private static StylesBuilder _borderStyles() {
        return new StylesBuilder().cellStyle("border").border().thin().end();
    }

    private static SpreadsheetMapper _ssmlMapper() {
        return SpreadsheetMapper.builder().stylesBuilder(_borderStyles()).build();
    }

    private static SpreadsheetMapper _poiMapper() {
        return SpreadsheetMapper.builder(
                new SpreadsheetFactory(XSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL))
                .stylesBuilder(_borderStyles())
                .build();
    }

    private static File _debugFile(final String name) throws IOException {
        Files.createDirectories(DEBUG_OUTPUT_DIR);
        return DEBUG_OUTPUT_DIR.resolve(name).toFile();
    }

    // (1) Vertical column merge — @DataColumn(merge = TRUE).

    @Data @NoArgsConstructor @AllArgsConstructor
    @DataGrid(columnStyle = "border", columnHeaderStyle = "border")
    static class VerticalMergeRow {
        @DataColumn(value = "group", merge = OptBoolean.TRUE) String group;
        @DataColumn("value") int value;
    }

    @Test
    void verticalColumnMerge_ssmlEqualsPoi() throws Exception {
        final List<VerticalMergeRow> data = Arrays.asList(
                new VerticalMergeRow("A", 1),
                new VerticalMergeRow("A", 2),
                new VerticalMergeRow("B", 3));

        final File ssmlFile = _debugFile("merge-inner-vertical-ssml.xlsx");
        final File poiFile = _debugFile("merge-inner-vertical-poi.xlsx");
        _ssmlMapper().writeValue(ssmlFile, data, VerticalMergeRow.class);
        _poiMapper().writeValue(poiFile, data, VerticalMergeRow.class);

        XlsxDomAssertions.assertPartEqualIgnoringDimension(
                poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // (2) Horizontal group-header merge + (3) multi-row header vertical merge
    // — same layout fires both paths: @DataColumnGroup horizontally merges
    // "Address" across zipcode/city, while {@code id}/{@code name} columns
    // vertically merge across the two-row header.

    @Data @NoArgsConstructor @AllArgsConstructor
    @DataGrid(columnStyle = "border", columnHeaderStyle = "border", groupHeaderStyle = "border")
    static class MultiRowHeaderRow {
        int id;
        String name;
        @DataColumnGroup("Address") Address address;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Address {
        @DataColumn("zipcode") String zipcode;
        @DataColumn("city") String city;
    }

    @Test
    void multiRowHeaderAndGroupMerge_ssmlEqualsPoi() throws Exception {
        final List<MultiRowHeaderRow> data = Arrays.asList(
                new MultiRowHeaderRow(1, "Alice", new Address("12345", "Seoul")),
                new MultiRowHeaderRow(2, "Bob", new Address("67890", "Busan")));

        final File ssmlFile = _debugFile("merge-inner-hdr-ssml.xlsx");
        final File poiFile = _debugFile("merge-inner-hdr-poi.xlsx");
        _ssmlMapper().writeValue(ssmlFile, data, MultiRowHeaderRow.class);
        _poiMapper().writeValue(poiFile, data, MultiRowHeaderRow.class);

        XlsxDomAssertions.assertPartEqualIgnoringDimension(
                poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // (2b) Group header merge with NO explicit groupHeaderStyle — the anchor
    // cell inherits columnHeaderStyle via the writeString path, and the inner
    // cells must inherit the same style. Without the fix the inner cells stay
    // empty and the border breaks along the group region.

    @Data @NoArgsConstructor @AllArgsConstructor
    @DataGrid(columnStyle = "border", columnHeaderStyle = "border")
    static class GroupHeaderInheritsRow {
        int id;
        String name;
        @DataColumnGroup("Address") Address address;
    }

    @Test
    void groupCellInheritsColumnHeaderStyle_ssmlEqualsPoi() throws Exception {
        final List<GroupHeaderInheritsRow> data = Arrays.asList(
                new GroupHeaderInheritsRow(1, "Alice", new Address("12345", "Seoul")),
                new GroupHeaderInheritsRow(2, "Bob", new Address("67890", "Busan")));

        final File ssmlFile = _debugFile("merge-inner-group-inherit-ssml.xlsx");
        final File poiFile = _debugFile("merge-inner-group-inherit-poi.xlsx");
        _ssmlMapper().writeValue(ssmlFile, data, GroupHeaderInheritsRow.class);
        _poiMapper().writeValue(poiFile, data, GroupHeaderInheritsRow.class);

        XlsxDomAssertions.assertPartEqualIgnoringDimension(
                poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // (4) Nested list outer-field merge — outer column merged across inner
    // list rows via mergeScopedColumns; fires the back-write path so the
    // inner-cell fill happens after inner items emit.

    @Data @NoArgsConstructor @AllArgsConstructor
    @DataGrid(columnStyle = "border", columnHeaderStyle = "border", groupHeaderStyle = "border")
    static class NestedListOrder {
        @DataColumn(value = "id", merge = OptBoolean.TRUE) int id;
        @DataColumnGroup("Items") List<LineItem> items;
        @DataColumn(value = "total", merge = OptBoolean.TRUE) double total;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class LineItem {
        @DataColumn("sku") String sku;
        @DataColumn("qty") int qty;
    }

    @Test
    void nestedListOuterMerge_ssmlEqualsPoi() throws Exception {
        final List<NestedListOrder> data = Arrays.asList(
                new NestedListOrder(1,
                        Arrays.asList(new LineItem("A1", 3), new LineItem("A2", 5)),
                        80.0));

        final File ssmlFile = _debugFile("merge-inner-nested-ssml.xlsx");
        final File poiFile = _debugFile("merge-inner-nested-poi.xlsx");
        _ssmlMapper().writeValue(ssmlFile, data, NestedListOrder.class);
        _poiMapper().writeValue(poiFile, data, NestedListOrder.class);

        XlsxDomAssertions.assertPartEqualIgnoringDimension(
                poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }
}
