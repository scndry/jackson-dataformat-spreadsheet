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
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.fasterxml.jackson.annotation.OptBoolean;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
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
        // SSML writer always emits <c s="0">, matching POI 5.2.3+ (bug-51037 fix).
        // POI 4.x ~ 5.2.2 omit default s, so DOM equality only holds on POI 5.2.3+.
        // Library policy: follow latest POI behavior; older-version drift is ignored. (#96)
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        File ssmlFile = _debugFile("dom-sheet-ssml.xlsx");
        File poiFile = _debugFile("dom-sheet-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, DATA, Entry.class);
        _poiMapper().writeValue(poiFile, DATA, Entry.class);

        _assertPartEqualIgnoringDimension(poiFile, ssmlFile, "/xl/worksheets/sheet1.xml");
    }

    // -- Nested list + merge=TRUE on outer fields ------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class MergedListItem {
        String product;
        int qty;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class MergedListOrder {
        @DataColumn(value = "id", merge = OptBoolean.TRUE) int id;
        List<MergedListItem> items;
        @DataColumn(value = "total", merge = OptBoolean.TRUE) double total;
    }

    private static final List<MergedListOrder> MERGED_LIST_DATA = Arrays.asList(
            new MergedListOrder(1,
                    Arrays.asList(new MergedListItem("Apple", 3), new MergedListItem("Banana", 5)),
                    10.0));

    @Test
    void sheetXmlDomEquivalent_listWithMerge() throws Exception {
        Assumptions.assumeTrue(PoiVersionProbe.isPoi523OrLater(),
                "DOM equivalence asserted only on POI 5.2.3+ — see #96");

        File ssmlFile = _debugFile("dom-list-merge-ssml.xlsx");
        File poiFile = _debugFile("dom-list-merge-poi.xlsx");

        new SpreadsheetMapper().writeValue(ssmlFile, MERGED_LIST_DATA, MergedListOrder.class);
        _poiMapper().writeValue(poiFile, MERGED_LIST_DATA, MergedListOrder.class);

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

    private static File _debugFile(final String name) throws IOException {
        Files.createDirectories(DEBUG_OUTPUT_DIR);
        return DEBUG_OUTPUT_DIR.resolve(name).toFile();
    }
}
