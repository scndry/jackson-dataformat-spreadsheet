package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import javax.xml.parsers.DocumentBuilderFactory;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

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

            final DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            dbf.setNamespaceAware(true);

            final Document expectedDoc = _parsePart(expectedPkg, partName, dbf);
            final Document actualDoc = _parsePart(actualPkg, partName, dbf);

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

            final DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            dbf.setNamespaceAware(true);

            final Document expectedDoc = _parsePart(expectedPkg, partName, dbf);
            final Document actualDoc = _parsePart(actualPkg, partName, dbf);

            assertThat(actualDoc.getDocumentElement().isEqualNode(
                    expectedDoc.getDocumentElement()))
                    .as("%s DOM equality", partName)
                    .isTrue();
        }
    }

    private static Document _parsePart(
            final OPCPackage pkg,
            final String partName,
            final DocumentBuilderFactory dbf) throws Exception {
        try (InputStream is = pkg.getPart(
                PackagingURIHelper.createPartName(partName)).getInputStream()) {
            return dbf.newDocumentBuilder().parse(is);
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
