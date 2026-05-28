package io.github.scndry.jackson.dataformat.spreadsheet;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.File;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * DOM-equality assertions for OOXML parts, normalising POI version-specific
 * variance so the same assertion holds across POI 4.x ~ 5.5.x.
 */
final class XlsxDomAssertions {

    private XlsxDomAssertions() {}

    /**
     * Asserts the named part is DOM-equal between {@code expected} and
     * {@code actual}, ignoring the {@code <dimension>} element and the
     * default {@code s="0"} cell-style attribute. Both sides may or may
     * not emit {@code s="0"} depending on POI version (5.2.3+ emits it
     * unconditionally via bug-51037 fix) and on whether the writer
     * assigns a default style to newly created cells; treating it as
     * equivalent to "no style" keeps the comparison stable across
     * configurations.
     */
    static void assertPartEqualIgnoringDimension(
            final File expected, final File actual, final String partName) throws Exception {
        try (OPCPackage expectedPkg = OPCPackage.open(expected);
             OPCPackage actualPkg = OPCPackage.open(actual)) {

            final Document expectedDoc = OpcXmlHelper.parsePart(expectedPkg, partName);
            final Document actualDoc = OpcXmlHelper.parsePart(actualPkg, partName);

            _removeDimensionElements(expectedDoc);
            _removeDimensionElements(actualDoc);

            _stripDefaultStyleAttribute(expectedDoc);
            _stripDefaultStyleAttribute(actualDoc);

            assertThat(actualDoc.getDocumentElement().isEqualNode(
                    expectedDoc.getDocumentElement()))
                    .as("%s DOM equality (ignoring dimension)", partName)
                    .isTrue();
        }
    }

    static void assertPartEqual(
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
                OpcXmlHelper.NS_SPREADSHEETML, "dimension");
        for (int i = dimensions.getLength() - 1; i >= 0; i--) {
            final Node node = dimensions.item(i);
            node.getParentNode().removeChild(node);
        }
    }

    private static void _stripDefaultStyleAttribute(final Document doc) {
        final NodeList cells = doc.getElementsByTagNameNS(
                OpcXmlHelper.NS_SPREADSHEETML, "c");
        for (int i = 0; i < cells.getLength(); i++) {
            final Element cell = (Element) cells.item(i);
            if ("0".equals(cell.getAttribute("s"))) {
                cell.removeAttribute("s");
            }
        }
    }
}
