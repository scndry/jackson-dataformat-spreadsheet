package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.XmlElementReader.Matcher;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Direct tests for {@link XmlElementReader} branches that integration
 * tests don't reach: navigation failure, namespaced attributes, rich-text
 * concatenation, skipping unknown children, all {@link CTCell} sub-element
 * paths, and the {@link Matcher} factory methods.
 */
class XmlElementReaderTest {

    private static XmlElementReader reader(String xml) {
        return new XmlElementReader(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)));
    }

    @Test
    void navigateTo_elementNotFound_throws() {
        XmlElementReader r = reader("<root><foo/></root>");
        assertThatThrownBy(() -> r.navigateTo("bar"))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("not found");
    }

    @Test
    void attribute_withNamespace_returnsValue() {
        XmlElementReader r = reader(
                "<root xmlns:r=\"http://example/ns\"><a r:id=\"123\"/></root>");
        r.navigateTo("a");
        assertThat(r.attribute("http://example/ns", "id")).isEqualTo("123");
    }

    @Test
    void intAttribute_returnsParsedInt() {
        XmlElementReader r = reader("<root><a count=\"42\"/></root>");
        r.navigateTo("a");
        assertThat(r.intAttribute("count")).isEqualTo(42);
    }

    @Test
    void readStringContent_singleText() {
        XmlElementReader r = reader("<root><si><t>hello</t></si></root>");
        r.navigateTo("si");
        assertThat(r.readStringContent()).isEqualTo("hello");
    }

    @Test
    void readStringContent_richTextRun_concatenated() {
        XmlElementReader r = reader(
                "<root><si><r><t>hello </t></r><r><t>world</t></r></si></root>");
        r.navigateTo("si");
        assertThat(r.readStringContent()).isEqualTo("hello world");
    }

    @Test
    void readStringContent_skipsUnknownChildren() {
        XmlElementReader r = reader(
                "<root><si><t>main</t><phoneticPr><foo/></phoneticPr></si></root>");
        r.navigateTo("si");
        assertThat(r.readStringContent()).isEqualTo("main");
    }

    @Test
    void collectCell_value() {
        XmlElementReader r = reader(
                "<root><c r=\"A1\" t=\"n\"><v>42</v></c></root>");
        r.navigateTo("c");
        CTCell cell = r.collectCell();
        assertThat(cell.getR()).isEqualTo("A1");
        assertThat(cell.getV()).isEqualTo("42");
        assertThat(cell.getIs()).isNull();
    }

    @Test
    void collectCell_formula() {
        XmlElementReader r = reader(
                "<root><c r=\"B2\"><f t=\"shared\">A1+1</f><v>5</v></c></root>");
        r.navigateTo("c");
        CTCell cell = r.collectCell();
        assertThat(cell.getR()).isEqualTo("B2");
        assertThat(cell.getFt()).isNotNull();
        assertThat(cell.getV()).isEqualTo("5");
    }

    @Test
    void collectCell_inlineString() {
        XmlElementReader r = reader(
                "<root><c r=\"C3\" t=\"inlineStr\"><is><t>inline</t></is></c></root>");
        r.navigateTo("c");
        CTCell cell = r.collectCell();
        assertThat(cell.getR()).isEqualTo("C3");
        assertThat(cell.getIs()).isEqualTo("inline");
    }

    @Test
    void collectCell_skipsUnknownChildren() {
        XmlElementReader r = reader(
                "<root><c r=\"D4\"><unknownChild/><v>9</v></c></root>");
        r.navigateTo("c");
        CTCell cell = r.collectCell();
        assertThat(cell.getR()).isEqualTo("D4");
        assertThat(cell.getV()).isEqualTo("9");
    }

    @Test
    void nextUntil_returnsNullAtStreamEnd() {
        XmlElementReader r = reader("<root><a/></root>");
        Matcher zMatcher = Matcher.startElement("z");  // never present
        assertThat(r.nextUntil(zMatcher)).isNull();
    }

    @Test
    void close_setsIsClosedTrue() {
        XmlElementReader r = reader("<root/>");
        assertThat(r.isClosed()).isFalse();
        r.close();
        assertThat(r.isClosed()).isTrue();
    }

    @Test
    void matcher_endElement_isEndElementTrue() {
        Matcher end = Matcher.endElement("foo");
        assertThat(end.isEndElement()).isTrue();
    }

    @Test
    void matcher_startElement_isEndElementFalse() {
        Matcher start = Matcher.startElement("foo");
        assertThat(start.isEndElement()).isFalse();
    }
}
