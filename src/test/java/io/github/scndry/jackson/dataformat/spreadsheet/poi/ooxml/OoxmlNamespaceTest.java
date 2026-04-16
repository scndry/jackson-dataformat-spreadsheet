package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import org.junit.jupiter.api.Test;

import javax.xml.namespace.QName;

import static org.assertj.core.api.Assertions.assertThat;

class OoxmlNamespaceTest {

    @Test
    void strictToTransitional() {
        String strict = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        String expected = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        assertThat(OoxmlNamespace.toTransitional(strict)).isEqualTo(expected);
    }

    @Test
    void transitionalToStrict() {
        String transitional = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        String expected = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        assertThat(OoxmlNamespace.toStrict(transitional)).isEqualTo(expected);
    }

    @Test
    void nonMatchingPassesThrough() {
        String unrelated = "http://example.com/something";
        assertThat(OoxmlNamespace.toTransitional(unrelated)).isEqualTo(unrelated);
        assertThat(OoxmlNamespace.toStrict(unrelated)).isEqualTo(unrelated);
    }

    @Test
    void qnameStrictToTransitional() {
        QName strict = new QName("http://purl.oclc.org/ooxml/spreadsheetml/main", "worksheet", "");
        QName result = OoxmlNamespace.toTransitional(strict);
        assertThat(result.getNamespaceURI()).isEqualTo("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        assertThat(result.getLocalPart()).isEqualTo("worksheet");
    }

    @Test
    void qnameAlreadyTransitionalReturnsSameInstance() {
        QName transitional = new QName("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "worksheet", "");
        assertThat(OoxmlNamespace.toTransitional(transitional)).isSameAs(transitional);
    }

    @Test
    void roundTrip() {
        String original = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        String strict = OoxmlNamespace.toStrict(original);
        String back = OoxmlNamespace.toTransitional(strict);
        assertThat(back).isEqualTo(original);
    }
}
