package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.util.regex.Pattern;
import javax.xml.namespace.QName;

/**
 * Utility for converting between OOXML Transitional and Strict namespace formats.
 * <p>
 * OOXML has two namespace families:
 * <ul>
 *   <li>Transitional: {@code http://schemas.openxmlformats.org/{area}/2006/{type}}</li>
 *   <li>Strict: {@code http://purl.oclc.org/ooxml/{area}/{type}}</li>
 * </ul>
 * Most Excel files use Transitional. Files saved in "Strict Open XML" mode use Strict.
 */
final class OoxmlNamespace {

    private static final Pattern TRANSITIONAL = Pattern
            .compile("http://schemas\\.openxmlformats\\.org/(\\w+)/2006/(\\w+)");
    private static final Pattern STRICT = Pattern
            .compile("http://purl\\.oclc\\.org/ooxml/(\\w+)/(\\w+)");

    private OoxmlNamespace() {
    }

    /**
     * Converts a Strict namespace URI to Transitional.
     * Returns the original URI if it is not in Strict format.
     */
    static String toTransitional(final String namespaceURI) {
        return STRICT
                .matcher(namespaceURI)
                .replaceFirst("http://schemas.openxmlformats.org/$1/2006/$2");
    }

    /**
     * Converts a Transitional namespace URI to Strict.
     * Returns the original URI if it is not in Transitional format.
     */
    static String toStrict(final String namespaceURI) {
        return TRANSITIONAL.matcher(namespaceURI).replaceFirst("http://purl.oclc.org/ooxml/$1/$2");
    }

    /**
     * Converts a QName with a Strict namespace to Transitional.
     * Returns the original QName if it is not in Strict format.
     */
    static QName toTransitional(final QName name) {
        final String converted = toTransitional(name.getNamespaceURI());
        if (converted.equals(name.getNamespaceURI())) {
            return name;
        }
        return new QName(converted, name.getLocalPart(), name.getPrefix());
    }
}
