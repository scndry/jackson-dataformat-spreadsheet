package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.InputStream;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec.CTCell;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec.STCellFormulaType;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec.STCellType;
import io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml.spec.SpreadsheetML;

/**
 * Lightweight StAX reader for OOXML SpreadsheetML.
 * Provides Matcher-based navigation, attribute access, and string content reading.
 */
final class XmlElementReader implements AutoCloseable {

    private static final XMLInputFactory XML_INPUT_FACTORY;

    static {
        XML_INPUT_FACTORY = XMLInputFactory.newInstance();
        XML_INPUT_FACTORY.setProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES, false);
        XML_INPUT_FACTORY.setProperty(XMLInputFactory.SUPPORT_DTD, false);
    }

    private final XMLStreamReader _xml;
    private boolean _closed;

    XmlElementReader(final InputStream src) {
        try {
            _xml = XML_INPUT_FACTORY.createXMLStreamReader(src);
        } catch (XMLStreamException e) {
            throw new IllegalStateException("Failed to initialize XML reader", e);
        }
    }

    // ---------------------------------------------------------------
    // Navigation
    // ---------------------------------------------------------------

    /** Navigate to the first occurrence of the named start element. */
    void navigateTo(final String localName) {
        try {
            while (_xml.hasNext()) {
                _xml.next();
                if (_xml.isStartElement() && localName.equals(_xml.getLocalName())) {
                    return;
                }
            }
        } catch (XMLStreamException e) {
            throw new IllegalStateException("Failed to navigate to <" + localName + ">", e);
        }
        throw new IllegalStateException("Element <" + localName + "> not found");
    }

    /** Advance until one of the matchers matches. Returns the matched Matcher,
            or {@code null} if stream ends. */
    Matcher nextUntil(final Matcher... matchers) {
        try {
            while (_xml.hasNext()) {
                final int event = _xml.next();
                if (event != XMLStreamConstants.START_ELEMENT &&
                        event != XMLStreamConstants.END_ELEMENT) continue;
                final String localName = _xml.getLocalName();
                for (final Matcher matcher : matchers) {
                    if (matcher.matches(event, localName)) {
                        return matcher;
                    }
                }
            }
            return null;
        } catch (XMLStreamException e) {
            throw new IllegalStateException("Failed to advance XML reader", e);
        }
    }

    // ---------------------------------------------------------------
    // Attribute access
    // ---------------------------------------------------------------

    String attribute(final String name) {
        return _xml.getAttributeValue(null, name);
    }

    String attribute(final String namespace, final String name) {
        return _xml.getAttributeValue(namespace, name);
    }

    int intAttribute(final String name) {
        return Integer.parseInt(_xml.getAttributeValue(null, name));
    }

    // ---------------------------------------------------------------
    // String content reading (CT_Rst content model)
    // ---------------------------------------------------------------

    /**
     * Read concatenated text from all {@code <t>} elements within the current element.
     * Works for both {@code <si>} (shared strings) and {@code <is>} (inline strings)
     * which share the same content model: plain {@code <t>} or rich text {@code <r><t>}.
     */
    String readStringContent() {
        try {
            final TextCollector collector = new TextCollector();
            _collectText(collector);
            return collector.result();
        } catch (XMLStreamException e) {
            throw new IllegalStateException("Failed to read string content", e);
        }
    }

    private void _collectText(final TextCollector collector) throws XMLStreamException {
        while (_xml.nextTag() == XMLStreamConstants.START_ELEMENT) {
            final String name = _xml.getLocalName();
            if (SpreadsheetML.TEXT.equals(name)) {
                collector.append(_xml.getElementText());
            } else if (SpreadsheetML.RICH_TEXT_RUN.equals(name)) {
                _collectText(collector);
            } else {
                skipElement();
            }
        }
    }

    private static final class TextCollector {
        private StringBuilder _sb;
        private String _first;

        void append(final String text) {
            if (_first == null) {
                _first = text;
            } else {
                if (_sb == null) _sb = new StringBuilder(_first);
                _sb.append(text);
            }
        }

        String result() {
            return _sb != null ? _sb.toString() : _first;
        }
    }

    // ---------------------------------------------------------------
    // Cell collection (CT_Cell — ECMA-376 §18.3.1.4)
    // ---------------------------------------------------------------

    /** Collect cell attributes and children into a lightweight {@link CTCell}. */
    CTCell collectCell() {
        try {
            final String r = _xml.getAttributeValue(null, SpreadsheetML.ATTR_REF);
            final STCellType t = STCellType
                    .of(_xml.getAttributeValue(null, SpreadsheetML.ATTR_TYPE));
            String v = null;
            STCellFormulaType ft = null;
            String is = null;

            while (_xml.nextTag() == XMLStreamConstants.START_ELEMENT) {
                switch (_xml.getLocalName()) {
                    case SpreadsheetML.VALUE:
                        v = _xml.getElementText();
                        break;
                    case SpreadsheetML.FORMULA:
                        ft = STCellFormulaType
                                .of(_xml.getAttributeValue(null, SpreadsheetML.ATTR_TYPE));
                        _xml.getElementText();
                        break;
                    case SpreadsheetML.INLINE_STRING:
                        is = readStringContent();
                        break;
                    default:
                        skipElement();
                        break;
                }
            }
            return new CTCell(r, t, v, ft, is);
        } catch (XMLStreamException e) {
            throw new IllegalStateException("Failed to read cell", e);
        }
    }

    // ---------------------------------------------------------------
    // Utilities
    // ---------------------------------------------------------------

    /** Skip the current element and all its children. */
    void skipElement() {
        try {
            int depth = 1;
            while (depth > 0) {
                switch (_xml.next()) {
                    case XMLStreamConstants.START_ELEMENT: depth++; break;
                    case XMLStreamConstants.END_ELEMENT:   depth--; break;
                }
            }
        } catch (XMLStreamException e) {
            throw new IllegalStateException("Failed to skip element", e);
        }
    }

    // ---------------------------------------------------------------
    // Lifecycle
    // ---------------------------------------------------------------

    boolean isClosed() {
        return _closed;
    }

    @Override
    public void close() {
        try {
            _xml.close();
            _closed = true;
        } catch (XMLStreamException e) {
            throw new IllegalStateException("Failed to close XML reader", e);
        }
    }

    // ---------------------------------------------------------------
    // Matcher
    // ---------------------------------------------------------------

    static final class Matcher {

        private final int _eventType;
        private final String _localName;

        private Matcher(final int eventType, final String localName) {
            _eventType = eventType;
            _localName = localName;
        }

        static Matcher startElement(final String localName) {
            return new Matcher(XMLStreamConstants.START_ELEMENT, localName);
        }

        static Matcher endElement(final String localName) {
            return new Matcher(XMLStreamConstants.END_ELEMENT, localName);
        }

        boolean matches(final int eventType, final String localName) {
            return _eventType == eventType && _localName.equals(localName);
        }

        boolean isEndElement() {
            return _eventType == XMLStreamConstants.END_ELEMENT;
        }
    }
}
