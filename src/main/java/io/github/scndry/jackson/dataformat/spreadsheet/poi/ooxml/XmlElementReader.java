package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import lombok.extern.slf4j.Slf4j;
import org.apache.xmlbeans.*;
import org.apache.xmlbeans.impl.common.StaxHelper;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import javax.xml.stream.events.*;
import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.util.Deque;
import java.util.Iterator;
import java.util.LinkedList;

import static org.apache.poi.schemas.ooxml.system.ooxml.TypeSystemHolder.typeSystem;

@Slf4j
final class XmlElementReader implements AutoCloseable {

    static final XmlOptions DEFAULT_XML_OPTIONS;

    static {
        DEFAULT_XML_OPTIONS = new XmlOptions();
        DEFAULT_XML_OPTIONS.setLoadDTDGrammar(false);
        DEFAULT_XML_OPTIONS.setLoadExternalDTD(false);
    }

    private final SchemaType _elementType;
    private final XMLEventReader _reader;
    private final Deque<XmlObject> _deque;
    private SchemaType _context;
    private XmlObject _lastObject;
    private boolean _closed;

    XmlElementReader(final InputStream src) {
        final XMLInputFactory factory = StaxHelper.newXMLInputFactory(DEFAULT_XML_OPTIONS);
        _deque = new LinkedList<>();
        try {
            final InputStream source = src.markSupported() ? src : new BufferedInputStream(src);
            source.mark(Integer.MAX_VALUE);
            _elementType = _findElementType(source, factory);
            source.reset();
            _reader = factory.createXMLEventReader(source);
            _deque.addFirst(typeSystem.newInstance(_elementType, DEFAULT_XML_OPTIONS));
        } catch (XMLStreamException e) {
            throw new IllegalStateException(e);
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    private static SchemaType _findElementType(final InputStream src, final XMLInputFactory factory) throws XMLStreamException {
        final XMLStreamReader reader = factory.createXMLStreamReader(src);
        while (!reader.isStartElement()) reader.next();
        final QName name = reader.getName();
        typeSystem.findType(name);
        SchemaType type = typeSystem.findDocumentType(name);
        if (type == null) {
            type = typeSystem.findType(name);
        }
        if (type == null) {
            throw new IllegalArgumentException("Cannot find type for " + name);
        }
        return type;
    }

    public SchemaType getElementType() {
        return _elementType;
    }

    public void markContext() {
        _context = _deque.getFirst().schemaType();
    }

    @SuppressWarnings("unchecked")
    public <T extends XmlObject> T current() {
        return (T) _deque.getFirst().copy();
    }

    @SuppressWarnings("unchecked")
    public <T extends XmlObject> T last() {
        return (T) _lastObject.copy();
    }

    public <T extends XmlObject> T collect() {
        markContext();
        nextUntil(_context.isDocumentType() ? Matcher.endDocumentOf(_context) : Matcher.endElementOf(_context));
        return last();
    }

    public XMLEvent nextUntil(final Matcher... matchers) {
        while (_reader.hasNext()) {
            final XMLEvent event = next();
            for (final Matcher matcher : matchers) {
                final XmlObject object = event.isEndElement() || event.isEndDocument() ? last() : current();
                if (matcher.matches(event, object)) {
                    return event;
                }
            }
        }
        return null;
    }

    public XMLEvent next() {
        try {
            final XMLEvent event = _reader.nextEvent();
            if (event.isStartDocument()) {
                _handle((StartDocument) event);
            } else if (event.isStartElement()) {
                _handle(event.asStartElement());
            } else if (event.isCharacters()) {
                _handle(event.asCharacters());
            } else if (event.isEndElement()) {
                _handle(event.asEndElement());
            } else if (event.isEndDocument()) {
                _handle((EndDocument) event);
            }
            return event;
        } catch (XMLStreamException e) {
            throw new IllegalStateException(e);
        }
    }

    public boolean isClosed() {
        return _closed;
    }

    @Override
    public void close() {
        try {
            _reader.close();
            _deque.clear();
            _lastObject = null;
            _closed = true;
        } catch (XMLStreamException e) {
            throw new IllegalStateException(e);
        }
    }

    private void _append(final QName name, final XmlObject object) {
        final XmlObject parent = _deque.getFirst();
        try (final XmlCursor cursor = object.newCursor();
             final XmlCursor parentCursor = parent.newCursor()) {
            parentCursor.toEndToken();
            parentCursor.beginElement(name);
            while (cursor.hasNextToken()) {
                if (cursor.isStartdoc()) cursor.toNextToken();
                else cursor.moveXml(parentCursor);
            }
        }
    }

    private void _handle(final StartDocument event) throws XMLStreamException {
        if (log.isTraceEnabled()) {
            log.trace("Start document: {}", event);
        }
    }

    private void _handle(final StartElement event) throws XMLStreamException {
        if (log.isTraceEnabled()) {
            log.trace("Start element: {}", event);
        }
        final QName name = event.getName();
        final XmlObject parent = _deque.getFirst();
        final SchemaProperty property = parent.schemaType().getElementProperty(name);
        SchemaType type = null;
        try {
            type = property == null ? null : property.getType();
        } catch (SchemaTypeLoaderException e) {
            log.info(e.getMessage());
        }
        if (type == null) {
            XMLEvent next;
            do {
                next = _reader.nextEvent(); // skip unknowns
            } while (!next.isEndElement() || !next.asEndElement().getName().equals(name));
            return;
        }
        final XmlObject object = typeSystem.newInstance(type, DEFAULT_XML_OPTIONS);
        try (final XmlCursor cursor = object.newCursor()) {
            cursor.toNextToken();
            for (final Iterator<?> attrs = event.getAttributes(); attrs.hasNext(); ) {
                final Attribute attribute = (Attribute) attrs.next();
                final QName attributeName = attribute.getName();
                final SchemaProperty attributeProperty = type.getAttributeProperty(attributeName);
                if (attributeProperty == null) continue;
                cursor.insertAttributeWithValue(attributeName, attribute.getValue());
            }
        }
        _deque.addFirst(object);
    }

    private void _handle(final Characters event) throws XMLStreamException {
        if (log.isTraceEnabled()) {
            log.trace("Characters: {}", event);
        }
        final XmlObject current = _deque.getFirst();
        if (current instanceof XmlAnySimpleType) {
            final XmlAnySimpleType object = (XmlAnySimpleType) current;
            object.setStringValue(object.getStringValue() + event.getData());
        }
    }

    private void _handle(final EndElement event) throws XMLStreamException {
        if (log.isTraceEnabled()) {
            log.trace("End element: {}", event);
        }
        final QName name = event.getName();
        _lastObject = _deque.removeFirst();
        if (_context != null) {
            if (_lastObject.schemaType().equals(_context)) {
                _context = null;
            } else {
                _append(name, _lastObject);
            }
        }
    }

    private void _handle(final EndDocument event) throws XMLStreamException {
        if (log.isTraceEnabled()) {
            log.trace("End document: {}", event);
        }
        _lastObject = _deque.removeFirst();
    }

    interface Matcher {

        static Matcher startElementOf(final SchemaType schemaType) {
            return create(StartElement.class, schemaType);
        }

        static Matcher endElementOf(final SchemaType schemaType) {
            return create(EndElement.class, schemaType);
        }

        static Matcher endDocumentOf(final SchemaType schemaType) {
            return create(EndDocument.class, schemaType);
        }

        static Matcher create(final Class<? extends XMLEvent> eventType, final SchemaType schemaType) {
            return (event, object) -> eventType.isInstance(event) && object.schemaType().equals(schemaType);
        }

        boolean matches(XMLEvent event, XmlObject object);
    }
}
