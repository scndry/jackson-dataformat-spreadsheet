package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.FormatSchema;
import com.fasterxml.jackson.core.JsonFactory;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.Version;
import com.fasterxml.jackson.databind.*;
import com.fasterxml.jackson.databind.deser.DataFormatReaders;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public final class SpreadsheetReader extends ObjectReader {

    SpreadsheetReader(final SpreadsheetMapper mapper, final DeserializationConfig config) {
        super(mapper, config);
    }

    SpreadsheetReader(final SpreadsheetMapper mapper, final DeserializationConfig config, final JavaType valueType,
                      final Object valueToUpdate, final FormatSchema schema,
                      final InjectableValues injectableValues) {
        super(mapper, config, valueType, valueToUpdate, schema, injectableValues);
    }

    @SuppressWarnings("java:S107")
    private SpreadsheetReader(final SpreadsheetReader base, final DeserializationConfig config, final JavaType valueType,
                              final JsonDeserializer<Object> rootDeser, final Object valueToUpdate,
                              final FormatSchema schema, final InjectableValues injectableValues,
                              final DataFormatReaders dataFormatReaders) {
        super(base, config, valueType, rootDeser, valueToUpdate, schema, injectableValues, dataFormatReaders);
    }

    private SpreadsheetReader(final SpreadsheetReader base, final DeserializationConfig config) {
        super(base, config);
    }

    private SpreadsheetReader(final SpreadsheetReader base, final SpreadsheetFactory f) {
        super(base, f);
    }

    @Override
    public Version version() {
        return PackageVersion.VERSION;
    }

    /*
    /**********************************************************
    /* Helper methods used internally for invoking constructors
    /**********************************************************
     */

    @Override
    protected SpreadsheetReader _new(final ObjectReader base, final JsonFactory f) {
        return new SpreadsheetReader((SpreadsheetReader) base, (SpreadsheetFactory) f);
    }

    @Override
    protected SpreadsheetReader _new(final ObjectReader base, final DeserializationConfig config) {
        return new SpreadsheetReader((SpreadsheetReader) base, config);
    }

    @Override
    protected SpreadsheetReader _new(final ObjectReader base, final DeserializationConfig config, final JavaType valueType,
                                     final JsonDeserializer<Object> rootDeser, final Object valueToUpdate,
                                     final FormatSchema schema, final InjectableValues injectableValues,
                                     final DataFormatReaders dataFormatReaders) {
        return new SpreadsheetReader((SpreadsheetReader) base, config, valueType,
                rootDeser, valueToUpdate, schema, injectableValues, dataFormatReaders);
    }

    @Override
    protected <T> SheetMappingIterator<T> _newIterator(final JsonParser p, final DeserializationContext ctxt,
                                                       final JsonDeserializer<?> deser, final boolean parserManaged) {
        return new SheetMappingIterator<>(_valueType, p, ctxt, deser, parserManaged, _valueToUpdate);
    }

    /*
    /**********************************************************
    /* Life-cycle, fluent factory methods, other
    /**********************************************************
     */

    @Override
    public SpreadsheetReader forType(final Class<?> valueType) {
        return (SpreadsheetReader) super.forType(valueType);
    }

    /*
    /**********************************************************
    /* Factory methods for creating SheetParsers
    /**********************************************************
     */

    public SheetParser createParser(final Sheet src) {
        return (SheetParser) _config.initialize(parserFactory().createParser(src), _schema);
    }

    public SheetParser createParser(final SheetInput<?> src) throws IOException {
        return (SheetParser) _config.initialize(parserFactory().createParser(src), _schema);
    }

    /*
    /**********************************************************
    /* Deserialization methods; others
    /**********************************************************
     */

    public <T> T readValue(final Sheet src) throws IOException {
        try (MappingIterator<T> iterator = readValues(src)) {
            return iterator.hasNext() ? iterator.next() : null;
        }
    }

    public <T> T readValue(final SheetInput<?> src) throws IOException {
        try (MappingIterator<T> iterator = readValues(src)) {
            return iterator.hasNext() ? iterator.next() : null;
        }
    }

    @Override
    public <T> T readValue(final File src) throws IOException {
        try (MappingIterator<T> iterator = readValues(src)) {
            return iterator.hasNext() ? iterator.next() : null;
        }
    }

    @Override
    public <T> T readValue(final InputStream src) throws IOException {
        try (MappingIterator<T> iterator = readValues(src)) {
            return iterator.hasNext() ? iterator.next() : null;
        }
    }

    /*
    /**********************************************************
    /* Deserialization methods; reading sequence of values
    /**********************************************************
     */

    @SuppressWarnings({"unchecked", "RedundantSuppression"})
    public <T> SheetMappingIterator<T> readValues(final Sheet src) throws IOException {
        return (SheetMappingIterator<T>) _bindAndReadValues(_considerFilter(createParser(src), true));
    }

    @SuppressWarnings({"unchecked", "RedundantSuppression"})
    public <T> SheetMappingIterator<T> readValues(final SheetInput<?> src) throws IOException {
        return (SheetMappingIterator<T>) _bindAndReadValues(_considerFilter(createParser(src), true));
    }

    @Override
    @SuppressWarnings({"unchecked", "RedundantSuppression"})
    public <T> SheetMappingIterator<T> readValues(final File src) throws IOException {
        return (SheetMappingIterator<T>) super.readValues(src);
    }

    @Override
    @SuppressWarnings({"unchecked", "RedundantSuppression"})
    public <T> SheetMappingIterator<T> readValues(final InputStream src) throws IOException {
        return (SheetMappingIterator<T>) super.readValues(src);
    }

    /*
    /**********************************************************
    /* Other public methods
    /**********************************************************
     */

    public SpreadsheetFactory parserFactory() {
        return (SpreadsheetFactory) _parserFactory;
    }
}
