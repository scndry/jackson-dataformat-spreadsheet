package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import com.fasterxml.jackson.core.*;
import com.fasterxml.jackson.core.base.ParserMinimalBase;
import com.fasterxml.jackson.core.exc.StreamReadException;
import com.fasterxml.jackson.core.io.ContentReference;
import com.fasterxml.jackson.core.io.IOContext;
import io.github.scndry.jackson.dataformat.spreadsheet.PackageVersion;
import io.github.scndry.jackson.dataformat.spreadsheet.SheetStreamContext;
import io.github.scndry.jackson.dataformat.spreadsheet.SheetStreamReadException;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellAddress;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.file.Files;
import java.util.Deque;
import java.util.Iterator;
import java.util.LinkedList;

@Slf4j
public final class SheetParser extends ParserMinimalBase {

    private final IOContext _ioContext;
    private final SheetReader _reader;
    private final Deque<JsonToken> _nextTokens;
    private final int _formatFeatures;
    private boolean _closed;
    private ObjectCodec _objectCodec;
    private SpreadsheetSchema _schema;
    private SheetStreamContext _parsingContext;
    private CellAddress _reference;
    private CellValue _value;

    public SheetParser(final IOContext ctxt, final int features, final ObjectCodec codec, final int formatFeatures, final SheetReader reader) {
        super(features);
        _ioContext = ctxt;
        _objectCodec = codec;
        _formatFeatures = formatFeatures;
        _reader = reader;
        _nextTokens = new LinkedList<>();
        _reference = new CellAddress(-1, -1);
    }

    public boolean isEnabled(final Feature f) {
        return (_formatFeatures & f.getMask()) != 0;
    }

    @Override
    public int getFormatFeatures() {
        return _formatFeatures;
    }

    @Override
    public Version version() {
        return PackageVersion.VERSION;
    }

    @Override
    public SpreadsheetSchema getSchema() {
        return _schema;
    }

    @Override
    public void setSchema(final FormatSchema schema) {
        _schema = (SpreadsheetSchema) schema;
        _parsingContext = SheetStreamContext.createRootContext(_schema);
    }

    @Override
    public boolean canUseSchema(final FormatSchema schema) {
        return schema instanceof SpreadsheetSchema;
    }

    public boolean isDate1904() {
        return _reader.isDate1904();
    }

    @Override
    public JsonToken nextToken() throws IOException {
        _checkSchemaSet();
        _prepareDeterministicNext();
        final JsonToken token = _nextTokens.removeFirst();
        if (log.isTraceEnabled()) {
            log.trace("{}", token);
        }
        if (token == null) {
            _currToken = null;
            return null;
        }
        switch (token) {
            case START_ARRAY:
                _parsingContext = _parsingContext.createChildArrayContext();
                break;
            case START_OBJECT:
                _parsingContext = _parsingContext.createChildObjectContext();
                break;
            case END_ARRAY:
            case END_OBJECT:
                _parsingContext = _parsingContext.clearAndGetParent();
                break;
            case FIELD_NAME:
                final Column column = _schema.getColumn(_reference);
                final ColumnPointer pointer = _parsingContext.relativePointer(column.getPointer());
                _parsingContext.setCurrentName(pointer.head().name());
                break;
            case VALUE_EMBEDDED_OBJECT:
            case VALUE_STRING:
            case VALUE_NUMBER_INT:
            case VALUE_NUMBER_FLOAT:
            case VALUE_TRUE:
            case VALUE_FALSE:
            case VALUE_NULL:
                break;
            case NOT_AVAILABLE:
                throw SheetStreamReadException.unexpected(this, token);
        }
        _currToken = token;
        return _currToken;
    }

    private void _prepareDeterministicNext() throws StreamReadException {
        while (_nextTokens.isEmpty() || _isStartObject()) {
            _prepareNext();
            _handleEmptyObject();
        }
    }

    private boolean _isStartObject() {
        return _nextTokens.size() == 1 && _nextTokens.peekFirst() == JsonToken.START_OBJECT;
    }

    private void _prepareNext() throws StreamReadException {
        final SheetToken token = _readNext();
        if (token == null) {
            _nextTokens.add(null);
            return;
        }
        switch (token) {
            case SHEET_DATA_START:
                _nextTokens.add(JsonToken.START_ARRAY);
                break;
            case ROW_START:
                _reference = new CellAddress(_reader.getRow(), -1);
                _nextTokens.add(JsonToken.START_OBJECT);
                break;
            case CELL_VALUE:
                _reference = _reader.getReference();
                _value = _reader.getCellValue();
                final Column column = _schema.getColumn(_reference);
                final ColumnPointer pointer = _parsingContext.relativePointer(column.getPointer().getParent());
                for (final ColumnPointer p : pointer) {
                    if (p.isParent()) {
                        _nextTokens.add(JsonToken.END_OBJECT);
                    } else {
                        _nextTokens.add(JsonToken.FIELD_NAME);
                        _nextTokens.add(JsonToken.START_OBJECT);
                    }
                }
                _nextTokens.add(JsonToken.FIELD_NAME);
                _nextTokens.add(_scalarValueToken());
                break;
            case ROW_END:
                final int depth = _parsingContext.currentPointer().depth();
                for (int i = 0; i < depth; i++) {
                    _nextTokens.add(JsonToken.END_OBJECT);
                }
                break;
            case SHEET_DATA_END:
                _nextTokens.add(JsonToken.END_ARRAY);
                break;
        }
    }

    private SheetToken _readNext() {
        if (!_reader.hasNext()) {
            return null;
        }
        SheetToken token = _reader.next();
        if (token == SheetToken.ROW_START) {
            while (!_schema.isInRowBounds(_reader.getRow())) {
                token = _reader.next();
            }
        }
        if (token == SheetToken.CELL_VALUE) {
            while (!_schema.isInColumnBounds(_reader.getColumn())) {
                token = _reader.next();
                if (token != SheetToken.CELL_VALUE) break;
            }
        }
        return token;
    }

    private JsonToken _scalarValueToken() throws StreamReadException {
        final CellType type = _value.getCellType();
        switch (type) {
            case NUMERIC:
                return _value.noFractionalPart() ? JsonToken.VALUE_NUMBER_INT : JsonToken.VALUE_NUMBER_FLOAT;
            case STRING:
                return JsonToken.VALUE_STRING;
            case BLANK:
                return JsonToken.VALUE_NULL;
            case BOOLEAN:
                return _value.getBooleanValue() ? JsonToken.VALUE_TRUE : JsonToken.VALUE_FALSE;
            case _NONE:
            case FORMULA:
            case ERROR:
        }
        throw SheetStreamReadException.unexpected(this, type);
    }

    private void _handleEmptyObject() {
        if (_nextTokens.size() != 2) return;
        final Iterator<JsonToken> iterator = _nextTokens.iterator();
        if (iterator.next() == JsonToken.START_OBJECT && iterator.next() == JsonToken.END_OBJECT) {
            _nextTokens.clear();
            if (isEnabled(Feature.BREAK_ON_BLANK_ROW)) {
                _nextTokens.add(null);
            } else if (isEnabled(Feature.BLANK_ROW_AS_NULL)) {
                _nextTokens.add(JsonToken.VALUE_NULL);
            }
        }
    }

    @Override
    protected void _handleEOF() throws JsonParseException {
        throw new UnsupportedOperationException();
    }

    @Override
    public String getCurrentName() throws IOException {
        return _parsingContext.getCurrentName();
    }

    @Override
    public ObjectCodec getCodec() {
        return _objectCodec;
    }

    @Override
    public void setCodec(final ObjectCodec oc) {
        _objectCodec = oc;
    }

    @Override
    public void close() throws IOException {
        _closed = true;
        if (_ioContext.isResourceManaged() || isEnabled(StreamReadFeature.AUTO_CLOSE_SOURCE)) {
            _reader.close();
        }
        final Object content = _ioContext.contentReference().getRawContent();
        if (content instanceof SheetInput) {
            final SheetInput<?> input = (SheetInput<?>) content;
            if (_ioContext.isResourceManaged() && input.isFile()) {
                Files.deleteIfExists(((File) input.getRaw()).toPath());
            }
        }
    }

    @Override
    public boolean isClosed() {
        return _closed;
    }

    @Override
    public SheetStreamContext getParsingContext() {
        return _parsingContext;
    }

    @Override
    public JsonLocation getCurrentLocation() {
        return new SheetLocation(_contentReference(), _reference.getRow(), _reference.getColumn());
    }

    @Override
    public JsonLocation getTokenLocation() {
        return new SheetLocation(_contentReference(), _reader.getRow(), _reader.getColumn());
    }

    @Override
    public void overrideCurrentName(final String name) {
        throw new UnsupportedOperationException();
    }

    @Override
    public String getText() throws IOException {
        return _value.getStringValue();
    }

    @Override
    public char[] getTextCharacters() throws IOException {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean hasTextCharacters() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Number getNumberValue() throws IOException {
        return _value.getNumberValue();
    }

    @Override
    public NumberType getNumberType() throws IOException {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getIntValue() throws IOException {
        return (int) _value.getNumberValue();
    }

    @Override
    public long getLongValue() throws IOException {
        return (long) _value.getNumberValue();
    }

    @Override
    public BigInteger getBigIntegerValue() throws IOException {
        return new BigInteger(_value.getStringValue());
    }

    @Override
    public float getFloatValue() throws IOException {
        return (float) _value.getNumberValue();
    }

    @Override
    public double getDoubleValue() throws IOException {
        return _value.getNumberValue();
    }

    @Override
    public BigDecimal getDecimalValue() throws IOException {
        return new BigDecimal(_value.getStringValue());
    }

    @Override
    public int getTextLength() throws IOException {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getTextOffset() throws IOException {
        throw new UnsupportedOperationException();
    }

    @Override
    public byte[] getBinaryValue(final Base64Variant b64variant) throws IOException {
        throw new UnsupportedOperationException();
    }

    private ContentReference _contentReference() {
        if (StreamReadFeature.INCLUDE_SOURCE_IN_LOCATION.enabledIn(_features)) {
            return _ioContext.contentReference();
        }
        return ContentReference.unknown();
    }

    private void _checkSchemaSet() throws IOException {
        if (_schema == null) {
            throw new SheetStreamReadException(this, "No schema of type '" + SpreadsheetSchema.SCHEMA_TYPE + "' set, can not parse");
        }
    }

    public enum Feature implements FormatFeature {
        BLANK_ROW_AS_NULL(true),
        BREAK_ON_BLANK_ROW(false),
        ;
        final boolean _defaultState;
        final int _mask;

        Feature(final boolean defaultState) {
            _defaultState = defaultState;
            _mask = 1 << ordinal();
        }

        public static int collectDefaults() {
            int flags = 0;
            for (Feature f : values()) {
                if (f.enabledByDefault()) {
                    flags |= f.getMask();
                }
            }
            return flags;
        }

        @Override
        public boolean enabledByDefault() {
            return _defaultState;
        }

        @Override
        public int getMask() {
            return _mask;
        }

        @Override
        public boolean enabledIn(final int flags) {
            return (flags & getMask()) != 0;
        }
    }
}
