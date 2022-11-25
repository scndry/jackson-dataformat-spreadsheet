package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import com.fasterxml.jackson.core.*;
import com.fasterxml.jackson.core.base.GeneratorBase;
import com.fasterxml.jackson.core.io.IOContext;
import io.github.scndry.jackson.dataformat.spreadsheet.PackageVersion;
import io.github.scndry.jackson.dataformat.spreadsheet.SheetStreamContext;
import io.github.scndry.jackson.dataformat.spreadsheet.SheetStreamWriteException;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;

@Slf4j
public final class SheetGenerator extends GeneratorBase {

    private static final String START_ARRAY = "Start an array";
    private static final String END_ARRAY = "End an array";
    private static final String START_OBJECT = "Start an object";
    private static final String END_OBJECT = "End an object";
    private static final String WRITE_STRING = "Write a string";
    private static final String WRITE_NUMBER = "Write a number";
    private static final String WRITE_BOOLEAN = "Write a boolean";
    private static final String WRITE_NULL = "Write a null";

    private final IOContext _ioContext;
    private final SheetWriter _writer;
    private SpreadsheetSchema _schema;
    private SheetStreamContext _outputContext;

    public SheetGenerator(final IOContext ctxt, final int features, final ObjectCodec codec, final SheetWriter writer) {
        super(features, codec);
        _ioContext = ctxt;
        _writer = writer;
    }

    @Override
    public Version version() {
        return PackageVersion.VERSION;
    }

    @Override
    public SheetStreamContext getOutputContext() {
        return _outputContext;
    }

    @Override
    public SpreadsheetSchema getSchema() {
        return _schema;
    }

    @Override
    public void setSchema(final FormatSchema schema) {
        if (_schema != null) return;
        _schema = (SpreadsheetSchema) schema;
        _writer.setSchema(_schema);
        _writer.writeHeaders();
        _outputContext = SheetStreamContext.createRootContext(_schema);
    }

    @Override
    public boolean canUseSchema(final FormatSchema schema) {
        return schema instanceof SpreadsheetSchema;
    }

    public boolean isDate1904() {
        return _writer.isDate1904();
    }

    @Override
    public void writeStartArray() throws IOException {
        writeStartArray(null, -1);
    }

    @Override
    public void writeStartArray(final Object forValue, final int size) throws IOException {
        _verifyValueWrite(START_ARRAY);
        _outputContext = _outputContext.createChildArrayContext(size);
    }

    @Override
    public void writeEndArray() throws IOException {
        _outputContext = _closeStruct(END_ARRAY);
    }

    @Override
    public void writeStartObject() throws IOException {
        writeStartObject(null);
    }

    @Override
    public void writeStartObject(final Object forValue) throws IOException {
        _verifyValueWrite(START_OBJECT);
        _outputContext = _outputContext.createChildObjectContext();
    }

    @Override
    public void writeEndObject() throws IOException {
        // final int size = _outputContext.size()
        _outputContext = _closeStruct(END_OBJECT);
        // final ColumnPointer pointer = _outputContext.currentPointer()
        // TODO support merge column in scope to optional features via annotation
        // _writer.mergeScopedColumns(pointer, _outputContext.getRow(), size)
    }

    @Override
    public void writeFieldName(final String name) throws IOException {
        _outputContext.setCurrentName(name);
    }

    @Override
    public void writeString(final String text) throws IOException {
        _verifyValueWrite(WRITE_STRING);
        _writer.writeString(text);
    }

    @Override
    public void writeString(final char[] buffer, final int offset, final int len) throws IOException {
        _verifyValueWrite(WRITE_STRING);
        _writer.writeString(new String(buffer, offset, len));
    }

    @Override
    public void writeRawUTF8String(final byte[] buffer, final int offset, final int len) throws IOException {
        _reportUnsupportedOperation();
    }

    @Override
    public void writeUTF8String(final byte[] buffer, final int offset, final int len) throws IOException {
        _reportUnsupportedOperation();
    }

    @Override
    public void writeRaw(final String text) throws IOException {
        _reportUnsupportedOperation();
    }

    @Override
    public void writeRaw(final String text, final int offset, final int len) throws IOException {
        _reportUnsupportedOperation();
    }

    @Override
    public void writeRaw(final char[] text, final int offset, final int len) throws IOException {
        _reportUnsupportedOperation();
    }

    @Override
    public void writeRaw(final char c) throws IOException {
        _reportUnsupportedOperation();
    }

    @Override
    public void writeBinary(final Base64Variant bv, final byte[] data, final int offset, final int len) throws IOException {
        _reportUnsupportedOperation();
    }

    @Override
    public void writeNumber(final int v) throws IOException {
        _verifyValueWrite(WRITE_NUMBER);
        _writer.writeNumeric(v);
    }

    @Override
    public void writeNumber(final long v) throws IOException {
        _verifyValueWrite(WRITE_NUMBER);
        _writer.writeNumeric(v);
    }

    @Override
    public void writeNumber(final BigInteger v) throws IOException {
        _verifyValueWrite(WRITE_NUMBER);
        _writer.writeString(v.toString());
    }

    @Override
    public void writeNumber(final double v) throws IOException {
        _verifyValueWrite(WRITE_NUMBER);
        _writer.writeNumeric(v);
    }

    @Override
    public void writeNumber(final float v) throws IOException {
        _verifyValueWrite(WRITE_NUMBER);
        _writer.writeNumeric(v);
    }

    @Override
    public void writeNumber(final BigDecimal v) throws IOException {
        _verifyValueWrite(WRITE_NUMBER);
        _writer.writeString(isEnabled(StreamWriteFeature.WRITE_BIGDECIMAL_AS_PLAIN) ? v.toPlainString() : v.toString());
    }

    @Override
    public void writeNumber(final String encodedValue) throws IOException {
        _verifyValueWrite(WRITE_NUMBER);
        _writer.writeString(encodedValue);
    }

    @Override
    public void writeBoolean(final boolean state) throws IOException {
        _verifyValueWrite(WRITE_BOOLEAN);
        _writer.writeBoolean(state);
    }

    @Override
    public void writeNull() throws IOException {
        _verifyValueWrite(WRITE_NULL);
        _writer.writeBlank();
    }

    @Override
    public void flush() throws IOException {
        // do nothing
    }

    @Override
    public void close() throws IOException {
        super.close();
        _writer.adjustColumnWidth();
        final Object content = _ioContext.contentReference().getRawContent();
        if (content instanceof SheetOutput) {
            final SheetOutput<?> output = (SheetOutput<?>) content;
            _writer.write(output.getOutputStream());
        }
        if (_ioContext.isResourceManaged() || isEnabled(StreamWriteFeature.AUTO_CLOSE_TARGET)) {
            _writer.close();
        }
    }

    @Override
    protected void _releaseBuffers() {
        _reportUnsupportedOperation();
    }

    @Override
    protected void _verifyValueWrite(final String typeMsg) throws IOException {
        _checkSchemaSet();
        _outputContext.writeValue();
        _writer.setReference(_outputContext.currentReference());
        if (log.isTraceEnabled()) {
            log.trace("{} {} {}", typeMsg, _outputContext.currentReference(), _outputContext.pathAsPointer(true));
        }
    }

    private SheetStreamContext _closeStruct(final String typeMsg) {
        final SheetStreamContext parent = _outputContext.clearAndGetParent();
        if (log.isTraceEnabled()) {
            log.trace("{} {} {}", typeMsg, parent.currentReference(), parent.pathAsPointer(true));
        }
        return parent;
    }

    private void _checkSchemaSet() throws IOException {
        if (_schema == null) {
            throw new SheetStreamWriteException("No schema of type '" + SpreadsheetSchema.SCHEMA_TYPE + "' set, can not generate", this);
        }
    }
}
