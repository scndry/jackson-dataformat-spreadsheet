package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JavaType;
import com.fasterxml.jackson.databind.JsonDeserializer;
import com.fasterxml.jackson.databind.MappingIterator;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetLocation;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

public final class SheetMappingIterator<T> extends MappingIterator<T> {

    SheetMappingIterator(final JavaType type, final JsonParser p, final DeserializationContext ctxt,
                         final JsonDeserializer<?> deser, final boolean managedParser, final Object valueToUpdate) {
        super(type, p, ctxt, deser, managedParser, valueToUpdate);
    }

    @Override
    public SheetParser getParser() {
        return (SheetParser) super.getParser();
    }

    @Override
    public SpreadsheetSchema getParserSchema() {
        return (SpreadsheetSchema) super.getParserSchema();
    }

    @Override
    public SheetLocation getCurrentLocation() {
        return (SheetLocation) super.getCurrentLocation();
    }
}
