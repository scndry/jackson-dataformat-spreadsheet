package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import java.io.IOException;
import java.util.function.BiFunction;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;

/**
 * Deserializer that converts Excel date serial numbers to Java date types using a configurable conversion function.
 */
public final class ExcelDateDeserializer<T> extends JsonDeserializer<T> {

    private final BiFunction<Double, Boolean, T> _function;

    public ExcelDateDeserializer(final BiFunction<Double, Boolean, T> function) {
        _function = function;
    }

    @Override
    public T deserialize(final JsonParser p, final DeserializationContext ctxt) throws IOException {
        final boolean date1904 = ((SheetParser) p).isDate1904();
        return _function.apply(p.getDoubleValue(), date1904);
    }
}
