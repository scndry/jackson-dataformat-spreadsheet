package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import java.io.IOException;
import java.util.function.BiFunction;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.JsonSerializer;
import com.fasterxml.jackson.databind.SerializerProvider;

/**
 * Serializer that converts Java date types to Excel date serial numbers using a configurable conversion function.
 */
public final class ExcelDateSerializer<T> extends JsonSerializer<T> {

    private final BiFunction<T, Boolean, Double> _function;

    public ExcelDateSerializer(final BiFunction<T, Boolean, Double> function) {
        _function = function;
    }

    @Override
    public void serialize(
            final T value,
            JsonGenerator gen,
            final SerializerProvider serializers) throws IOException {
        final boolean date1904 = ((SheetGenerator) gen).isDate1904();
        gen.writeNumber(_function.apply(value, date1904));
    }
}
