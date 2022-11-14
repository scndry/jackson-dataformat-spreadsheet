package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.JsonSerializer;
import com.fasterxml.jackson.databind.SerializerProvider;
import lombok.RequiredArgsConstructor;

import java.io.IOException;
import java.util.function.BiFunction;

@RequiredArgsConstructor
public final class ExcelDateSerializer<T> extends JsonSerializer<T> {

    private final BiFunction<T, Boolean, Double> function;

    @Override
    public void serialize(final T value, JsonGenerator gen, final SerializerProvider serializers) throws IOException {
        final boolean date1904 = ((SheetGenerator) gen).isDate1904();
        gen.writeNumber(function.apply(value, date1904));
    }
}
