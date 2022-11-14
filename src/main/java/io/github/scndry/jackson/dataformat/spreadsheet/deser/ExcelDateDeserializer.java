package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;
import lombok.RequiredArgsConstructor;

import java.io.IOException;
import java.util.function.BiFunction;

@RequiredArgsConstructor
public final class ExcelDateDeserializer<T> extends JsonDeserializer<T> {

    private final BiFunction<Double, Boolean, T> function;

    @Override
    public T deserialize(final JsonParser p, final DeserializationContext ctxt) throws IOException {
        final boolean date1904 = ((SheetParser) p).isDate1904();
        return function.apply(p.getDoubleValue(), date1904);
    }
}
