package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.exc.StreamReadException;
import com.fasterxml.jackson.core.util.RequestPayload;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;

@SuppressWarnings("java:S110")
public final class SheetStreamReadException extends StreamReadException {

    public SheetStreamReadException(final JsonParser p, final String msg) {
        super(p, msg);
    }

    public static SheetStreamReadException unexpected(final SheetParser p, final Object value) {
        return new SheetStreamReadException(p, "Unexpected value: " + value);
    }

    @Override
    public SheetStreamReadException withParser(final JsonParser p) {
        _processor = p;
        return this;
    }

    @Override
    public SheetStreamReadException withRequestPayload(final RequestPayload payload) {
        _requestPayload = payload;
        return this;
    }
}
