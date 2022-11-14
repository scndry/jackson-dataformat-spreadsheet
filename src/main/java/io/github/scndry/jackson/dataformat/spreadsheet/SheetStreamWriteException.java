package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.exc.StreamWriteException;

@SuppressWarnings("java:S110")
public final class SheetStreamWriteException extends StreamWriteException {

    public SheetStreamWriteException(final String msg, final JsonGenerator g) {
        super(msg, g);
    }

    @Override
    public StreamWriteException withGenerator(final JsonGenerator g) {
        _processor = g;
        return this;
    }
}
