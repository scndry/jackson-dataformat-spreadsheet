package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import io.github.scndry.jackson.dataformat.spreadsheet.SheetContent;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.ToString;

import java.io.File;
import java.io.OutputStream;

@Getter
@ToString
@EqualsAndHashCode
public final class SheetOutput<T> implements SheetContent<T> {

    private final T raw;
    private final String name;

    private SheetOutput(final T raw, final String name) {
        this.raw = raw;
        this.name = name;
    }

    public static SheetOutput<File> target(final File raw) {
        return new SheetOutput<>(raw, null);
    }

    public static SheetOutput<File> target(final File raw, final String sheetName) {
        return new SheetOutput<>(raw, sheetName);
    }

    public static SheetOutput<OutputStream> target(final OutputStream raw) {
        return new SheetOutput<>(raw, null);
    }

    public static SheetOutput<OutputStream> target(final OutputStream raw, final String sheetName) {
        return new SheetOutput<>(raw, sheetName);
    }
}
