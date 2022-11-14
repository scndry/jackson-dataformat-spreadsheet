package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import io.github.scndry.jackson.dataformat.spreadsheet.SheetContent;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.ToString;

import java.io.File;
import java.io.InputStream;

@Getter
@ToString
@EqualsAndHashCode
public final class SheetInput<T> implements SheetContent<T> {

    private final T raw;
    private final String name;
    private final int index;

    private SheetInput(final T raw, final int index) {
        this.raw = raw;
        this.name = null;
        this.index = index;
    }

    private SheetInput(final T raw, final String name) {
        this.raw = raw;
        this.name = name;
        this.index = -1;
    }

    public static SheetInput<File> source(final File raw) {
        return new SheetInput<>(raw, 0);
    }

    public static SheetInput<File> source(final File raw, final int sheetIndex) {
        return new SheetInput<>(raw, sheetIndex);
    }

    public static SheetInput<File> source(final File raw, final String sheetName) {
        return new SheetInput<>(raw, sheetName);
    }

    public static SheetInput<InputStream> source(final InputStream raw) {
        return new SheetInput<>(raw, 0);
    }

    public static SheetInput<InputStream> source(final InputStream raw, final int sheetIndex) {
        return new SheetInput<>(raw, sheetIndex);
    }

    public static SheetInput<InputStream> source(final InputStream raw, final String sheetName) {
        return new SheetInput<>(raw, sheetName);
    }
}
