package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import java.io.File;
import java.io.InputStream;

import lombok.EqualsAndHashCode;

import io.github.scndry.jackson.dataformat.spreadsheet.SheetContent;

/**
 * Immutable descriptor for a spreadsheet read source. Wraps a
 * {@link File} or {@link InputStream} together with an optional
 * sheet selector (by name or by zero-based index).
 * <p>
 * Use the {@code source} factory methods to create instances.
 *
 * @param <T> type of the underlying raw source
 * @see SheetContent
 * @see SheetParser
 */
@EqualsAndHashCode
public final class SheetInput<T> implements SheetContent<T> {

    private final T _raw;
    private final String _name;
    private final int _index;

    private SheetInput(final T raw, final int index) {
        _raw = raw;
        _name = null;
        _index = index;
    }

    private SheetInput(final T raw, final String name) {
        _raw = raw;
        _name = name;
        _index = -1;
    }

    /**
     * Creates a {@code SheetInput} reading from the given file,
     * targeting the first sheet (index 0).
     */
    public static SheetInput<File> source(final File raw) {
        return new SheetInput<>(raw, 0);
    }

    /**
     * Creates a {@code SheetInput} reading from the given file,
     * targeting the sheet at the specified index.
     */
    public static SheetInput<File> source(final File raw, final int sheetIndex) {
        return new SheetInput<>(raw, sheetIndex);
    }

    /**
     * Creates a {@code SheetInput} reading from the given file,
     * targeting the sheet with the specified name.
     */
    public static SheetInput<File> source(final File raw, final String sheetName) {
        return new SheetInput<>(raw, sheetName);
    }

    /**
     * Creates a {@code SheetInput} reading from the given
     * stream, targeting the first sheet (index 0).
     */
    public static SheetInput<InputStream> source(final InputStream raw) {
        return new SheetInput<>(raw, 0);
    }

    /**
     * Creates a {@code SheetInput} reading from the given
     * stream, targeting the sheet at the specified index.
     */
    public static SheetInput<InputStream> source(final InputStream raw, final int sheetIndex) {
        return new SheetInput<>(raw, sheetIndex);
    }

    /**
     * Creates a {@code SheetInput} reading from the given
     * stream, targeting the sheet with the specified name.
     */
    public static SheetInput<InputStream> source(final InputStream raw, final String sheetName) {
        return new SheetInput<>(raw, sheetName);
    }

    @Override
    public T getRaw() { return _raw; }

    @Override
    public String getName() { return _name; }

    public int getIndex() { return _index; }

    @Override
    public String toString() {
        return "SheetInput(raw=" + _raw + ", name=" + _name + ", index=" + _index + ")";
    }

}
