package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import java.io.File;
import java.io.OutputStream;

import lombok.EqualsAndHashCode;

import io.github.scndry.jackson.dataformat.spreadsheet.SheetContent;

/**
 * Immutable descriptor for a spreadsheet write target. Wraps a
 * {@link File} or {@link OutputStream} together with an optional
 * sheet name.
 * <p>
 * Use the {@code target} factory methods to create instances.
 *
 * @param <T> type of the underlying raw target
 * @see SheetContent
 * @see SheetGenerator
 */
@EqualsAndHashCode
public final class SheetOutput<T> implements SheetContent<T> {

    private final T _raw;
    private final String _name;

    private SheetOutput(final T raw, final String name) {
        _raw = raw;
        _name = name;
    }

    /**
     * Creates a {@code SheetOutput} writing to the given file
     * with an auto-generated sheet name.
     */
    public static SheetOutput<File> target(final File raw) {
        return new SheetOutput<>(raw, null);
    }

    /**
     * Creates a {@code SheetOutput} writing to the given file
     * with the specified sheet name.
     */
    public static SheetOutput<File> target(final File raw, final String sheetName) {
        return new SheetOutput<>(raw, sheetName);
    }

    /**
     * Creates a {@code SheetOutput} writing to the given stream
     * with an auto-generated sheet name.
     */
    public static SheetOutput<OutputStream> target(final OutputStream raw) {
        return new SheetOutput<>(raw, null);
    }

    /**
     * Creates a {@code SheetOutput} writing to the given stream
     * with the specified sheet name.
     */
    public static SheetOutput<OutputStream> target(final OutputStream raw, final String sheetName) {
        return new SheetOutput<>(raw, sheetName);
    }

    @Override
    public T getRaw() { return _raw; }

    @Override
    public String getName() { return _name; }

    @Override
    public String toString() {
        return "SheetOutput(raw=" + _raw + ", name=" + _name + ")";
    }

}
