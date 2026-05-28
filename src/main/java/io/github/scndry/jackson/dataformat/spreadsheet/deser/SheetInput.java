package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Path;

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
    private final String _password;

    private SheetInput(final T raw, final String name, final int index, final String password) {
        _raw = raw;
        _name = name;
        _index = index;
        _password = password;
    }

    private SheetInput(final T raw, final int index) {
        this(raw, null, index, null);
    }

    private SheetInput(final T raw, final String name) {
        this(raw, name, -1, null);
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
     * Creates a {@code SheetInput} reading from the given path,
     * targeting the first sheet (index 0).
     */
    public static SheetInput<File> source(final Path raw) {
        return source(raw.toFile());
    }

    /**
     * Creates a {@code SheetInput} reading from the given path,
     * targeting the sheet at the specified index.
     */
    public static SheetInput<File> source(final Path raw, final int sheetIndex) {
        return source(raw.toFile(), sheetIndex);
    }

    /**
     * Creates a {@code SheetInput} reading from the given path,
     * targeting the sheet with the specified name.
     */
    public static SheetInput<File> source(final Path raw, final String sheetName) {
        return source(raw.toFile(), sheetName);
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

    /**
     * Returns a copy of this {@code SheetInput} with the given password for
     * OOXML file-level (agile) decryption. Pass {@code null} to clear.
     *
     * <p>The decrypted plaintext is materialised in a POSIX owner-only temp
     * file that is deleted when the resulting {@link SheetParser} closes;
     * always close the parser (try-with-resources or {@code SheetMappingIterator})
     * to release it promptly.
     *
     * <p>For {@code InputStream} sources, the caller still owns the stream;
     * this library does not close it. Password verification reads the source
     * before mapping, so a wrong password fails fast with
     * {@link org.apache.poi.EncryptedDocumentException}.
     */
    public SheetInput<T> withPassword(final String password) {
        return new SheetInput<>(_raw, _name, _index, password);
    }

    @Override
    public T getRaw() { return _raw; }

    @Override
    public String getName() { return _name; }

    public int getIndex() { return _index; }

    public String getPassword() { return _password; }

    @Override
    public String toString() {
        return "SheetInput(raw=" + _raw + ", name=" + _name + ", index=" + _index
                + ", password=" + (_password == null ? "null" : "***") + ")";
    }

}
