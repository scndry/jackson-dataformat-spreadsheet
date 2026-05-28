package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import java.io.File;
import java.io.OutputStream;
import java.nio.file.Path;

import lombok.EqualsAndHashCode;
import org.apache.poi.ss.util.WorkbookUtil;

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
    private final String _password;

    private SheetOutput(final T raw, final String name, final String password) {
        _raw = raw;
        _name = name;
        _password = password;
    }

    private SheetOutput(final T raw, final String name) {
        this(raw, name, null);
    }

    private static void _validateSheetName(final String name) {
        if (name != null) WorkbookUtil.validateSheetName(name);
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
     *
     * @throws IllegalArgumentException if {@code sheetName} violates Excel constraints
     *         (1-31 chars, no {@code / \ ? * ] [}, no leading/trailing apostrophe).
     */
    public static SheetOutput<File> target(final File raw, final String sheetName) {
        _validateSheetName(sheetName);
        return new SheetOutput<>(raw, sheetName);
    }

    /**
     * Creates a {@code SheetOutput} writing to the given path
     * with an auto-generated sheet name.
     */
    public static SheetOutput<File> target(final Path raw) {
        return target(raw.toFile());
    }

    /**
     * Creates a {@code SheetOutput} writing to the given path
     * with the specified sheet name.
     *
     * @throws IllegalArgumentException if {@code sheetName} violates Excel constraints
     *         (1-31 chars, no {@code / \ ? * ] [}, no leading/trailing apostrophe).
     */
    public static SheetOutput<File> target(final Path raw, final String sheetName) {
        return target(raw.toFile(), sheetName);
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
     *
     * @throws IllegalArgumentException if {@code sheetName} violates Excel constraints
     *         (1-31 chars, no {@code / \ ? * ] [}, no leading/trailing apostrophe).
     */
    public static SheetOutput<OutputStream> target(final OutputStream raw, final String sheetName) {
        _validateSheetName(sheetName);
        return new SheetOutput<>(raw, sheetName);
    }

    /**
     * Returns a copy of this {@code SheetOutput} with the given password for
     * OOXML file-level encryption. Pass {@code null} to clear.
     *
     * <p>Encryption uses agile mode with AES-256 + SHA-512 (PBKDF2, 100K spin).
     * The writer streams plain OOXML into an {@code EncryptedTempData}
     * (AES-128-CBC, in-memory random key) so the intermediate bytes on disk
     * cannot be decrypted even if the temp survives a crash. The encrypted
     * output lands on the original target only after the generator closes;
     * if the write fails mid-flight, the partial target file is deleted.
     */
    public SheetOutput<T> withPassword(final String password) {
        return new SheetOutput<>(_raw, _name, password);
    }

    @Override
    public T getRaw() { return _raw; }

    @Override
    public String getName() { return _name; }

    public String getPassword() { return _password; }

    @Override
    public String toString() {
        return "SheetOutput(raw=" + _raw + ", name=" + _name
                + ", password=" + (_password == null ? "null" : "***") + ")";
    }

}
