package io.github.scndry.jackson.dataformat.spreadsheet.ser;

import java.io.File;
import java.io.OutputStream;
import java.nio.file.Path;

import com.fasterxml.jackson.annotation.JsonIgnore;

import lombok.EqualsAndHashCode;
import org.apache.poi.ss.util.WorkbookUtil;

import io.github.scndry.jackson.dataformat.spreadsheet.EncryptionSpec;
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
    private final EncryptionSpec _encryption;

    private SheetOutput(final T raw, final String name, final String password,
                        final EncryptionSpec encryption) {
        _raw = raw;
        _name = name;
        _password = password;
        _encryption = encryption;
    }

    private SheetOutput(final T raw, final String name) {
        this(raw, name, null, null);
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
     * Returns a copy with the given password for OOXML file-level encryption,
     * using {@link EncryptionSpec#strong()} (agile AES-256 + SHA-512). Pass
     * {@code null} to clear.
     *
     * <p>For {@link java.io.File} targets the encrypted output is written to a
     * sibling temp and atomically renamed onto the target on close, so
     * mid-write failures leave the original target untouched. For
     * {@link java.io.OutputStream} targets the bytes are written directly and
     * the stream is not closed by this library.
     */
    public SheetOutput<T> withPassword(final String password) {
        return new SheetOutput<>(_raw, _name, password, null);
    }

    /**
     * Returns a copy with the given password and {@link EncryptionSpec}. The
     * spec selects the strength / compatibility / cipher trade-off. A
     * {@code null} spec falls back to {@link EncryptionSpec#strong()}. A
     * {@code null} password clears both.
     */
    public SheetOutput<T> withPassword(final String password, final EncryptionSpec encryption) {
        return new SheetOutput<>(_raw, _name, password, password == null ? null : encryption);
    }

    @Override
    public T getRaw() { return _raw; }

    @Override
    public String getName() { return _name; }

    @JsonIgnore
    public String getPassword() { return _password; }

    @JsonIgnore
    public EncryptionSpec getEncryption() { return _encryption; }

    @Override
    public String toString() {
        return "SheetOutput(raw=" + _raw + ", name=" + _name
                + ", password=" + (_password == null ? "null" : "***")
                + ", encryption=" + (_encryption == null ? "default" : _encryption) + ")";
    }

}
