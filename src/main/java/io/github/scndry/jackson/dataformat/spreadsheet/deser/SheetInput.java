package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import io.github.scndry.jackson.dataformat.spreadsheet.SheetContent;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.ToString;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.util.TempFile;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

@Getter
@ToString
@EqualsAndHashCode
@SuppressWarnings({"unchecked", "java:S1452"})
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
        return (SheetInput<InputStream>) source(raw, false);
    }

    public static SheetInput<InputStream> source(final InputStream raw, final int sheetIndex) {
        return (SheetInput<InputStream>) source(raw, sheetIndex, false);
    }

    public static SheetInput<InputStream> source(final InputStream raw, final String sheetName) {
        return (SheetInput<InputStream>) source(raw, sheetName, false);
    }

    public static SheetInput<?> source(final InputStream raw, final boolean useTempFile) {
        return new SheetInput<>(useTempFile ? createTempFile(raw) : raw, 0);
    }

    public static SheetInput<?> source(final InputStream raw, final int sheetIndex, final boolean useTempFile) {
        return new SheetInput<>(useTempFile ? createTempFile(raw) : raw, sheetIndex);
    }

    public static SheetInput<?> source(final InputStream raw, final String sheetName, final boolean useTempFile) {
        return new SheetInput<>(useTempFile ? createTempFile(raw) : raw, sheetName);
    }

    private static File createTempFile(final InputStream raw) {
        File file = null;
        try {
            file = TempFile.createTempFile("sheet-input", ".xlsx");
            Files.copy(raw, file.toPath(), StandardCopyOption.REPLACE_EXISTING);
            return file;
        } catch (IOException e) {
            FileUtils.deleteQuietly(file);
            throw new UncheckedIOException(e);
        } finally {
            IOUtils.closeQuietly(raw);
        }
    }
}
