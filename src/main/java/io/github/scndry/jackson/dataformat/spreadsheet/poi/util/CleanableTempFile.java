package io.github.scndry.jackson.dataformat.spreadsheet.poi.util;

import org.apache.poi.util.DefaultTempFileCreationStrategy;
import org.apache.poi.util.TempFileCreationStrategy;

import java.io.File;
import java.io.IOException;
import java.util.Objects;

public final class CleanableTempFile {

    private static TempFileCreationStrategy strategy = new DefaultTempFileCreationStrategy();
    private static TempFileCleaner cleaner = new DefaultTempFileCleaner();

    private CleanableTempFile() {
    }

    public static void setTempFileCreationStrategy(final TempFileCreationStrategy strategy) {
        Objects.requireNonNull(strategy, "'strategy' must not be null.");
        CleanableTempFile.strategy = strategy;
    }

    public static void setTempFileCleaner(final TempFileCleaner cleaner) {
        Objects.requireNonNull(strategy, "'cleaner' must not be null.");
        CleanableTempFile.cleaner = cleaner;
    }

    public static File createTempFile(final String prefix, final String suffix) throws IOException {
        final File newFile = strategy.createTempFile(prefix, suffix);
        newFile.deleteOnExit();
        cleaner.register(newFile);
        return newFile;
    }
}
