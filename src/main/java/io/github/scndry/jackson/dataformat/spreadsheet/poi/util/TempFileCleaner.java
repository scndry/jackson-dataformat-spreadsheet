package io.github.scndry.jackson.dataformat.spreadsheet.poi.util;

import java.io.File;

public interface TempFileCleaner {

    default void register(final File referent) {
        register(referent.getPath(), referent);
    }

    void register(String pathname, Object referent);

    void shutdown();

    boolean isShutdown();
}
