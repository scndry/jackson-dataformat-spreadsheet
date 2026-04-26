package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;

/**
 * Common interface for spreadsheet read/write source descriptors.
 * Provides access to the underlying raw source and an optional
 * sheet name selector.
 *
 * @param <T> type of the underlying raw source
 * @see io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput
 * @see io.github.scndry.jackson.dataformat.spreadsheet.ser.SheetOutput
 */
public interface SheetContent<T> {

    T getRaw();

    String getName();

    default boolean isNamed() {
        return getName() != null;
    }

    default boolean isFile() {
        return getRaw() instanceof File;
    }
}
