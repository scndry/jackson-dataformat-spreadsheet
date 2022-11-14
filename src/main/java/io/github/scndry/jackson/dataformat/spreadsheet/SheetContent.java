package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;

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
