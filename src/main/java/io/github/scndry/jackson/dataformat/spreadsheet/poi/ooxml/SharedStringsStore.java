package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.IOException;

/**
 * Writer-side shared string store abstraction.
 * Provides deduplication and index-based retrieval for building sharedStrings.xml.
 */
interface SharedStringsStore extends AutoCloseable {

    int getOrAdd(String value);

    int size();

    int length(int index);

    char charAt(int index, int offset);

    String get(int index);

    @Override
    void close() throws IOException;
}
