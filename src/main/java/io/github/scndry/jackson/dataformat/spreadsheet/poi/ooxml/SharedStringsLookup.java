package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

/**
 * Lightweight shared string lookup abstraction.
 * Returns plain {@code String} instead of POI's {@code RichTextString}
 * since only text content is needed.
 */
interface SharedStringsLookup extends AutoCloseable {

    String getItemAt(int idx);

    @Override
    void close() throws Exception;
}
