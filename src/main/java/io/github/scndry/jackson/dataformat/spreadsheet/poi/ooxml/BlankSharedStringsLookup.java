package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

/**
 * Sentinel {@link SharedStringsLookup} for workbooks without a shared string table.
 * Returns an empty string for any index.
 */
enum BlankSharedStringsLookup implements SharedStringsLookup {
    INSTANCE;

    @Override
    public String getItemAt(final int idx) {
        return "";
    }

    @Override
    public void close() {
        // no-op
    }
}
