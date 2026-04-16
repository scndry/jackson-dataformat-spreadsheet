package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

/**
 * Sentinel {@link SharedStringLookup} for workbooks without a shared string table.
 * Returns an empty string for any index.
 */
enum BlankSharedStringLookup implements SharedStringLookup {
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
