package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

/**
 * Sentinel {@link StylesLookup} for workbooks without a styles part.
 * Returns {@code numFmtId 0} and a {@code null} format string for any index.
 */
enum BlankStylesLookup implements StylesLookup {
    INSTANCE;

    @Override
    public int getNumFmtId(final int xfIdx) {
        return 0;
    }

    @Override
    public String getFormatString(final int numFmtId) {
        return null;
    }
}
