package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

/**
 * Lightweight style lookup abstraction for resolving a cell's style index
 * to a number-format pattern, isolating SSML read code from {@code styles.xml}
 * representation choices.
 */
interface StylesLookup {

    int getNumFmtId(int xfIdx);

    String getFormatString(int numFmtId);
}
