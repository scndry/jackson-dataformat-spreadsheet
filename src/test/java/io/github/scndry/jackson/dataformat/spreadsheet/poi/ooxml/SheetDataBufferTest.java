package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Direct tests for {@link SheetDataBuffer} — bit packing round-trip,
 * row-linked back-write insertion, rolling row directory, and growth.
 */
class SheetDataBufferTest {

    @Test
    void appendNumeric_emitsCellWithValueAttribute() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(0, 0, 2, 3.5);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).isEqualTo("<row r=\"1\"><c r=\"A1\" s=\"2\" t=\"n\"><v>3.5</v></c></row>");
    }

    @Test
    void appendString_emitsSharedStringIndex() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendString(1, 2, 0, 42);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).isEqualTo("<row r=\"2\"><c r=\"C2\" s=\"0\" t=\"s\"><v>42</v></c></row>");
    }

    @Test
    void appendBoolean_emitsZeroOrOne() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendBoolean(0, 0, 0, true);
        buf.appendBoolean(0, 1, 0, false);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).isEqualTo(
                "<row r=\"1\">"
                        + "<c r=\"A1\" s=\"0\" t=\"b\"><v>1</v></c>"
                        + "<c r=\"B1\" s=\"0\" t=\"b\"><v>0</v></c>"
                        + "</row>");
    }

    @Test
    void appendBlank_emitsSelfClosingCell() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendBlank(0, 0, 5);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).isEqualTo("<row r=\"1\"><c r=\"A1\" s=\"5\"/></row>");
    }

    @Test
    void appendsSameRow_emitsCellsInAppendOrder() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(0, 0, 0, 1);
        buf.appendNumeric(0, 1, 0, 2);
        buf.appendNumeric(0, 2, 0, 3);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).contains(
                "<v>1.0</v></c><c r=\"B1\"",
                "<v>2.0</v></c><c r=\"C1\"",
                "<v>3.0</v></c></row>");
    }

    @Test
    void appendsForwardRows_emitsRowsInOrder() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(0, 0, 0, 1);
        buf.appendNumeric(1, 0, 0, 2);
        buf.appendNumeric(2, 0, 0, 3);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).isEqualTo(
                "<row r=\"1\"><c r=\"A1\" s=\"0\" t=\"n\"><v>1.0</v></c></row>"
                        + "<row r=\"2\"><c r=\"A2\" s=\"0\" t=\"n\"><v>2.0</v></c></row>"
                        + "<row r=\"3\"><c r=\"A3\" s=\"0\" t=\"n\"><v>3.0</v></c></row>");
    }

    @Test
    void backWrite_appendsToPastRowTail() {
        // Simulates back-write: inner list rows 0..1, then outer cell on row 0
        // lands at the tail of row 0 — emitted in cell-list order, not column order.
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(0, 1, 0, 10);   // inner row 0
        buf.appendNumeric(1, 1, 0, 20);   // inner row 1
        buf.appendNumeric(0, 3, 0, 99);   // back-write: outer cell on row 0

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).isEqualTo(
                "<row r=\"1\">"
                        + "<c r=\"B1\" s=\"0\" t=\"n\"><v>10.0</v></c>"
                        + "<c r=\"D1\" s=\"0\" t=\"n\"><v>99.0</v></c>"
                        + "</row>"
                        + "<row r=\"2\">"
                        + "<c r=\"B2\" s=\"0\" t=\"n\"><v>20.0</v></c>"
                        + "</row>");
    }

    @Test
    void rowBaseTracksFirstAppendedRow() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(5, 0, 0, 1);
        assertThat(buf.maxRowSeen()).isEqualTo(5);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).startsWith("<row r=\"6\">");
    }

    @Test
    void rowBeforeBase_throwsIllegalStateException() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(5, 0, 0, 1);

        assertThatThrownBy(() -> buf.appendNumeric(3, 0, 0, 2))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("precedes rowBase");
    }

    @Test
    void emptyRowsInRange_skippedOnFlush() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(0, 0, 0, 1);
        buf.appendNumeric(3, 0, 0, 2);   // rows 1, 2 stay empty

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).isEqualTo(
                "<row r=\"1\"><c r=\"A1\" s=\"0\" t=\"n\"><v>1.0</v></c></row>"
                        + "<row r=\"4\"><c r=\"A4\" s=\"0\" t=\"n\"><v>2.0</v></c></row>");
    }

    @Test
    void byteSize_isUpperBoundOfFlushOutput() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(0, 0, 0, 1);
        buf.appendNumeric(0, 1, 0, 2);
        buf.appendNumeric(1, 0, 0, 3);

        // 3 cells × 68 + 2 row tags × 23 = 250 — upper bound
        final long predicted = buf.byteSize();
        assertThat(predicted).isEqualTo(3L * 68 + 2L * 23);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat((long) sb.length()).isLessThanOrEqualTo(predicted);
    }

    @Test
    void flushTo_resetsBuffer() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(0, 0, 0, 1);
        assertThat(buf.size()).isEqualTo(1);

        buf.flushTo(new StringBuilder());

        assertThat(buf.isEmpty()).isTrue();
        assertThat(buf.size()).isZero();
        assertThat(buf.maxRowSeen()).isEqualTo(-1);
    }

    @Test
    void emptyBuffer_flushEmitsNothing() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).isEmpty();
    }

    @Test
    void growth_handles10000Cells() {
        // Beyond INITIAL_CELL_CAPACITY (64) and INITIAL_ROW_CAPACITY (16) —
        // exercises 1.5× growth on both cell arrays and row directory.
        SheetDataBuffer buf = new SheetDataBuffer(8);
        for (int row = 0; row < 1_000; row++) {
            for (int col = 0; col < 10; col++) {
                buf.appendNumeric(row, col, 0, row + col);
            }
        }
        assertThat(buf.size()).isEqualTo(10_000);
        assertThat(buf.maxRowSeen()).isEqualTo(999);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).startsWith("<row r=\"1\">");
        assertThat(sb.toString()).endsWith("</row>");
    }

    @Test
    void colLetterCache_handlesColumnsBeyondCacheSize() {
        // cache size 2, col 5 falls back to direct convertNumToColString.
        SheetDataBuffer buf = new SheetDataBuffer(2);
        buf.appendNumeric(0, 5, 0, 1);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        assertThat(sb.toString()).contains("r=\"F1\"");
    }

    @Test
    void colLetterCache_reusesEntryAcrossCalls() {
        SheetDataBuffer buf = new SheetDataBuffer(4);
        buf.appendNumeric(0, 1, 0, 1);
        buf.appendNumeric(1, 1, 0, 2);
        buf.appendNumeric(2, 1, 0, 3);

        StringBuilder sb = new StringBuilder();
        buf.flushTo(sb);
        // All three cells reference column B — cache hit on entries 2 and 3.
        assertThat(sb.toString()).contains("r=\"B1\"", "r=\"B2\"", "r=\"B3\"");
    }

    @Test
    void typeCodeMatchesPoiCellType() {
        // POI CellType.code: NUMERIC=0, STRING=1, FORMULA=2, BLANK=3, BOOLEAN=4, ERROR=5
        // Reserved codes (2, 5) are intentionally unmapped.
        assertThat(SheetDataBuffer.TYPE_NUMERIC).isEqualTo((byte) 0);
        assertThat(SheetDataBuffer.TYPE_STRING).isEqualTo((byte) 1);
        assertThat(SheetDataBuffer.TYPE_BLANK).isEqualTo((byte) 3);
        assertThat(SheetDataBuffer.TYPE_BOOLEAN).isEqualTo((byte) 4);
    }
}
