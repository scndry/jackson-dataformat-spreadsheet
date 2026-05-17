package io.github.scndry.jackson.dataformat.spreadsheet.poi.ooxml;

import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.ss.util.CellReference;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection;

/**
 * Sheet-data accumulator using Structure of Arrays layout.
 *
 * <p>Cell metadata packs into 16 bytes ({@code long _packed} + {@code long _values}),
 * with a row-linked list ({@code int _next}) for O(1) back-write insertion into
 * past rows. A rolling row directory tracks per-row head/tail cell indices, so
 * cells appended out-of-row-order (back-write) land at the row's tail without
 * scanning the buffer.
 *
 * <p>Bit layout of {@code _packed}:
 * <pre>
 *   bit 63..44  row     (20 bits, Excel max 1,048,576)
 *   bit 43..30  col     (14 bits, Excel max 16,384)
 *   bit 29..15  style   (15 bits, Excel max 32,767)
 *   bit 14..12  type    (3 bits, matches org.apache.poi.ss.usermodel.CellType.code)
 *   bit 11..0   reserved
 * </pre>
 *
 * <p>{@code _values}:
 * <ul>
 *   <li>{@link #TYPE_NUMERIC}: {@link Double#doubleToRawLongBits(double)} of the value</li>
 *   <li>{@link #TYPE_STRING}: shared-string index (low 32 bits)</li>
 *   <li>{@link #TYPE_BOOLEAN}: 0 or 1</li>
 *   <li>{@link #TYPE_BLANK}: 0 (slot unused)</li>
 * </ul>
 *
 * <p>Package-private — implementation detail of {@link SSMLSheetWriter}.
 */
final class SheetDataBuffer {

    // matches org.apache.poi.ss.usermodel.CellType.code (POI 3.15 beta 3+).
    // FORMULA (2) and ERROR (5) are reserved — not emitted by the streaming writer.
    static final byte TYPE_NUMERIC = 0;
    static final byte TYPE_STRING  = 1;
    static final byte TYPE_BLANK   = 3;
    static final byte TYPE_BOOLEAN = 4;

    // Lazy capacity inflated on first append, then grown 1.5× — same shape
    // as java.util.ArrayList. Covers typical Excel schemas (≤ 16 columns)
    // without grow, while keeping idle buffers minimal.
    private static final int DEFAULT_CAPACITY = 16;

    private static final long[] EMPTY_LONG_ARRAY = new long[0];
    private static final int[] EMPTY_INT_ARRAY = new int[0];

    private static final int TYPE_BITS = 3;
    private static final int STYLE_BITS = 15;
    private static final int COL_BITS = 14;
    private static final int ROW_BITS = 20;

    private static final int TYPE_SHIFT = 12;
    private static final int STYLE_SHIFT = TYPE_SHIFT + TYPE_BITS;        // 15
    private static final int COL_SHIFT = STYLE_SHIFT + STYLE_BITS;        // 30
    private static final int ROW_SHIFT = COL_SHIFT + COL_BITS;            // 44

    private static final long TYPE_MASK = (1L << TYPE_BITS) - 1;
    private static final long STYLE_MASK = (1L << STYLE_BITS) - 1;
    private static final long COL_MASK = (1L << COL_BITS) - 1;
    private static final long ROW_MASK = (1L << ROW_BITS) - 1;

    private long[] _packed = EMPTY_LONG_ARRAY;
    private long[] _values = EMPTY_LONG_ARRAY;
    private int[] _next = EMPTY_INT_ARRAY;
    private int _size;

    private int _rowBase;
    private int _rowSpan;
    private int[] _rowHead = EMPTY_INT_ARRAY;
    private int[] _rowTail = EMPTY_INT_ARRAY;

    private final String[] _colLetters;

    SheetDataBuffer(final int colLetterCacheSize) {
        if (colLetterCacheSize < 1) {
            throw new IllegalArgumentException(
                    "colLetterCacheSize must be >= 1, got " + colLetterCacheSize);
        }
        _colLetters = new String[colLetterCacheSize];
    }

    void appendNumeric(final int row, final int col, final int style, final double value) {
        _append(row, col, style, TYPE_NUMERIC, Double.doubleToRawLongBits(value));
    }

    void appendString(final int row, final int col, final int style, final int sstIndex) {
        // Low 32 bits of the value slot hold the index; high 32 bits remain reserved.
        _append(row, col, style, TYPE_STRING, sstIndex);
    }

    void appendBoolean(final int row, final int col, final int style, final boolean value) {
        _append(row, col, style, TYPE_BOOLEAN, value ? 1L : 0L);
    }

    void appendBlank(final int row, final int col, final int style) {
        _append(row, col, style, TYPE_BLANK, 0L);
    }

    int size() {
        return _size;
    }

    boolean isEmpty() {
        return _size == 0;
    }

    int maxRowSeen() {
        return _rowSpan == 0 ? -1 : _rowBase + _rowSpan - 1;
    }

    /** Upper-bound heap footprint of the cell SoA arrays and row directory.
     *  Used by the back-write runtime monitor as a fail-fast guard against
     *  unbounded accumulation while flush is suspended. The output XML is a
     *  separate concern — its size is bounded by
     *  {@link io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection}
     *  on the same memory basis. */
    long byteSize() {
        return (long) _size * BackWriteProjection.CELL_MEMORY_BYTES
                + (long) _rowSpan * BackWriteProjection.ROW_MEMORY_BYTES;
    }

    /** Sink invoked after each XML fragment is appended to {@code sb} so
     *  the writer can keep its buffer within its flush threshold. Matches
     *  the existing fragment-level check pattern in
     *  {@code SSMLSheetWriter._append}. */
    @FunctionalInterface
    interface FlushSink {
        void afterFragment() throws IOException;
    }

    private static final FlushSink NO_OP_SINK = () -> {};

    /** Convenience overload for callers that do not need per-fragment
     *  draining (tests, in-memory accumulation). */
    void flushTo(final StringBuilder sb) {
        try {
            flushTo(sb, NO_OP_SINK);
        } catch (IOException e) {
            throw new AssertionError("NO_OP_SINK never throws", e);
        }
    }

    /** Emit all buffered cells, row by row, into {@code sb}, invoking
     *  {@code sink} after each cell or row-tag fragment so callers can
     *  drain {@code sb} before it exceeds its flush threshold. Resets
     *  the buffer on return. */
    void flushTo(final StringBuilder sb, final FlushSink sink) throws IOException {
        for (int offset = 0; offset < _rowSpan; offset++) {
            int cellIdx = _rowHead[offset];
            if (cellIdx < 0) continue;
            final int row = _rowBase + offset;
            sb.append("<row r=\"").append(row + 1).append("\">");
            sink.afterFragment();
            while (cellIdx >= 0) {
                _appendCell(sb, _packed[cellIdx], _values[cellIdx]);
                sink.afterFragment();
                cellIdx = _next[cellIdx];
            }
            sb.append("</row>");
            sink.afterFragment();
        }
        _reset();
    }

    private void _append(final int row, final int col, final int style,
                         final byte type, final long valueSlot) {
        final long packed =
                ((long) row << ROW_SHIFT)
                        | ((long) col << COL_SHIFT)
                        | ((long) style << STYLE_SHIFT)
                        | ((long) type << TYPE_SHIFT);

        final int cellIdx = _size;
        _ensureCellCapacity(cellIdx + 1);
        _packed[cellIdx] = packed;
        _values[cellIdx] = valueSlot;
        _next[cellIdx] = -1;

        if (_rowSpan == 0) {
            _rowBase = row;
            _ensureRowCapacity(1);
            _rowHead[0] = cellIdx;
            _rowTail[0] = cellIdx;
            _rowSpan = 1;
        } else if (row < _rowBase) {
            throw new IllegalStateException(
                    "Row " + row + " precedes rowBase " + _rowBase
                            + " — back-write target was flushed prematurely");
        } else {
            final int rowOffset = row - _rowBase;
            if (rowOffset >= _rowSpan) {
                _ensureRowCapacity(rowOffset + 1);
                for (int i = _rowSpan; i <= rowOffset; i++) {
                    _rowHead[i] = -1;
                    _rowTail[i] = -1;
                }
                _rowSpan = rowOffset + 1;
            }
            if (_rowHead[rowOffset] < 0) {
                _rowHead[rowOffset] = cellIdx;
                _rowTail[rowOffset] = cellIdx;
            } else {
                final int tailCol = (int) ((_packed[_rowTail[rowOffset]] >>> COL_SHIFT) & COL_MASK);
                if (col >= tailCol) {
                    // Fast path — forward column-ascending append (typical case).
                    _next[_rowTail[rowOffset]] = cellIdx;
                    _rowTail[rowOffset] = cellIdx;
                } else {
                    // Slow path — out-of-order column insertion (merge inner-cell
                    // fill emitted after inner items). OOXML strict readers
                    // require ascending cell-reference order within a row.
                    final int head = _rowHead[rowOffset];
                    final int headCol = (int) ((_packed[head] >>> COL_SHIFT) & COL_MASK);
                    if (col < headCol) {
                        _next[cellIdx] = head;
                        _rowHead[rowOffset] = cellIdx;
                    } else {
                        int prev = head;
                        while (true) {
                            final int nxt = _next[prev];
                            if (nxt < 0) break;
                            final int nxtCol = (int) ((_packed[nxt] >>> COL_SHIFT) & COL_MASK);
                            if (col < nxtCol) break;
                            prev = nxt;
                        }
                        _next[cellIdx] = _next[prev];
                        _next[prev] = cellIdx;
                    }
                }
            }
        }
        _size++;
    }

    private void _appendCell(final StringBuilder sb, final long packed, final long value) {
        final int row = (int) ((packed >>> ROW_SHIFT) & ROW_MASK);
        final int col = (int) ((packed >>> COL_SHIFT) & COL_MASK);
        final int style = (int) ((packed >>> STYLE_SHIFT) & STYLE_MASK);
        final byte type = (byte) ((packed >>> TYPE_SHIFT) & TYPE_MASK);

        sb.append("<c r=\"")
                .append(_colLetter(col))
                .append(row + 1)
                .append("\" s=\"")
                .append(style);
        switch (type) {
            case TYPE_NUMERIC:
                sb.append("\" t=\"n\"><v>")
                        .append(Double.longBitsToDouble(value))
                        .append("</v></c>");
                break;
            case TYPE_STRING:
                sb.append("\" t=\"s\"><v>")
                        .append((int) value)
                        .append("</v></c>");
                break;
            case TYPE_BOOLEAN:
                sb.append("\" t=\"b\"><v>")
                        .append(value == 0L ? '0' : '1')
                        .append("</v></c>");
                break;
            case TYPE_BLANK:
                sb.append("\"/>");
                break;
            default:
                throw new IllegalStateException("Unknown cell type code " + type);
        }
    }

    private String _colLetter(final int col) {
        if (col < 0 || col >= _colLetters.length) {
            return CellReference.convertNumToColString(col);
        }
        String letter = _colLetters[col];
        if (letter == null) {
            letter = CellReference.convertNumToColString(col);
            _colLetters[col] = letter;
        }
        return letter;
    }

    private void _reset() {
        _size = 0;
        _rowBase = 0;
        _rowSpan = 0;
    }

    private void _ensureCellCapacity(final int required) {
        if (required <= _packed.length) return;
        final int newCap = _grow(_packed.length, required);
        _packed = Arrays.copyOf(_packed, newCap);
        _values = Arrays.copyOf(_values, newCap);
        _next = Arrays.copyOf(_next, newCap);
    }

    private void _ensureRowCapacity(final int required) {
        if (required <= _rowHead.length) return;
        final int newCap = _grow(_rowHead.length, required);
        _rowHead = Arrays.copyOf(_rowHead, newCap);
        _rowTail = Arrays.copyOf(_rowTail, newCap);
    }

    private static int _grow(final int oldCap, final int required) {
        if (oldCap == 0) {
            return Math.max(DEFAULT_CAPACITY, required);
        }
        return Math.max(oldCap + (oldCap >> 1), required);
    }
}
