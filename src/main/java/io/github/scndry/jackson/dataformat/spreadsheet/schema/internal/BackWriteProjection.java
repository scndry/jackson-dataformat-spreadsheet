package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.WeakHashMap;

import com.fasterxml.jackson.databind.JavaType;

import lombok.extern.slf4j.Slf4j;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Internal helpers for the SSML back-write scenario — cell-XML upper bound,
 * runtime limit, schema-level scope detection, and projected size. Used by
 * {@code SheetDataBuffer} (byte upper bound for buffered cells) and
 * {@code SheetGenerator} (pre-list fail-fast check on {@code writeStartArray}).
 *
 * <p>Not part of the public API. Classes under
 * {@code io.github.scndry.jackson.dataformat.spreadsheet.schema.internal}
 * may change without notice between releases — do not invoke directly
 * from application code.
 */
@Slf4j
public final class BackWriteProjection {

    // <c r="REF" s="STYLE" t="TYPE"><v>VALUE</v></c>
    //   fixed = "<c r=\"" (6) + "\" s=\"" (4) + "\" t=\"" (4) + "\"><v>" (5)
    //         + "</v></c>" (8) = 27 bytes
    //   ref   max = "XFD1048576" = 10 chars
    //   style max = 32767 digits = 5 chars
    //   type      = "n" | "s" | "b" = 1 char (writeBlank emits no t attribute)
    private static final int CELL_FIXED_TAGS_BYTES = 27;
    private static final int CELL_REF_MAX = 10;
    private static final int CELL_STYLE_MAX = 5;
    private static final int CELL_TYPE_MAX = 1;

    // <row r="N">..</row>
    //   open  = "<row r=\"" (8) + N digits + "\">" (2) = 10 + N
    //   close = "</row>"                                    = 6
    // Excel row max = 1048576 (7 digits) → worst total = 23 bytes.
    private static final int ROW_TAG_BYTES = 23;

    /** Cache for {@link #requiresBackWriteScope(SpreadsheetSchema)} keyed on
     *  schema identity. WeakHashMap so schemas reclaimed by GC drop their
     *  entry — typical applications hold a single schema for their lifetime,
     *  edge-case ClassLoader reload / dynamically generated schemas do not
     *  leak. */
    private static final Map<SpreadsheetSchema, Boolean> SCOPE_CACHE =
            Collections.synchronizedMap(new WeakHashMap<>());

    private BackWriteProjection() {
    }

    /** Upper-bound bytes for one cell XML of the given column.
     *
     *  <p>Cell-XML output paths in this library:
     *  <ul>
     *    <li>All numeric overloads ({@code writeNumber(int|long|float|double)})
     *        route through {@code SheetGenerator → _writer.writeNumeric(double)},
     *        which calls {@code StringBuilder.append(double)} — every numeric
     *        primitive emits a {@code Double.toString}-formatted value
     *        (worst case ~25 chars), regardless of the source type.</li>
     *    <li>{@code boolean} emits "0" or "1" (1 char).</li>
     *    <li>{@code String} / {@code Enum} / {@code BigDecimal} /
     *        {@code BigInteger} route through {@code SheetGenerator →
     *        _writer.writeString → SharedStringsStore}, leaving only a
     *        shared-string index in the cell (max 10 digits for
     *        {@code Integer.MAX_VALUE}). Their content lives in
     *        {@code sharedStrings.xml}, a separate stream not gated by
     *        back-write buffer limits.</li>
     *    <li>{@code Date} / {@code java.time.*} route through
     *        {@code writeNumber(double serial)}, same 25-char bound.</li>
     *  </ul>
     *  Every supported Java type has a bounded cell XML size. */
    public static int cellMaxBytes(final Column column) {
        final JavaType type = column.getType();
        final Class<?> raw = type.getRawClass();
        final int valueMax;
        if (raw == boolean.class || raw == Boolean.class) {
            valueMax = 1;
        } else if (raw == String.class || raw.isEnum()
                || raw == java.math.BigDecimal.class
                || raw == java.math.BigInteger.class) {
            valueMax = 10;                                    // shared string index
        } else {
            valueMax = 25;
        }
        return CELL_FIXED_TAGS_BYTES + CELL_REF_MAX + CELL_STYLE_MAX + CELL_TYPE_MAX + valueMax;
    }

    /** Upper-bound bytes for one inner row (all columns inside the given
     *  array pointer + the {@code <row>} tag overhead). Every supported
     *  Java type has a bounded per-cell size — see {@link #cellMaxBytes}. */
    public static long innerRowMaxBytes(
            final SpreadsheetSchema schema, final ColumnPointer arrayPointer) {
        final List<Column> inners = schema.getColumns(arrayPointer);
        long bytes = ROW_TAG_BYTES;
        for (final Column c : inners) {
            bytes += cellMaxBytes(c);
        }
        return bytes;
    }

    /** Projected back-write buffer max for a nested list of
     *  {@code listSize} elements rooted at {@code arrayPointer}. */
    public static long project(
            final SpreadsheetSchema schema, final ColumnPointer arrayPointer, final int listSize) {
        if (listSize <= 0) return 0;
        return (long) listSize * innerRowMaxBytes(schema, arrayPointer);
    }

    /** Back-write buffer limit (bytes). Default {@code max(1 MB, heap/128)}.
     *
     *  <p>{@code SheetDataBuffer} grows its SoA arrays by 1.5×
     *  ({@code Arrays.copyOf}), so during a grow the old (≈ ⅔ × limit) and
     *  new (limit) arrays coexist — peak heap during a single grow
     *  ≈ 1.667 × current buffer size. heap/128 limit + 1.667× grow peak
     *  ≈ heap/77 transient peak, leaving the remaining heap (≈ 76/77) for
     *  other allocations and GC headroom. The 1 MB floor keeps the limit
     *  meaningful on small heaps (Lambda / container).
     *
     *  <p>No external override knob is exposed; a future need can introduce
     *  one on {@code SpreadsheetFactory} without breaking callers. */
    public static long backWriteBufferLimit() {
        return Math.max(1L * 1024 * 1024, Runtime.getRuntime().maxMemory() / 128);
    }

    /** Human-readable byte size formatting for diagnostics. */
    public static String formatBytes(final long bytes) {
        if (bytes < 1024L) return bytes + " B";
        if (bytes < 1024L * 1024) return (bytes / 1024) + " KB";
        if (bytes < 1024L * 1024 * 1024) {
            return String.format("%.1f MB", bytes / (1024.0 * 1024));
        }
        return String.format("%.2f GB", bytes / (1024.0 * 1024 * 1024));
    }

    /** Returns true when the schema's column order has any outer
     *  (non-array-scope) column appearing after at least one inner
     *  (array-scope) column — the trigger for SSML back-write. */
    public static boolean hasOuterFieldAfterList(final SpreadsheetSchema schema) {
        boolean seenInner = false;
        for (final Column c : schema) {
            if (c == null) continue;
            final boolean inner = c.getPointer().contains(ColumnPointer.array());
            if (inner) {
                seenInner = true;
            } else if (seenInner) {
                return true;
            }
        }
        return false;
    }

    /** Cached form of {@link #hasOuterFieldAfterList(SpreadsheetSchema)}.
     *  Writers call this once per nested array to gate flush suspension —
     *  the cache avoids recomputing the result across every record. */
    public static boolean requiresBackWriteScope(final SpreadsheetSchema schema) {
        final Boolean cached = SCOPE_CACHE.get(schema);
        if (cached != null) return cached;
        final boolean v = hasOuterFieldAfterList(schema);
        SCOPE_CACHE.put(schema, v);
        return v;
    }

    /** Logs a warn at schema build time when {@link #hasOuterFieldAfterList}
     *  returns true. Call once after schema construction. */
    public static void warnIfScenario(final SpreadsheetSchema schema) {
        if (!hasOuterFieldAfterList(schema)) return;
        log.warn(
                "Schema has an outer field declared after a List<T> field —"
                + " nested list contents are buffered in memory until the"
                + " outer field can be back-written. Large nested lists may"
                + " cause OOM in streaming mode."
                + " Consider USE_POI_USER_MODEL, splitting data, or"
                + " reordering fields so the outer field comes before the list.");
    }
}
