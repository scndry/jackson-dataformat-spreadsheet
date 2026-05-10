package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import java.util.List;

import lombok.extern.slf4j.Slf4j;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Internal helpers for projecting the SSML back-write buffer size from
 * a schema + a runtime nested-list size, and for detecting whether a
 * given schema can trigger the back-write code path.
 *
 * <p>Not part of the public API. Classes under
 * {@code io.github.scndry.jackson.dataformat.spreadsheet.schema.internal}
 * may change without notice between releases — do not invoke directly
 * from application code.
 *
 * @see SpreadsheetSchema#cellMaxBytes(Column)
 * @see SpreadsheetSchema#backWriteBufferLimit()
 */
@Slf4j
public final class BackWriteProjection {

    private static final int ROW_TAG_BYTES = 22;

    private BackWriteProjection() {
        // utility class — no instances
    }

    /** Upper-bound bytes for one inner row (all columns inside the given
     *  array pointer + the {@code <row>} tag overhead). Every supported
     *  Java type has a bounded per-cell size — see
     *  {@link SpreadsheetSchema#cellMaxBytes(Column)} for the breakdown. */
    public static long innerRowMaxBytes(
            final SpreadsheetSchema schema, final ColumnPointer arrayPointer) {
        final List<Column> inners = schema.getColumns(arrayPointer);
        long bytes = ROW_TAG_BYTES;
        for (final Column c : inners) {
            bytes += SpreadsheetSchema.cellMaxBytes(c);
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
