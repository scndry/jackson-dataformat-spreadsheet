package io.github.scndry.jackson.dataformat.spreadsheet.schema.internal;

import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.WeakHashMap;

import lombok.extern.slf4j.Slf4j;

import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;

/**
 * Internal helpers for the SSML back-write scenario — per-cell heap
 * footprint, runtime limit, schema-level scope detection, and projected
 * size. All sizes are SoA <strong>internal memory</strong> bytes (not
 * output XML bytes); the back-write limit's intent is heap protection
 * while flush is suspended inside an array scope. Used by
 * {@code SheetDataBuffer} (per-cell memory) and {@code SheetGenerator}
 * (pre-list fail-fast check on {@code writeStartArray}).
 *
 * <p>Not part of the public API. Classes under
 * {@code io.github.scndry.jackson.dataformat.spreadsheet.schema.internal}
 * may change without notice between releases — do not invoke directly
 * from application code.
 */
@Slf4j
public final class BackWriteProjection {

    /** SoA cell record footprint — {@code long _packed (8) + long _values (8) + int _next (4)}.
     *  Single source of truth shared with {@code SheetDataBuffer.byteSize()}. */
    public static final int CELL_MEMORY_BYTES = 20;

    /** SoA row directory entry footprint — {@code int _rowHead (4) + int _rowTail (4)}.
     *  Single source of truth shared with {@code SheetDataBuffer.byteSize()}. */
    public static final int ROW_MEMORY_BYTES = 8;

    /** Cache for {@link #requiresBackWriteScope(SpreadsheetSchema)} keyed on
     *  schema identity. WeakHashMap so schemas reclaimed by GC drop their
     *  entry — typical applications hold a single schema for their lifetime,
     *  edge-case ClassLoader reload / dynamically generated schemas do not
     *  leak. */
    private static final Map<SpreadsheetSchema, Boolean> SCOPE_CACHE =
            Collections.synchronizedMap(new WeakHashMap<>());

    private BackWriteProjection() {
    }

    /** SoA heap footprint of one inner row — every column's cell record
     *  ({@link #CELL_MEMORY_BYTES}) plus the row directory entry
     *  ({@link #ROW_MEMORY_BYTES}). */
    public static long innerRowMaxBytes(
            final SpreadsheetSchema schema, final ColumnPointer arrayPointer) {
        final List<Column> inners = schema.getColumns(arrayPointer);
        return (long) ROW_MEMORY_BYTES + (long) inners.size() * CELL_MEMORY_BYTES;
    }

    /** Projected back-write buffer max for a nested list of
     *  {@code listSize} elements rooted at {@code arrayPointer}. */
    public static long project(
            final SpreadsheetSchema schema, final ColumnPointer arrayPointer, final int listSize) {
        if (listSize <= 0) return 0;
        return (long) listSize * innerRowMaxBytes(schema, arrayPointer);
    }

    /** Back-write buffer limit (bytes of SoA internal memory).
     *  Default {@code max(1 MB, heap/128)}.
     *
     *  <p>{@code SheetDataBuffer} grows its SoA arrays by 1.5×
     *  ({@code Arrays.copyOf}), so during a grow the old (≈ ⅔ × limit) and
     *  new (limit) arrays coexist — peak heap during a single grow
     *  ≈ 1.667 × current buffer size. heap/128 limit + 1.667× grow peak
     *  ≈ heap/77 transient peak, leaving the remaining heap (≈ 76/77) for
     *  other allocations and GC headroom. The 1 MB floor keeps the limit
     *  meaningful on small heaps (Lambda / container).
     *
     *  <p>The limit measures SoA cell records, not output XML. The output
     *  XML stream ({@code _sb} in {@code SSMLSheetWriter}) is bounded by
     *  its own flush threshold and is flushed cell-by-cell during
     *  {@code flushTo} via a streaming sink.
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

    /** Returns true when any record scope (root or nested) has an
     *  outer field of that scope appearing after a child array of the
     *  same scope — the trigger for SSML back-write at that scope. */
    public static boolean hasOuterFieldAfterList(final SpreadsheetSchema schema) {
        final Set<ColumnPointer> scopes = new LinkedHashSet<>();
        scopes.add(ColumnPointer.empty());
        for (final Column c : schema) {
            if (c == null) continue;
            ColumnPointer head = ColumnPointer.empty();
            for (final ColumnPointer seg : c.getPointer()) {
                head = head.resolve(seg);
                if (seg.equals(ColumnPointer.array())) scopes.add(head);
            }
        }
        for (final ColumnPointer scope : scopes) {
            if (_outerAfterListAt(schema, scope)) return true;
        }
        return false;
    }

    private static boolean _outerAfterListAt(
            final SpreadsheetSchema schema, final ColumnPointer scope) {
        boolean seenChildArray = false;
        for (final Column c : schema) {
            if (c == null) continue;
            final ColumnPointer myScope = _immediateScope(c.getPointer());
            if (!myScope.equals(scope) && !myScope.startsWith(scope)) continue;
            if (myScope.equals(scope)) {
                if (seenChildArray) return true;
            } else {
                seenChildArray = true;
            }
        }
        return false;
    }

    private static ColumnPointer _immediateScope(final ColumnPointer pointer) {
        ColumnPointer head = ColumnPointer.empty();
        ColumnPointer last = ColumnPointer.empty();
        for (final ColumnPointer seg : pointer) {
            head = head.resolve(seg);
            if (seg.equals(ColumnPointer.array())) last = head;
        }
        return last;
    }

    /** Returns true when any record scope (root or nested) has more than
     *  one immediate child array — sibling lists share the parent row,
     *  so the second list must back-write into rows already produced by
     *  the first list. */
    public static boolean hasMultipleSiblingLists(final SpreadsheetSchema schema) {
        final Map<ColumnPointer, Set<ColumnPointer>> childArraysByParent = new HashMap<>();
        for (final Column c : schema) {
            if (c == null) continue;
            ColumnPointer head = ColumnPointer.empty();
            ColumnPointer parent = ColumnPointer.empty();
            for (final ColumnPointer seg : c.getPointer()) {
                head = head.resolve(seg);
                if (seg.equals(ColumnPointer.array())) {
                    childArraysByParent.computeIfAbsent(parent, k -> new HashSet<>()).add(head);
                    parent = head;
                }
            }
        }
        for (final Set<ColumnPointer> childArrays : childArraysByParent.values()) {
            if (childArrays.size() > 1) return true;
        }
        return false;
    }

    /** Cached form of {@link #hasOuterFieldAfterList(SpreadsheetSchema)} or
     *  {@link #hasMultipleSiblingLists(SpreadsheetSchema)}. Writers call this
     *  once per nested array to gate flush suspension — the cache avoids
     *  recomputing the result across every record. */
    public static boolean requiresBackWriteScope(final SpreadsheetSchema schema) {
        final Boolean cached = SCOPE_CACHE.get(schema);
        if (cached != null) return cached;
        final boolean v = hasOuterFieldAfterList(schema) || hasMultipleSiblingLists(schema);
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
