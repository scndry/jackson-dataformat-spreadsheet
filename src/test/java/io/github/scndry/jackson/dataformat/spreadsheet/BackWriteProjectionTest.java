package io.github.scndry.jackson.dataformat.spreadsheet;

import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection;

import static io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection.CELL_MEMORY_BYTES;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection.ROW_MEMORY_BYTES;
import static org.assertj.core.api.Assertions.assertThat;

/**
 * Unit-level checks for the back-write projection — guards the
 * linear scaling of {@link BackWriteProjection#project} against list
 * size, the heap-aware limit defaults, and the schema-level scope
 * detection.
 *
 * <p>A regression here would silently weaken the back-write OOM guard.
 */
class BackWriteProjectionTest {

    // ----------------------------------------------------------------
    // project — linear in list size, sums every inner column's bound
    // ----------------------------------------------------------------

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Inner {
        @DataColumn("v") int v;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Outer {
        @DataColumn(value = "id", merge = OptBoolean.TRUE) int id;
        List<Inner> items;
        @DataColumn(value = "total", merge = OptBoolean.TRUE) int total;
    }

    @Test
    void project_oneRow_equalsInnerRowMaxBytes() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(Outer.class);
        final ColumnPointer arrayPointer = _findInnerArrayPointer(schema);

        // Inner has 1 column → cellMemory 20 + rowMemory 8 = 28 bytes.
        final long projected = BackWriteProjection.project(schema, arrayPointer, 1);
        assertThat(projected).isEqualTo(ROW_MEMORY_BYTES + CELL_MEMORY_BYTES);
    }

    @Test
    void project_scalesLinearlyWithListSize() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(Outer.class);
        final ColumnPointer arrayPointer = _findInnerArrayPointer(schema);

        final long perRow = BackWriteProjection.project(schema, arrayPointer, 1);
        assertThat(BackWriteProjection.project(schema, arrayPointer, 1_000))
                .isEqualTo(perRow * 1_000);
        assertThat(BackWriteProjection.project(schema, arrayPointer, 100_000_000))
                .isEqualTo(perRow * 100_000_000L);
    }

    @Test
    void project_zeroOrNegativeSize_returnsZero() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(Outer.class);
        final ColumnPointer arrayPointer = _findInnerArrayPointer(schema);

        assertThat(BackWriteProjection.project(schema, arrayPointer, 0)).isZero();
        assertThat(BackWriteProjection.project(schema, arrayPointer, -1)).isZero();
    }

    // ----------------------------------------------------------------
    // backWriteBufferLimit — max(1 MB, heap/128)
    // ----------------------------------------------------------------

    @Test
    void backWriteBufferLimit_isAtLeastOneMegaFloor() {
        assertThat(BackWriteProjection.backWriteBufferLimit())
                .isGreaterThanOrEqualTo(1L * 1024 * 1024);
    }

    @Test
    void backWriteBufferLimit_matchesHeapOver128Formula() {
        // Limit is max(1 MB, heap/128). The cap matches max — so the limit
        // is never above heap/128 except when heap/128 < 1 MB (small heap),
        // in which case the floor takes over.
        final long heap = Runtime.getRuntime().maxMemory();
        final long expected = Math.max(1L * 1024 * 1024, heap / 128);
        assertThat(BackWriteProjection.backWriteBufferLimit()).isEqualTo(expected);
    }

    // ----------------------------------------------------------------
    // requiresBackWriteScope — detection over column order
    // ----------------------------------------------------------------

    @Test
    void requiresBackWriteScope_outerAfterListReturnsTrue() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(Outer.class);
        assertThat(BackWriteProjection.requiresBackWriteScope(schema)).isTrue();
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class OuterFirst {
        @DataColumn("id") int id;
        @DataColumn("total") int total;
        List<Inner> items;
    }

    @Test
    void requiresBackWriteScope_outerFirstReturnsFalse() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(OuterFirst.class);
        assertThat(BackWriteProjection.requiresBackWriteScope(schema)).isFalse();
    }

    // ----------------------------------------------------------------
    // Helpers
    // ----------------------------------------------------------------

    private static SpreadsheetSchema _schemaFor(final Class<?> type) throws Exception {
        return new SpreadsheetMapper().sheetSchemaFor(type);
    }

    /** Finds the array pointer ({@code items/[]}) for an inner column —
     *  the same pointer the writer hands to {@code project()} at
     *  {@code writeStartArray}. */
    private static ColumnPointer _findInnerArrayPointer(final SpreadsheetSchema schema) {
        for (final Column c : schema) {
            if (c == null) continue;
            final ColumnPointer p = c.getPointer();
            if (p.contains(ColumnPointer.array())) {
                return p.getParent();
            }
        }
        throw new AssertionError("schema has no inner column");
    }

}
