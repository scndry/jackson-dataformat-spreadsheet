package io.github.scndry.jackson.dataformat.spreadsheet;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Date;
import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;
import com.fasterxml.jackson.databind.type.SimpleType;

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

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Unit-level checks for the back-write projection — guards the
 * cell-XML upper bound, the linear scaling against list size, the
 * heap-aware limit defaults, and the schema-level scope detection.
 *
 * <p>These rules are documented in
 * {@link BackWriteProjection#cellMaxBytes(Column)} and
 * {@link BackWriteProjection#backWriteBufferLimit()}; a regression here
 * would silently weaken the back-write OOM guard.
 */
class BackWriteProjectionTest {

    // SoA internal memory units — must stay in sync with
    // BackWriteProjection.{CELL_MEMORY_BYTES, ROW_MEMORY_BYTES}.
    private static final int CELL_MEMORY_BYTES = 20;
    private static final int ROW_MEMORY_BYTES = 8;

    // ----------------------------------------------------------------
    // cellMaxBytes — SoA cell record is a fixed 20 bytes regardless of
    // Java type (long _packed + long _values + int _next).
    // ----------------------------------------------------------------

    @Test
    void cellMaxBytes_isConstantAcrossAllSupportedTypes() {
        assertThat(BackWriteProjection.cellMaxBytes(_column(boolean.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(Boolean.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(String.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(SampleEnum.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(BigDecimal.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(BigInteger.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(int.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(long.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(float.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(double.class))).isEqualTo(CELL_MEMORY_BYTES);
        assertThat(BackWriteProjection.cellMaxBytes(_column(Date.class))).isEqualTo(CELL_MEMORY_BYTES);
    }

    enum SampleEnum { A }

    private static Column _column(final Class<?> raw) {
        return new Column(
                ColumnPointer.empty().resolve("x"),
                DataColumn.Value.empty(),
                SimpleType.constructUnsafe(raw));
    }

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
