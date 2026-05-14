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

    // Constants mirror BackWriteProjection — duplicated here as a test
    // contract: if the writer-side numbers change, this test must change.
    private static final int CELL_FIXED_OVERHEAD = 27 + 10 + 5 + 1;   // tags + ref + style + type
    private static final int ROW_TAG_BYTES = 23;

    // ----------------------------------------------------------------
    // cellMaxBytes — every supported Java type has a bounded cell XML
    // ----------------------------------------------------------------

    @Test
    void cellMaxBytes_booleanIsOneCharValue() {
        assertThat(BackWriteProjection.cellMaxBytes(_column(boolean.class)))
                .isEqualTo(CELL_FIXED_OVERHEAD + 1);
        assertThat(BackWriteProjection.cellMaxBytes(_column(Boolean.class)))
                .isEqualTo(CELL_FIXED_OVERHEAD + 1);
    }

    @Test
    void cellMaxBytes_sharedStringTypesAreTenCharIndex() {
        // String / Enum / BigDecimal / BigInteger all route through
        // SharedStringsStore, so only the int index (max 10 digits) lands
        // in the cell XML.
        final int expected = CELL_FIXED_OVERHEAD + 10;
        assertThat(BackWriteProjection.cellMaxBytes(_column(String.class)))
                .isEqualTo(expected);
        assertThat(BackWriteProjection.cellMaxBytes(_column(SampleEnum.class)))
                .isEqualTo(expected);
        assertThat(BackWriteProjection.cellMaxBytes(_column(BigDecimal.class)))
                .isEqualTo(expected);
        assertThat(BackWriteProjection.cellMaxBytes(_column(BigInteger.class)))
                .isEqualTo(expected);
    }

    @Test
    void cellMaxBytes_numericAndDateUseDoubleWorstCase() {
        // All numeric overloads route through writeNumeric(double) →
        // StringBuilder.append(double) — Double.toString worst case 25 chars.
        // Date / java.time.* route through writeNumber(double serial), same bound.
        final int expected = CELL_FIXED_OVERHEAD + 25;
        assertThat(BackWriteProjection.cellMaxBytes(_column(int.class))).isEqualTo(expected);
        assertThat(BackWriteProjection.cellMaxBytes(_column(long.class))).isEqualTo(expected);
        assertThat(BackWriteProjection.cellMaxBytes(_column(float.class))).isEqualTo(expected);
        assertThat(BackWriteProjection.cellMaxBytes(_column(double.class))).isEqualTo(expected);
        assertThat(BackWriteProjection.cellMaxBytes(_column(Date.class))).isEqualTo(expected);
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

        // Inner has 1 numeric column → cellMax 68 + row tag 23 = 91 bytes.
        final long projected = BackWriteProjection.project(schema, arrayPointer, 1);
        assertThat(projected).isEqualTo(ROW_TAG_BYTES + (CELL_FIXED_OVERHEAD + 25));
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
