package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.ColumnPointer;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection;

import static io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection.CELL_MEMORY_BYTES;
import static io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.BackWriteProjection.ROW_MEMORY_BYTES;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

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

    // Mid level holds the outer-after-list pattern; root is clean
    // (id + items[] only). Triggers SSML back-write at the items/[] scope
    // — root-only detection misses this.
    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidWithTrailing {
        @DataColumn(value = "midId", merge = OptBoolean.TRUE) String midId;
        List<Inner> details;
        @DataColumn(value = "midName", merge = OptBoolean.TRUE) String midName;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NestedOuterAfterList {
        @DataColumn(value = "id", merge = OptBoolean.TRUE) int id;
        List<MidWithTrailing> items;
    }

    @Test
    void requiresBackWriteScope_nestedScopeOuterAfterListReturnsTrue() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(NestedOuterAfterList.class);
        assertThat(BackWriteProjection.requiresBackWriteScope(schema)).isTrue();
    }

    // Mid level keeps outer-before-list ordering at every scope; back-write
    // not required anywhere.
    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidOuterFirst {
        @DataColumn(value = "midId", merge = OptBoolean.TRUE) String midId;
        @DataColumn(value = "midName", merge = OptBoolean.TRUE) String midName;
        List<Inner> details;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NestedOuterFirst {
        @DataColumn(value = "id", merge = OptBoolean.TRUE) int id;
        List<MidOuterFirst> items;
    }

    @Test
    void requiresBackWriteScope_nestedScopeOuterFirstReturnsFalse() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(NestedOuterFirst.class);
        assertThat(BackWriteProjection.requiresBackWriteScope(schema)).isFalse();
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Extra {
        @DataColumn("e") int e;
    }

    // Root scope holds two sibling lists — top-level detection already
    // catches this. Baseline test the existing schema lacked.
    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class RootSiblingLists {
        @DataColumn("id") int id;
        List<Inner> items;
        List<Extra> extras;
    }

    @Test
    void requiresBackWriteScope_rootMultiSiblingReturnsTrue() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(RootSiblingLists.class);
        assertThat(BackWriteProjection.requiresBackWriteScope(schema)).isTrue();
    }

    // Sibling lists live inside a nested record (mids[]/{as[], bs[]}) —
    // top-level head count is 1 (mids/[]), so top-level-only detection
    // misses this.
    @Data @NoArgsConstructor @AllArgsConstructor
    static class NestedMultiSiblingMid {
        @DataColumn("mid") String mid;
        List<Inner> as;
        List<Extra> bs;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NestedMultiSiblingTop {
        @DataColumn("id") int id;
        List<NestedMultiSiblingMid> mids;
    }

    @Test
    void requiresBackWriteScope_nestedScopeMultiSiblingReturnsTrue() throws Exception {
        final SpreadsheetSchema schema = _schemaFor(NestedMultiSiblingTop.class);
        assertThat(BackWriteProjection.requiresBackWriteScope(schema)).isTrue();
    }

    // ----------------------------------------------------------------
    // Integration — writeStartArray's fail-fast on projected overflow
    // ----------------------------------------------------------------

    @TempDir File tempDir;

    @Test
    void writeStartArray_projectionExceedsLimit_throwsWithMitigationGuidance() throws Exception {
        // Outer has back-write scope (list followed by outer field). Compute
        // the boundary list size that just trips the projection limit and
        // verify the message contract — operator size, the projected and
        // limit byte figures, and the three mitigation options.
        final SpreadsheetMapper mapper = new SpreadsheetMapper();
        final SpreadsheetSchema schema = mapper.sheetSchemaFor(Outer.class);
        final ColumnPointer arrayPointer = _findInnerArrayPointer(schema);
        final long perRow = BackWriteProjection.project(schema, arrayPointer, 1);
        final long limit = BackWriteProjection.backWriteBufferLimit();
        final int overflowSize = Math.toIntExact(limit / perRow + 1);

        final List<Inner> items = new ArrayList<>(overflowSize);
        for (int i = 0; i < overflowSize; i++) items.add(new Inner(i));
        final Outer outer = new Outer(1, items, 100);

        final File file = new File(tempDir, "overflow.xlsx");

        assertThatThrownBy(() -> mapper.writeValue(
                        file, Collections.singletonList(outer), Outer.class))
                .rootCause()
                .isInstanceOf(SheetStreamWriteException.class)
                .hasMessageContaining("Nested list size " + overflowSize)
                .hasMessageContaining("projected back-write")
                .hasMessageContaining("exceeds limit")
                .hasMessageContaining("USE_POI_USER_MODEL");
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
