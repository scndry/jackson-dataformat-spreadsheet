package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import java.util.List;

import com.fasterxml.jackson.core.JsonToken;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.SheetStreamReadException;
import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.Column;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.internal.SpreadsheetSchemaImpl;

import static org.assertj.core.api.Assertions.assertThatCode;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * Direct unit checks on the buffer fail-fast — exercises RecordTreeBuffer
 * with a small bufferLimitBytes so the per-cell footprint quickly trips
 * the limit. Bypasses SheetParser to keep the test boilerplate-free.
 */
class RecordTreeBufferLimitTest {

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Inner {
        @DataColumn("x") int x;
        @DataColumn("y") int y;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Outer {
        @DataColumn(value = "id", anchor = true) int id;
        List<Inner> items;
    }

    @Test
    void onRowEnd_bufferUnderLimit_doesNotThrow() throws Exception {
        // 21 cells × 20 bytes = 420 → well under 10_000
        assertThatCode(_runNRows(10_000L, 10)).doesNotThrowAnyException();
    }

    @Test
    void onRowEnd_bufferOverLimit_throws() throws Exception {
        // small limit (~5 cells) — first inner row already trips it
        assertThatThrownBy(_runNRows(100L, 5))
                .isInstanceOf(SheetStreamReadException.class)
                .hasMessageContaining("Nested-list buffer")
                .hasMessageContaining("exceeds limit")
                .hasMessageContaining("USE_POI_USER_MODEL");
    }

    private static ThrowingRunnable _runNRows(final long bufferLimitBytes, final int innerRows)
            throws Exception {
        final SpreadsheetSchemaImpl schema = (SpreadsheetSchemaImpl) new SpreadsheetMapper().sheetSchemaFor(Outer.class);
        final Column idCol = _findColumn(schema, "id");
        final Column xCol = _findColumn(schema, "x");
        final Column yCol = _findColumn(schema, "y");

        return () -> {
            final RecordTreeBuffer alg = new RecordTreeBuffer(schema, bufferLimitBytes);
            final CountingEmitter out = new CountingEmitter();
            alg.onSheetDataStart(out);
            // anchor row carries id + first inner cells
            alg.onRowStart();
            alg.onCellValue(idCol, new CellValue(1.0));
            alg.onCellValue(xCol, new CellValue(1.0));
            alg.onCellValue(yCol, new CellValue(2.0));
            alg.onRowEnd(out);
            // continuation rows accumulate inner cells without freeing
            for (int i = 1; i < innerRows; i++) {
                alg.onRowStart();
                alg.onCellValue(xCol, new CellValue(1.0 + i));
                alg.onCellValue(yCol, new CellValue(2.0 + i));
                alg.onRowEnd(out);
            }
            alg.onSheetDataEnd(out);
        };
    }

    private static Column _findColumn(final SpreadsheetSchemaImpl schema, final String name) {
        for (final Column c : schema) {
            if (c != null && name.equals(c.getName())) return c;
        }
        throw new AssertionError("no column " + name);
    }

    interface ThrowingRunnable extends org.assertj.core.api.ThrowableAssert.ThrowingCallable {
        void run() throws Exception;
        @Override default void call() throws Throwable { run(); }
    }

    private static final class CountingEmitter implements RecordTreeBuffer.Emitter {
        @Override public void token(final JsonToken t) { /* ignore */ }
        @Override public void fieldName(final String n) { /* ignore */ }
        @Override public void scalar(final CellValue v, final JsonToken s) { /* ignore */ }
    }
}
