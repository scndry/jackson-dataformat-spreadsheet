package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * Builds a nested-list bench fixture — outer record count is
 * {@code innerRowCount / ITEMS_PER_OUTER}; each outer carries
 * {@link #ITEMS_PER_OUTER} {@link NestedLineItem}s, yielding
 * {@code innerRowCount} total data rows (file-size parity with the
 * flat {@code BenchWorkbookBuilder}).
 */
public final class NestedBenchWorkbookBuilder {

    public static final int ITEMS_PER_OUTER = 10;

    private NestedBenchWorkbookBuilder() {}

    public static File createSampleFile(final String prefix, final int innerRowCount) throws IOException {
        if (innerRowCount % ITEMS_PER_OUTER != 0) {
            throw new IllegalArgumentException(
                    "innerRowCount must be a multiple of " + ITEMS_PER_OUTER
                            + " (got " + innerRowCount + ")");
        }
        final int outerCount = innerRowCount / ITEMS_PER_OUTER;
        File file = File.createTempFile(prefix, ".xlsx");
        file.deleteOnExit();
        SpreadsheetMapper mapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS));
        List<NestedRow> data = NestedRow.sample(outerCount, ITEMS_PER_OUTER);
        mapper.writeValue(file, data, NestedRow.class);
        return file;
    }
}
