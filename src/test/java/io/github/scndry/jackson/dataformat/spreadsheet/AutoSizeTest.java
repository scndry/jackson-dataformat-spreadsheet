package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.annotation.OptBoolean;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for {@code @DataColumn(autoSize = TRUE)} measurement coverage.
 * <p>
 * Before this change, autoSize was scoped to the SXSSF random-access window or a
 * hardcoded last-100-row range for non-SXSSF — content outside that range was silently
 * ignored, producing column widths that didn't fit data the user could see in the file.
 * After the change, every cell is measured at write time and the per-column max is
 * applied at close, so a wide value in row 0 keeps influencing the final width even
 * after thousands of rows of short content.
 */
class AutoSizeTest {

    private static final String LONG_VALUE = "extra-long-value-that-should-widen-this-column";
    private static final String SHORT_VALUE = "x";

    @TempDir File tempDir;

    @Data @NoArgsConstructor @AllArgsConstructor
    @DataGrid(autoSizeColumn = OptBoolean.TRUE)
    static class Item {
        @DataColumn("Name")
        private String name;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    @DataGrid(autoSizeColumn = OptBoolean.TRUE)
    static class Pair {
        @DataColumn("First")
        private String first;
        @DataColumn("Second")
        private String second;
    }

    @Test
    void sxssfCapturesWidthFromFlushedRows() throws Exception {
        File file = new File(tempDir, "as-sxssf.xlsx");

        // SXSSFWorkbook default window = 100. The first data row's long value gets flushed long
        // before close(); without inline measurement it would be invisible to autoSize.
        List<Item> data = new ArrayList<>();
        data.add(new Item(LONG_VALUE));
        for (int i = 1; i < 250; i++) {
            data.add(new Item(SHORT_VALUE));
        }

        SpreadsheetMapper mapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new, SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            int width = sheet.getColumnWidth(0);
            // The long value is ~46 chars. A correct autoSize should reach ~12000+ (units = 1/256 char).
            // A broken autoSize (window-bounded, sees only "x") would be near 256-512.
            assertThat(width).as("column width must reflect the long value in the flushed row").isGreaterThan(20 * 256);
        }
    }

    @Test
    void xssfCapturesWidthBeyondLast100Rows() throws Exception {
        File file = new File(tempDir, "as-xssf.xlsx");

        // Row 0 has the long value; rows 1..199 have short values.
        // The previous implementation hardcoded firstRow = lastRow - 100, so row 0 was outside the scan range.
        List<Item> data = new ArrayList<>();
        data.add(new Item(LONG_VALUE));
        for (int i = 1; i < 200; i++) {
            data.add(new Item(SHORT_VALUE));
        }

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            int width = sheet.getColumnWidth(0);
            assertThat(width).as("column width must reflect row 0 even when last 100 rows are short").isGreaterThan(20 * 256);
        }
    }

    @Test
    void ssmlDefaultDoesNotApplyAutoSize() throws Exception {
        File file = new File(tempDir, "as-ssml.xlsx");

        List<Item> data = new ArrayList<>();
        data.add(new Item(LONG_VALUE));
        for (int i = 1; i < 10; i++) {
            data.add(new Item(SHORT_VALUE));
        }

        // Default streaming mode (SSMLSheetWriter) emits XML directly, so autoSize is logged and skipped.
        new SpreadsheetMapper().writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            int width = sheet.getColumnWidth(0);
            // If autoSize were applied here, the long value would push width well above 20*256.
            // Default streaming leaves the column at the workbook default — pin the upper bound.
            assertThat(width).as("SSML should not apply autoSize — long content must not widen the column").isLessThan(20 * 256);
        }
    }

    @Test
    void samplingCapturesOutlierAtSamplePoint() throws Exception {
        File file = new File(tempDir, "as-sample-hit.xlsx");

        // Default origin: header at row 0, data starts at row 1. Sampling beyond the first 100
        // data rows hits relative rows 100, 200, ... — i.e., absolute rows 101, 201, ...
        // Place the long value at list index 200 (absolute row 201) so it lands on a sample point.
        List<Item> data = new ArrayList<>();
        for (int i = 0; i < 300; i++) {
            data.add(new Item(i == 200 ? LONG_VALUE : SHORT_VALUE));
        }

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            int width = wb.getSheetAt(0).getColumnWidth(0);
            assertThat(width).as("sampling must capture the outlier landing on a sample point").isGreaterThan(20 * 256);
        }
    }

    @Test
    void multipleColumnsTrackedIndependently() throws Exception {
        File file = new File(tempDir, "as-multi-col.xlsx");

        // Different long values in different columns at different rows. Both within the
        // first-100 full-scan range so sampling doesn't interfere with the assertion.
        List<Pair> data = new ArrayList<>();
        data.add(new Pair(LONG_VALUE, SHORT_VALUE));   // wide for col 0 only
        data.add(new Pair(SHORT_VALUE, LONG_VALUE));   // wide for col 1 only
        for (int i = 2; i < 50; i++) {
            data.add(new Pair(SHORT_VALUE, SHORT_VALUE));
        }

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        mapper.writeValue(file, data, Pair.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getColumnWidth(0)).as("col 0 must reflect its own long value").isGreaterThan(20 * 256);
            assertThat(sheet.getColumnWidth(1)).as("col 1 must reflect its own long value").isGreaterThan(20 * 256);
        }
    }

    @Test
    void samplingIsRelativeToDataStart() throws Exception {
        File file = new File(tempDir, "as-high-origin.xlsx");

        // Origin "A105" puts header at row 104, data start at row 105. The first data cell sits
        // at absolute row 105 — outside both the absolute first-100 range and the modulo-100 grid
        // (105 % 100 != 0). Only data-relative sampling sees this row as relRow 0 → fully scanned.
        List<Item> data = new ArrayList<>();
        data.add(new Item(LONG_VALUE));
        for (int i = 1; i < 50; i++) {
            data.add(new Item(SHORT_VALUE));
        }

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .origin("A105")
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            int width = wb.getSheetAt(0).getColumnWidth(0);
            assertThat(width).as("data-relative sampling must measure the first data row regardless of high origin")
                    .isGreaterThan(20 * 256);
        }
    }

    @Test
    void shortDataYieldsNarrowColumn() throws Exception {
        File file = new File(tempDir, "as-baseline.xlsx");

        // All-short baseline: confirms the wide column in the other tests is caused by the long row,
        // not by some incidental floor (header text, min-width, etc.).
        List<Item> data = new ArrayList<>();
        for (int i = 0; i < 200; i++) {
            data.add(new Item(SHORT_VALUE));
        }

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        mapper.writeValue(file, data, Item.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            int width = sheet.getColumnWidth(0);
            // Header "Name" (4 chars) is the widest content. autoSize should be small.
            assertThat(width).as("baseline width with only short content").isLessThan(20 * 256);
        }
    }
}
