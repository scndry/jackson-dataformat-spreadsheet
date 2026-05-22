package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.annotation.OptBoolean;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import static org.assertj.core.api.Assertions.assertThatThrownBy;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for merge column feature via @DataColumn(merge = TRUE).
 */
class MergeTest {

    SpreadsheetMapper mapper;
    @TempDir Path tempDir;

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new, SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Outer {
        @DataColumn(value = "A", merge = OptBoolean.TRUE) int a;
        List<Inner> inners;
        @DataColumn(value = "E", merge = OptBoolean.TRUE) int e;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Inner {
        @DataColumn("B") int b;
        @DataColumn("C") int c;
    }

    @Test
    void mergeColumnsAreApplied() throws Exception {
        File file = tempFile("merge.xlsx");

        Outer value = new Outer(1, Arrays.asList(
                new Inner(2, 3),
                new Inner(4, 5)),
                10);

        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            // A column (col 0): merged across 2 inner rows
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastColumn()).isEqualTo(0);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(2);
            });

            // E column (col 3): merged across 2 inner rows
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(3);
                assertThat(r.getLastColumn()).isEqualTo(3);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(2);
            });
        }
    }

    @Test
    void noMergeWithoutAnnotation() throws Exception {
        File file = tempFile("no-merge.xlsx");

        Outer value = new Outer(1, Arrays.asList(
                new Inner(2, 3),
                new Inner(4, 5)),
                10);

        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            // B (col 1) and C (col 2) should NOT be merged
            assertThat(merged).noneSatisfy(r ->
                    assertThat(r.getFirstColumn()).isIn(1, 2));
        }
    }

    @Test
    void singleInnerNoMerge() throws Exception {
        File file = tempFile("single-inner.xlsx");

        Outer value = new Outer(1, Arrays.asList(new Inner(2, 3)), 10);
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            // size=1 → no merge needed
            assertThat(sheet.getMergedRegions()).isEmpty();
        }
    }

    @Test
    void mergeExceedsSXSSFWindowSize_autoAdjusted() throws Exception {
        // SXSSFWorkbook default window = 100 rows.
        // Library auto-adjusts window when array size exceeds it.
        File file = tempFile("large-merge-auto.xlsx");

        int innerCount = 200;
        Inner[] inners = new Inner[innerCount];
        for (int i = 0; i < innerCount; i++) {
            inners[i] = new Inner(i, i * 10);
        }

        Outer value = new Outer(1, Arrays.asList(inners), 99);
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(innerCount);
            });

            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(3);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(innerCount);
            });

            assertThat((int) sheet.getRow(1).getCell(1).getNumericCellValue()).isEqualTo(0);
            assertThat((int) sheet.getRow(200).getCell(1).getNumericCellValue()).isEqualTo(199);
            assertThat((int) sheet.getRow(1).getCell(3).getNumericCellValue()).isEqualTo(99);
        }
    }

    @Test
    void mergeExceedsSXSSFWindowSize_largerWindow() throws Exception {
        // Workaround: increase SXSSF window via WorkbookProvider
        File file = tempFile("large-merge.xlsx");
        SpreadsheetMapper largeWindowMapper = new SpreadsheetMapper(
                new SpreadsheetFactory(() -> new SXSSFWorkbook(300),
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS));

        int innerCount = 200;
        Inner[] inners = new Inner[innerCount];
        for (int i = 0; i < innerCount; i++) {
            inners[i] = new Inner(i, i * 10);
        }

        Outer value = new Outer(1, Arrays.asList(inners), 99);
        largeWindowMapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(innerCount);
            });

            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(3);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(innerCount);
            });

            assertThat((int) sheet.getRow(1).getCell(1).getNumericCellValue()).isEqualTo(0);
            assertThat((int) sheet.getRow(200).getCell(1).getNumericCellValue()).isEqualTo(199);
        }
    }

    @Test
    void mergeExceedsSXSSFWindowSize_dynamicWindow() throws Exception {
        // POC: SXSSFSheet.setRandomAccessWindowSize() is public.
        // Dynamically increase window before writing, then restore.
        File file = tempFile("dynamic-merge.xlsx");

        int innerCount = 200;
        Inner[] inners = new Inner[innerCount];
        for (int i = 0; i < innerCount; i++) {
            inners[i] = new Inner(i, i * 10);
        }

        Outer value = new Outer(1, Arrays.asList(inners), 99);

        try (SXSSFWorkbook wb = new SXSSFWorkbook()) { // default window 100
            SXSSFSheet sheet = wb.createSheet("test");

            // Dynamically increase window before write
            sheet.setRandomAccessWindowSize(innerCount + 10);

            mapper.writeValue(sheet, value);

            // Restore default window for subsequent rows
            sheet.setRandomAccessWindowSize(SXSSFWorkbook.DEFAULT_WINDOW_SIZE);

            try (java.io.OutputStream os = new java.io.FileOutputStream(file)) {
                wb.write(os);
            }
            wb.dispose();
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(innerCount);
            });

            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(3);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(innerCount);
            });

            // Data integrity
            assertThat((int) sheet.getRow(1).getCell(1).getNumericCellValue()).isEqualTo(0);
            assertThat((int) sheet.getRow(200).getCell(1).getNumericCellValue()).isEqualTo(199);
            assertThat((int) sheet.getRow(1).getCell(3).getNumericCellValue()).isEqualTo(99);
        }
    }

    @Test
    void emptyInnerList() throws Exception {
        File file = tempFile("empty-inner.xlsx");

        Outer value = new Outer(1, Arrays.asList(), 99);
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getMergedRegions()).isEmpty();
            // a and e should still be written on the same row
            assertThat((int) sheet.getRow(1).getCell(0).getNumericCellValue()).isEqualTo(1);
            assertThat((int) sheet.getRow(1).getCell(3).getNumericCellValue()).isEqualTo(99);
        }
    }

    @Test
    void multipleOutersWithLargeInnerList() throws Exception {
        // Window restored after first Outer, second Outer should work too
        File file = tempFile("multi-outer.xlsx");

        int innerCount = 150; // exceeds default window 100
        Inner[] inners = new Inner[innerCount];
        for (int i = 0; i < innerCount; i++) {
            inners[i] = new Inner(i, i * 10);
        }

        List<Outer> values = Arrays.asList(
                new Outer(1, Arrays.asList(inners), 10),
                new Outer(2, Arrays.asList(inners), 20));
        mapper.writeValue(file, values, Outer.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            // First Outer: rows 1-150
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getFirstRow()).isEqualTo(1);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(innerCount);
            });

            // Second Outer: rows 151-300
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getFirstRow()).isEqualTo(1 + innerCount);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(innerCount);
            });

            // Data integrity: second Outer's values
            int secondOuterRow = 1 + innerCount;
            assertThat((int) sheet.getRow(secondOuterRow).getCell(0)
                    .getNumericCellValue()).isEqualTo(2);
            assertThat((int) sheet.getRow(secondOuterRow).getCell(3)
                    .getNumericCellValue()).isEqualTo(20);
        }
    }

    // -- Nested list inside nested list --

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class DeepOuter {
        @DataColumn(value = "ID", merge = OptBoolean.TRUE) int id;
        List<MidLevel> items;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidLevel {
        @DataColumn(value = "Group", merge = OptBoolean.TRUE) String group;
        List<Inner> details;
    }

    @Test
    void nestedListInsideNestedList() throws Exception {
        File file = tempFile("deep-nested.xlsx");

        DeepOuter value = new DeepOuter(1, Arrays.asList(
                new MidLevel("X", Arrays.asList(new Inner(1, 2), new Inner(3, 4))),
                new MidLevel("Y", Arrays.asList(new Inner(5, 6)))));
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            // ID column: merged across all 3 rows (2 + 1)
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(3);
            });

            // Group "X": merged across 2 rows
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(2);
            });
        }
    }

    @Test
    void nestedListWithLargeInnerDetails() throws Exception {
        // Inner details array (150) exceeds default window (100).
        // ensureRowWindow triggers for details, not for items.
        File file = tempFile("deep-large.xlsx");

        int detailCount = 150;
        Inner[] details = new Inner[detailCount];
        for (int i = 0; i < detailCount; i++) {
            details[i] = new Inner(i, i * 10);
        }

        DeepOuter value = new DeepOuter(1, Arrays.asList(
                new MidLevel("X", Arrays.asList(details)),
                new MidLevel("Y", Arrays.asList(new Inner(99, 99)))));
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            // ID: merged across all rows (150 + 1 = 151)
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(detailCount + 1);
            });

            // Group "X": merged across 150 rows
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(1);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(detailCount);
            });

            // Data integrity: last detail of X, and Y's single detail
            assertThat((int) sheet.getRow(150).getCell(2).getNumericCellValue()).isEqualTo(149);
            assertThat((int) sheet.getRow(151).getCell(1).getStringCellValue().charAt(0)).isEqualTo('Y');
        }
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class DeepOuterWithTrailing {
        @DataColumn(value = "ID", merge = OptBoolean.TRUE) int id;
        List<MidLevelWithDetails> items;
        @DataColumn(value = "Total", merge = OptBoolean.TRUE) int total; // AFTER list
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class MidLevelWithDetails {
        @DataColumn("Group") String group;
        List<Inner> details;
    }

    @Test
    void deepNestedExceedsWindow_clearErrorMessage() {
        // 2 MidLevels × 80 details = 160 total rows.
        // ensureRowWindow(2) for items → no adjustment (2 < 100).
        // ensureRowWindow(80) for details → no adjustment (80 < 100).
        // But total 160 rows exceeds window 100 → row 1 flushed
        // → writing "total" to row 1 fails with clear message.
        File file = tempFile("deep-exceed.xlsx");

        Inner[] details = new Inner[80];
        for (int i = 0; i < 80; i++) {
            details[i] = new Inner(i, i);
        }

        DeepOuterWithTrailing value = new DeepOuterWithTrailing(1, Arrays.asList(
                new MidLevelWithDetails("A", Arrays.asList(details)),
                new MidLevelWithDetails("B", Arrays.asList(details))),
                99);

        assertThatThrownBy(() -> mapper.writeValue(file, value))
                .isInstanceOf(Exception.class)
                .hasMessageContaining("SXSSF row window")
                .hasMessageContaining("WorkbookProvider");
    }

    // -- Class-level merge default --

    @Data @NoArgsConstructor @AllArgsConstructor
    @DataGrid(mergeColumn = OptBoolean.TRUE)
    static class ClassLevelMerge {
        int id;
        List<Inner> items;
        int total;
    }

    @Test
    void classLevelMergeDefault() throws Exception {
        File file = tempFile("class-merge.xlsx");

        ClassLevelMerge value = new ClassLevelMerge(1, Arrays.asList(
                new Inner(2, 3), new Inner(4, 5)), 10);
        mapper.writeValue(file, value);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            // Both id (col 0) and total (col 3) merged via class-level default
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(2);
            });
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(3);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(2);
            });
        }
    }

    @Test
    void windowRestoredAfterLargeWrite() throws Exception {
        // After writing a large nested list, subsequent normal rows
        // should flush normally with the restored window size.
        File file = tempFile("window-restored.xlsx");

        int innerCount = 150;
        Inner[] inners = new Inner[innerCount];
        for (int i = 0; i < innerCount; i++) {
            inners[i] = new Inner(i, i * 10);
        }

        // Write 5 Outers: first has 150 inners, rest have 2 each
        List<Outer> values = Arrays.asList(
                new Outer(1, Arrays.asList(inners), 10),
                new Outer(2, Arrays.asList(new Inner(1, 2), new Inner(3, 4)), 20),
                new Outer(3, Arrays.asList(new Inner(5, 6), new Inner(7, 8)), 30),
                new Outer(4, Arrays.asList(new Inner(9, 10), new Inner(11, 12)), 40),
                new Outer(5, Arrays.asList(new Inner(13, 14), new Inner(15, 16)), 50));
        mapper.writeValue(file, values, Outer.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);

            // Last Outer's data should be intact
            int lastOuterRow = 1 + innerCount + 2 + 2 + 2; // 157
            assertThat((int) sheet.getRow(lastOuterRow).getCell(0)
                    .getNumericCellValue()).isEqualTo(5);
            assertThat((int) sheet.getRow(lastOuterRow).getCell(3)
                    .getNumericCellValue()).isEqualTo(50);
        }
    }

    private File tempFile(String name) {
        return tempDir.resolve(name).toFile();
    }

    // -- Sibling nested lists reproduction (mergeRange off-by-one bug) --

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Item {
        @DataColumn("ItemB") int b;
        @DataColumn("ItemC") int c;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Note {
        @DataColumn("NoteB") int b;
        @DataColumn("NoteC") int c;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class TwoListsOuter {
        @DataColumn(value = "ID", merge = OptBoolean.TRUE) int id;
        List<Item> items;
        List<Note> notes;
        @DataColumn(value = "Total", merge = OptBoolean.TRUE) int total;
    }

    @Test
    void siblingNestedLists_mergeRangeMustNotOverflow() throws Exception {
        File file = new File("/tmp/jackson-spreadsheet-multi-list-bug.xlsx");
        if (file.exists() && !file.delete()) {
            throw new IllegalStateException("Failed to delete " + file);
        }

        // Two outers, each with items[2] + notes[2] (same length M=N=2)
        TwoListsOuter outer1 = new TwoListsOuter(
                1,
                Arrays.asList(new Item(11, 12), new Item(13, 14)),
                Arrays.asList(new Note(15, 16), new Note(17, 18)),
                99);
        TwoListsOuter outer2 = new TwoListsOuter(
                2,
                Arrays.asList(new Item(21, 22), new Item(23, 24)),
                Arrays.asList(new Note(25, 26), new Note(27, 28)),
                199);

        mapper.writeValue(file, Arrays.asList(outer1, outer2), TwoListsOuter.class);

        System.out.println("[Reproduction] Written to: " + file.getAbsolutePath());

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            System.out.println("[Reproduction] Merged regions:");
            for (CellRangeAddress r : merged) {
                System.out.println("  " + r);
            }

            // Expected: each outer's id (col 0) mergeRange spans max(M,N) = 2 rows.
            // Bug: spans M+N-1 = 3 rows → overlaps the next outer's first row.
            assertThat(merged)
                    .filteredOn(r -> r.getFirstColumn() == 0)
                    .allSatisfy(r -> {
                        int rowSpan = r.getLastRow() - r.getFirstRow() + 1;
                        assertThat(rowSpan)
                                .as("id mergeRange row span (expected max(M,N)=2, "
                                        + "bug spans M+N-1=3)")
                                .isEqualTo(2);
                    });
        }
    }

    @Test
    void siblingNestedLists_mergeRangeMustNotOverflow_ssmlMode() throws Exception {
        // SSML emits XML directly, so the bug surfaces as overlapping
        // mergedRegions instead of an immediate throw (POI mode).
        File file = new File("/tmp/jackson-spreadsheet-multi-list-bug-ssml.xlsx");
        if (file.exists() && !file.delete()) {
            throw new IllegalStateException("Failed to delete " + file);
        }

        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();

        TwoListsOuter outer1 = new TwoListsOuter(
                1,
                Arrays.asList(new Item(11, 12), new Item(13, 14)),
                Arrays.asList(new Note(15, 16), new Note(17, 18)),
                99);
        TwoListsOuter outer2 = new TwoListsOuter(
                2,
                Arrays.asList(new Item(21, 22), new Item(23, 24)),
                Arrays.asList(new Note(25, 26), new Note(27, 28)),
                199);

        ssmlMapper.writeValue(file, Arrays.asList(outer1, outer2), TwoListsOuter.class);

        System.out.println("[SSML Reproduction] Written to: " + file.getAbsolutePath());

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            System.out.println("[SSML Reproduction] Merged regions:");
            for (CellRangeAddress r : merged) {
                System.out.println("  " + r);
            }

            assertThat(merged)
                    .filteredOn(r -> r.getFirstColumn() == 0)
                    .allSatisfy(r -> {
                        int rowSpan = r.getLastRow() - r.getFirstRow() + 1;
                        assertThat(rowSpan)
                                .as("id mergeRange row span (expected max(M,N)=2, "
                                        + "bug spans M+N-1=3)")
                                .isEqualTo(2);
                    });
        }
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class TwoListsWithMiddleField {
        @DataColumn(value = "ID", merge = OptBoolean.TRUE) int id;
        List<Item> items;
        @DataColumn(value = "Mid", merge = OptBoolean.TRUE) int middle;   // between lists
        List<Note> notes;
        @DataColumn(value = "Total", merge = OptBoolean.TRUE) int total;
    }

    @Test
    void siblingNestedLists_middleField_poiMode() throws Exception {
        File file = tempFile("middle-field-poi.xlsx");

        TwoListsWithMiddleField outer1 = new TwoListsWithMiddleField(
                1,
                Arrays.asList(new Item(11, 12), new Item(13, 14)),
                50,
                Arrays.asList(new Note(15, 16), new Note(17, 18)),
                99);
        TwoListsWithMiddleField outer2 = new TwoListsWithMiddleField(
                2,
                Arrays.asList(new Item(21, 22), new Item(23, 24)),
                150,
                Arrays.asList(new Note(25, 26), new Note(27, 28)),
                199);

        mapper.writeValue(file, Arrays.asList(outer1, outer2),
                TwoListsWithMiddleField.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();
            assertThat(merged)
                    .filteredOn(r -> r.getFirstColumn() == 0)
                    .allSatisfy(r -> {
                        int rowSpan = r.getLastRow() - r.getFirstRow() + 1;
                        assertThat(rowSpan).isEqualTo(2);
                    });
        }
    }

    @Test
    void siblingNestedLists_middleField_ssmlMode() throws Exception {
        File file = new File("/tmp/jackson-spreadsheet-multi-list-middle-bug-ssml.xlsx");
        if (file.exists() && !file.delete()) {
            throw new IllegalStateException("Failed to delete " + file);
        }

        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();

        TwoListsWithMiddleField outer1 = new TwoListsWithMiddleField(
                1,
                Arrays.asList(new Item(11, 12), new Item(13, 14)),
                50,
                Arrays.asList(new Note(15, 16), new Note(17, 18)),
                99);
        TwoListsWithMiddleField outer2 = new TwoListsWithMiddleField(
                2,
                Arrays.asList(new Item(21, 22), new Item(23, 24)),
                150,
                Arrays.asList(new Note(25, 26), new Note(27, 28)),
                199);

        ssmlMapper.writeValue(file, Arrays.asList(outer1, outer2),
                TwoListsWithMiddleField.class);

        System.out.println("[SSML Middle] Written to: " + file.getAbsolutePath());

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            System.out.println("[SSML Middle] Merged regions:");
            for (CellRangeAddress r : merged) {
                System.out.println("  " + r);
            }

            assertThat(merged)
                    .filteredOn(r -> r.getFirstColumn() == 0)
                    .allSatisfy(r -> {
                        int rowSpan = r.getLastRow() - r.getFirstRow() + 1;
                        assertThat(rowSpan)
                                .as("id mergeRange row span (expected max(M,N)=2)")
                                .isEqualTo(2);
                    });
        }
    }

    @Test
    void siblingNestedLists_lengthMismatch_throws() {
        File file = tempFile("length-mismatch.xlsx");

        TwoListsOuter outer = new TwoListsOuter(
                1,
                Arrays.asList(new Item(11, 12), new Item(13, 14)),
                Arrays.asList(new Note(15, 16), new Note(17, 18), new Note(19, 20)),
                99);

        assertThatThrownBy(() -> mapper.writeValue(file, outer))
                .rootCause()
                .hasMessageContaining("Sibling nested List<T> fields")
                .hasMessageContaining("Previous list size: 2")
                .hasMessageContaining("current list size: 3");
    }

    @Test
    void siblingNestedLists_lengthMismatch_throws_ssmlMode() {
        File file = tempFile("length-mismatch-ssml.xlsx");
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();

        TwoListsOuter outer = new TwoListsOuter(
                1,
                Arrays.asList(new Item(11, 12), new Item(13, 14)),
                Arrays.asList(new Note(15, 16), new Note(17, 18), new Note(19, 20)),
                99);

        assertThatThrownBy(() -> ssmlMapper.writeValue(file, outer))
                .rootCause()
                .hasMessageContaining("Sibling nested List<T> fields")
                .hasMessageContaining("Previous list size: 2")
                .hasMessageContaining("current list size: 3");
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class Detail {
        @DataColumn("DetailX") int x;
        @DataColumn("DetailY") int y;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    static class WithNested {
        @DataColumn("Yi") int yi;
        List<Detail> zs;
        @DataColumn("Yj") int yj;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class SiblingWithNestedOuter {
        @DataColumn(value = "ID", merge = OptBoolean.TRUE) int id;
        List<Item> xs;            // flat element
        List<WithNested> ys;      // nested element (contains List<Detail>)
    }

    @Test
    void siblingNestedLists_innerNested_throws() {
        // xs row span = 2 (flat).
        // ys = 1 WithNested with 3 details → row span = 3.
        // Length mismatch caught by sibling length check at ys close.
        File file = tempFile("sibling-with-nested.xlsx");

        SiblingWithNestedOuter outer = new SiblingWithNestedOuter(
                1,
                Arrays.asList(new Item(11, 12), new Item(13, 14)),
                Arrays.asList(new WithNested(20,
                        Arrays.asList(new Detail(21, 22), new Detail(23, 24), new Detail(25, 26)),
                        29)));

        assertThatThrownBy(() -> mapper.writeValue(file, outer))
                .rootCause()
                .hasMessageContaining("Sibling nested List<T> fields")
                .hasMessageContaining("Previous list size: 2")
                .hasMessageContaining("current list size: 3");
    }

    @Test
    void siblingNestedLists_innerNested_throws_ssmlMode() {
        File file = tempFile("sibling-with-nested-ssml.xlsx");
        SpreadsheetMapper ssmlMapper = new SpreadsheetMapper();

        SiblingWithNestedOuter outer = new SiblingWithNestedOuter(
                1,
                Arrays.asList(new Item(11, 12), new Item(13, 14)),
                Arrays.asList(new WithNested(20,
                        Arrays.asList(new Detail(21, 22), new Detail(23, 24), new Detail(25, 26)),
                        29)));

        assertThatThrownBy(() -> ssmlMapper.writeValue(file, outer))
                .rootCause()
                .hasMessageContaining("Sibling nested List<T> fields")
                .hasMessageContaining("Previous list size: 2")
                .hasMessageContaining("current list size: 3");
    }
}
