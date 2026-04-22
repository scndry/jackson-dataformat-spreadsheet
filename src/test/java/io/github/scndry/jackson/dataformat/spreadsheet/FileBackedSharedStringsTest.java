package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Tests for FILE_BACKED_SHARED_STRINGS feature using H2 MVStore.
 */
class FileBackedSharedStringsTest {

    @TempDir
    Path tempDir;

    @DataGrid
    static class StringRow {
        public String a;
        public String b;
        public String c;

        public StringRow() {}
        public StringRow(String a, String b, String c) {
            this.a = a;
            this.b = b;
            this.c = c;
        }
    }

    @Test
    void readWithFileBackedSharedStrings() throws Exception {
        // Create XLSX with shared string table (XSSFWorkbook uses SST)
        File file = tempDir.resolve("sst.xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("a");
            header.createCell(1).setCellValue("b");
            header.createCell(2).setCellValue("c");
            for (int i = 0; i < 100; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue("alpha-" + i);
                row.createCell(1).setCellValue("beta-" + i);
                row.createCell(2).setCellValue("gamma-" + i);
            }
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SheetParser.Feature.FILE_BACKED_SHARED_STRINGS)
                .build();

        List<StringRow> rows = mapper.readValues(file, StringRow.class);
        assertThat(rows).hasSize(100);
        assertThat(rows.get(0).a).isEqualTo("alpha-0");
        assertThat(rows.get(0).b).isEqualTo("beta-0");
        assertThat(rows.get(0).c).isEqualTo("gamma-0");
        assertThat(rows.get(99).a).isEqualTo("alpha-99");
        assertThat(rows.get(99).c).isEqualTo("gamma-99");
    }

    @Test
    void readWithFileBackedVsInMemory_sameResult() throws Exception {
        // Both strategies should produce identical results
        File file = tempDir.resolve("compare.xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("a");
            header.createCell(1).setCellValue("b");
            header.createCell(2).setCellValue("c");
            for (int i = 0; i < 50; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue("value-" + i);
                row.createCell(1).setCellValue("shared"); // repeated string
                row.createCell(2).setCellValue("data-" + (i % 10)); // partially repeated
            }
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetMapper inMemoryMapper = new SpreadsheetMapper();
        SpreadsheetMapper fileBackedMapper = SpreadsheetMapper.builder()
                .enable(SheetParser.Feature.FILE_BACKED_SHARED_STRINGS)
                .build();

        List<StringRow> inMemory = inMemoryMapper.readValues(file, StringRow.class);
        List<StringRow> fileBacked = fileBackedMapper.readValues(file, StringRow.class);

        assertThat(fileBacked).hasSize(inMemory.size());
        for (int i = 0; i < inMemory.size(); i++) {
            assertThat(fileBacked.get(i).a).isEqualTo(inMemory.get(i).a);
            assertThat(fileBacked.get(i).b).isEqualTo(inMemory.get(i).b);
            assertThat(fileBacked.get(i).c).isEqualTo(inMemory.get(i).c);
        }
    }

    @Test
    void tempFileCleanedUpAfterRead() throws Exception {
        File file = tempDir.resolve("cleanup.xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("a");
            header.createCell(1).setCellValue("b");
            header.createCell(2).setCellValue("c");
            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue("x");
            row.createCell(1).setCellValue("y");
            row.createCell(2).setCellValue("z");
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SheetParser.Feature.FILE_BACKED_SHARED_STRINGS)
                .build();

        // Read and close — temp file should be cleaned up
        List<StringRow> rows = mapper.readValues(file, StringRow.class);
        assertThat(rows).hasSize(1);

        // Verify no jackson-sst-*.mv files remain in temp dir
        File tempDir = new File(System.getProperty("java.io.tmpdir"));
        File[] sstFiles = tempDir.listFiles((dir, name) ->
                name.startsWith("jackson-sst-") && name.endsWith(".mv"));
        assertThat(sstFiles == null || sstFiles.length == 0)
                .as("Temp MVStore files should be cleaned up")
                .isTrue();
    }

    @Test
    void readWithEncryptedFileBacked() throws Exception {
        File file = tempDir.resolve("encrypted.xlsx").toFile();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("a");
            header.createCell(1).setCellValue("b");
            header.createCell(2).setCellValue("c");
            for (int i = 0; i < 100; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue("alpha-" + i);
                row.createCell(1).setCellValue("shared");
                row.createCell(2).setCellValue("gamma-" + (i % 10));
            }
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
        }

        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SheetParser.Feature.FILE_BACKED_SHARED_STRINGS)
                .enable(SheetParser.Feature.ENCRYPT_FILE_BACKED_STORE)
                .build();

        List<StringRow> rows = mapper.readValues(file, StringRow.class);
        assertThat(rows).hasSize(100);
        assertThat(rows.get(0).a).isEqualTo("alpha-0");
        assertThat(rows.get(0).b).isEqualTo("shared");
        assertThat(rows.get(99).c).isEqualTo("gamma-9");
    }
}
