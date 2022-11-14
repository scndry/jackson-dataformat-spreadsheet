package io.github.scndry.jackson.dataformat.spreadsheet;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import support.fixture.Entry;

import java.io.File;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

class SchemaGeneratorTest {

    @TempDir
    Path tempDir;
    File out;
    Row row;

    @BeforeEach
    void setUp() throws Exception {
        out = tempDir.resolve("test.xlsx").toFile();
    }

    @Test
    void changeOriginAddress() throws Exception {
        final SpreadsheetMapper mapper = new SpreadsheetMapper().setOrigin("B2");
        mapper.writeValue(out, Entry.VALUE);
        try (XSSFWorkbook workbook = new XSSFWorkbook(out)) {
            final XSSFSheet sheet = workbook.getSheetAt(0);
            row = sheet.getRow(0);
            assertThat(row).isNull();
            row = sheet.getRow(1);
            assertThat(row.getCell(0)).isNull();
            assertCellValue(1, "a");
            assertCellValue(2, "b");
            row = sheet.getRow(2);
            assertThat(row.getCell(0)).isNull();
            assertCellValue(1, 1);
            assertCellValue(2, 2);
            row = sheet.getRow(3);
            assertThat(row).isNull();
        }
    }

    @Test
    void overwriteColumnNames() throws Exception {
        final SpreadsheetMapper mapper = new SpreadsheetMapper()
                .setColumnNameResolver(prop -> prop.getName().toUpperCase());
        mapper.writeValue(out, Entry.VALUE);
        try (XSSFWorkbook workbook = new XSSFWorkbook(out)) {
            final XSSFSheet sheet = workbook.getSheetAt(0);
            row = sheet.getRow(0);
            assertCellValue(0, Entry.Fields.A.toUpperCase());
            assertCellValue(1, Entry.Fields.B.toUpperCase());
        }
    }

    void assertCellValue(final int cellnum, String expected) {
        assertThat(row.getCell(cellnum).getStringCellValue()).isEqualTo(expected);
    }

    void assertCellValue(final int cellnum, double expected) {
        assertThat(row.getCell(cellnum).getNumericCellValue()).isEqualTo(expected);
    }
}
