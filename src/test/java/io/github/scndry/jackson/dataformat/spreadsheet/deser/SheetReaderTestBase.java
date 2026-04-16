package io.github.scndry.jackson.dataformat.spreadsheet.deser;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.io.TempDir;

import static org.assertj.core.api.Assertions.*;

/**
 * Base test for {@link SheetReader} implementations.
 * Verifies the token sequence contract for a 2-row, 2-column
 * headless spreadsheet.
 */
public class SheetReaderTestBase {

    @TempDir protected File tempDir;

    /** Strict OOXML fixture — cannot be generated programmatically. */
    protected final InputStream strictSource =
            getClass().getClassLoader().getResourceAsStream(
                    "support/fixture/entries-headless-strict.xlsx");

    protected SheetReader reader;

    /**
     * Creates a 2-row, 2-column headless XLSX file
     * programmatically (Transitional format).
     */
    protected File createHeadlessFixture() throws IOException {
        File file = new File(tempDir, "headless.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            for (int i = 0; i < 2; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue(i * 2 + 1);
                row.createCell(1).setCellValue(i * 2 + 2);
            }
            try (java.io.FileOutputStream fos =
                    new java.io.FileOutputStream(file)) {
                wb.write(fos);
            }
        }
        return file;
    }

    protected void testSheetReader() throws Exception {
        assertNext(SheetToken.SHEET_DATA_START);
        assertThat(reader.getRow()).isNegative();
        assertThat(reader.getCellValue()).isNull();

        for (int i = 0; i < 2; i++) {
            assertNext(SheetToken.ROW_START);
            assertThat(reader.getRow()).isNotNegative();

            for (int j = 0; j < 2; j++) {
                assertNext(SheetToken.CELL_VALUE);
                assertThat(reader.getReference()).isNotNull();
                assertThat(reader.getCellValue()).isNotNull();
            }

            assertNext(SheetToken.ROW_END);
        }

        assertNext(SheetToken.SHEET_DATA_END);
        assertThat(reader).isExhausted();
        assertThatThrownBy(reader::next)
                .isInstanceOf(NoSuchElementException.class);

        reader.close();
        assertThat(reader.isClosed()).isTrue();
    }

    private void assertNext(final SheetToken expected) {
        assertThat(reader).hasNext();
        assertThat(reader.next()).isEqualTo(expected);
    }
}
