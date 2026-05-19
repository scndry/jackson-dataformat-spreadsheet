package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * BigInteger cross-library interop — POI emits numeric cells via
 * {@code Double.toString}, which can produce scientific notation (e.g. "5.0E10")
 * for large values. {@link java.math.BigInteger#BigInteger(String)} rejects
 * scientific notation; the parser routes the raw cell text through
 * {@link java.math.BigDecimal} instead so values written by any tool that
 * lands in a numeric cell are readable as BigInteger.
 */
class BigIntegerInteropTest {

    @TempDir File tempDir;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class Row1 {
        BigInteger big;
    }

    @Test
    void readsBigIntegerFromPoiNumericCellWithScientificNotation() throws Exception {
        File file = new File(tempDir, "biginteger.xlsx");
        BigInteger value = BigInteger.valueOf(50_000L * 1_000_000L); // 5e10 — POI emits "5.0E10"
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue("big");
            sheet.createRow(1).createCell(0).setCellValue(value.doubleValue());
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }

        SpreadsheetMapper mapper = new SpreadsheetMapper();
        List<Row1> rows = mapper.readValues(file, Row1.class);
        assertThat(rows).hasSize(1);
        assertThat(rows.get(0).getBig()).isEqualTo(value);
    }

    @Test
    void readsBigIntegerFromPoiNumericCellWithIntegerForm() throws Exception {
        File file = new File(tempDir, "biginteger-small.xlsx");
        BigInteger value = BigInteger.valueOf(12345L); // POI emits "12345.0"
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue("big");
            sheet.createRow(1).createCell(0).setCellValue(value.doubleValue());
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }

        SpreadsheetMapper mapper = new SpreadsheetMapper();
        List<Row1> rows = mapper.readValues(file, Row1.class);
        assertThat(rows.get(0).getBig()).isEqualTo(value);
    }
}
