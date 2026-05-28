package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.file.Path;
import java.util.Arrays;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import com.fasterxml.jackson.annotation.JsonFormat;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

import static org.assertj.core.api.Assertions.assertThat;

class BigNumberCellTypeTest {

    @TempDir Path tempDir;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NumericDefault {
        @DataColumn("amount") BigDecimal amount;
        @DataColumn("ticks") BigInteger ticks;
    }

    @Test
    void bigDecimalAndBigInteger_writeAsNumericByDefault() throws Exception {
        File file = tempDir.resolve("default.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file, Arrays.asList(new NumericDefault(
                new BigDecimal("1234.56"), BigInteger.valueOf(987))), NumericDefault.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).getRow(1).getCell(0).getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(wb.getSheetAt(0).getRow(1).getCell(0).getNumericCellValue()).isEqualTo(1234.56);
            assertThat(wb.getSheetAt(0).getRow(1).getCell(1).getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(wb.getSheetAt(0).getRow(1).getCell(1).getNumericCellValue()).isEqualTo(987.0);
        }
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class StringOptIn {
        @JsonFormat(shape = JsonFormat.Shape.STRING)
        @DataColumn("amount") BigDecimal amount;
        @JsonFormat(shape = JsonFormat.Shape.STRING)
        @DataColumn("ticks") BigInteger ticks;
    }

    @Test
    void jsonFormatStringOptIn_writesAsStringCell() throws Exception {
        File file = tempDir.resolve("optin.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file, Arrays.asList(new StringOptIn(
                new BigDecimal("1234.56"), BigInteger.valueOf(987))), StringOptIn.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).getRow(1).getCell(0).getCellType()).isEqualTo(CellType.STRING);
            assertThat(wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue()).isEqualTo("1234.56");
            assertThat(wb.getSheetAt(0).getRow(1).getCell(1).getCellType()).isEqualTo(CellType.STRING);
            assertThat(wb.getSheetAt(0).getRow(1).getCell(1).getStringCellValue()).isEqualTo("987");
        }
    }
}
