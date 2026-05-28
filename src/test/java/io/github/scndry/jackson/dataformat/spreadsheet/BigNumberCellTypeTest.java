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
import org.apache.poi.ss.usermodel.Cell;

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

    @Test
    void bigDecimalPrecisionLoss_silentInDefaultPath() throws Exception {
        final BigDecimal original = new BigDecimal("1.234567890123456789");
        final double expected = original.doubleValue();
        assertThat(new BigDecimal(Double.toString(expected))).isNotEqualTo(original);

        File file = tempDir.resolve("precision-bigdecimal.xlsx").toFile();
        new SpreadsheetMapper().writeValue(file,
                Arrays.asList(new NumericDefault(original, BigInteger.ZERO)), NumericDefault.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Cell cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertThat(cell.getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(cell.getNumericCellValue()).isEqualTo(expected);
        }
    }

    @Test
    void bigIntegerBeyondDoubleRange_silentInDefaultPath() throws Exception {
        final BigInteger original = BigInteger.valueOf(2).pow(60).add(BigInteger.ONE);
        final double expected = original.doubleValue();
        assertThat(java.math.BigDecimal.valueOf(expected).toBigInteger()).isNotEqualTo(original);

        File file = tempDir.resolve("precision-biginteger.xlsx").toFile();
        new SpreadsheetMapper().writeValue(file,
                Arrays.asList(new NumericDefault(BigDecimal.ZERO, original)), NumericDefault.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Cell cell = wb.getSheetAt(0).getRow(1).getCell(1);
            assertThat(cell.getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(cell.getNumericCellValue()).isEqualTo(expected);
        }
    }

    @Test
    void poiUserModel_alsoEmitsNumeric() throws Exception {
        File file = tempDir.resolve("poi-user-model.xlsx").toFile();
        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        poiMapper.writeValue(file, Arrays.asList(new NumericDefault(
                new BigDecimal("1234.56"), BigInteger.valueOf(987))), NumericDefault.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            assertThat(wb.getSheetAt(0).getRow(1).getCell(0).getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(wb.getSheetAt(0).getRow(1).getCell(0).getNumericCellValue()).isEqualTo(1234.56);
            assertThat(wb.getSheetAt(0).getRow(1).getCell(1).getCellType()).isEqualTo(CellType.NUMERIC);
            assertThat(wb.getSheetAt(0).getRow(1).getCell(1).getNumericCellValue()).isEqualTo(987.0);
        }
    }

    // -- Round-trip ----------------------------------------------------

    @Test
    void roundTrip_defaultNumeric_smallScale() throws Exception {
        final NumericDefault original = new NumericDefault(
                new BigDecimal("1234.56"), BigInteger.valueOf(987));
        File file = tempDir.resolve("rt-default-small.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file, Arrays.asList(original), NumericDefault.class);

        java.util.List<NumericDefault> read = mapper.readValues(file, NumericDefault.class);
        assertThat(read).singleElement().satisfies(r -> {
            assertThat(r.getAmount()).isEqualByComparingTo(original.getAmount());
            assertThat(r.getTicks()).isEqualTo(original.getTicks());
        });
    }

    @Test
    void roundTrip_defaultNumeric_precisionLossOnLargeScale() throws Exception {
        final BigDecimal original = new BigDecimal("1.234567890123456789");
        File file = tempDir.resolve("rt-default-large.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file,
                Arrays.asList(new NumericDefault(original, BigInteger.ZERO)), NumericDefault.class);

        java.util.List<NumericDefault> read = mapper.readValues(file, NumericDefault.class);
        BigDecimal roundTripped = read.get(0).getAmount();
        assertThat(roundTripped).isEqualByComparingTo(BigDecimal.valueOf(original.doubleValue()));
        assertThat(roundTripped).isNotEqualByComparingTo(original);
    }

    @Test
    void roundTrip_defaultNumeric_bigIntegerBeyondDoubleRange() throws Exception {
        final BigInteger original = BigInteger.valueOf(2).pow(60).add(BigInteger.ONE);
        File file = tempDir.resolve("rt-default-bigint.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file,
                Arrays.asList(new NumericDefault(BigDecimal.ZERO, original)), NumericDefault.class);

        java.util.List<NumericDefault> read = mapper.readValues(file, NumericDefault.class);
        BigInteger roundTripped = read.get(0).getTicks();
        assertThat(roundTripped).isEqualTo(BigDecimal.valueOf((double) original.doubleValue()).toBigInteger());
        assertThat(roundTripped).isNotEqualTo(original);
    }

    @Test
    void roundTrip_stringOptIn_preservesPrecision() throws Exception {
        final BigDecimal amount = new BigDecimal("1.234567890123456789");
        final BigInteger ticks = BigInteger.valueOf(2).pow(60).add(BigInteger.ONE);
        File file = tempDir.resolve("rt-optin.xlsx").toFile();
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file, Arrays.asList(new StringOptIn(amount, ticks)), StringOptIn.class);

        java.util.List<StringOptIn> read = mapper.readValues(file, StringOptIn.class);
        assertThat(read).singleElement().satisfies(r -> {
            assertThat(r.getAmount()).isEqualByComparingTo(amount);
            assertThat(r.getTicks()).isEqualTo(ticks);
        });
    }

    @Test
    void roundTrip_poiUserModel_defaultNumeric() throws Exception {
        final NumericDefault original = new NumericDefault(
                new BigDecimal("1234.56"), BigInteger.valueOf(987));
        File file = tempDir.resolve("rt-poi.xlsx").toFile();
        SpreadsheetMapper poiMapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL)
                .build();
        poiMapper.writeValue(file, Arrays.asList(original), NumericDefault.class);

        java.util.List<NumericDefault> read = poiMapper.readValues(file, NumericDefault.class);
        assertThat(read).singleElement().satisfies(r -> {
            assertThat(r.getAmount()).isEqualByComparingTo(original.getAmount());
            assertThat(r.getTicks()).isEqualTo(original.getTicks());
        });
    }
}
