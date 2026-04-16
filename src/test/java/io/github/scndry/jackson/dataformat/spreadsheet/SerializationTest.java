package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Serialization type coverage for SheetGenerator.
 */
class SerializationTest {

    SpreadsheetMapper mapper;
    @TempDir Path tempDir;

    @BeforeEach
    void setUp() {
        mapper = new SpreadsheetMapper();
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class NumericTypes {
        int intVal;
        long longVal;
        float floatVal;
        double doubleVal;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class BigTypes {
        BigInteger bigInt;
        BigDecimal bigDec;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class MixedRow {
        String text;
        int num;
        boolean flag;
    }

    @Test
    void numericTypes() throws Exception {
        File file = tempFile("numeric.xlsx");
        List<NumericTypes> input = Arrays.asList(
                new NumericTypes(42, 123456789L, 1.5f, 3.14),
                new NumericTypes(-1, 0L, 0.0f, -99.9)
        );

        mapper.writeValue(file, input, NumericTypes.class);
        List<NumericTypes> output = mapper.readValues(file, NumericTypes.class);

        assertThat(output).hasSize(2);
        assertThat(output.get(0).getIntVal()).isEqualTo(42);
        assertThat(output.get(0).getLongVal()).isEqualTo(123456789L);
        assertThat(output.get(0).getDoubleVal()).isEqualTo(3.14);
        assertThat(output.get(1).getIntVal()).isEqualTo(-1);
    }

    @Test
    void bigIntegerAndBigDecimal() throws Exception {
        File file = tempFile("big.xlsx");
        List<BigTypes> input = Arrays.asList(
                new BigTypes(new BigInteger("123456789012345"), new BigDecimal("99999.99"))
        );

        mapper.writeValue(file, input, BigTypes.class);
        List<BigTypes> output = mapper.readValues(file, BigTypes.class);

        assertThat(output).hasSize(1);
        assertThat(output.get(0).getBigInt()).isEqualTo(new BigInteger("123456789012345"));
        assertThat(output.get(0).getBigDec()).isEqualByComparingTo(new BigDecimal("99999.99"));
    }

    @Test
    void emptyString() throws Exception {
        File file = tempFile("empty-str.xlsx");
        mapper.writeValue(file, Arrays.asList(new MixedRow("", 0, false)), MixedRow.class);

        List<MixedRow> output = mapper.readValues(file, MixedRow.class);
        assertThat(output.get(0).getText()).isEqualTo("");
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    static class DateRow {
        LocalDate date;
        LocalDateTime dateTime;
    }

    @Test
    void dateRoundTrip() throws Exception {
        SpreadsheetMapper dateMapper = SpreadsheetMapper.builder().build();
        dateMapper.registerModule(new ExcelDateModule());

        File file = tempFile("date.xlsx");
        List<DateRow> input = Arrays.asList(
                new DateRow(LocalDate.of(2026, 4, 15), LocalDateTime.of(2026, 4, 15, 10, 30, 0))
        );

        dateMapper.writeValue(file, input, DateRow.class);
        List<DateRow> output = dateMapper.readValues(file, DateRow.class);

        assertThat(output).hasSize(1);
        assertThat(output.get(0).getDate()).isEqualTo(LocalDate.of(2026, 4, 15));
        assertThat(output.get(0).getDateTime()).isEqualTo(LocalDateTime.of(2026, 4, 15, 10, 30, 0));
    }

    private File tempFile(String name) {
        return tempDir.resolve(name).toFile();
    }
}
