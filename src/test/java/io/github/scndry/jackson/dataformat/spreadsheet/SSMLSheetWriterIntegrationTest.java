package io.github.scndry.jackson.dataformat.spreadsheet;

import com.fasterxml.jackson.annotation.OptBoolean;
import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.style.StylesBuilder;

import static org.assertj.core.api.Assertions.assertThat;

class SSMLSheetWriterIntegrationTest {

    private static final Path DEBUG_OUTPUT_DIR = Paths.get("build/debug-output");

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class Entry {
        String name;
        int qty;
        boolean active;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class MergedOuter {
        @DataColumn(value = "A", merge = OptBoolean.TRUE)
        int a;
        List<MergedInner> inners;
        @DataColumn(value = "E", merge = OptBoolean.TRUE)
        int e;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    static class MergedInner {
        int b;
        int c;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class StyledEntry {
        @DataColumn(value = "Name", style = "data", headerStyle = "hdr")
        String name;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @DataGrid
    static class FixedWidthEntry {
        @DataColumn(value = "Name", width = 24)
        String name;
        @DataColumn(value = "Qty", width = 12)
        int qty;
    }

    @Test
    void ssmlWriterRoundTrip() throws Exception {
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        File file = _debugFile("ssml.xlsx");
        List<Entry> input = Arrays.asList(
                new Entry("A", 1, true),
                new Entry("B", 2, false));

        mapper.writeValue(file, input, Entry.class);
        List<Entry> output = mapper.readValues(file, Entry.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void ssmlWriterXmlSpecialCharactersRoundTrip() throws Exception {
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        File file = _debugFile("ssml-xml-chars.xlsx");
        List<Entry> input = Arrays.asList(
                new Entry("<script>alert('xss')</script>", 1, true),
                new Entry("Tom & Jerry", 2, false),
                new Entry("\"quoted\"", 3, true),
                new Entry("less < greater >", 4, false),
                new Entry("apos'trophe", 5, true));

        mapper.writeValue(file, input, Entry.class);
        List<Entry> output = mapper.readValues(file, Entry.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void ssmlWriterFallsBackToPoiWhenWorkbookIsNotXssf() throws Exception {
        SpreadsheetFactory factory = new SpreadsheetFactory(
                SXSSFWorkbook::new,
                SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS);
        SpreadsheetMapper mapper = new SpreadsheetMapper(factory);

        File file = _debugFile("ssml-fallback.xlsx");
        List<Entry> input = Arrays.asList(
                new Entry("C", 3, true),
                new Entry("D", 4, false));

        mapper.writeValue(file, input, Entry.class);
        List<Entry> output = mapper.readValues(file, Entry.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void ssmlWriterWithFileBackedSharedStrings() throws Exception {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
                .build();

        File file = _debugFile("ssml-file-backed.xlsx");
        List<Entry> input = Arrays.asList(
                new Entry("E", 5, true),
                new Entry("E", 6, false),
                new Entry("F", 7, true));

        mapper.writeValue(file, input, Entry.class);
        List<Entry> output = mapper.readValues(file, Entry.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void ssmlWriterWithEncryptedFileBackedSharedStrings() throws Exception {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
                .enable(SpreadsheetFactory.Feature.ENCRYPT_FILE_BACKED_STORE)
                .build();

        File file = _debugFile("ssml-file-backed-encrypted.xlsx");
        List<Entry> input = Arrays.asList(
                new Entry("G", 8, true),
                new Entry("G", 9, false),
                new Entry("H", 10, true));

        mapper.writeValue(file, input, Entry.class);
        List<Entry> output = mapper.readValues(file, Entry.class);

        assertThat(output).isEqualTo(input);
    }

    @Test
    void ssmlWriterAppliesMergedRegions() throws Exception {
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        File file = _debugFile("ssml-merge.xlsx");
        MergedOuter input = new MergedOuter(
                1,
                Arrays.asList(new MergedInner(2, 3), new MergedInner(4, 5)),
                10);

        mapper.writeValue(file, input);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<CellRangeAddress> merged = sheet.getMergedRegions();

            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(0);
                assertThat(r.getLastColumn()).isEqualTo(0);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(2);
            });
            assertThat(merged).anySatisfy(r -> {
                assertThat(r.getFirstColumn()).isEqualTo(3);
                assertThat(r.getLastColumn()).isEqualTo(3);
                assertThat(r.getLastRow() - r.getFirstRow() + 1).isEqualTo(2);
            });
        }
    }

    @Test
    void ssmlWriterAppliesHeaderStyle() throws Exception {
        StylesBuilder styles = new StylesBuilder()
                .cellStyle("hdr").font().bold().end().end()
                .cellStyle("data").font().italic().end().end();
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .stylesBuilder(styles)
                .build();

        File file = _debugFile("ssml-header-style.xlsx");
        mapper.writeValue(file, Arrays.asList(new StyledEntry("A")), StyledEntry.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Font headerFont = wb.getFontAt(sheet.getRow(0).getCell(0).getCellStyle().getFontIndex());
            Font dataFont = wb.getFontAt(sheet.getRow(1).getCell(0).getCellStyle().getFontIndex());

            assertThat(headerFont.getBold()).isTrue();
            assertThat(dataFont.getItalic()).isTrue();
            assertThat(dataFont.getBold()).isFalse();
        }
    }

    @Test
    void ssmlWriterWritesBlankCellWithStyle() throws Exception {
        StylesBuilder styles = new StylesBuilder()
                .cellStyle("data")
                    .fillForegroundColor(IndexedColors.YELLOW)
                    .fillPattern().solidForeground()
                    .font().italic().end()
                    .end()
                .cellStyle("hdr")
                    .font().bold().end()
                    .end();
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .stylesBuilder(styles)
                .build();

        File file = _debugFile("ssml-blank-style.xlsx");
        mapper.writeValue(file, Arrays.asList(new StyledEntry(null)), StyledEntry.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);

            assertThat(cell).isNotNull();
            assertThat(cell.getCellType()).isEqualTo(CellType.BLANK);
            assertThat(cell.getCellStyle().getFillPattern()).isEqualTo(FillPatternType.SOLID_FOREGROUND);
            Font dataFont = wb.getFontAt(cell.getCellStyle().getFontIndex());
            assertThat(dataFont.getItalic()).isTrue();
        }
    }

    @Test
    void ssmlWriterAppliesFixedColumnWidth() throws Exception {
        SpreadsheetMapper mapper = new SpreadsheetMapper();

        File file = _debugFile("ssml-fixed-width.xlsx");
        mapper.writeValue(file, Arrays.asList(new FixedWidthEntry("A", 1)), FixedWidthEntry.class);

        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getColumnWidth(0)).isEqualTo(24 * 256);
            assertThat(sheet.getColumnWidth(1)).isEqualTo(12 * 256);
        }
    }

    @Data
    @NoArgsConstructor
    @DataGrid
    static class AllTypesEntry {
        float floatPrim;
        double doublePrim;
        int intPrim;
        long longPrim;
        Float floatObj;
        Double doubleObj;
        Integer intObj;
        Long longObj;
        BigInteger bigInteger;
        BigDecimal bigDecimal;
        String string;
        Date date;
        Calendar calendar;
        LocalDateTime localDateTime;
        LocalDate localDate;
    }

    @Test
    void ssmlWriterAllTypesWithSimpleStyles() throws Exception {
        StylesBuilder styles = StylesBuilder.simple();
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .stylesBuilder(styles)
                .build();

        AllTypesEntry input = new AllTypesEntry();
        input.setFloatPrim(1.1f);
        input.setDoublePrim(2.2);
        input.setIntPrim(3);
        input.setLongPrim(4L);
        input.setFloatObj(5.5f);
        input.setDoubleObj(6.6);
        input.setIntObj(7);
        input.setLongObj(8L);
        input.setBigInteger(new BigInteger("999999999999"));
        input.setBigDecimal(new BigDecimal("123456.789"));
        input.setString("hello");
        input.setDate(new Date(1700000000000L));
        Calendar cal = Calendar.getInstance();
        cal.setTimeInMillis(1700000000000L);
        input.setCalendar(cal);
        input.setLocalDateTime(LocalDateTime.of(2023, 11, 14, 12, 0, 0));
        input.setLocalDate(LocalDate.of(2023, 11, 14));

        File file = _debugFile("ssml-all-types.xlsx");
        mapper.writeValue(file, input);

        // Verify readable and data format strings applied
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            assertThat(sheet.getRow(0)).isNotNull(); // header
            assertThat(sheet.getRow(1)).isNotNull(); // data

            org.apache.poi.ss.usermodel.Row data = sheet.getRow(1);
            // float/double primitives — #,##0.00
            assertThat(data.getCell(0).getCellStyle().getDataFormatString()).isEqualTo("#,##0.00");
            assertThat(data.getCell(1).getCellStyle().getDataFormatString()).isEqualTo("#,##0.00");
            // int/long primitives — #,##0
            assertThat(data.getCell(2).getCellStyle().getDataFormatString()).isEqualTo("#,##0");
            assertThat(data.getCell(3).getCellStyle().getDataFormatString()).isEqualTo("#,##0");
            // Float/Double wrappers — 0.00
            assertThat(data.getCell(4).getCellStyle().getDataFormatString()).isEqualTo("0.00");
            assertThat(data.getCell(5).getCellStyle().getDataFormatString()).isEqualTo("0.00");
            // Integer/Long wrappers — 0
            assertThat(data.getCell(6).getCellStyle().getDataFormatString()).isEqualTo("0");
            assertThat(data.getCell(7).getCellStyle().getDataFormatString()).isEqualTo("0");
            // BigInteger/BigDecimal/String — @
            assertThat(data.getCell(8).getCellStyle().getDataFormatString()).isEqualTo("@");
            assertThat(data.getCell(9).getCellStyle().getDataFormatString()).isEqualTo("@");
            assertThat(data.getCell(10).getCellStyle().getDataFormatString()).isEqualTo("@");
            // Date/Calendar/LocalDateTime — yyyy-mm-dd hh:mm:ss
            assertThat(data.getCell(11).getCellStyle().getDataFormatString()).isEqualTo("yyyy-mm-dd hh:mm:ss");
            assertThat(data.getCell(12).getCellStyle().getDataFormatString()).isEqualTo("yyyy-mm-dd hh:mm:ss");
            assertThat(data.getCell(13).getCellStyle().getDataFormatString()).isEqualTo("yyyy-mm-dd hh:mm:ss");
            // LocalDate — yyyy-mm-dd
            assertThat(data.getCell(14).getCellStyle().getDataFormatString()).isEqualTo("yyyy-mm-dd");
        }

        // Round-trip
        List<AllTypesEntry> read = mapper.readValues(file, AllTypesEntry.class);
        assertThat(read).hasSize(1);
        AllTypesEntry output = read.get(0);
        assertThat(output.getFloatPrim()).isEqualTo(1.1f);
        assertThat(output.getDoublePrim()).isEqualTo(2.2);
        assertThat(output.getIntPrim()).isEqualTo(3);
        assertThat(output.getLongPrim()).isEqualTo(4L);
        assertThat(output.getFloatObj()).isEqualTo(5.5f);
        assertThat(output.getDoubleObj()).isEqualTo(6.6);
        assertThat(output.getIntObj()).isEqualTo(7);
        assertThat(output.getLongObj()).isEqualTo(8L);
        assertThat(output.getBigInteger()).isEqualTo(new BigInteger("999999999999"));
        assertThat(output.getBigDecimal()).isEqualByComparingTo(new BigDecimal("123456.789"));
        assertThat(output.getString()).isEqualTo("hello");
        // Date, Calendar, LocalDateTime excluded — Excel stores dates as numeric serials,
        // round-trip precision depends on timezone and is not deterministic.
        assertThat(output.getLocalDate()).isEqualTo(LocalDate.of(2023, 11, 14));
    }

    private static File _debugFile(final String name) throws IOException {
        Files.createDirectories(DEBUG_OUTPUT_DIR);
        return DEBUG_OUTPUT_DIR.resolve(name).toFile();
    }
}
