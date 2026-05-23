package io.github.scndry.jackson.dataformat.spreadsheet.internal;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

/**
 * Cell-type overhead probe — eight 10-column schemas, one per common
 * Java type, against identical row count. Surfaces per-cell read/write
 * cost differences (FloatingDecimal vs SST lookup vs Excel-serial date
 * conversion vs BigDecimal parse) without nested-list overhead.
 *
 * <p>Fixtures are written via POI directly for the read side
 * (`setCellValue` per type) and via the library mapper for the write
 * side. Date/datetime cells carry a data format style so the library's
 * {@code ExcelDateModule} recognises them.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class CellTypeOverheadBenchmark {

    @Param({"10000"})
    int rowCount;

    SpreadsheetMapper mapper;

    File intFile, longFile, doubleFile, booleanFile, stringFile,
            bigDecimalFile, localDateFile, localDateTimeFile;

    List<IntRow> intData;
    List<LongRow> longData;
    List<DoubleRow> doubleData;
    List<BooleanRow> booleanData;
    List<StringRow> stringData;
    List<BigDecimalRow> bigDecimalData;
    List<LocalDateRow> localDateData;
    List<LocalDateTimeRow> localDateTimeData;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class IntRow {
        @DataColumn private int c0; @DataColumn private int c1;
        @DataColumn private int c2; @DataColumn private int c3;
        @DataColumn private int c4; @DataColumn private int c5;
        @DataColumn private int c6; @DataColumn private int c7;
        @DataColumn private int c8; @DataColumn private int c9;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class LongRow {
        @DataColumn private long c0; @DataColumn private long c1;
        @DataColumn private long c2; @DataColumn private long c3;
        @DataColumn private long c4; @DataColumn private long c5;
        @DataColumn private long c6; @DataColumn private long c7;
        @DataColumn private long c8; @DataColumn private long c9;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class DoubleRow {
        @DataColumn private double c0; @DataColumn private double c1;
        @DataColumn private double c2; @DataColumn private double c3;
        @DataColumn private double c4; @DataColumn private double c5;
        @DataColumn private double c6; @DataColumn private double c7;
        @DataColumn private double c8; @DataColumn private double c9;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class BooleanRow {
        @DataColumn private boolean c0; @DataColumn private boolean c1;
        @DataColumn private boolean c2; @DataColumn private boolean c3;
        @DataColumn private boolean c4; @DataColumn private boolean c5;
        @DataColumn private boolean c6; @DataColumn private boolean c7;
        @DataColumn private boolean c8; @DataColumn private boolean c9;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class StringRow {
        @DataColumn private String c0; @DataColumn private String c1;
        @DataColumn private String c2; @DataColumn private String c3;
        @DataColumn private String c4; @DataColumn private String c5;
        @DataColumn private String c6; @DataColumn private String c7;
        @DataColumn private String c8; @DataColumn private String c9;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class BigDecimalRow {
        @DataColumn private BigDecimal c0; @DataColumn private BigDecimal c1;
        @DataColumn private BigDecimal c2; @DataColumn private BigDecimal c3;
        @DataColumn private BigDecimal c4; @DataColumn private BigDecimal c5;
        @DataColumn private BigDecimal c6; @DataColumn private BigDecimal c7;
        @DataColumn private BigDecimal c8; @DataColumn private BigDecimal c9;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class LocalDateRow {
        @DataColumn private LocalDate c0; @DataColumn private LocalDate c1;
        @DataColumn private LocalDate c2; @DataColumn private LocalDate c3;
        @DataColumn private LocalDate c4; @DataColumn private LocalDate c5;
        @DataColumn private LocalDate c6; @DataColumn private LocalDate c7;
        @DataColumn private LocalDate c8; @DataColumn private LocalDate c9;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class LocalDateTimeRow {
        @DataColumn private LocalDateTime c0; @DataColumn private LocalDateTime c1;
        @DataColumn private LocalDateTime c2; @DataColumn private LocalDateTime c3;
        @DataColumn private LocalDateTime c4; @DataColumn private LocalDateTime c5;
        @DataColumn private LocalDateTime c6; @DataColumn private LocalDateTime c7;
        @DataColumn private LocalDateTime c8; @DataColumn private LocalDateTime c9;
    }

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        mapper = new SpreadsheetMapper();

        intData = new ArrayList<>(rowCount);
        longData = new ArrayList<>(rowCount);
        doubleData = new ArrayList<>(rowCount);
        booleanData = new ArrayList<>(rowCount);
        stringData = new ArrayList<>(rowCount);
        bigDecimalData = new ArrayList<>(rowCount);
        localDateData = new ArrayList<>(rowCount);
        localDateTimeData = new ArrayList<>(rowCount);

        final LocalDate dateBase = LocalDate.of(2020, 1, 1);
        final LocalDateTime dateTimeBase = LocalDateTime.of(2020, 1, 1, 0, 0);

        for (int i = 0; i < rowCount; i++) {
            intData.add(new IntRow(i, i+1, i+2, i+3, i+4, i+5, i+6, i+7, i+8, i+9));
            longData.add(new LongRow(i, i+1, i+2, i+3, i+4, i+5, i+6, i+7, i+8, i+9));
            doubleData.add(new DoubleRow(i+.1, i+.2, i+.3, i+.4, i+.5, i+.6, i+.7, i+.8, i+.9, i+1.0));
            booleanData.add(new BooleanRow(
                    i%2==0, i%2==1, i%3==0, i%3==1, i%5==0, i%5==1, i%7==0, i%7==1, i%11==0, i%11==1));
            stringData.add(new StringRow(
                    "c0-"+i, "c1-"+i, "c2-"+i, "c3-"+i, "c4-"+i,
                    "c5-"+i, "c6-"+i, "c7-"+i, "c8-"+i, "c9-"+i));
            bigDecimalData.add(new BigDecimalRow(
                    BigDecimal.valueOf(i), BigDecimal.valueOf(i+1), BigDecimal.valueOf(i+2),
                    BigDecimal.valueOf(i+3), BigDecimal.valueOf(i+4), BigDecimal.valueOf(i+5),
                    BigDecimal.valueOf(i+6), BigDecimal.valueOf(i+7), BigDecimal.valueOf(i+8),
                    BigDecimal.valueOf(i+9)));
            localDateData.add(new LocalDateRow(
                    dateBase.plusDays(i), dateBase.plusDays(i+1), dateBase.plusDays(i+2),
                    dateBase.plusDays(i+3), dateBase.plusDays(i+4), dateBase.plusDays(i+5),
                    dateBase.plusDays(i+6), dateBase.plusDays(i+7), dateBase.plusDays(i+8),
                    dateBase.plusDays(i+9)));
            localDateTimeData.add(new LocalDateTimeRow(
                    dateTimeBase.plusMinutes(i), dateTimeBase.plusMinutes(i+1),
                    dateTimeBase.plusMinutes(i+2), dateTimeBase.plusMinutes(i+3),
                    dateTimeBase.plusMinutes(i+4), dateTimeBase.plusMinutes(i+5),
                    dateTimeBase.plusMinutes(i+6), dateTimeBase.plusMinutes(i+7),
                    dateTimeBase.plusMinutes(i+8), dateTimeBase.plusMinutes(i+9)));
        }

        intFile = _writeNumericFile("bench-celltype-int-", rowCount, (row, j, i) ->
                row.createCell(j).setCellValue((double) (i + j)));
        longFile = _writeNumericFile("bench-celltype-long-", rowCount, (row, j, i) ->
                row.createCell(j).setCellValue((double) (i + j)));
        doubleFile = _writeNumericFile("bench-celltype-double-", rowCount, (row, j, i) ->
                row.createCell(j).setCellValue(i + (j + 1) / 10.0));
        booleanFile = _writeBooleanFile("bench-celltype-boolean-", rowCount);
        stringFile = _writeStringFile("bench-celltype-string-", rowCount);
        bigDecimalFile = _writeNumericFile("bench-celltype-bigdecimal-", rowCount, (row, j, i) ->
                row.createCell(j).setCellValue((double) (i + j)));
        localDateFile = _writeDateFile("bench-celltype-localdate-", rowCount, "yyyy-mm-dd", dateBase, false);
        localDateTimeFile = _writeDateFile(
                "bench-celltype-localdatetime-", rowCount, "yyyy-mm-dd hh:mm:ss", dateTimeBase, true);
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        for (File f : new File[]{intFile, longFile, doubleFile, booleanFile, stringFile,
                bigDecimalFile, localDateFile, localDateTimeFile}) {
            if (f != null) f.delete();
        }
    }

    @FunctionalInterface
    private interface CellWriter {
        void write(Row row, int col, int rowIdx);
    }

    private static File _writeNumericFile(
            final String prefix, final int rows, final CellWriter writer) throws IOException {
        File file = File.createTempFile(prefix, ".xlsx");
        file.deleteOnExit();
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)) {
            Sheet sheet = wb.createSheet("Sheet1");
            Row h = sheet.createRow(0);
            for (int c = 0; c < 10; c++) h.createCell(c).setCellValue("c" + c);
            for (int i = 0; i < rows; i++) {
                Row r = sheet.createRow(i + 1);
                for (int c = 0; c < 10; c++) writer.write(r, c, i);
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
            wb.dispose();
        }
        return file;
    }

    private static File _writeBooleanFile(final String prefix, final int rows) throws IOException {
        File file = File.createTempFile(prefix, ".xlsx");
        file.deleteOnExit();
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)) {
            Sheet sheet = wb.createSheet("Sheet1");
            Row h = sheet.createRow(0);
            for (int c = 0; c < 10; c++) h.createCell(c).setCellValue("c" + c);
            for (int i = 0; i < rows; i++) {
                Row r = sheet.createRow(i + 1);
                for (int c = 0; c < 10; c++) r.createCell(c).setCellValue((i + c) % 2 == 0);
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
            wb.dispose();
        }
        return file;
    }

    private static File _writeStringFile(final String prefix, final int rows) throws IOException {
        File file = File.createTempFile(prefix, ".xlsx");
        file.deleteOnExit();
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)) {
            Sheet sheet = wb.createSheet("Sheet1");
            Row h = sheet.createRow(0);
            for (int c = 0; c < 10; c++) h.createCell(c).setCellValue("c" + c);
            for (int i = 0; i < rows; i++) {
                Row r = sheet.createRow(i + 1);
                for (int c = 0; c < 10; c++) r.createCell(c).setCellValue("c" + c + "-" + i);
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
            wb.dispose();
        }
        return file;
    }

    private static File _writeDateFile(
            final String prefix, final int rows, final String format,
            final Object base, final boolean dateTime) throws IOException {
        File file = File.createTempFile(prefix, ".xlsx");
        file.deleteOnExit();
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)) {
            Sheet sheet = wb.createSheet("Sheet1");
            DataFormat fmt = wb.createDataFormat();
            CellStyle style = wb.createCellStyle();
            style.setDataFormat(fmt.getFormat(format));
            Row h = sheet.createRow(0);
            for (int c = 0; c < 10; c++) h.createCell(c).setCellValue("c" + c);
            for (int i = 0; i < rows; i++) {
                Row r = sheet.createRow(i + 1);
                for (int c = 0; c < 10; c++) {
                    Cell cell = r.createCell(c);
                    if (dateTime) cell.setCellValue(((LocalDateTime) base).plusMinutes(i + c));
                    else cell.setCellValue(((LocalDate) base).plusDays(i + c));
                    cell.setCellStyle(style);
                }
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
            wb.dispose();
        }
        return file;
    }

    // Read benchmarks

    @Benchmark public void intRead(Blackhole bh) throws IOException {
        bh.consume(mapper.readValues(intFile, IntRow.class));
    }
    @Benchmark public void longRead(Blackhole bh) throws IOException {
        bh.consume(mapper.readValues(longFile, LongRow.class));
    }
    @Benchmark public void doubleRead(Blackhole bh) throws IOException {
        bh.consume(mapper.readValues(doubleFile, DoubleRow.class));
    }
    @Benchmark public void booleanRead(Blackhole bh) throws IOException {
        bh.consume(mapper.readValues(booleanFile, BooleanRow.class));
    }
    @Benchmark public void stringRead(Blackhole bh) throws IOException {
        bh.consume(mapper.readValues(stringFile, StringRow.class));
    }
    @Benchmark public void bigDecimalRead(Blackhole bh) throws IOException {
        bh.consume(mapper.readValues(bigDecimalFile, BigDecimalRow.class));
    }
    @Benchmark public void localDateRead(Blackhole bh) throws IOException {
        bh.consume(mapper.readValues(localDateFile, LocalDateRow.class));
    }
    @Benchmark public void localDateTimeRead(Blackhole bh) throws IOException {
        bh.consume(mapper.readValues(localDateTimeFile, LocalDateTimeRow.class));
    }

    // Write benchmarks

    @Benchmark public void intWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, intData, IntRow.class);
        bh.consume(out);
    }
    @Benchmark public void longWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, longData, LongRow.class);
        bh.consume(out);
    }
    @Benchmark public void doubleWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, doubleData, DoubleRow.class);
        bh.consume(out);
    }
    @Benchmark public void booleanWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, booleanData, BooleanRow.class);
        bh.consume(out);
    }
    @Benchmark public void stringWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, stringData, StringRow.class);
        bh.consume(out);
    }
    @Benchmark public void bigDecimalWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, bigDecimalData, BigDecimalRow.class);
        bh.consume(out);
    }
    @Benchmark public void localDateWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, localDateData, LocalDateRow.class);
        bh.consume(out);
    }
    @Benchmark public void localDateTimeWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, localDateTimeData, LocalDateTimeRow.class);
        bh.consume(out);
    }
}
