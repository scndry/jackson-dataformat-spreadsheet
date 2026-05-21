package io.github.scndry.jackson.dataformat.spreadsheet.internal;

import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Probe — 10 string columns, unique per row. Same column count as the shipped
 * {@code BenchRow} schema but 100% string, isolating the SharedStrings code
 * path from numeric/date deserialization.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class StringOnlyBenchmark {

    @Param({"10000", "50000", "100000"})
    int rowCount;

    File file;
    File writeTarget;
    SpreadsheetMapper mapper;
    List<StringRow> writeData;

    @Data @NoArgsConstructor @DataGrid
    public static class StringRow {
        private String s1; private String s2; private String s3; private String s4; private String s5;
        private String s6; private String s7; private String s8; private String s9; private String s10;
    }

    private static final String[] HEADERS = {
            "s1","s2","s3","s4","s5","s6","s7","s8","s9","s10"
    };

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        mapper = new SpreadsheetMapper();
        file = File.createTempFile("bench-stringonly-", ".xlsx");
        file.deleteOnExit();

        writeData = new ArrayList<>();
        for (int i = 0; i < rowCount; i++) {
            StringRow w = new StringRow();
            w.setS1("str-" + i + "-1");
            w.setS2("str-" + i + "-2");
            w.setS3("str-" + i + "-3");
            w.setS4("str-" + i + "-4");
            w.setS5("str-" + i + "-5");
            w.setS6("str-" + i + "-6");
            w.setS7("str-" + i + "-7");
            w.setS8("str-" + i + "-8");
            w.setS9("str-" + i + "-9");
            w.setS10("str-" + i + "-10");
            writeData.add(w);
        }

        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)) {
            Sheet sheet = wb.createSheet("StringOnly");
            Row header = sheet.createRow(0);
            for (int c = 0; c < HEADERS.length; c++) {
                header.createCell(c).setCellValue(HEADERS[c]);
            }
            for (int i = 0; i < rowCount; i++) {
                StringRow w = writeData.get(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(w.getS1());
                row.createCell(1).setCellValue(w.getS2());
                row.createCell(2).setCellValue(w.getS3());
                row.createCell(3).setCellValue(w.getS4());
                row.createCell(4).setCellValue(w.getS5());
                row.createCell(5).setCellValue(w.getS6());
                row.createCell(6).setCellValue(w.getS7());
                row.createCell(7).setCellValue(w.getS8());
                row.createCell(8).setCellValue(w.getS9());
                row.createCell(9).setCellValue(w.getS10());
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
            wb.dispose();
        }
    }

    @Setup(Level.Invocation)
    public void setUpWriteTarget() throws IOException {
        writeTarget = File.createTempFile("bench-stringonly-write-", ".xlsx");
        writeTarget.deleteOnExit();
    }

    @TearDown(Level.Invocation)
    public void tearDownWriteTarget() {
        writeTarget.delete();
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        file.delete();
    }

    @Benchmark
    public void readJackson(Blackhole bh) throws IOException {
        List<StringRow> values = mapper.readValues(file, StringRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void readPoiUserModel(Blackhole bh) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(file)) {
            Sheet sheet = wb.getSheetAt(0);
            List<StringRow> entries = new ArrayList<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                StringRow w = new StringRow();
                w.setS1(row.getCell(0).getStringCellValue());
                w.setS2(row.getCell(1).getStringCellValue());
                w.setS3(row.getCell(2).getStringCellValue());
                w.setS4(row.getCell(3).getStringCellValue());
                w.setS5(row.getCell(4).getStringCellValue());
                w.setS6(row.getCell(5).getStringCellValue());
                w.setS7(row.getCell(6).getStringCellValue());
                w.setS8(row.getCell(7).getStringCellValue());
                w.setS9(row.getCell(8).getStringCellValue());
                w.setS10(row.getCell(9).getStringCellValue());
                entries.add(w);
            }
            bh.consume(entries);
        }
    }

    @Benchmark
    public void readFastExcel(Blackhole bh) throws IOException {
        try (ReadableWorkbook wb = new ReadableWorkbook(file)) {
            List<StringRow> entries = new ArrayList<>();
            wb.getFirstSheet().openStream().forEach(row -> {
                if (row.getRowNum() == 1) return;
                StringRow w = new StringRow();
                w.setS1(row.getCellText(0));
                w.setS2(row.getCellText(1));
                w.setS3(row.getCellText(2));
                w.setS4(row.getCellText(3));
                w.setS5(row.getCellText(4));
                w.setS6(row.getCellText(5));
                w.setS7(row.getCellText(6));
                w.setS8(row.getCellText(7));
                w.setS9(row.getCellText(8));
                w.setS10(row.getCellText(9));
                entries.add(w);
            });
            bh.consume(entries);
        }
    }

    @Benchmark
    public void writeJackson(Blackhole bh) throws IOException {
        mapper.writeValue(writeTarget, writeData, StringRow.class);
        bh.consume(writeTarget);
    }

    @Benchmark
    public void writePoiSxssf(Blackhole bh) throws IOException {
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true);
             OutputStream os = new FileOutputStream(writeTarget)) {
            Sheet sheet = wb.createSheet("StringOnly");
            Row header = sheet.createRow(0);
            for (int c = 0; c < HEADERS.length; c++) {
                header.createCell(c).setCellValue(HEADERS[c]);
            }
            for (int i = 0; i < writeData.size(); i++) {
                StringRow w = writeData.get(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(w.getS1());
                row.createCell(1).setCellValue(w.getS2());
                row.createCell(2).setCellValue(w.getS3());
                row.createCell(3).setCellValue(w.getS4());
                row.createCell(4).setCellValue(w.getS5());
                row.createCell(5).setCellValue(w.getS6());
                row.createCell(6).setCellValue(w.getS7());
                row.createCell(7).setCellValue(w.getS8());
                row.createCell(8).setCellValue(w.getS9());
                row.createCell(9).setCellValue(w.getS10());
            }
            wb.write(os);
            wb.dispose();
        }
        bh.consume(writeTarget);
    }

    @Benchmark
    public void writeFastExcel(Blackhole bh) throws IOException {
        try (OutputStream os = new FileOutputStream(writeTarget);
             Workbook wb = new Workbook(os, "bench", "1.0")) {
            Worksheet ws = wb.newWorksheet("StringOnly");
            for (int c = 0; c < HEADERS.length; c++) {
                ws.value(0, c, HEADERS[c]);
            }
            for (int i = 0; i < writeData.size(); i++) {
                StringRow w = writeData.get(i);
                int row = i + 1;
                ws.value(row, 0, w.getS1());
                ws.value(row, 1, w.getS2());
                ws.value(row, 2, w.getS3());
                ws.value(row, 3, w.getS4());
                ws.value(row, 4, w.getS5());
                ws.value(row, 5, w.getS6());
                ws.value(row, 6, w.getS7());
                ws.value(row, 7, w.getS8());
                ws.value(row, 8, w.getS9());
                ws.value(row, 9, w.getS10());
            }
        }
        bh.consume(writeTarget);
    }
}
