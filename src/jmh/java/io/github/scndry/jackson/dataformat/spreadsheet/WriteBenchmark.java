package io.github.scndry.jackson.dataformat.spreadsheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.fesod.sheet.FesodSheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class WriteBenchmark {

    @Param({"1000", "10000", "50000", "100000"})
    int rowCount;

    List<BenchRow> data;
    File file;
    SpreadsheetMapper mapper;
    SpreadsheetMapper mapperPoi;

    @Setup(Level.Trial)
    public void setUp() {
        mapper = new SpreadsheetMapper();
        mapperPoi = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new, SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        data = new ArrayList<>();
        for (int i = 0; i < rowCount; i++) {
            data.add(BenchRow.create(i));
        }
    }

    @Setup(Level.Invocation)
    public void setUpFile() throws IOException {
        file = File.createTempFile("bench-write-", ".xlsx");
        file.deleteOnExit();
    }

    @TearDown(Level.Invocation)
    public void tearDown() {
        file.delete();
    }

    @Benchmark
    public void jacksonSpreadsheet(Blackhole bh) throws IOException {
        mapper.writeValue(file, data, BenchRow.class);
        bh.consume(file);
    }

    @Benchmark
    public void jacksonSpreadsheetPOI(Blackhole bh) throws IOException {
        mapperPoi.writeValue(file, data, BenchRow.class);
        bh.consume(file);
    }

    @Benchmark
    public void poiSXSSF(Blackhole bh) throws IOException {
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            for (int c = 0; c < BenchRow.HEADERS.length; c++) {
                header.createCell(c).setCellValue(BenchRow.HEADERS[c]);
            }
            for (int i = 0; i < data.size(); i++) {
                BenchRow r = data.get(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(r.getId());
                row.createCell(1).setCellValue(r.getName());
                row.createCell(2).setCellValue(r.getCategory());
                row.createCell(3).setCellValue(r.getStatus());
                row.createCell(4).setCellValue(r.getQuantity());
                row.createCell(5).setCellValue(r.getPrice());
                row.createCell(6).setCellValue(r.getAmount().doubleValue());
                row.createCell(7).setCellValue(r.getDueDate());
                row.createCell(8).setCellValue(r.getDescription());
                row.createCell(9).setCellValue(r.getCreatedAt());
            }
            try (OutputStream os = new FileOutputStream(file)) {
                wb.write(os);
            }
            wb.dispose();
            bh.consume(file);
        }
    }

    @Benchmark
    public void fesod(Blackhole bh) throws IOException {
        FesodSheet.write(file, BenchRow.class).sheet().doWrite(data);
        bh.consume(file);
    }

    @Benchmark
    public void fastExcel(Blackhole bh) throws IOException {
        try (OutputStream os = new FileOutputStream(file);
             Workbook wb = new Workbook(os, "bench", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            for (int c = 0; c < BenchRow.HEADERS.length; c++) {
                ws.value(0, c, BenchRow.HEADERS[c]);
            }
            for (int i = 0; i < data.size(); i++) {
                BenchRow r = data.get(i);
                int row = i + 1;
                ws.value(row, 0, r.getId());
                ws.value(row, 1, r.getName());
                ws.value(row, 2, r.getCategory());
                ws.value(row, 3, r.getStatus());
                ws.value(row, 4, r.getQuantity());
                ws.value(row, 5, r.getPrice());
                ws.value(row, 6, r.getAmount());
                ws.value(row, 7, r.getDueDate().atStartOfDay());
                ws.value(row, 8, r.getDescription());
                ws.value(row, 9, r.getCreatedAt());
            }
            bh.consume(file);
        }
    }
}
