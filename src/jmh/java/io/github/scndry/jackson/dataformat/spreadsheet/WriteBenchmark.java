package io.github.scndry.jackson.dataformat.spreadsheet;

import org.apache.fesod.sheet.FesodSheet;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class WriteBenchmark {

    @Param({"1000", "10000", "50000", "100000"})
    int rowCount;

    List<EntryData> data;
    File file;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class EntryData {
        public String name;
        public int quantity;
        public double price;
        public String description;
    }

    @Setup(Level.Trial)
    public void setUp() {
        data = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            data.add(new EntryData("item-" + i, i, i * 1.5, "description of item " + i));
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
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(file, data, EntryData.class);
        bh.consume(file);
    }

    @Benchmark
    public void jacksonSpreadsheetFileBacked(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper(
                new SpreadsheetFactory().enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS));
        mapper.writeValue(file, data, EntryData.class);
        bh.consume(file);
    }

    @Benchmark
    public void jacksonSpreadsheetFileBackedEncrypted(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper(
                new SpreadsheetFactory()
                        .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
                        .enable(SpreadsheetFactory.Feature.ENCRYPT_FILE_BACKED_STORE));
        mapper.writeValue(file, data, EntryData.class);
        bh.consume(file);
    }

    @Benchmark
    public void jacksonSpreadsheetPOI(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new, SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        mapper.writeValue(file, data, EntryData.class);
        bh.consume(file);
    }

    @Benchmark
    public void poiSXSSF(Blackhole bh) throws IOException {
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("name");
            header.createCell(1).setCellValue("quantity");
            header.createCell(2).setCellValue("price");
            header.createCell(3).setCellValue("description");
            for (int i = 0; i < data.size(); i++) {
                EntryData e = data.get(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(e.name);
                row.createCell(1).setCellValue(e.quantity);
                row.createCell(2).setCellValue(e.price);
                row.createCell(3).setCellValue(e.description);
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
        FesodSheet.write(file, EntryData.class).sheet().doWrite(data);
        bh.consume(file);
    }

    @Benchmark
    public void fastExcel(Blackhole bh) throws IOException {
        try (OutputStream os = new FileOutputStream(file);
             Workbook wb = new Workbook(os, "bench", "1.0")) {
            Worksheet ws = wb.newWorksheet("Sheet1");
            ws.value(0, 0, "name");
            ws.value(0, 1, "quantity");
            ws.value(0, 2, "price");
            ws.value(0, 3, "description");
            for (int i = 0; i < data.size(); i++) {
                EntryData e = data.get(i);
                ws.value(i + 1, 0, e.name);
                ws.value(i + 1, 1, e.quantity);
                ws.value(i + 1, 2, e.price);
                ws.value(i + 1, 3, e.description);
            }
            bh.consume(file);
        }
    }
}
