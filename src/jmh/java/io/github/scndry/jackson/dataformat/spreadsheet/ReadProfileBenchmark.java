package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetInput;
import io.github.scndry.jackson.dataformat.spreadsheet.schema.SpreadsheetSchema;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Isolates costs: schema generation vs full read (with reused mapper/schema).
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class ReadProfileBenchmark {

    @Param({"1000", "10000", "50000"})
    int rowCount;

    File file;
    SpreadsheetMapper mapper;
    SpreadsheetSchema schema;

    @DataGrid
    public static class Entry {
        public String name;
        public int quantity;
        public double price;
        public String description;

        public Entry() {}
    }

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        file = File.createTempFile("bench-profile-", ".xlsx");
        file.deleteOnExit();
        try (SXSSFWorkbook wb = new SXSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("name");
            header.createCell(1).setCellValue("quantity");
            header.createCell(2).setCellValue("price");
            header.createCell(3).setCellValue("description");
            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue("item-" + i);
                row.createCell(1).setCellValue(i);
                row.createCell(2).setCellValue(i * 1.5);
                row.createCell(3).setCellValue("description of item " + i);
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
            wb.dispose();
        }
        mapper = new SpreadsheetMapper();
        schema = mapper.sheetSchemaFor(Entry.class);
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        file.delete();
    }

    /** Full cost: new mapper + schema generation + parsing every time. */
    @Benchmark
    public void coldRead(Blackhole bh) throws IOException {
        SpreadsheetMapper cold = new SpreadsheetMapper();
        List<Entry> values = cold.readValues(file, Entry.class);
        bh.consume(values);
    }

    /** Reuse mapper (schema cached internally on first call). */
    @Benchmark
    public void warmMapper(Blackhole bh) throws IOException {
        List<Entry> values = mapper.readValues(file, Entry.class);
        bh.consume(values);
    }

    /** Reuse mapper + pre-built schema. Pure parsing cost. */
    @Benchmark
    public void warmSchema(Blackhole bh) throws IOException {
        @SuppressWarnings("unchecked")
        List<Entry> values = (List<Entry>) (List<?>) mapper.reader(schema).forType(Entry.class)
                .readValues(SheetInput.source(file)).readAll();
        bh.consume(values);
    }

    /** Schema generation cost only (no file I/O). */
    @Benchmark
    public void schemaOnly(Blackhole bh) throws Exception {
        SpreadsheetSchema s = mapper.sheetSchemaFor(Entry.class);
        bh.consume(s);
    }
}
