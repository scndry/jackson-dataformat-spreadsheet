package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetFactory;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Compares InMemorySharedStrings (default) vs FileBackedSharedStrings (H2 MVStore)
 * across both read and write paths.
 * <p>
 * Uses string-heavy data (4 string columns, mostly unique values) to stress the
 * shared string table. Read benchmarks parse a pre-generated file; write benchmarks
 * serialize the same data to a target file.
 */
@State(Scope.Benchmark)
@BenchmarkMode({Mode.AverageTime})
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class SharedStringsBenchmark {

    @Param({"10000", "50000", "100000"})
    int rowCount;

    File readFile;
    File writeFile;
    List<StringEntry> data;

    @DataGrid
    public static class StringEntry {
        public String col1;
        public String col2;
        public String col3;
        public String col4;

        public StringEntry() {}
    }

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        // Read source — pre-generated file with string-heavy data.
        // XSSFWorkbook (not SXSSF) to ensure shared string table is used.
        readFile = File.createTempFile("bench-sst-read-", ".xlsx");
        readFile.deleteOnExit();
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("col1");
            header.createCell(1).setCellValue("col2");
            header.createCell(2).setCellValue("col3");
            header.createCell(3).setCellValue("col4");
            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue("category-" + (i % 100));
                row.createCell(1).setCellValue("product-" + i);
                row.createCell(2).setCellValue("description of product " + i + " in category " + (i % 100));
                row.createCell(3).setCellValue("note-" + i + "-" + i);
            }
            try (FileOutputStream fos = new FileOutputStream(readFile)) {
                wb.write(fos);
            }
        }

        // In-memory data for write benchmarks (same string-heavy distribution).
        data = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            StringEntry e = new StringEntry();
            e.col1 = "category-" + (i % 100);
            e.col2 = "product-" + i;
            e.col3 = "description of product " + i + " in category " + (i % 100);
            e.col4 = "note-" + i + "-" + i;
            data.add(e);
        }
    }

    @Setup(Level.Invocation)
    public void setUpWriteFile() throws IOException {
        writeFile = File.createTempFile("bench-sst-write-", ".xlsx");
        writeFile.deleteOnExit();
    }

    @TearDown(Level.Invocation)
    public void tearDownWriteFile() {
        if (writeFile != null) writeFile.delete();
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        readFile.delete();
    }

    // ─── Read benchmarks ─────────────────────────────────────────

    @Benchmark
    public void inMemoryRead(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        List<StringEntry> values = mapper.readValues(readFile, StringEntry.class);
        bh.consume(values);
    }

    @Benchmark
    public void fileBackedRead(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
                .build();
        List<StringEntry> values = mapper.readValues(readFile, StringEntry.class);
        bh.consume(values);
    }

    @Benchmark
    public void fileBackedEncryptedRead(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
                .enable(SpreadsheetFactory.Feature.ENCRYPT_FILE_BACKED_STORE)
                .build();
        List<StringEntry> values = mapper.readValues(readFile, StringEntry.class);
        bh.consume(values);
    }

    // ─── Write benchmarks ────────────────────────────────────────

    @Benchmark
    public void inMemoryWrite(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        mapper.writeValue(writeFile, data, StringEntry.class);
        bh.consume(writeFile);
    }

    @Benchmark
    public void fileBackedWrite(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
                .build();
        mapper.writeValue(writeFile, data, StringEntry.class);
        bh.consume(writeFile);
    }

    @Benchmark
    public void fileBackedEncryptedWrite(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SpreadsheetFactory.Feature.FILE_BACKED_SHARED_STRINGS)
                .enable(SpreadsheetFactory.Feature.ENCRYPT_FILE_BACKED_STORE)
                .build();
        mapper.writeValue(writeFile, data, StringEntry.class);
        bh.consume(writeFile);
    }
}
