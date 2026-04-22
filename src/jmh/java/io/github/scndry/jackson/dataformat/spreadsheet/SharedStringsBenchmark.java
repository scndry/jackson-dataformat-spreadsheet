package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
import io.github.scndry.jackson.dataformat.spreadsheet.deser.SheetParser;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * Compares InMemorySharedStrings (default) vs FileBackedSharedStrings (H2 MVStore).
 * <p>
 * Uses string-heavy data to stress the shared string table.
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

    File file;

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
        file = File.createTempFile("bench-sst-", ".xlsx");
        file.deleteOnExit();
        // XSSFWorkbook (not SXSSF) to ensure shared string table is used
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
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
        }
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        file.delete();
    }

    @Benchmark
    public void inMemory(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        List<StringEntry> values = mapper.readValues(file, StringEntry.class);
        bh.consume(values);
    }

    @Benchmark
    public void fileBacked(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SheetParser.Feature.FILE_BACKED_SHARED_STRINGS)
                .build();
        List<StringEntry> values = mapper.readValues(file, StringEntry.class);
        bh.consume(values);
    }

    @Benchmark
    public void fileBackedEncrypted(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = SpreadsheetMapper.builder()
                .enable(SheetParser.Feature.FILE_BACKED_SHARED_STRINGS)
                .enable(SheetParser.Feature.ENCRYPT_FILE_BACKED_STORE)
                .build();
        List<StringEntry> values = mapper.readValues(file, StringEntry.class);
        bh.consume(values);
    }
}
