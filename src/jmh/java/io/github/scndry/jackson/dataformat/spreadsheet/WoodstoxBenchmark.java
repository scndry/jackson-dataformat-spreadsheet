package io.github.scndry.jackson.dataformat.spreadsheet;

import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;
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
 * Compares JDK built-in StAX vs Woodstox StAX implementation.
 * Woodstox is auto-detected via Java SPI when on the classpath.
 * <p>
 * Run with Woodstox on classpath (default in this project):
 * <pre>./gradlew jmh -Pjmh.includes="WoodstoxBenchmark"</pre>
 * <p>
 * To compare, run once with Woodstox excluded (requires manual dependency change).
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class WoodstoxBenchmark {

    @Param({"10000", "50000"})
    int rowCount;

    File file;

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
        file = File.createTempFile("bench-woodstox-", ".xlsx");
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

        // Log which StAX implementation is active
        javax.xml.stream.XMLInputFactory factory = javax.xml.stream.XMLInputFactory.newInstance();
        System.out.println("StAX implementation: " + factory.getClass().getName());
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        file.delete();
    }

    @Benchmark
    public void read(Blackhole bh) throws IOException {
        SpreadsheetMapper mapper = new SpreadsheetMapper();
        List<Entry> values = mapper.readValues(file, Entry.class);
        bh.consume(values);
    }
}
