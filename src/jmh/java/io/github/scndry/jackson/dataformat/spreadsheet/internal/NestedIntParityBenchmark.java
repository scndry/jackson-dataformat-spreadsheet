package io.github.scndry.jackson.dataformat.spreadsheet.internal;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import com.fasterxml.jackson.annotation.OptBoolean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openjdk.jmh.annotations.*;
import org.openjdk.jmh.infra.Blackhole;

import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetFactory;
import io.github.scndry.jackson.dataformat.spreadsheet.SpreadsheetMapper;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumn;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataColumnGroup;
import io.github.scndry.jackson.dataformat.spreadsheet.annotation.DataGrid;

/**
 * Type-controlled cell-count-parity probe — flat 10 int columns vs
 * nested 6 int outer + 4 int inner. All cells share the same primitive
 * type so per-cell cost variance from BigDecimal/LocalDateTime/String
 * formatting drops out. The remaining gap measures the pure record-tree
 * algorithm overhead (RecordNode + LinkedHashMap + Cell wrapper alloc on
 * read; back-write buffer + scope traversal on write).
 *
 * <p>Cell-count parity: flat = targetCells / 10 rows; nested inner =
 * round_to_10(targetCells / 4.6). itemsPerOuter = 10.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Mode.AverageTime)
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1)
@Measurement(iterations = 5, time = 1)
@Fork(1)
public class NestedIntParityBenchmark {

    static final int ITEMS_PER_OUTER = 10;

    @Param({"100000", "1000000"})
    int targetCells;

    File flatFile;
    File nestedFile;
    List<IntRow> flatData;
    List<IntOuter> nestedData;
    SpreadsheetMapper mapper;
    SpreadsheetMapper mapperPoiRead;
    SpreadsheetMapper mapperPoiWrite;

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class IntRow {
        @DataColumn private int c0;
        @DataColumn private int c1;
        @DataColumn private int c2;
        @DataColumn private int c3;
        @DataColumn private int c4;
        @DataColumn private int c5;
        @DataColumn private int c6;
        @DataColumn private int c7;
        @DataColumn private int c8;
        @DataColumn private int c9;
    }

    @Data @NoArgsConstructor @AllArgsConstructor
    public static class IntItem {
        @DataColumn private int i0;
        @DataColumn private int i1;
        @DataColumn private int i2;
        @DataColumn private int i3;
    }

    @Data @NoArgsConstructor @AllArgsConstructor @DataGrid
    public static class IntOuter {
        @DataColumn(anchor = OptBoolean.TRUE) private int id;
        @DataColumn private int o1;
        @DataColumn private int o2;
        @DataColumn private int o3;
        @DataColumn private int o4;
        @DataColumn private int o5;
        @DataColumnGroup("Items") private List<IntItem> items;
    }

    @Setup(Level.Trial)
    public void setUp() throws IOException {
        final int flatRows = targetCells / 10;
        final int nestedInner = Math.round(targetCells / 4.6f / ITEMS_PER_OUTER) * ITEMS_PER_OUTER;
        final int outerCount = nestedInner / ITEMS_PER_OUTER;

        mapper = new SpreadsheetMapper();
        mapperPoiRead = new SpreadsheetMapper(
                new SpreadsheetFactory(XSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));
        mapperPoiWrite = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS)
                        .enable(SpreadsheetFactory.Feature.USE_POI_USER_MODEL));

        flatData = new ArrayList<>(flatRows);
        for (int i = 0; i < flatRows; i++) {
            flatData.add(new IntRow(i, i + 1, i + 2, i + 3, i + 4, i + 5, i + 6, i + 7, i + 8, i + 9));
        }
        nestedData = new ArrayList<>(outerCount);
        for (int i = 0; i < outerCount; i++) {
            List<IntItem> items = new ArrayList<>(ITEMS_PER_OUTER);
            final int base = i * ITEMS_PER_OUTER;
            for (int j = 0; j < ITEMS_PER_OUTER; j++) {
                items.add(new IntItem(base + j, base + j + 1, base + j + 2, base + j + 3));
            }
            nestedData.add(new IntOuter(i, i + 1, i + 2, i + 3, i + 4, i + 5, items));
        }

        flatFile = _writeFlatFile(flatData);
        nestedFile = _writeNestedFile(nestedData);
    }

    @TearDown(Level.Trial)
    public void tearDown() {
        flatFile.delete();
        nestedFile.delete();
    }

    private File _writeFlatFile(final List<IntRow> data) throws IOException {
        File file = File.createTempFile("bench-int-parity-flat-", ".xlsx");
        file.deleteOnExit();
        try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)) {
            Sheet sheet = wb.createSheet("Sheet1");
            Row header = sheet.createRow(0);
            for (int c = 0; c < 10; c++) header.createCell(c).setCellValue("c" + c);
            for (int i = 0; i < data.size(); i++) {
                IntRow r = data.get(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(r.getC0());
                row.createCell(1).setCellValue(r.getC1());
                row.createCell(2).setCellValue(r.getC2());
                row.createCell(3).setCellValue(r.getC3());
                row.createCell(4).setCellValue(r.getC4());
                row.createCell(5).setCellValue(r.getC5());
                row.createCell(6).setCellValue(r.getC6());
                row.createCell(7).setCellValue(r.getC7());
                row.createCell(8).setCellValue(r.getC8());
                row.createCell(9).setCellValue(r.getC9());
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                wb.write(fos);
            }
            wb.dispose();
        }
        return file;
    }

    private File _writeNestedFile(final List<IntOuter> data) throws IOException {
        File file = File.createTempFile("bench-int-parity-nested-", ".xlsx");
        file.deleteOnExit();
        // Use the library writer for nested layout — column ordering is authoritative.
        SpreadsheetMapper writer = new SpreadsheetMapper(
                new SpreadsheetFactory(SXSSFWorkbook::new,
                        SpreadsheetFactory.DEFAULT_SHEET_PARSER_FEATURE_FLAGS));
        writer.writeValue(file, data, IntOuter.class);
        return file;
    }

    @Benchmark
    public void flatRead(Blackhole bh) throws IOException {
        List<IntRow> values = mapper.readValues(flatFile, IntRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void flatReadPoi(Blackhole bh) throws IOException {
        List<IntRow> values = mapperPoiRead.readValues(flatFile, IntRow.class);
        bh.consume(values);
    }

    @Benchmark
    public void flatWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, flatData, IntRow.class);
        bh.consume(out);
    }

    @Benchmark
    public void flatWritePoi(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapperPoiWrite.writeValue(out, flatData, IntRow.class);
        bh.consume(out);
    }

    @Benchmark
    public void nestedRead(Blackhole bh) throws IOException {
        List<IntOuter> values = mapper.readValues(nestedFile, IntOuter.class);
        bh.consume(values);
    }

    @Benchmark
    public void nestedReadPoi(Blackhole bh) throws IOException {
        List<IntOuter> values = mapperPoiRead.readValues(nestedFile, IntOuter.class);
        bh.consume(values);
    }

    @Benchmark
    public void nestedWrite(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapper.writeValue(out, nestedData, IntOuter.class);
        bh.consume(out);
    }

    @Benchmark
    public void nestedWritePoi(Blackhole bh) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        mapperPoiWrite.writeValue(out, nestedData, IntOuter.class);
        bh.consume(out);
    }
}
